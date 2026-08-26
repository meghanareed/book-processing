"""
reenrich_existing.py
--------------------
Fills in missing metadata (PageCount, LengthCategory, Genre, Tropes, Triggers,
ASIN, Description, AgeRange, ISBNs) for rows that are already in books_output.xlsx
but were never fully enriched.

Uses the same lookup pipeline as books.py (Google Books → Open Library → Amazon
ASIN → OpenAI AI enrichment) without touching any row that is already complete.

Skip logic (both configurable in launcher_config.json → "reenrich" section):
  - Rows enriched within the last N days are skipped (default 180 = 6 months)
  - Rows already >= X% complete across scored fields are skipped (default 70%)

Usage:
    python reenrich_existing.py [options]

Options:
    --dry-run             Show what would change without writing anything
    --limit N             Only process the first N eligible books (default: all)
    --missing FIELD       Only process rows missing this specific field
                          e.g. --missing PageCount  (default: any missing field)
    --sheet NAME          Sheet name to read/write (default: All Books)
    --min-age-days N      Override: skip if enriched within N days (0 = ignore date)
    --skip-threshold 0.8  Override: skip if fill % >= this value (0–1)
    --force               Ignore all skip conditions and process every row

Examples:
    # Re-enrich everything (respects skip rules from config)
    python reenrich_existing.py

    # Just fill in page counts — cheap, no OpenAI cost
    python reenrich_existing.py --missing PageCount

    # Test on 10 books first
    python reenrich_existing.py --limit 10 --dry-run

    # Force re-enrich everything regardless of age or fill %
    python reenrich_existing.py --force
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import shutil
import sys
import time
import datetime

# Force UTF-8 output on Windows — the default cp1252 console codec can't
# encode emoji/arrows and would crash on the first ✅ or → in a log line.
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
import traceback
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
# SCRIPT_DIR holds the code (and launcher_config.json); DATA_DIR holds the
# spreadsheets and logs.  They are the same folder in the old layout and
# separate once the code is cloned out of OneDrive.
SCRIPT_DIR   = Path(__file__).resolve().parent
DATA_DIR     = Path(r"C:\Users\megha\OneDrive\Documents\Reading")

EXCEL_PATH   = DATA_DIR / "books_output.xlsx"
if not EXCEL_PATH.exists():
    EXCEL_PATH = SCRIPT_DIR / "books_output.xlsx"   # old layout: data beside the code
BOOKS_PY     = SCRIPT_DIR / "books.py"
SHEET_NAME   = "All Books"

# Load optional overrides from launcher_config.json
_launcher_cfg: dict = {}
_cfg_path = SCRIPT_DIR / "launcher_config.json"
if _cfg_path.exists():
    try:
        _launcher_cfg = json.loads(_cfg_path.read_text(encoding="utf-8"))
    except Exception:
        pass
_re_cfg = _launcher_cfg.get("reenrich", {})

# Skip a row if it was enriched within this many days (0 = ignore date check)
CFG_MIN_AGE_DAYS    = int(_re_cfg.get("min_age_days", 180))
# Skip a row if fill% across scored fields >= this value (0.0–1.0)
CFG_SKIP_THRESHOLD  = float(_re_cfg.get("skip_threshold", 0.70))

# Fields counted when calculating fill percentage (mirrors books.py)
ENRICHMENT_SCORED_FIELDS = [
    "ISBN_13", "ASIN", "Description", "Genre",
    "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
]

# Fields that this script can fill in / check for --missing
FILLABLE_FIELDS = [
    "PageCount", "LengthCategory", "Genre", "Tropes", "Triggers",
    "AgeRange", "Description", "ASIN", "ISBN_13", "ISBN_10",
]

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
LOG_DIR  = EXCEL_PATH.parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_PATH = LOG_DIR / f"reenrich_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
_log_fh  = None

def _open_log() -> None:
    global _log_fh
    try:
        _log_fh = open(LOG_PATH, "w", encoding="utf-8", buffering=1)
    except Exception as e:
        print(f"[WARN] Could not open log file: {e}", flush=True)

def log(msg: str = "") -> None:
    ts   = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    if _log_fh:
        try:
            _log_fh.write(line + "\n")
        except Exception:
            pass

def _close_log() -> None:
    if _log_fh:
        try:
            _log_fh.close()
        except Exception:
            pass

# ---------------------------------------------------------------------------
# Load books.py helpers
# ---------------------------------------------------------------------------
def _load_books_module():
    if not BOOKS_PY.exists():
        sys.exit(f"ERROR: books.py not found at {BOOKS_PY}\n"
                 "Place this script in the same folder as books.py.")
    spec = importlib.util.spec_from_file_location("books", BOOKS_PY)
    mod  = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
        return mod
    except RuntimeError as e:
        if "OPENAI_API_KEY" in str(e):
            print("=" * 60, flush=True)
            print("ERROR: OpenAI API key is not set.", flush=True)
            print("", flush=True)
            print("To fix this, open Settings in the launcher and", flush=True)
            print("enter your API key in the 'OpenAI API Key' field.", flush=True)
            print("", flush=True)
            print("The Fill Missing Metadata step needs the API key", flush=True)
            print("to fill in Genre, Tropes, and Triggers via AI.", flush=True)
            print("", flush=True)
            print("If you only want page counts (no AI), run from", flush=True)
            print("the command line with:", flush=True)
            print("  python reenrich_existing.py --missing PageCount", flush=True)
            print("(that step only uses Google Books -- no API key needed)", flush=True)
            print("=" * 60, flush=True)
            sys.exit(1)
        raise

books = _load_books_module()

lookup_book_metadata          = books.lookup_book_metadata
page_count_to_length_category = books.page_count_to_length_category
normalize_csv_list            = books.normalize_csv_list
clean_text                    = books.clean_text
AI_ENRICH_SLEEP               = getattr(books, "AI_ENRICH_SLEEP_SECONDS", 0.3)

# Pull scored fields + thresholds from books.py if available, else use our own
_SCORED_FIELDS   = getattr(books, "ENRICHMENT_SCORED_FIELDS", ENRICHMENT_SCORED_FIELDS)
_MIN_AGE_DAYS    = CFG_MIN_AGE_DAYS    # command-line can override later
_SKIP_THRESHOLD  = CFG_SKIP_THRESHOLD

# ---------------------------------------------------------------------------
# Skip-logic helpers
# ---------------------------------------------------------------------------

def fill_pct(row: pd.Series) -> float:
    """Fraction of scored enrichment fields that are non-empty (0.0–1.0)."""
    if not _SCORED_FIELDS:
        return 0.0
    filled = 0
    for field in _SCORED_FIELDS:
        val = clean_text(row.get(field, ""))
        if field == "PageCount":
            try:
                if int(float(val)) > 0:
                    filled += 1
                continue
            except Exception:
                pass
        if val:
            filled += 1
    return filled / len(_SCORED_FIELDS)


def last_enriched_date(row: pd.Series) -> "datetime.date | None":
    """Parse the Last Enriched cell; return a date or None."""
    raw = clean_text(row.get("Last Enriched", ""))
    if not raw:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


def should_skip(row: pd.Series, min_age_days: int, skip_threshold: float, force: bool) -> tuple[bool, str]:
    """
    Return (skip, reason).
    skip=True means don't re-enrich this row.
    """
    if force:
        return False, ""

    # Date check
    if min_age_days > 0:
        last = last_enriched_date(row)
        if last is not None:
            age = (datetime.date.today() - last).days
            if age < min_age_days:
                return True, f"enriched {age}d ago (< {min_age_days}d)"

    # Fill% check
    pct = fill_pct(row)
    if pct >= skip_threshold:
        return True, f"{pct:.0%} complete (>= {skip_threshold:.0%} threshold)"

    return False, ""

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def fetch_amazon_page_count(asin: str) -> int:
    """
    Fetch page count from Amazon product page for a given ASIN.
    Looks for 'Print length' in the product details — Amazon Kindle books
    always list this even when Google Books has no data.
    Returns 0 if not found.
    """
    if not asin:
        return 0
    import urllib.request, re as _re
    url = f"https://www.amazon.com/dp/{asin}"
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept-Language": "en-US,en;q=0.9",
        })
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8", errors="replace")
        # "Print length" appears in product details table
        m = _re.search(r'Print length.*?(\d{2,4})\s*pages', html, _re.IGNORECASE | _re.DOTALL)
        if m:
            return int(m.group(1))
        # Fallback: "X pages" near "Kindle Edition"
        m = _re.search(r'(\d{2,4})\s*pages', html)
        if m:
            val = int(m.group(1))
            if 10 <= val <= 5000:
                return val
    except Exception:
        pass
    return 0
    """True if a cell value should be treated as missing."""
    v = clean_text(value)
    return v == "" or v.lower() in ("nan", "none")


def is_empty(value) -> bool:
    """True if a cell value should be treated as missing."""
    v = clean_text(value)
    return v == "" or v.lower() in ("nan", "none")


def needs_work(row: pd.Series, missing_field: str | None) -> bool:
    """
    True if the row has at least one fillable field that's empty.
    When --missing FIELD is given, only that field is checked.
    (Skip-age and fill-% checks are handled separately by should_skip.)
    """
    if missing_field:
        val = row.get(missing_field, "")
        if missing_field == "PageCount":
            try:
                return int(float(val)) == 0
            except Exception:
                return is_empty(val)
        return is_empty(val)

    for field in FILLABLE_FIELDS:
        val = row.get(field, "")
        if field == "PageCount":
            try:
                if int(float(val)) == 0:
                    return True
                continue
            except Exception:
                pass
        if is_empty(val):
            return True
    return False


def apply_result(row: pd.Series, result: dict) -> tuple[pd.Series, list[str]]:
    """
    Merge lookup result into row, only filling genuinely empty cells.
    Returns (updated_row, list_of_changed_fields).
    """
    changed = []
    row = row.copy()

    for field, new_val in result.items():
        if field not in row.index:
            continue
        new_str = clean_text(new_val)
        if not new_str:
            continue

        current = clean_text(row[field])

        # PageCount: only overwrite 0 / blank
        if field == "PageCount":
            try:
                if int(float(current)) > 0:
                    continue  # already has a real value
            except Exception:
                pass  # current is blank/nan — fall through to overwrite
            try:
                if int(float(new_str)) <= 0:
                    continue  # new value is also useless
            except Exception:
                continue
            row[field] = float(new_str)  # float64 to match pandas column dtype
            changed.append(field)
            continue

        # LengthCategory: derive from PageCount if possible
            if current:
                continue  # already set
            row[field] = new_str
            changed.append(field)
            continue

        # All other fields: only fill if currently empty
        if current:
            continue

        # Normalise list fields before storing
        if field in ("Genre", "Tropes", "Triggers"):
            new_str = normalize_csv_list(new_str)
        row[field] = new_str
        changed.append(field)

    # Always (re-)derive LengthCategory if PageCount was just filled
    if "PageCount" in changed and not clean_text(row.get("LengthCategory", "")):
        lc = page_count_to_length_category(row["PageCount"])
        if lc:
            row["LengthCategory"] = lc
            if "LengthCategory" not in changed:
                changed.append("LengthCategory")

    # Stamp Last Enriched whenever any field was filled
    if changed:
        row["Last Enriched"] = datetime.date.today().isoformat()
        if "Last Enriched" not in changed:
            changed.append("Last Enriched")

    return row, changed


def save_excel(df: pd.DataFrame) -> None:
    """Atomic write: temp file → backup → rename."""
    import os
    tmp_path = EXCEL_PATH.with_suffix(".tmp.xlsx")
    bak_path = EXCEL_PATH.with_suffix(".bak.xlsx")

    # Preserve all other sheets
    try:
        all_sheets: dict = pd.read_excel(EXCEL_PATH, sheet_name=None)
    except Exception:
        all_sheets = {}
    all_sheets[SHEET_NAME] = df

    try:
        with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
            for name, sheet_df in all_sheets.items():
                sheet_df.to_excel(writer, sheet_name=name, index=False)
    except Exception as e:
        tmp_path.unlink(missing_ok=True)
        raise RuntimeError(f"Write failed (original untouched): {e}") from e

    try:
        if EXCEL_PATH.exists():
            shutil.copy2(EXCEL_PATH, bak_path)
    except Exception:
        pass

    tmp_path.replace(EXCEL_PATH)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="Re-enrich existing books_output.xlsx rows")
    parser.add_argument("--dry-run",        action="store_true", help="Show changes without saving")
    parser.add_argument("--limit",          type=int,   default=0,                  help="Max books to process (0=all)")
    parser.add_argument("--missing",        type=str,   default="",                 help="Only process rows missing this field")
    parser.add_argument("--sheet",          type=str,   default=SHEET_NAME,         help="Sheet name")
    parser.add_argument("--min-age-days",   type=int,   default=CFG_MIN_AGE_DAYS,   help="Skip if enriched within N days (0=off)")
    parser.add_argument("--skip-threshold", type=float, default=CFG_SKIP_THRESHOLD, help="Skip if fill%% >= this value 0–1 (default 0.70)")
    parser.add_argument("--force",          action="store_true", help="Ignore all skip conditions")
    args = parser.parse_args()

    missing_field = args.missing.strip() or None
    if missing_field and missing_field not in FILLABLE_FIELDS:
        sys.exit(f"ERROR: --missing must be one of: {', '.join(FILLABLE_FIELDS)}")

    _open_log()
    try:
        _run(args, missing_field)
    except KeyboardInterrupt:
        log("Interrupted by user.")
    except Exception as e:
        log(f"FATAL: {e}")
        log(traceback.format_exc())
    finally:
        _close_log()


def _run(args, missing_field: str | None) -> None:
    global SHEET_NAME
    SHEET_NAME = args.sheet

    min_age_days   = args.min_age_days
    skip_threshold = args.skip_threshold
    force          = args.force

    log("=" * 60)
    log("Re-Enrichment Script for books_output.xlsx")
    log(f"Excel      : {EXCEL_PATH}")
    log(f"Filter     : {'missing ' + missing_field if missing_field else 'any missing field'}")
    log(f"Limit      : {args.limit or 'no limit'}")
    log(f"Skip age   : {'disabled' if min_age_days == 0 else f'< {min_age_days} days'}")
    log(f"Skip fill% : {skip_threshold:.0%}+ complete")
    log(f"Force      : {force}")
    log(f"Dry run    : {args.dry_run}")
    log("=" * 60)

    if not EXCEL_PATH.exists():
        sys.exit(f"ERROR: {EXCEL_PATH} not found.")

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

    # Ensure Last Enriched column exists as the last column
    if "Last Enriched" not in df.columns:
        df["Last Enriched"] = ""
        log("Added 'Last Enriched' column to sheet.")
        if not args.dry_run:
            save_excel(df)
            log("Saved column addition to file.")

    # Force string columns that pandas may have inferred as float64
    # (happens when a column is entirely empty/NaN on load).
    # Without this, writing a date string like '2026-06-17' crashes.
    STRING_COLS = [
        "Last Enriched", "Description", "Genre", "Tropes", "Triggers",
        "AgeRange", "LengthCategory", "ISBN_13", "ISBN_10", "ASIN",
        "Lookup Source", "Metadata Enriched",
    ]
    for col in STRING_COLS:
        if col in df.columns and df[col].dtype != object:
            df[col] = df[col].astype(object)

    log(f"Loaded {len(df)} rows from '{SHEET_NAME}'")

    # Two-pass eligibility: needs_work (has gaps) AND not skipped by date/fill%
    has_gaps_indices = [
        i for i, row in df.iterrows()
        if clean_text(row.get("Title", "")) and needs_work(row, missing_field)
    ]
    log(f"Rows with gaps    : {len(has_gaps_indices)}")

    eligible_indices = []
    skip_date_count  = 0
    skip_pct_count   = 0
    for i in has_gaps_indices:
        skip, reason = should_skip(df.loc[i], min_age_days, skip_threshold, force)
        if skip:
            if "ago" in reason:
                skip_date_count += 1
            else:
                skip_pct_count += 1
        else:
            eligible_indices.append(i)

    log(f"Skipped (recent)  : {skip_date_count}")
    log(f"Skipped (fill %)  : {skip_pct_count}")
    log(f"Will process      : {len(eligible_indices)}")

    if args.limit:
        eligible_indices = eligible_indices[: args.limit]
        log(f"Limited to        : {len(eligible_indices)} (--limit {args.limit})")

    if not eligible_indices:
        log("Nothing to do.")
        return

    # Process
    processed = 0
    updated   = 0
    skipped   = 0
    errors    = 0
    save_interval = 10  # save every N updates so a crash loses minimal progress

    for idx in eligible_indices:
        row    = df.loc[idx]
        title  = clean_text(row.get("Title", ""))
        author = clean_text(row.get("Author", ""))
        processed += 1

        pct = fill_pct(row)
        log(f"\n[{processed}/{len(eligible_indices)}] {title} | {author}  ({pct:.0%} complete)")

        missing_now = [f for f in FILLABLE_FIELDS if needs_work(row, f)]
        log(f"  Missing: {', '.join(missing_now) or 'nothing?'}")

        if args.dry_run:
            log("  [DRY RUN] Would call lookup_book_metadata -- skipping.")
            continue

        try:
            result  = lookup_book_metadata(title, author)

            # If Google Books/OpenLibrary found nothing for PageCount but we have
            # an ASIN, try Amazon directly — Kindle books always have "Print length"
            if not result.get("PageCount") or str(result.get("PageCount","")) in ("","0"):
                asin = clean_text(row.get("ASIN",""))
                if asin:
                    amz_pages = fetch_amazon_page_count(asin)
                    if amz_pages > 0:
                        result["PageCount"] = float(amz_pages)
                        log(f"  [AMAZON] Found PageCount={amz_pages} via ASIN")

            new_row, changed = apply_result(row, result)

            if changed:
                # Assign column-by-column to avoid pandas dtype conflicts
                # (assigning a whole Series with df.loc[idx] = new_row forces
                # pandas to reconcile mixed types and crashes on float64 columns
                # receiving strings like dates, or string columns receiving floats)
                for col in changed:
                    val = new_row[col]
                    if col == "PageCount":
                        try:
                            val = float(val)
                        except Exception:
                            continue
                    df.at[idx, col] = val
                updated += 1
                display_changed = [c for c in changed if c != "Last Enriched"]
                new_pct = fill_pct(new_row)
                log(f"  [OK] Filled: {', '.join(display_changed)}  ({pct:.0%} -> {new_pct:.0%})")
            else:
                # Nothing new found — stamp date so this row is skipped for
                # the next min_age_days days rather than retried pointlessly.
                df.at[idx, "Last Enriched"] = datetime.date.today().isoformat()
                skipped += 1
                log(f"  -- No new data found (date stamped, will skip for {args.min_age_days}d)")

            rows_touched = updated + skipped
            if rows_touched > 0 and rows_touched % save_interval == 0:
                log(f"\n  [AUTO-SAVE] Saving after {rows_touched} rows processed...")
                save_excel(df)
                log(f"  [AUTO-SAVE] Saved.")

            time.sleep(AI_ENRICH_SLEEP)

        except Exception as e:
            errors += 1
            log(f"  [ERROR] {e}")
            log(f"  {traceback.format_exc().strip()}")
            time.sleep(1)

    if not args.dry_run and (updated + skipped) > 0:
        log(f"\nSaving final results...")
        save_excel(df)
        log(f"Saved: {EXCEL_PATH}")

    log("")
    log("=" * 60)
    log(f"DONE")
    log(f"Processed : {processed}")
    log(f"Updated   : {updated}")
    log(f"No change : {skipped}")
    log(f"Errors    : {errors}")
    if not args.dry_run:
        log(f"Log file  : {LOG_PATH}")
    log("=" * 60)


if __name__ == "__main__":
    main()
