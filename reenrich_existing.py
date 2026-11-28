"""
reenrich_existing.py
--------------------
Fills in missing metadata (PageCount, LengthCategory, Genre, Tropes, Triggers,
ASIN, Description, AgeRange, ISBNs) for rows that are already in books_output.xlsx
but were never fully enriched.

Uses the same lookup pipeline as books.py (Google Books → Open Library → Amazon
ASIN → OpenAI AI enrichment) without touching any row that is already complete.

Skip logic (all configurable in launcher_config.json → "reenrich" section):
  - Books with a selector decision of Read / Ignored / Removed are skipped
    (skip_decided, default true) — they are never offered again, so enriching
    them spends a lookup and an OpenAI call on nothing.
  - Books whose Read column is Yes are skipped (skip_read, default true). Turn
    this one off if your Read column over-reports: earlier versions of
    apply_reading_log.py set Read = Yes for Ignored and Removed as well.
  - Rows enriched within the last N days are skipped (default 180 = 6 months)
  - Rows already >= X% complete across scored fields are skipped (default 70%)

Category re-grading (recategorize, default true):
  Before enriching anything, LengthCategory is re-derived from PageCount for
  every row that has a page count, so the tags stay in step with the buckets the
  book selector filters on (X-Short <100, Short <200, Medium 200-400, Long
  400-600, Epic 600+). It is local arithmetic — no lookup, no model call — but
  books.py still needs OPENAI_API_KEY present just to import.
  --recategorize does only this pass and exits; --no-recategorize skips it.

  --force overrides the last two skip rules only. The first two are opted out of with
  --include-decided and --include-read, so a forced run never quietly spends
  money on books you have finished with.

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
    --force               Ignore the age and fill% rules (read/removed still skipped)
    --retry-after-days N  Skip books tried in the last N days that are still
                          missing something (0 = off). Use it with --limit
                          --force: without it every batch starts by retrying
                          the books the previous batch could not fill.
    --include-decided     Also process books marked Read / Ignored / Removed
    --include-read        Also process books whose Read column is Yes
    --recategorize        Only re-grade LengthCategory from existing page counts
    --no-recategorize     Skip the category re-grade before enrichment and process every row
    --ai-only             Only run the model on rows with a description but no tropes or
                          spice tags (no lookups)

Examples:
    # Re-enrich everything (respects skip rules from config)
    python reenrich_existing.py

    # Just fill in page counts. Note --force: the default skip rules (enriched
    # in the last 180 days, or already 70% complete) leave out almost every row
    # whose only gap is a page count. A row missing nothing the model can
    # supply skips the OpenAI call, and one that already has an ASIN skips the
    # Amazon search, so this is the cheap direction — but still try --limit 25
    # first to gauge the hit rate.
    python reenrich_existing.py --missing PageCount --force

    # Fix categories that disagree with their page count — free, no lookups
    python reenrich_existing.py --recategorize

    # Test on 10 books first
    python reenrich_existing.py --limit 10 --dry-run

    # Work through every missing description in batches
    python reenrich_existing.py --missing Description --limit 200 --force --retry-after-days 7

    # Force re-enrich everything regardless of age or fill %
    python reenrich_existing.py --force

    # Include the books you have already read or dismissed
    python reenrich_existing.py --force --include-decided --include-read
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
# Skip rows you have already dealt with: a book you will never be offered again
# is a lookup and an OpenAI call spent on nothing. Two separate switches because
# the two columns are not equally trustworthy — "Selector Decision" is written
# only by apply_reading_log.py, while "Read" has historical noise (earlier
# versions set it to Yes for Ignored and Removed too, and StoryGraph sync sets
# it as well). Turn skip_read off if your Read column over-reports.
CFG_SKIP_DECIDED    = bool(_re_cfg.get("skip_decided", True))
CFG_SKIP_READ       = bool(_re_cfg.get("skip_read", True))
# Re-grade LengthCategory from PageCount for rows that already have a page
# count. Costs nothing — no lookup, no model call — and is what brings a sheet
# back in step after the category boundaries move.
CFG_RECATEGORIZE    = bool(_re_cfg.get("recategorize", True))

# Fields counted when calculating fill percentage (mirrors books.py)
ENRICHMENT_SCORED_FIELDS = [
    "ISBN_13", "ASIN", "Description", "Genre",
    "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
]

# Fields that this script can fill in / check for --missing
FILLABLE_FIELDS = [
    "PageCount", "LengthCategory", "Genre", "Tropes", "Triggers", "Spice Tags",
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
            print("The key is required even for a --missing PageCount run", flush=True)
            print("or --recategorize: books.py needs it just to import,", flush=True)
            print("whether or not the run ends up calling the model.", flush=True)
            print("=" * 60, flush=True)
            sys.exit(1)
        raise

books = _load_books_module()

lookup_book_metadata          = books.lookup_book_metadata
page_count_to_length_category = books.page_count_to_length_category
normalize_csv_list            = books.normalize_csv_list
normalize_spice_tags          = getattr(books, "normalize_spice_tags", None)
fill_spice_from_tropes        = getattr(books, "fill_spice_from_tropes", None)
is_vague_author               = getattr(books, "is_vague_author", lambda a: not clean_text(a))
clean_text                    = books.clean_text
AI_ENRICH_SLEEP               = getattr(books, "AI_ENRICH_SLEEP_SECONDS", 0.3)

# Pull scored fields + thresholds from books.py if available, else use our own
_SCORED_FIELDS   = getattr(books, "ENRICHMENT_SCORED_FIELDS", ENRICHMENT_SCORED_FIELDS)
_MIN_AGE_DAYS    = CFG_MIN_AGE_DAYS    # command-line can override later
_SKIP_THRESHOLD  = CFG_SKIP_THRESHOLD

# Written by apply_reading_log.py: "Selector Decision" records what you chose
# in the book selector, and "Read" is Yes only for a genuine Read (StoryGraph
# sync sets it too). Older sheets have neither column — a missing column just
# means nothing has been decided yet.
DECISION_COL = "Selector Decision"
READ_COL     = "Read"
DISMISSED_DECISIONS = {"read", "ignored", "removed"}

# ---------------------------------------------------------------------------
# Skip-logic helpers
# ---------------------------------------------------------------------------
def is_yes(value) -> bool:
    return clean_text(value).lower() in ("yes", "y", "true", "1")


def finished_reason(row: pd.Series, skip_decided: bool, skip_read: bool) -> str:
    """Why this book is already dealt with, or "" if it still needs offering."""
    decision = clean_text(row.get(DECISION_COL, ""))
    if skip_decided and decision.lower() in DISMISSED_DECISIONS:
        return f"selector decision: {decision}"
    if skip_read and is_yes(row.get(READ_COL, "")):
        return "marked Read"
    return ""



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


# When each book was last tried, kept beside the workbook rather than in it. Last
# Enriched cannot do this job: a failed row is deliberately not stamped there.
ATTEMPTS_PATH = EXCEL_PATH.parent / "_reenrich_attempts.json"

# Amazon's page window, carried across runs so a rest holds however many runs
# there are. Amazon blocks on how many pages come in a stretch and lets up
# after a rest, so the limit is pages per window, then hours off.
AMAZON_USAGE_PATH = EXCEL_PATH.parent / "_amazon_pages.json"
DEFAULT_AMAZON_PAGES_PER_WINDOW = 50
DEFAULT_AMAZON_REST_HOURS = 2.5
# After a captcha or a run of blocked pages, rest for this many windows' rest.
AMAZON_BLOCK_REST_FACTOR = 2
_AMAZON_STATE_LOADED = False


def _epoch(text) -> float:
    return datetime.datetime.fromisoformat(text).timestamp() if text else 0.0


def _iso(epoch: float) -> str:
    return datetime.datetime.fromtimestamp(epoch).isoformat(timespec="seconds") if epoch else ""


def load_amazon_state() -> tuple[int, float, float]:
    """(pages used in the current window, rest-until, last page), epoch seconds or 0."""
    try:
        data = json.loads(AMAZON_USAGE_PATH.read_text(encoding="utf-8"))
        return (int(data.get("window_used", 0)), _epoch(data.get("rest_until")),
                _epoch(data.get("last_page")))
    except FileNotFoundError:
        return 0, 0.0, 0.0
    except Exception as e:
        log(f"[WARN] Ignoring unreadable {AMAZON_USAGE_PATH.name}: {e}")
        return 0, 0.0, 0.0


def save_amazon_usage() -> None:
    tmp = AMAZON_USAGE_PATH.with_suffix(".tmp")
    if (books.AMAZON_WINDOW_PAGES and not books.AMAZON_REST_UNTIL
            and books.AMAZON_WINDOW_USED >= books.AMAZON_WINDOW_PAGES):
        # Filled on the run's last Amazon book: start the rest now, so the file
        # says when it ends.
        books.amazon_skip_reason()
    try:
        tmp.write_text(json.dumps({
            "window_used": books.AMAZON_WINDOW_USED,
            "rest_until": _iso(books.AMAZON_REST_UNTIL),
            "last_page": _iso(books.AMAZON_LAST_PAGE),
        }), encoding="utf-8")
        tmp.replace(AMAZON_USAGE_PATH)
    except Exception as e:
        log(f"[WARN] Could not save {AMAZON_USAGE_PATH.name}: {e}")


def attempt_key(row: pd.Series, missing_field: str | None) -> str:
    # The author is part of the key, so correcting it makes the row eligible again.
    return "|".join([
        missing_field or "any",
        books.build_duplicate_key(clean_text(row.get("Title", "")),
                                  clean_text(row.get("Author", ""))),
    ])


def load_attempts() -> dict:
    try:
        data = json.loads(ATTEMPTS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except Exception as e:
        log(f"[WARN] Ignoring unreadable {ATTEMPTS_PATH.name}: {e}")
        return {}


def save_attempts(attempts: dict) -> None:
    tmp = ATTEMPTS_PATH.with_suffix(".tmp")
    try:
        tmp.write_text(json.dumps(attempts, indent=0, sort_keys=True), encoding="utf-8")
        tmp.replace(ATTEMPTS_PATH)
    except Exception as e:
        log(f"[WARN] Could not save {ATTEMPTS_PATH.name}: {e}")


def free_sources_key(row: pd.Series, missing_field: str | None) -> str:
    """Where to note that Google, Open Library, Apple and Hardcover had nothing.

    Kept apart from the attempt itself: such a book has not been tried, since
    Amazon was not asked, but asking the free sources again within the week
    spends Google's daily quota on an answer we already have.
    """
    return "free sources empty|" + attempt_key(row, missing_field)


def attempted_recently(attempts: dict, key: str, days: int) -> bool:
    try:
        last = datetime.date.fromisoformat(attempts[key])
    except (KeyError, TypeError, ValueError):
        return False
    return (datetime.date.today() - last).days < days


def should_skip(row: pd.Series, min_age_days: int, skip_threshold: float,
                force: bool, skip_decided: bool = True,
                skip_read: bool = True) -> tuple[bool, str, str]:
    """
    Return (skip, reason, kind) where kind is "finished", "date" or "fill".
    skip=True means don't re-enrich this row.
    """
    # Checked ahead of --force on purpose: force means "ignore the freshness and
    # fill-% thresholds", not "spend a lookup and a model call on a book that
    # will never be offered again". --include-read/--include-decided opt out.
    reason = finished_reason(row, skip_decided, skip_read)
    if reason:
        return True, reason, "finished"

    if force:
        return False, "", ""

    # Date check
    if min_age_days > 0:
        last = last_enriched_date(row)
        if last is not None:
            age = (datetime.date.today() - last).days
            if age < min_age_days:
                return True, f"enriched {age}d ago (< {min_age_days}d)", "date"

    # Fill% check
    pct = fill_pct(row)
    if pct >= skip_threshold:
        return True, f"{pct:.0%} complete (>= {skip_threshold:.0%} threshold)", "fill"

    return False, "", ""

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


def recategorize(df: pd.DataFrame) -> list[tuple[str, str, str, str]]:
    """Re-derive LengthCategory from PageCount wherever the two disagree.

    Local arithmetic only — no lookup and no model call — so it runs over the
    whole sheet regardless of the skip rules. Rows without a page count are left
    alone: there is nothing to grade them against.

    Returns the rows it changed as (title, page count, old category, new).
    """
    if "LengthCategory" not in df.columns or "PageCount" not in df.columns:
        return []

    changes = []
    for idx, row in df.iterrows():
        derived = page_count_to_length_category(row.get("PageCount", ""))
        if not derived:
            continue  # no usable page count — leave whatever is there
        current = clean_text(row.get("LengthCategory", ""))
        if current == derived:
            continue
        df.at[idx, "LengthCategory"] = derived
        changes.append((
            clean_text(row.get("Title", "")),
            clean_text(row.get("PageCount", "")),
            current or "(blank)",
            derived,
        ))
    return changes


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

        # LengthCategory needs no special case: an empty cell is filled by the
        # generic path below, and the block after this loop re-derives it from
        # the page count whenever one was just filled in.

        # All other fields: only fill if currently empty
        if current:
            continue

        # Normalise list fields before storing
        if field in ("Genre", "Tropes", "Triggers"):
            new_str = normalize_csv_list(new_str)
        elif field == "Spice Tags" and normalize_spice_tags:
            new_str = normalize_spice_tags(new_str)
            if not new_str:
                continue
        row[field] = new_str
        changed.append(field)

    # Always (re-)derive LengthCategory if PageCount was just filled — including
    # over an existing value, which is how a row that a missing page count once
    # stamped "Short" gets the category its real page count deserves.
    if "PageCount" in changed:
        lc = page_count_to_length_category(row["PageCount"])
        if lc and lc != clean_text(row.get("LengthCategory", "")):
            row["LengthCategory"] = lc
            if "LengthCategory" not in changed:
                changed.append("LengthCategory")

    # Stamp Last Enriched whenever any field was filled — except for a row with
    # no author, which is not finished, it is waiting on you. Stamping it would
    # hide it behind the freshness rule for months, and adding the author (the
    # one thing that would let it enrich properly) would appear to do nothing.
    if changed and not is_vague_author(row.get("Author", "")):
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
    parser.add_argument("--force",          action="store_true", help="Ignore the age and fill%% skip rules (read/removed books are still skipped)")
    parser.add_argument("--retry-after-days", type=int, default=0, help="Skip books tried within N days that are still missing something (0=off)")
    parser.add_argument("--include-decided", action="store_true", help="Also process books with a selector decision of Read/Ignored/Removed")
    parser.add_argument("--include-read",    action="store_true", help="Also process books whose Read column is Yes")
    parser.add_argument("--recategorize",    action="store_true", help="Only re-grade LengthCategory from existing page counts, then exit (no lookups, no API cost)")
    parser.add_argument("--no-recategorize", action="store_true", help="Skip the category re-grade that normally runs before enrichment")
    parser.add_argument("--no-amazon",       action="store_true", help="Do not open any Amazon page in this run")
    parser.add_argument("--amazon-pages-per-window", type=int, default=DEFAULT_AMAZON_PAGES_PER_WINDOW,
                        help=f"Amazon pages to open before resting it, across runs (default {DEFAULT_AMAZON_PAGES_PER_WINDOW}; 0 = none)")
    parser.add_argument("--amazon-rest-hours", type=float, default=DEFAULT_AMAZON_REST_HOURS,
                        help=f"Hours to leave Amazon alone after each window (default {DEFAULT_AMAZON_REST_HOURS}; doubled after a block)")
    parser.add_argument("--ai-only",         action="store_true", help="Only run the model on rows that have a description but no tropes or spice tags; no book lookups")
    args = parser.parse_args()

    missing_field = args.missing.strip() or None
    if missing_field and missing_field not in FILLABLE_FIELDS:
        sys.exit(f"ERROR: --missing must be one of: {', '.join(FILLABLE_FIELDS)}")
    if args.ai_only and missing_field:
        sys.exit("ERROR: --ai-only picks its own rows; leave out --missing")

    _open_log()
    try:
        _run(args, missing_field)
    except KeyboardInterrupt:
        log("Interrupted by user.")
    except Exception as e:
        log(f"FATAL: {e}")
        log(traceback.format_exc())
    finally:
        # Pages opened count against the window whatever happened to the run,
        # dry runs and interrupted runs included.
        if _AMAZON_STATE_LOADED:
            save_amazon_usage()
            log(f"Amazon pages opened this run: {books.AMAZON_PAGES_USED}")
        _close_log()


def _run(args, missing_field: str | None) -> None:
    global SHEET_NAME
    SHEET_NAME = args.sheet

    min_age_days   = args.min_age_days
    skip_threshold = args.skip_threshold
    force          = args.force
    skip_decided   = CFG_SKIP_DECIDED and not args.include_decided
    skip_read      = CFG_SKIP_READ    and not args.include_read
    do_recategorize = (CFG_RECATEGORIZE or args.recategorize) and not args.no_recategorize

    log("=" * 60)
    log("Re-Enrichment Script for books_output.xlsx")
    log(f"Excel      : {EXCEL_PATH}")
    log(f"Filter     : {'missing ' + missing_field if missing_field else 'any missing field'}")
    log(f"Limit      : {args.limit or 'no limit'}")
    log(f"Skip age   : {'disabled' if min_age_days == 0 else f'< {min_age_days} days'}")
    log(f"Skip fill% : {skip_threshold:.0%}+ complete")
    log(f"Skip decided: {'yes — read/ignored/removed left alone' if skip_decided else 'no'}")
    log(f"Skip read  : {'yes — Read=Yes left alone' if skip_read else 'no'}")
    log(f"Recategorize: {'yes' if do_recategorize else 'no'}")
    log(f"Force      : {force}")
    log(f"Retry after: {f'{args.retry_after_days} days' if args.retry_after_days > 0 else 'off'}")
    log(f"Dry run    : {args.dry_run}")

    global _AMAZON_STATE_LOADED
    per_window = getattr(args, "amazon_pages_per_window", DEFAULT_AMAZON_PAGES_PER_WINDOW)
    rest_hours = getattr(args, "amazon_rest_hours", DEFAULT_AMAZON_REST_HOURS)
    books.AMAZON_DISABLED = bool(getattr(args, "no_amazon", False))
    books.AMAZON_WINDOW_PAGES = max(0, per_window)
    books.AMAZON_REST_SECONDS = max(0.0, rest_hours) * 3600
    books.AMAZON_BLOCK_REST_SECONDS = books.AMAZON_REST_SECONDS * AMAZON_BLOCK_REST_FACTOR
    (books.AMAZON_WINDOW_USED, books.AMAZON_REST_UNTIL,
     books.AMAZON_LAST_PAGE) = load_amazon_state()
    books.AMAZON_PAGES_USED = 0
    _AMAZON_STATE_LOADED = True
    if books.AMAZON_DISABLED:
        log("Amazon     : off for this run")
    elif books.AMAZON_REST_UNTIL > time.time():
        log("Amazon     : resting until "
            + datetime.datetime.fromtimestamp(books.AMAZON_REST_UNTIL).strftime("%a %H:%M"))
    else:
        books.amazon_skip_reason()   # lets a finished rest or a long gap reset the window
        left = max(0, books.AMAZON_WINDOW_PAGES - books.AMAZON_WINDOW_USED)
        log(f"Amazon     : up to {left} pages, then rest {rest_hours:g}h")
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

    # Spice Tags arrived after most sheets were written, and --missing cannot
    # look for a column that is not there yet.
    if "Spice Tags" not in df.columns:
        df["Spice Tags"] = ""
        log("Added 'Spice Tags' column to sheet.")
        if not args.dry_run:
            save_excel(df)
            log("Saved column addition to file.")

    # Force string columns that pandas may have inferred as float64
    # (happens when a column is entirely empty/NaN on load).
    # Without this, writing a date string like '2026-06-17' crashes.
    STRING_COLS = [
        "Last Enriched", "Description", "Genre", "Tropes", "Triggers", "Spice Tags",
        "AgeRange", "LengthCategory", "ISBN_13", "ISBN_10", "ASIN",
        "Lookup Source", "Metadata Enriched",
    ]
    for col in STRING_COLS:
        if col in df.columns and df[col].dtype != object:
            df[col] = df[col].astype(object)

    log(f"Loaded {len(df)} rows from '{SHEET_NAME}'")

    # Free, like the re-grade below: blank Spice Tags read off the row's own
    # Tropes/Triggers, so rows the model may never revisit still get filters.
    if fill_spice_from_tropes:
        filled = fill_spice_from_tropes(df)
        log(f"Spice from tropes : {filled} row(s)")
        if filled and not args.dry_run:
            save_excel(df)

    # Re-grade categories before anything else: it is free, it needs no network,
    # and it is the only thing that fixes rows whose page count was always fine
    # but whose category was written under different boundaries.
    if do_recategorize:
        changes = recategorize(df)
        log(f"Recategorized     : {len(changes)} row(s)")
        for title, pages, was, now in changes[:20]:
            log(f"  {title[:45]:45} {pages:>6} pg  {was} -> {now}")
        if len(changes) > 20:
            log(f"  ... and {len(changes) - 20} more")
        if changes and not args.dry_run:
            save_excel(df)
            log("Saved category changes.")

    if args.recategorize:
        log("--recategorize given: category pass only, nothing else to do.")
        return

    # Two-pass eligibility: needs_work (has gaps) AND not skipped by date/fill%
    if getattr(args, "ai_only", False):
        # A description but no tropes or spice tags: the model has not read
        # this description, or ran before Spice Tags existed. Cells that are
        # already filled stay as they are.
        has_gaps_indices = [
            i for i, row in df.iterrows()
            if clean_text(row.get("Title", "")) and clean_text(row.get("Description", ""))
            and (is_empty(row.get("Tropes", "")) or is_empty(row.get("Spice Tags", "")))
            and not is_vague_author(row.get("Author", ""))
        ]
    else:
        has_gaps_indices = [
            i for i, row in df.iterrows()
            if clean_text(row.get("Title", "")) and needs_work(row, missing_field)
        ]
    log(f"Rows with gaps    : {len(has_gaps_indices)}")

    eligible_indices = []
    skipped_by_kind = {"finished": 0, "date": 0, "fill": 0, "attempt": 0, "no author": 0}
    attempts = load_attempts()
    for i in has_gaps_indices:
        skip, reason, kind = should_skip(
            df.loc[i], min_age_days, skip_threshold, force, skip_decided, skip_read
        )
        if not skip and args.retry_after_days > 0 and attempted_recently(
                attempts, attempt_key(df.loc[i], missing_field), args.retry_after_days):
            skip, kind = True, "attempt"
        # No description is taken for a row with no author, so a run after only
        # descriptions would spend a lookup on it for nothing.
        if not skip and missing_field == "Description" and is_vague_author(
                df.at[i, "Author"] if "Author" in df.columns else ""):
            skip, kind = True, "no author"
        if skip:
            skipped_by_kind[kind] += 1
        else:
            eligible_indices.append(i)

    log(f"Skipped (read/removed): {skipped_by_kind['finished']}")
    log(f"Skipped (recent)  : {skipped_by_kind['date']}")
    log(f"Skipped (fill %)  : {skipped_by_kind['fill']}")
    if args.retry_after_days > 0:
        log(f"Skipped (tried)   : {skipped_by_kind['attempt']}")
    if skipped_by_kind["no author"]:
        log(f"Skipped (no author): {skipped_by_kind['no author']} — add an author to get a description")
    log(f"Will process      : {len(eligible_indices)}")

    # Books the free sources had nothing for this week can only be filled by
    # Amazon. Re-asking the free sources each run spent the run, and Google's
    # quota, on the same books from the front of the sheet. So they go first
    # while Amazon has pages left in its window, straight to Amazon, and wait
    # otherwise; the rest of the run goes to books not yet looked at.
    amazon_only: set = set()
    if missing_field == "Description" and args.retry_after_days > 0:
        waiting = [i for i in eligible_indices if attempted_recently(
            attempts, free_sources_key(df.loc[i], missing_field), args.retry_after_days)]
        if waiting:
            if books.amazon_skip_reason():
                left = 0
            elif books.AMAZON_WINDOW_PAGES is None:
                left = len(waiting) * 3
            else:
                left = max(0, books.AMAZON_WINDOW_PAGES - books.AMAZON_WINDOW_USED)
            first = []
            for i in waiting:
                # The description page, plus a search when there is no ASIN and
                # a page-count page when that is missing too.
                cost = (1 + (not clean_text(df.at[i, "ASIN"] if "ASIN" in df.columns else ""))
                        + needs_work(df.loc[i], "PageCount"))
                if cost <= left:
                    first.append(i)
                    left -= cost
            held = set(waiting) - set(first)
            eligible_indices = first + [i for i in eligible_indices
                                        if i not in held and i not in set(first)]
            amazon_only = set(first)
            log(f"Amazon only       : {len(first)} (free sources had nothing this week)")
            if held:
                log(f"Waiting for Amazon: {len(held)} (left for a run with Amazon pages free)")

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
    vague_author_rows: list[str] = []
    lookup_errors     = 0
    consecutive_errors = 0
    # A quota block fails every row identically. Without a circuit breaker the
    # run keeps going for hours, learns nothing, and stamps each row on the way
    # past so none of them are retried for six months.
    ERROR_ABORT_STREAK = 25
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
        if is_vague_author(author):
            log("  No author — identifiers only, no description or AI-derived fields. "
                "Add an author and re-run to fill them.")
            vague_author_rows.append(title)

        if args.dry_run:
            log("  [DRY RUN] Would call lookup_book_metadata -- skipping.")
            continue

        try:
            # Hand the lookup what this row is actually missing, and the ASIN it
            # already holds. Both let it skip work that cannot change the row:
            # the model call only ever returns Genre/AgeRange/Tropes/Triggers,
            # and an Amazon search for a known ASIN re-derives what we passed in.
            if getattr(args, "ai_only", False):
                # No lookups: the model reads the description already in the row.
                result = {**books.ai_content_enrichment(
                    title=title, author=author,
                    description=clean_text(row.get("Description", "")),
                    genre=clean_text(row.get("Genre", "")),
                    age_range=clean_text(row.get("AgeRange", "")),
                    page_count=row.get("PageCount", ""),
                ), "Lookup Error": ""}
            else:
                result = lookup_book_metadata(
                    title, author,
                    need_fields=set(missing_now),
                    known_asin=clean_text(row.get("ASIN", "")),
                    free_sources=idx not in amazon_only,
                )

            # If Google Books/OpenLibrary found nothing for PageCount but we have
            # an ASIN, try Amazon directly — Kindle books always have "Print length"
            # Only for rows that lack one: apply_result would discard it anyway,
            # and each needless product page brings Amazon's captcha closer.
            if ("PageCount" in missing_now and not getattr(args, "ai_only", False)
                    and not books.amazon_skip_reason()
                    and str(result.get("PageCount") or "") in ("", "0")):
                asin = clean_text(row.get("ASIN",""))
                if asin:
                    books.note_amazon_page()
                    amz_pages = fetch_amazon_page_count(asin)
                    if amz_pages > 0:
                        result["PageCount"] = float(amz_pages)
                        log(f"  [AMAZON] Found PageCount={amz_pages} via ASIN")

            lookup_error  = clean_text(result.get("Lookup Error", ""))
            lookup_failed = bool(lookup_error)
            # Amazon showing a different book means the other sources answered
            # and had nothing, so the book was tried. Any other error - a quota,
            # an outage, a captcha or a blocked page - says nothing about the
            # book and must not hide it.
            if not lookup_failed or lookup_error.startswith("Amazon page for"):
                attempts[attempt_key(row, missing_field)] = datetime.date.today().isoformat()
            # The error starts "Amazon skipped" only when every free source
            # answered and none had a blurb. A quota or an outage comes first
            # in the error, so a book is never parked on an answer we did not get.
            if (missing_field == "Description" and idx not in amazon_only
                    and lookup_error.startswith("Amazon skipped")):
                attempts[free_sources_key(row, missing_field)] = datetime.date.today().isoformat()
            if lookup_failed:
                lookup_errors += 1
                # Once Amazon is paused the other sources still work, so a row
                # only Amazon could fill is not a sign that the run is stuck.
                if not lookup_error.startswith("Amazon skipped"):
                    consecutive_errors += 1
                log(f"  [LOOKUP FAILED] {result['Lookup Error']}")
            else:
                consecutive_errors = 0

            new_row, changed = apply_result(row, result)

            if lookup_failed:
                # Do not stamp Last Enriched. The row was never actually looked
                # up, so hiding it behind the freshness rule would be recording
                # a result we never got.
                if changed:
                    for col in changed:
                        if col != "Last Enriched":
                            df.at[idx, col] = new_row[col]
                skipped += 1
                if consecutive_errors >= ERROR_ABORT_STREAK:
                    log("")
                    log(f"ABORTING: {consecutive_errors} lookups failed in a row.")
                    log(f"  Last error: {result['Lookup Error']}")
                    log("  Nothing has been date-stamped, so re-running once the")
                    log("  cause is cleared will pick these rows up again.")
                    break
            elif changed:
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
                # A row with no author is the exception: it stays eligible so
                # that filling the Author column is enough to get it enriched.
                if is_vague_author(author):
                    skipped += 1
                    log("  -- No new data found (no author — left eligible for re-run)")
                else:
                    df.at[idx, "Last Enriched"] = datetime.date.today().isoformat()
                    skipped += 1
                    log(f"  -- No new data found (date stamped, will skip for {args.min_age_days}d)")

            rows_touched = updated + skipped
            if rows_touched > 0 and rows_touched % save_interval == 0:
                log(f"\n  [AUTO-SAVE] Saving after {rows_touched} rows processed...")
                save_excel(df)
                save_attempts(attempts)
                save_amazon_usage()
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
        save_attempts(attempts)
        log(f"Saved: {EXCEL_PATH}")

    log("")
    log("=" * 60)
    log(f"DONE")
    log(f"Processed : {processed}")
    log(f"Updated   : {updated}")
    log(f"No change : {skipped}")
    log(f"Errors    : {errors}")
    if lookup_errors:
        log("")
        log(f"Lookups failed : {lookup_errors} — these rows were NOT date-stamped")
        log("                 and stay eligible for a later run.")
    if vague_author_rows:
        log("")
        log(f"No author : {len(vague_author_rows)} — identifiers only, no AI-derived")
        log("            fields. Fill the Author column for these and re-run:")
        for t in vague_author_rows[:20]:
            log(f"              - {t}")
        if len(vague_author_rows) > 20:
            log(f"              ... and {len(vague_author_rows) - 20} more")
    if not args.dry_run:
        log(f"Log file  : {LOG_PATH}")
    log("=" * 60)


if __name__ == "__main__":
    main()
