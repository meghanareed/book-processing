# amazon_owned_books.py
#
# Scrapes https://www.amazon.com/hz/mycd/digital-console/contentlist/booksAll/dateDsc/
# Skips any book whose card contains the READ badge.
# For every unread book:
#   - Already in "All Books" (by DuplicateKey OR ASIN)  -> set Owned = "Yes"
#   - Brand new                                          -> enrich via books.py pipeline,
#                                                           append with Owned = "Yes"
#
# Selectors verified against real Amazon library HTML (March 2026):
#   Card root  : div.digital_entity_details   (one per book)
#   Title      : div[id^="content-title-"]    (first child of card root)
#   Author     : div[id^="content-author-"]   (sibling inside card root)
#   ASIN       : extracted from element id, e.g. "content-title-B0CKJSB3XN"
#   READ badge : div#content-read-badge       (sibling of card root, OUTSIDE it,
#                                              inside the same DigitalEntitySummary container)
#   Next page  : a#page-RIGHT_PAGE

import re
import sys
import time
import importlib.util
from pathlib import Path

import pandas as pd
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
    Error as PlaywrightError,
)

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------
import datetime
import traceback
import sys as _sys

# Windows consoles and pipes default to cp1252, which cannot encode the ✓ ✗ →
# characters logged below; printing one raises UnicodeEncodeError mid-run.
# Force UTF-8 so this holds however the script is started.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

_log_file = None   # opened immediately after LOG_PATH is defined (see below)
_log_buffer: list[str] = []  # captures messages emitted before the file is open

def log(msg: str = "", *, also_print: bool = True) -> None:
    """Write msg to stdout (flushed for live streaming) and the log file.

    The launcher runs this script with `python -u` and reads its stdout
    line-by-line into the Output Log panel, so a single flushed print() is
    all that's needed. Previously this also printed to stderr, which — with
    the launcher merging stderr into stdout — made every line appear twice
    in the panel.
    """
    ts   = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    if also_print:
        # Single flushed write to stdout — appears in the launcher panel
        # (and in a console, if run directly) in real time.
        try:
            print(line, flush=True)
        except Exception:
            pass
    if _log_file:
        try:
            _log_file.write(line + "\n")
            _log_file.flush()
        except Exception:
            pass
    else:
        # File not open yet — buffer so nothing is silently lost
        _log_buffer.append(line)

def _open_log() -> None:
    """No-op: log file is opened at module level right after LOG_PATH is defined.
    Kept for backward compatibility. Falls back to opening the file if it failed earlier."""
    global _log_file
    if _log_file is not None:
        return  # Already open — don't re-open and lose early log lines
    try:
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        _log_file = open(LOG_PATH, "a", encoding="utf-8", buffering=1)
        for _buffered_line in _log_buffer:
            _log_file.write(_buffered_line + "\n")
        _log_buffer.clear()
        log(f"Log file (late open): {LOG_PATH}")
    except Exception as e:
        log(f"WARNING: Could not open log file: {e}")

def _close_log() -> None:
    if _log_file:
        try:
            _log_file.close()
        except Exception:
            pass

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
import json

# SCRIPT_DIR holds the code (and launcher_config.json); DATA_DIR holds the
# spreadsheets and logs.  They are the same folder in the old layout and
# separate once the code is cloned out of OneDrive.
SCRIPT_DIR  = Path(__file__).resolve().parent
DATA_DIR    = Path(r"C:\Users\megha\OneDrive\Documents\Reading")

EXCEL_PATH  = DATA_DIR / "books_output.xlsx"
BOOKS_PY    = SCRIPT_DIR / "books.py"
SHEET_NAME  = "All Books"

AMAZON_LIBRARY_URL = (
    "https://www.amazon.com/hz/mycd/digital-console/contentlist/booksAll/dateDsc/"
)

# Load settings from launcher config if available
_launcher_config = {}
# book_launcher.py writes the config next to itself, so look there first.
# Fall back to DATA_DIR for the old layout where code and data shared a folder.
_config_path = SCRIPT_DIR / "launcher_config.json"
if not _config_path.exists():
    _config_path = DATA_DIR / "launcher_config.json"
if _config_path.exists():
    try:
        _launcher_config = json.loads(_config_path.read_text(encoding="utf-8"))
        print(f"[CONFIG] Loaded settings from launcher_config.json")
    except Exception as e:
        print(f"[CONFIG] Could not load launcher config: {e}")

HEADLESS        = False
TEST_PAGE_LIMIT = 0      # 0 = all pages; positive int = stop early for testing
START_PAGE      = _launcher_config.get("amazon_owned_books", {}).get("start_page", 1)
_max_pages      = _launcher_config.get("amazon_owned_books", {}).get("max_pages", 0)
END_PAGE        = (START_PAGE + _max_pages - 1) if _max_pages > 0 else 0  # 0 = no limit (scrape everything)
PAGE_LOAD_WAIT  = 3.5    # seconds after clicking Next
LOGIN_TIMEOUT   = 300    # seconds the user has to log in (5 min)
BATCH_SIZE      = 10     # scrape this many pages, then merge+save, then continue in same session

# Create logs folder if it doesn't exist
LOGS_DIR = EXCEL_PATH.parent / "logs"
LOGS_DIR.mkdir(exist_ok=True)

# Log file written next to the Excel — timestamped so runs never overwrite each other
LOG_PATH = LOGS_DIR / f"amazon_scrape_{__import__('datetime').datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

# Open the log file immediately so ALL log() calls — including import-time ones
# from _load_books_module() below — are written to disk straight away.
# This is the main fix for "nothing in the launcher log panel" under pythonw.
try:
    _log_file = open(LOG_PATH, "w", encoding="utf-8", buffering=1)  # line-buffered
    for _buffered_line in _log_buffer:
        _log_file.write(_buffered_line + "\n")
    _log_buffer.clear()
    _log_file.write(
        f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Log file: {LOG_PATH}\n"
    )
    _log_file.flush()
except Exception as _log_open_err:
    print(f"[WARNING] Could not open log file: {_log_open_err}", flush=True)

# ---------------------------------------------------------------------------
# IMPORT SHARED HELPERS FROM books.py
# ---------------------------------------------------------------------------
# We load books.py as a module so we can reuse its enrichment pipeline
# without copy-pasting.  The OpenAI client and config flags inside books.py
# are initialised on import; make sure OPENAI_API_KEY is set in the env.

def _load_books_module():
    if not BOOKS_PY.exists():
        log(f"WARNING: books.py not found at {BOOKS_PY}")
        log("         Metadata enrichment will be skipped.")
        return None
    spec = importlib.util.spec_from_file_location("books", BOOKS_PY)
    mod  = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
        return mod
    except Exception as e:
        log(f"WARNING: Could not load books.py ({e}). Enrichment disabled.")
        return None

books_mod = _load_books_module()

def _books_fn(name):
    """Return a function from books.py, or None if unavailable."""
    return getattr(books_mod, name, None) if books_mod else None

lookup_book_metadata    = _books_fn("lookup_book_metadata")
normalize_csv_list      = _books_fn("normalize_csv_list")
page_count_to_length    = _books_fn("page_count_to_length_category")
is_enriched_enough      = _books_fn("is_enriched_enough")
AI_ENRICH_SLEEP         = getattr(books_mod, "AI_ENRICH_SLEEP_SECONDS", 0.2) if books_mod else 0.2

# ---------------------------------------------------------------------------
# LOCAL HELPERS  (mirror books.py helpers so we work even if import fails)
# ---------------------------------------------------------------------------

def clean(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def normalize(value: str) -> str:
    v = clean(value).lower()
    v = re.sub(r"[^\w\s]", "", v)
    v = re.sub(r"\s+", " ", v)
    return v.strip()


def build_duplicate_key(title: str, author: str) -> str:
    return f"{normalize(title)}|{normalize(author)}"


def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Add any missing columns this script writes to."""
    for col in ["Owned", "Read", "Skip Storygraph"]:
        if col not in df.columns:
            df[col] = ""
    return df


def save_excel(df: pd.DataFrame) -> None:
    """
    Write All Books sheet back while preserving all other sheets.
    Uses atomic write (temp file → rename) so the original is never
    left as 0kb if the write crashes or is interrupted.
    Also keeps a .bak of the last good save.
    """
    import shutil, tempfile, os

    # Read all existing sheets
    try:
        all_sheets: dict = pd.read_excel(EXCEL_PATH, sheet_name=None)
    except Exception:
        all_sheets = {}
    all_sheets[SHEET_NAME] = df

    # Write to a temp file in the same directory
    tmp_path = EXCEL_PATH.with_suffix(".tmp.xlsx")
    try:
        with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
            for name, sheet_df in all_sheets.items():
                sheet_df.to_excel(writer, sheet_name=name, index=False)
    except Exception as e:
        # Clean up bad temp file, leave original untouched
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise RuntimeError(f"Excel write failed (original intact): {e}") from e

    # Back up the current good file before replacing it
    bak_path = EXCEL_PATH.with_suffix(".bak.xlsx")
    try:
        if EXCEL_PATH.exists():
            shutil.copy2(EXCEL_PATH, bak_path)
    except Exception:
        pass  # backup failure is non-fatal

    # Atomic replace — on Windows replace() is not truly atomic but is
    # far safer than writing directly to the target file
    try:
        tmp_path.replace(EXCEL_PATH)
    except Exception as e:
        raise RuntimeError(f"Could not replace Excel file: {e}") from e

# ---------------------------------------------------------------------------
# ENRICHMENT  (calls books.py pipeline, falls back gracefully)
# ---------------------------------------------------------------------------

def enrich_row(row: dict) -> dict:
    """
    Fill metadata fields using the same pipeline as books.py.
    `row` is mutated in place and also returned.
    """
    if lookup_book_metadata is None:
        return row

    title  = clean(row.get("Title"))
    author = clean(row.get("Author"))

    # Skip if already well-enriched
    if is_enriched_enough and is_enriched_enough(row):
        row["Metadata Enriched"] = "Yes"
        return row

    log(f"    Enriching: {title} | {author}")
    try:
        result = lookup_book_metadata(title, author)
    except Exception as e:
        log(f"    Enrichment error: {e}")
        return row

    for field in [
        "ISBN_10", "ISBN_13", "ASIN", "Lookup Source", "Description",
        "Genre", "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
    ]:
        current = clean(row.get(field))
        new_val = clean(result.get(field))
        if not current and new_val:
            row[field] = new_val

    # Normalise list fields
    if normalize_csv_list:
        for f in ("Genre", "Tropes", "Triggers"):
            row[f] = normalize_csv_list(clean(row.get(f)))

    # Derive length category if still missing
    if not clean(row.get("LengthCategory")) and page_count_to_length:
        row["LengthCategory"] = page_count_to_length(row.get("PageCount", ""))

    if is_enriched_enough and is_enriched_enough(row):
        row["Metadata Enriched"] = "Yes"

    time.sleep(AI_ENRICH_SLEEP)
    return row

# ---------------------------------------------------------------------------
# LOGIN DETECTION
# ---------------------------------------------------------------------------

def is_logged_in(page) -> bool:
    url = page.url.lower()
    if any(s in url for s in ("signin", "ap/signin", "ap/register", "authentication")):
        return False
    try:
        if page.locator("form[name='signIn'], input#ap_email").count() > 0:
            return False
    except Exception:
        pass
    try:
        # Library content present = logged in
        if page.locator("div.digital_entity_details").count() > 0:
            return True
    except Exception:
        pass
    return False


def library_url(page_num: int = 1) -> str:
    """Build the Amazon library URL for a given page number."""
    if page_num <= 1:
        return AMAZON_LIBRARY_URL
    return f"{AMAZON_LIBRARY_URL}?pageNumber={page_num}"


def wait_for_login(page, start_page: int = 1) -> None:
    """
    Navigate to the Amazon library (at start_page) and ensure cards are loaded.

    Amazon's library is a SPA -- calling goto() twice causes a navigation
    conflict ("interrupted by another navigation").  The fix is a single goto,
    then wait up to LOGIN_TIMEOUT seconds for the book cards to appear.
    If they never appear (not logged in) we pause and let the user sign in,
    then wait again for the cards -- still no second goto needed because
    Amazon redirects back to the library URL automatically after login.
    """
    target = library_url(start_page)
    log(f"Opening Amazon library (page {start_page})...")
    try:
        page.goto(target, wait_until="domcontentloaded")
    except (PlaywrightError, PlaywrightTimeoutError):
        pass  # SPA may fire its own internal redirect; that is fine

    # Phase 1: wait up to 8s to see if we are already logged in
    try:
        page.wait_for_selector("div.digital_entity_details", timeout=8_000)
        log(f"Already logged in -- cards visible on page {start_page}.")
        return
    except PlaywrightTimeoutError:
        pass

    # Phase 2: not logged in -- prompt user and then watch for cards automatically.
    # IMPORTANT: do NOT use input() here -- under pythonw there is no console,
    # so input() blocks forever with no way for the user to respond.
    log("\n" + "=" * 60)
    log("  Amazon sign-in required.")
    log("  Please log into Amazon in the browser window that just opened.")
    log("  The script will continue automatically once your library loads.")
    log("  (You have up to 5 minutes to complete sign-in.)")
    log("=" * 60)

    # Phase 3: poll for cards — fires automatically once the user finishes signing in.
    # Amazon redirects back to the library page after login, so we just wait for cards.
    if start_page > 1:
        log(f"Navigating to start page {start_page}...")
        try:
            page.goto(target, wait_until="domcontentloaded")
        except (PlaywrightError, PlaywrightTimeoutError):
            pass

    log("Waiting for library cards (up to 5 minutes)...")
    try:
        page.wait_for_selector("div.digital_entity_details", timeout=LOGIN_TIMEOUT * 1000)
        log(f"Library loaded — starting scrape from page {start_page}.")
    except PlaywrightTimeoutError:
        log("WARNING: Book cards still not visible after 5 minutes — scraping will proceed anyway.")
# ---------------------------------------------------------------------------
# SCRAPING  — exact selectors from real Amazon HTML
# ---------------------------------------------------------------------------
# Structure per book (simplified):
#
#  div.DigitalEntitySummary-module__container_*          ← outer wrapper
#    div.DigitalEntitySummary-module__entity_information_container_*
#      div.digital_entity_details                        ← CARD ROOT (our anchor)
#        div#content-title-{ASIN}.digital_entity_title   ← title wrapper
#          div[role="heading"]                           ← actual title text
#        div#content-author-{ASIN}.information_row       ← author text
#    div#content-read-badge.information_row.readBadgeText  ← READ badge (OUTSIDE card root)
#
# Strategy:
#   1. Select every div.digital_entity_details
#   2. Extract title from first div[id^="content-title-"] inside it
#   3. Extract author from first div[id^="content-author-"] inside it
#   4. Extract ASIN from the id attribute (e.g. "content-title-B0CKJSB3XN")
#   5. Walk UP to the DigitalEntitySummary container and look for content-read-badge there

ASIN_RE = re.compile(r"[A-Z0-9]{10}")


def _safe_inner_text(locator) -> str:
    try:
        if locator.count() > 0:
            return clean(locator.first.inner_text())
    except Exception:
        pass
    return ""


def scrape_page_via_js(page) -> list[dict]:
    """
    Use page.evaluate to extract all book data in a single JS call —
    much faster and immune to Playwright per-element timeouts.
    Returns list of {title, author, asin, is_read}.
    """
    script = """
    () => {
        const books = [];
        // Each card root
        document.querySelectorAll('div.digital_entity_details').forEach(card => {
            // Title element: div[id^="content-title-"]
            const titleEl = card.querySelector('div[id^="content-title-"]');
            // Actual text is in the inner heading div
            const titleText = titleEl
                ? (titleEl.querySelector('div[role="heading"]') || titleEl).innerText.trim()
                : '';

            // Author element: div[id^="content-author-"]
            const authorEl = card.querySelector('div[id^="content-author-"]');
            const authorText = authorEl ? authorEl.innerText.trim() : '';

            // ASIN from the title element id, e.g. "content-title-B0CKJSB3XN"
            let asin = '';
            if (titleEl) {
                const m = (titleEl.id || '').match(/[A-Z0-9]{10}/);
                if (m) asin = m[0];
            }

            // READ badge is OUTSIDE the card root — walk up to the
            // DigitalEntitySummary container then search within it.
            // The container is typically 2–3 levels up from digital_entity_details.
            let container = card.parentElement;
            for (let i = 0; i < 5; i++) {
                if (!container) break;
                // Stop at the outermost summary container
                if (container.className && container.className.includes('DigitalEntitySummary-module__container')) break;
                container = container.parentElement;
            }
            const readBadge = container
                ? container.querySelector('div#content-read-badge')
                : null;
            const isRead = readBadge
                ? readBadge.innerText.trim().toUpperCase().includes('READ')
                : false;

            if (titleText) {
                books.push({ title: titleText, author: authorText, asin: asin, is_read: isRead });
            }
        });
        return books;
    }
    """
    try:
        return page.evaluate(script) or []
    except (PlaywrightError, PlaywrightTimeoutError) as e:
        log(f"    JS evaluation error: {e}")
        return []


def scrape_current_page(page) -> list[dict]:
    """Wait for content, scroll to load ALL lazy items, then extract via JS."""
    try:
        page.wait_for_selector("div.digital_entity_details", timeout=12_000)
    except PlaywrightTimeoutError:
        log("    Timed out waiting for book cards — page may be empty or slow.")
        return []

    # Scroll to the absolute bottom in steps to trigger all lazy-loaded cards,
    # then back to top.  Use document height rather than a fixed pixel count so
    # long pages are fully covered.
    try:
        # Get total page height and scroll in 500px chunks
        total_height = page.evaluate("document.body.scrollHeight")
        scrolled = 0
        step = 500
        while scrolled < total_height:
            page.evaluate(f"window.scrollBy(0, {step})")
            time.sleep(0.2)
            scrolled += step
            # Re-check height in case new content loaded and extended the page
            total_height = page.evaluate("document.body.scrollHeight")
        page.evaluate("window.scrollTo(0, 0)")
        time.sleep(0.3)
    except (PlaywrightError, PlaywrightTimeoutError):
        pass  # non-fatal

    books = scrape_page_via_js(page)
    log(f"    [SCRAPE] Extracted {len(books)} cards from current page.")
    return books


def go_to_next_page(page) -> bool:
    """
    Click the Next pagination link and wait for new cards to appear.
    Primary selector: a#page-RIGHT_PAGE (confirmed from real HTML).

    Diagnosis notes printed so you can see exactly why pagination stopped.
    """
    selectors = [
        "a#page-RIGHT_PAGE",                          # primary — confirmed
        "a[aria-label='Next']",
        "a.page-link[aria-label='Next']",
        "a:has-text('»')",
        "span.a-last:not(.a-disabled) a",
    ]

    for sel in selectors:
        try:
            btn = page.locator(sel).first
            if btn.count() == 0:
                continue
            if not btn.is_visible(timeout=2_000):
                log(f"    [PAGINATION] '{sel}' found but not visible — skipping.")
                continue

            # Only treat aria-disabled="true" as disabled — NOT class substring match,
            # which was a false-positive bug (Amazon class names contain "disabled" as
            # part of colour/style tokens on perfectly active buttons).
            aria_disabled = (btn.get_attribute("aria-disabled") or "").lower()
            if aria_disabled == "true":
                log(f"    [PAGINATION] '{sel}' is aria-disabled=true — last page reached.")
                return False

            log(f"    [PAGINATION] Clicking '{sel}'…")
            btn.click()

            # Wait for new cards rather than domcontentloaded — Amazon is a SPA and
            # domcontentloaded can fire before the new page content renders, OR it can
            # be interrupted by the SPA's own internal navigation causing a crash.
            try:
                page.wait_for_selector("div.digital_entity_details", timeout=12_000)
            except PlaywrightTimeoutError:
                log("    [PAGINATION] Cards did not appear after click — may be last page.")
                return False

            time.sleep(PAGE_LOAD_WAIT)
            return True

        except (PlaywrightError, PlaywrightTimeoutError) as e:
            log(f"    [PAGINATION] Error on selector '{sel}': {e}")
            continue

    log("    [PAGINATION] No Next button found in DOM — treating as last page.")
    return False


def scrape_amazon_library(page, start_page: int = 1, stop_after_page: int = 0) -> list[dict]:
    """
    Scrape pages start_page..stop_after_page (0 = no limit).
    The browser session must already be on the correct start page.
    """
    all_books: list[dict] = []
    page_num = start_page - 1   # incremented at the top of each iteration

    while True:
        page_num += 1
        log(f"\n  ── Page {page_num} ──────────────────────────────")

        books = scrape_current_page(page)
        read_ct   = sum(1 for b in books if b["is_read"])
        unread_ct = len(books) - read_ct
        log(f"  Page {page_num}: {len(books)} cards  ({read_ct} READ, {unread_ct} unread)  |  batch running total: {len(all_books) + len(books)}")
        all_books.extend(books)

        if TEST_PAGE_LIMIT and page_num >= TEST_PAGE_LIMIT:
            log(f"  TEST_PAGE_LIMIT={TEST_PAGE_LIMIT} — stopping early.")
            break

        if stop_after_page and page_num >= stop_after_page:
            log(f"  Batch end (page {stop_after_page}) — batch complete, browser stays open for next batch.")
            break

        advanced = go_to_next_page(page)
        if not advanced:
            log(f"  Pagination stopped after page {page_num} — see [PAGINATION] lines above for reason.")
            break

    return all_books

# ---------------------------------------------------------------------------
# MERGE LOGIC
# ---------------------------------------------------------------------------

def _build_lookups(df: pd.DataFrame) -> tuple[dict, dict]:
    """Return (dk_to_idx, asin_to_idx) position maps for the dataframe."""
    dk_to_idx: dict[str, list] = {}
    for idx, row in df.iterrows():
        dk = clean(row.get("DuplicateKey", "")).lower()
        if dk:
            dk_to_idx.setdefault(dk, []).append(idx)

    asin_to_idx: dict[str, list] = {}
    for idx, row in df.iterrows():
        a = clean(row.get("ASIN", "")).lower()
        if a:
            asin_to_idx.setdefault(a, []).append(idx)

    return dk_to_idx, asin_to_idx


def _find_matches(dk: str, asin: str, dk_to_idx: dict, asin_to_idx: dict) -> list:
    if dk in dk_to_idx:
        return dk_to_idx[dk]
    if asin and asin.lower() in asin_to_idx:
        return asin_to_idx[asin.lower()]
    return []


def merge_amazon_books(
    df: pd.DataFrame,
    unread: list[dict],
    read: list[dict],
) -> tuple[pd.DataFrame, int, int, int, int]:
    """
    Pass 1 — Unread books:
      • Already in sheet (DuplicateKey OR ASIN match)
            → set Owned = "Yes"  |  skip enrichment
      • Brand new
            → enrich, append with Owned = "Yes"

    Pass 2 — Read books (run after Pass 1 so newly appended rows are included):
      • Already in sheet (DuplicateKey OR ASIN match)
            → set Read = "Yes", Skip Storygraph = "Yes"
      • Not in sheet  → do nothing (we only track owned/unread books)

    Returns (updated_df, unread_updated, unread_appended, read_updated, read_skipped).
    """

    # ── Pass 1: unread ────────────────────────────────────────────────────────
    dk_to_idx, asin_to_idx = _build_lookups(df)

    unread_updated  = 0
    unread_appended = 0
    seen_dks: set[str] = set()
    new_rows: list[dict] = []

    for book in unread:
        title  = clean(book["title"])
        author = re.sub(r"^by\s+", "", clean(book["author"]), flags=re.IGNORECASE).strip()
        asin   = clean(book["asin"])
        dk     = build_duplicate_key(title, author)

        if dk in seen_dks:
            continue
        seen_dks.add(dk)

        matched = _find_matches(dk, asin, dk_to_idx, asin_to_idx)

        if matched:
            # Exists — update Owned only, NO enrichment
            for pos in matched:
                if clean(df.at[pos, "Owned"]) != "Yes":
                    df.at[pos, "Owned"] = "Yes"
                    unread_updated += 1
                if asin and not clean(df.at[pos, "ASIN"]):
                    df.at[pos, "ASIN"] = asin
            log(f"  [UPDATE]  {title} | {author}")
        else:
            # New — enrich then append
            log(f"  [NEW]     {title} | {author}  (ASIN: {asin or '?'})")
            new_row: dict = {col: "" for col in df.columns}
            new_row.update({
                "Title":        title,
                "Author":       author,
                "ASIN":         asin,
                "Owned":        "Yes",
                "DuplicateKey": dk,
                "Needs Review": "Yes",
                "Confidence":   1.0,
            })
            new_row = enrich_row(new_row)
            new_rows.append(new_row)
            unread_appended += 1

            # Register phantom so duplicates later in the same scrape hit UPDATE
            phantom = f"__new_{unread_appended}"
            dk_to_idx.setdefault(dk, []).append(phantom)
            if asin:
                asin_to_idx.setdefault(asin.lower(), []).append(phantom)

    if new_rows:
        df = pd.concat(
            [df, pd.DataFrame(new_rows, columns=df.columns)],
            ignore_index=True,
        )

    # ── Pass 2: read books ────────────────────────────────────────────────────
    # Rebuild lookups to include any newly appended rows from Pass 1
    dk_to_idx, asin_to_idx = _build_lookups(df)

    read_updated = 0
    read_skipped = 0
    seen_dks_read: set[str] = set()

    log(f"\n  Processing {len(read)} READ books from Amazon...")
    for book in read:
        title  = clean(book["title"])
        author = re.sub(r"^by\s+", "", clean(book["author"]), flags=re.IGNORECASE).strip()
        asin   = clean(book["asin"])
        dk     = build_duplicate_key(title, author)

        if dk in seen_dks_read:
            continue
        seen_dks_read.add(dk)

        matched = _find_matches(dk, asin, dk_to_idx, asin_to_idx)

        if matched:
            for pos in matched:
                changed = False
                if clean(df.at[pos, "Read"]) != "Yes":
                    df.at[pos, "Read"] = "Yes"
                    changed = True
                if clean(df.at[pos, "Skip Storygraph"]) != "Yes":
                    df.at[pos, "Skip Storygraph"] = "Yes"
                    changed = True
                if asin and not clean(df.at[pos, "ASIN"]):
                    df.at[pos, "ASIN"] = asin
                if changed:
                    read_updated += 1
            log(f"  [READ]    {title} | {author}")
        else:
            # Not in sheet — book was read but never tracked as owned/unread
            read_skipped += 1

    return df, unread_updated, unread_appended, read_updated, read_skipped

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main() -> None:
    # NOTE: The stdout/stderr Tee that used to live here has been removed.
    # It existed for the old pythonw (no-console) launch mode and, combined
    # with log() writing to the file directly, caused every line to be written
    # to the log file 2–3 times. The launcher now runs this script with
    # `python -u` and captures stdout directly, and log() writes each line to
    # the file exactly once, so no Tee is needed.
    _open_log()
    log(f"{'='*60}")
    log(f"Amazon Owned Books scraper started")
    log(f"START_PAGE={START_PAGE}  END_PAGE={END_PAGE}  ({('ALL PAGES' if END_PAGE == 0 else f'{END_PAGE - START_PAGE + 1} pages')})")
    log(f"BATCH_SIZE={BATCH_SIZE}  TEST_PAGE_LIMIT={TEST_PAGE_LIMIT}")
    log(f"Excel: {EXCEL_PATH}")
    log(f"{'='*60}")

    try:
        _main()
    except Exception as e:
        log(f"FATAL ERROR: {e}")
        log(traceback.format_exc())
        raise
    finally:
        log("Scraper finished.")
        _close_log()


def _navigate_to_batch_start(page, first_page: int) -> None:
    """
    In the existing browser session, navigate to the correct start page for
    this batch.  The session (and login cookies) are already alive — we just
    do a plain goto() to jump to the right page number.
    """
    target = library_url(first_page)
    log(f"  Navigating to page {first_page}: {target}")
    try:
        page.goto(target, wait_until="domcontentloaded")
    except (PlaywrightError, PlaywrightTimeoutError):
        pass  # SPA internal redirect — fine
    try:
        page.wait_for_selector("div.digital_entity_details", timeout=12_000)
        log(f"  Ready — cards visible on page {first_page}.")
    except PlaywrightTimeoutError:
        log(f"  WARNING: cards not visible on page {first_page} — proceeding anyway.")


def _scrape_batch(page, first_page: int, last_page: int) -> list[dict]:
    """
    Using the already-open browser page, navigate to first_page and scrape
    through to last_page.  The browser stays open between batches so the
    Amazon session (login cookies) is preserved.
    """
    _navigate_to_batch_start(page, first_page)
    log(f"  Scraping pages {first_page} to {last_page}...")
    return scrape_amazon_library(
        page,
        start_page=first_page,
        stop_after_page=last_page,
    )


def _load_df() -> pd.DataFrame:
    """Load the Excel sheet, ensure columns, fill missing DuplicateKeys."""
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df = ensure_columns(df)
    if "DuplicateKey" not in df.columns:
        df["DuplicateKey"] = ""
    missing_dk = df["DuplicateKey"].apply(clean) == ""
    df.loc[missing_dk, "DuplicateKey"] = df.loc[missing_dk].apply(
        lambda r: build_duplicate_key(clean(r.get("Title", "")), clean(r.get("Author", ""))),
        axis=1,
    )
    return df


def _main() -> None:
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Excel not found: {EXCEL_PATH}")

    batch_size   = max(1, BATCH_SIZE)
    current_page = START_PAGE
    end_page     = END_PAGE

    log(f"Loading: {EXCEL_PATH}")
    df = _load_df()
    log(f"Loaded {len(df)} existing rows.")

    grand_total      = 0
    grand_u_updated  = 0
    grand_u_appended = 0
    grand_r_updated  = 0
    grand_r_skipped  = 0
    batch_num        = 0

    # ── Single browser session for the entire run ─────────────────────────
    # Keeping the browser open between batches preserves the Amazon login
    # session — no re-authentication needed between batches.
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=HEADLESS)
        page    = browser.new_context().new_page()

        try:
            # Login once at the very start
            wait_for_login(page, start_page=START_PAGE)

            while True:
                batch_num  += 1
                batch_first = current_page
                batch_last  = current_page + batch_size - 1
                if end_page and batch_last > end_page:
                    batch_last = end_page

                log(f"")
                log(f"{'='*60}")
                log(f"BATCH {batch_num}  —  pages {batch_first} to {batch_last}")
                log(f"Current config: END_PAGE={end_page} ({'NO LIMIT - will scrape all pages' if end_page == 0 else f'will stop at page {end_page}'})")
                log(f"{'='*60}")

                # ── Scrape in existing session ────────────────────────────
                batch_books = _scrape_batch(page, batch_first, batch_last)
                grand_total += len(batch_books)

                unread = [b for b in batch_books if not b["is_read"]]
                read   = [b for b in batch_books if b["is_read"]]
                log(f"Batch {batch_num}: {len(batch_books)} books  |  {len(read)} READ  |  {len(unread)} unread")

                # ── Merge + enrich + save ─────────────────────────────────
                if batch_books:
                    log(f"Reloading Excel before merge...")
                    df = _load_df()
                    log(f"Merging batch {batch_num}...")
                    try:
                        df, u_up, u_ap, r_up, r_sk = merge_amazon_books(df, unread, read)
                        grand_u_updated  += u_up
                        grand_u_appended += u_ap
                        grand_r_updated  += r_up
                        grand_r_skipped  += r_sk
                        log(f"Batch {batch_num}: {u_up} updated, {u_ap} appended, "
                            f"{r_up} read-flagged, {r_sk} read-skipped")
                        log(f"Saving Excel file...")
                        save_excel(df)
                        log(f"Saved after batch {batch_num}: {EXCEL_PATH}")
                    except Exception as e:
                        log(f"ERROR during merge/save: {e}")
                        import traceback
                        log(traceback.format_exc())
                        raise
                else:
                    log(f"Batch {batch_num}: empty — end of library reached.")

                # ── Stop conditions ───────────────────────────────────────
                log(f"Checking stop conditions: end_page={end_page}, batch_last={batch_last}, batch_books={'empty' if not batch_books else f'{len(batch_books)} books'}")
                
                if end_page and batch_last >= end_page:
                    log(f"STOP: END_PAGE={end_page} reached (batch_last={batch_last}) — stopping.")
                    break
                if not batch_books:
                    log("STOP: Empty batch — end of library reached.")
                    break

                current_page = batch_last + 1
                log(f"✓ Batch {batch_num} complete. Continuing to next batch starting at page {current_page}...")
                log("")  # Blank line before next batch

        except KeyboardInterrupt:
            log("Keyboard interrupt — saving progress and stopping.")
        finally:
            try:
                browser.close()
            except Exception:
                pass

    log(f"")
    log(f"{'='*60}")
    log(f"ALL BATCHES COMPLETE")
    log(f"Total scraped  : {grand_total} books across {batch_num} batch(es)")
    log(f"Unread updated : {grand_u_updated}  |  appended+enriched: {grand_u_appended}")
    log(f"Read flagged   : {grand_r_updated}  |  not in sheet: {grand_r_skipped}")
    log(f"Final Excel    : {EXCEL_PATH}")
    log(f"{'='*60}")


if __name__ == "__main__":
    main()
