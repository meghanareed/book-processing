import re
import time
import unicodedata
import difflib
import datetime
import json
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
import sys

# Windows consoles and pipes default to cp1252, which cannot encode the ✓ ✗ →
# characters logged below; printing one raises UnicodeEncodeError mid-run.
# Force UTF-8 so this holds however the script is started.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _is_browser_closed(e: Exception) -> bool:
    """Return True if the exception means the browser/page was closed."""
    msg = str(e).lower()
    return any(x in msg for x in [
        "target closed", "target page", "browser has been closed",
        "page closed", "context closed", "connection closed",
    ])


# SCRIPT_DIR holds the code (and launcher_config.json); DATA_DIR holds the
# spreadsheets and logs.  They are the same folder in the old layout and
# separate once the code is cloned out of OneDrive.
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(r"C:\Users\megha\OneDrive\Documents\Reading")

EXCEL_PATH = DATA_DIR / "books_output.xlsx"
SHEET_NAME = "All Books"
BASE_URL = "https://app.thestorygraph.com"

# Load settings from launcher config if available.  book_launcher.py writes the
# config next to itself, so look there first; fall back to DATA_DIR for the old
# layout where code and data shared a folder.
_launcher_config = {}
_config_path = SCRIPT_DIR / "launcher_config.json"
if not _config_path.exists():
    _config_path = DATA_DIR / "launcher_config.json"
if _config_path.exists():
    try:
        _launcher_config = json.loads(_config_path.read_text(encoding="utf-8"))
        print(f"[CONFIG] Loaded settings from launcher_config.json")
    except Exception as e:
        print(f"[CONFIG] Could not load launcher config: {e}")

_SG_CFG = _launcher_config.get("storygraph", {})


def _cfg_days(key: str, default: int) -> int:
    """Read a non-negative day count from the launcher config."""
    try:
        return max(0, int(_SG_CFG.get(key, default)))
    except (TypeError, ValueError):
        return default


HEADLESS = False
TEST_LIMIT = _SG_CFG.get("max_books", 0)  # 0 = no limit
SLEEP_BETWEEN_BOOKS = 1.5
TITLE_MATCH_THRESHOLD = 0.60  # Lowered from 0.65 for better matching

# --- Reprocessing cooldowns (both in days, 0 disables) ---------------------
# SKIP_IF_PROCESSED_WITHIN_DAYS is the blanket guard: a book touched inside
# this window is left alone whatever happened last time, so running the tool
# twice in a month doesn't redo the same work.
# RETRY_FAILED_AFTER_DAYS is how long a *failed* attempt waits before it
# becomes eligible again.  Successes never come back regardless — they are
# excluded by their terminal status.
SKIP_IF_PROCESSED_WITHIN_DAYS = _cfg_days("skip_if_processed_within_days", 30)
RETRY_FAILED_AFTER_DAYS       = _cfg_days("retry_failed_after_days", 30)

# Saving rewrites every sheet and copies a backup, which on OneDrive costs far
# longer than processing a book.  Doing that per book dominated the run, so
# save in batches instead; progress is still flushed on exit, interrupt or
# error, so at most this many books' results are at risk.
SAVE_EVERY_N_BOOKS = max(1, _cfg_days("save_every_n_books", 25))

# A real, reused browser profile rather than a blank one per run.  StoryGraph's
# bot check hands out a clearance cookie; with a throwaway profile there is
# nowhere to keep it, so every run starts the challenge from scratch.  Kept
# beside the code, not in OneDrive — profiles are large and hold lock files
# that OneDrive fights over.
USER_DATA_DIR = SCRIPT_DIR / "browser_profile" / "storygraph"

# Prefer a real Chrome install; its fingerprint clears checks that the bundled
# Chromium build trips.  Falls back to Chromium when Chrome isn't present.
BROWSER_CHANNEL = _SG_CFG.get("browser_channel", "chrome")

# Logging
LOG_FILE = None

def _open_log() -> None:
    """Create timestamped log file in logs/ subfolder."""
    global LOG_FILE
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Create logs folder if it doesn't exist
    logs_dir = EXCEL_PATH.parent / "logs"
    logs_dir.mkdir(exist_ok=True)
    
    LOG_FILE = logs_dir / f"storygraph_{timestamp}.log"
    LOG_FILE.write_text("", encoding="utf-8")
    print(f"Logging to: {LOG_FILE}")
    log(f"StoryGraph To-Read Processor - Log started")

def log(msg: str) -> None:
    """Print to console AND write to log file."""
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    formatted = f"[{timestamp}] {msg}"
    print(formatted, flush=True)  # flush=True forces immediate output
    
    if LOG_FILE:
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(formatted + "\n")
                f.flush()  # Also flush file output
        except Exception:
            pass

STATUS_COL        = "StoryGraph Status"
MATCHED_QUERY_COL = "StoryGraph Matched Query"
NOTES_COL         = "StoryGraph Notes"
COMPLETED_COL     = "StoryGraph Completed"
SKIP_COL          = "Skip Storygraph"
OWNED_COL         = "Owned"                    # sourced from Amazon scraper
OWNED_SG_COL      = "StoryGraph Owned"         # tracks whether we marked owned on SG
ADDED_DATE_COL    = "StoryGraph Date"            # date this row was last queued for StoryGraph processing
SHORT_TITLE_COL   = "Short Title"                 # title stripped of subtitle (everything before first :)
READ_COL          = "Read"                       # set to "Yes" when SG shows book as already Read
DECISION_COL      = "Selector Decision"           # Read / Ignored / Removed, set by apply_reading_log.py


def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def clean_isbn(value) -> str:
    """Normalise an ISBN cell, dropping ones Excel has mangled into floats.

    Excel stores a 13-digit ISBN as a number, so it comes back as
    '9781390000000.0' — the '.0' is cosmetic but the zeros are real digits
    lost to float precision. Searching for that finds nothing and costs a
    page load, so a value whose digits have clearly been rounded away is
    dropped rather than queried.
    """
    text = clean_text(value)
    if not text:
        return ""
    if text.endswith(".0"):
        text = text[:-2]
    text = text.replace("-", "").replace(" ", "")
    if not re.fullmatch(r"\d{9}[\dXx]|\d{13}", text):
        return ""
    # 9781390000000 and friends: precision loss leaves a run of trailing zeros
    # no real ISBN has, since the last digit is a checksum.
    if re.search(r"0{5,}$", text):
        return ""
    return text


def ensure_storygraph_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in [STATUS_COL, MATCHED_QUERY_COL, NOTES_COL, COMPLETED_COL, SKIP_COL, OWNED_SG_COL, ADDED_DATE_COL, SHORT_TITLE_COL, DECISION_COL]:
        if col not in df.columns:
            df[col] = ""
    return df


STATUS_FAILED = "Failed"

# Notes written by earlier versions, which marked failures Skipped/Completed=Yes
# and so excluded them forever.  Matching on the note text lets those rows come
# back through the retry cooldown instead of staying stuck.
FAILURE_NOTE_MARKERS = (
    "to read button not found",
    "no asin/isbn/title available",
    "no matching storygraph result",
)


def is_failed_attempt(row: pd.Series) -> bool:
    """True if the last attempt failed, rather than succeeding or finding the
    book already on the shelf.  Failures are retried once their cooldown is up."""
    if clean_text(row.get(STATUS_COL)) in {STATUS_FAILED, "Not Found", "Error"}:
        return True
    notes = clean_text(row.get(NOTES_COL)).lower()
    return any(marker in notes for marker in FAILURE_NOTE_MARKERS)


def is_terminal_status(row: pd.Series) -> bool:
    """True if this book is done for good and should never be reprocessed.

    A failed attempt is never terminal — however it was recorded — so that the
    retry cooldown decides when it comes back.
    """
    if is_failed_attempt(row):
        return False
    status    = clean_text(row.get(STATUS_COL))
    completed = clean_text(row.get(COMPLETED_COL)).lower()
    return completed == "yes" or status in {"Added", "Skipped"}


def days_since(value) -> float | None:
    """Days since a date cell.  None when blank or unparseable (never processed)."""
    text = clean_text(value)
    if not text:
        return None
    try:
        parsed = pd.to_datetime(text)
    except Exception:
        return None
    if pd.isna(parsed):
        return None
    return (datetime.datetime.now() - parsed.to_pydatetime()).total_seconds() / 86400.0


def apply_cooldowns(df: pd.DataFrame) -> pd.DataFrame:
    """Drop books that were processed too recently to be worth redoing."""
    if ADDED_DATE_COL not in df.columns or df.empty:
        return df

    ages   = df[ADDED_DATE_COL].apply(days_since)
    failed = df.apply(is_failed_attempt, axis=1)

    if SKIP_IF_PROCESSED_WITHIN_DAYS:
        # The blanket guard is about not redoing work that succeeded.  A failed
        # attempt is not work done, so it answers to the retry cooldown below
        # instead — otherwise a crashed run would be locked out for a month.
        recent = ~failed & ages.apply(
            lambda a: a is not None and a < SKIP_IF_PROCESSED_WITHIN_DAYS
        )
        if recent.any():
            log(f"  [COOLDOWN] Skipping {int(recent.sum())} book(s) processed in the "
                f"last {SKIP_IF_PROCESSED_WITHIN_DAYS} day(s)")
            df   = df[~recent].copy()
            ages = ages[~recent]

    if RETRY_FAILED_AFTER_DAYS and not df.empty:
        too_soon = df.apply(is_failed_attempt, axis=1) & ages.apply(
            lambda a: a is not None and a < RETRY_FAILED_AFTER_DAYS
        )
        if too_soon.any():
            log(f"  [COOLDOWN] {int(too_soon.sum())} failed book(s) not due for retry "
                f"yet (retry after {RETRY_FAILED_AFTER_DAYS} days)")
            df = df[~too_soon].copy()

    return df


DECISION_REMOVED = "Removed"


def needs_removal(row: pd.Series) -> bool:
    """True if the selector said Removed — take it off the To-Read pile."""
    return clean_text(row.get(DECISION_COL)).lower() == DECISION_REMOVED.lower()


def should_skip_storygraph(row: pd.Series) -> bool:
    if clean_text(row.get(SKIP_COL)).lower() == "yes":
        return True
    decision = clean_text(row.get(DECISION_COL))
    if not decision:
        return False
    # Read and Ignored just mean "don't offer this again", so there is nothing
    # to do on StoryGraph.  Removed is a job: go and take it off the pile.
    return not needs_removal(row)


def is_owned(row: pd.Series) -> bool:
    """Return True if this book came from Amazon (Owned = Yes)."""
    return clean_text(row.get(OWNED_COL)).lower() == "yes"


def short_title(title: str) -> str:
    """Return everything before the first ' : ' or ': ' in a title."""
    t = clean_text(title)
    for sep in [": ", " : "]:
        if sep in t:
            return t[:t.index(sep)].strip()
    return t  # no subtitle — return as-is


def build_search_queries(row: pd.Series) -> list[str]:
    asin   = clean_text(row.get("ASIN"))
    isbn13 = clean_isbn(row.get("ISBN_13"))
    isbn10 = clean_isbn(row.get("ISBN_10"))
    title  = clean_text(row.get("Title"))
    author = clean_text(row.get("Author"))
    st     = clean_text(row.get(SHORT_TITLE_COL)) or short_title(title)

    queries = []
    # Priority 1: IDs (most specific)
    if asin:   queries.append(asin)
    if isbn13: queries.append(isbn13)
    if isbn10: queries.append(isbn10)
    
    # Priority 2: Title + Author
    if title and author:
        queries.append(f"{title} {author}")
    elif title:
        queries.append(title)
    
    # Priority 3: Short title + author (helps when full title fails to match)
    if st and st.lower() != title.lower():
        if author:
            queries.append(f"{st} {author}")
        else:
            queries.append(st)
    
    # Priority 4: Title-only fallback (search by title, verify author in HTML)
    # This catches cases where author name is slightly different in StoryGraph
    if title and author:  # Only add if we have both (otherwise already added above)
        queries.append(title)
    
    # Deduplicate while preserving order
    out, seen = [], set()
    for q in queries:
        key = q.lower()
        if key not in seen:
            seen.add(key)
            out.append(q)
    return out


def launch_browser(p):
    """Open a persistent browser profile and return the context.

    Persisting the profile is what makes the security check survive between
    runs: the clearance cookie and your StoryGraph session are stored on disk
    instead of being thrown away when the browser closes.  The first run still
    needs you to clear the check and log in by hand; later runs should come up
    already verified and signed in.
    """
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

    options = dict(
        user_data_dir=str(USER_DATA_DIR),
        headless=HEADLESS,
        viewport={"width": 1400, "height": 900},
        locale="en-US",
        # Playwright otherwise launches with --enable-automation, which the
        # check reads as "this is not a person".
        args=["--disable-blink-features=AutomationControlled"],
        ignore_default_args=["--enable-automation"],
    )

    if BROWSER_CHANNEL:
        try:
            context = p.chromium.launch_persistent_context(channel=BROWSER_CHANNEL, **options)
            log(f"  [BROWSER] Using {BROWSER_CHANNEL} with profile: {USER_DATA_DIR}")
            return context
        except Exception as e:
            log(f"  [BROWSER] {BROWSER_CHANNEL} unavailable ({e}); falling back to Chromium")

    context = p.chromium.launch_persistent_context(**options)
    log(f"  [BROWSER] Using bundled Chromium with profile: {USER_DATA_DIR}")
    return context


def wait_for_user_login(page):
    """
    Open StoryGraph and wait for user to log in.
    Shows a GUI dialog instead of console input so it works when
    launched from the GUI launcher.
    """
    import tkinter as tk
    from tkinter import messagebox
    
    page.goto(BASE_URL, wait_until="domcontentloaded")
    
    # Create a hidden root window for the messagebox
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    root.attributes('-topmost', True)  # Make dialog appear on top
    
    messagebox.showinfo(
        "StoryGraph Login",
        "In the browser window that just opened:\n\n"
        "1. If you see \"Performing security verification\", wait for it to "
        "finish. Don't refresh — refreshing restarts the check.\n"
        "2. Log in if you aren't already.\n\n"
        "Click OK once you can see your library.\n\n"
        "The browser profile is saved, so later runs should come up already "
        "verified and signed in.",
        parent=root
    )
    
    root.destroy()
    log("User confirmed login complete, proceeding...")


def check_session_alive(page) -> bool:
    """
    Check if the browser page is still alive and accessible.
    Returns False if page is closed or unresponsive.
    """
    try:
        # Try to get the current URL - this will fail if page is closed
        _ = page.url
        return True
    except Exception as e:
        log(f"    [SESSION] Browser session is dead: {e}")
        return False


def open_search(page):
    """Navigate to a StoryGraph search page."""
    if not check_session_alive(page):
        log(f"    [SEARCH] Cannot navigate - browser is closed")
        return False
    
    for url in [f"{BASE_URL}/browse", f"{BASE_URL}/search", BASE_URL]:
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=15000)
            page.wait_for_timeout(1500)
            return True
        except Exception as e:
            log(f"    [SEARCH] Navigation to {url} failed: {e}")
            pass
    
    log(f"    [SEARCH] ERROR: Could not navigate to any search page")
    return False


def submit_search(page, query: str) -> bool:
    """Submit a search query on StoryGraph."""
    log(f"    [SEARCH] Searching for: {query!r}")
    
    # Check if browser is still alive
    if not check_session_alive(page):
        log(f"    [SEARCH] ERROR: Browser session lost - cannot search")
        log(f"    [SEARCH] You may have closed the browser or been logged out")
        return False
    
    # Navigate to search page
    if not open_search(page):
        return False
    
    # Try to find and fill search box
    selectors = [
        'input[type="search"]',
        'input[placeholder*="Search" i]',
        'input[aria-label*="Search" i]',
        'input[name*="search" i]',
        'input[id*="search" i]',
    ]
    for selector in selectors:
        try:
            box = page.locator(selector).first
            if box.is_visible(timeout=3000):
                box.click()
                box.fill(query)
                box.press("Enter")
                page.wait_for_load_state("domcontentloaded")
                page.wait_for_timeout(2000)
                log(f"    [SEARCH] Search submitted, landed on: {page.url}")
                return True
        except Exception as e:
            log(f"    [SEARCH] Selector '{selector}' failed: {e}")
            pass
    
    log(f"    [SEARCH] ERROR: Could not find search box")


# Letters that carry no combining mark, so NFKD alone will not fold them.
_FOLD_EXTRA = str.maketrans({
    "ø": "o", "đ": "d", "ð": "d", "ł": "l", "ħ": "h", "ı": "i",
    "ß": "ss", "æ": "ae", "œ": "oe", "þ": "th",
})


def fold(text) -> str:
    """Lowercase and strip accents for comparison.

    Comparisons were already lowercased; the miss was diacritics.  StoryGraph
    shows 'Blanka Lipińska' while the spreadsheet holds 'Blanka Lipinska', and
    'lipinska' is not a substring of 'lipińska' — so the author check rejected
    a page whose every author link was the right person.
    """
    lowered = clean_text(text).lower().translate(_FOLD_EXTRA)
    return "".join(
        c for c in unicodedata.normalize("NFKD", lowered)
        if not unicodedata.combining(c)
    )


# Reading text from a locator that matches nothing makes Playwright wait out
# its 30s default.  The sibling probes below try two such selectors per
# candidate, which is where the 60-second stalls in the log came from.
ELEMENT_TEXT_TIMEOUT_MS = 2000


def safe_inner_text(locator, timeout_ms: int = ELEMENT_TEXT_TIMEOUT_MS) -> str:
    """inner_text() that gives up quickly instead of blocking the whole run."""
    try:
        if locator.count() == 0:
            return ""
        return clean_text(locator.inner_text(timeout=timeout_ms))
    except Exception:
        return ""


BOOK_HREF_RE = re.compile(r"^/books/[0-9a-f-]{36}/?$", re.IGNORECASE)

# Anchor text that is page furniture rather than a book title.  These were
# being scored against the title and could outrank the real result: on an ASIN
# search for Desperate Measures, '29 editions' scored 0.375 and the actual
# title scored 0.154.
NON_BOOK_LINK_TEXT = {
    "add a book", "manually add a book to our database",
    "mark as owned", "add to \"up next\"", "to read", "29 editions",
}


def is_book_result_link(text: str, href: str) -> bool:
    """True only for a link that leads to an actual book page.

    Rejects /books/new, /books/<id>/editions, mark-as-owned helpers and
    empty anchors — all of which appear on a results page and none of which
    are the book.
    """
    href = clean_text(href)
    if not BOOK_HREF_RE.match(href.split("?")[0]):
        return False
    text = clean_text(text)
    if not text:
        return False
    low = text.lower()
    if low in NON_BOOK_LINK_TEXT or low.endswith(" editions"):
        return False
    return True


def is_identifier_query(query: str) -> bool:
    """True for an ASIN or ISBN — a search term that identifies one book.

    A hit from one of these needs no title check: the identifier already
    picked the book out. Requiring a fuzzy title match on top is what made
    an ASIN search for 'Wicked Villain series' reject the correct result,
    because the spreadsheet holds the series name, not the book's title.
    """
    q = clean_text(query).replace("-", "").replace(" ", "")
    if q.endswith(".0"):
        q = q[:-2]
    return bool(
        re.fullmatch(r"B[0-9A-Z]{9}", q, re.IGNORECASE)   # ASIN
        or re.fullmatch(r"\d{9}[\dXx]", q)                # ISBN-10
        or re.fullmatch(r"\d{13}", q)                     # ISBN-13
    )


def _title_score(candidate: str, target: str) -> float:
    """
    Fuzzy similarity between candidate and target title.
    Strips subtitles (after ': ' or ' (') before comparing so
    'God of Ruin' matches 'God of Ruin: A Dark College Romance...'.
    """
    def core(t: str) -> str:
        t = fold(t)   # accent-insensitive, same as the author check
        for sep in [": ", " - ", " ("]:
            if sep in t:
                t = t[:t.index(sep)]
        return t.strip()
    a, b = core(candidate), core(target)
    if not a or not b:
        return 0.0
    # Also give full credit if one is a substring of the other
    if a in b or b in a:
        return 1.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def pick_first_matching_result(page, title: str, author: str,
                               trust_identifier: bool = False) -> bool:
    title  = clean_text(title)
    author = clean_text(author)
    
    # For multi-author books, extract all last names
    author_last_names = []
    if author:
        # Split by comma or "and" to get individual authors
        authors = [a.strip() for a in author.replace(" and ", ", ").split(",")]
        for auth in authors:
            if auth:
                last_name = fold(auth.split()[-1])
                author_last_names.append(last_name)
    
    log(f"    [PICK] Looking for: title={title!r}, author={author!r}")
    log(f"    [PICK] Author last names to match: {author_last_names}")
    log(f"    [PICK] Current URL: {page.url}")
    
    # Try multiple selector strategies
    selectors = [
        'a[href*="/books/"]',  # Broadest - any book link
        'main a[href*="/books/"]',  # Book links in main content
        '.book-pane a',  # Book pane links (if structure exists)
        '[class*="book"] a',  # Any element with "book" in class
        f'a:has-text("{title[:30]}")' if len(title) >= 5 else None,  # Partial title match
    ]

    best_score   = 0.0
    best_locator = None
    best_text    = ""
    best_href    = ""
    
    all_candidates = []  # For debugging

    for selector in selectors:
        if not selector:
            continue
        try:
            links = page.locator(selector)
            count = links.count()
            log(f"    [PICK] Selector '{selector}': {count} candidates found")
            
            # Sample up to 20 results to find best match
            for i in range(min(count, 20)):
                link = links.nth(i)
                try:
                    text = safe_inner_text(link)
                    href = clean_text(link.get_attribute("href"))
                except Exception as e:
                    log(f"    [PICK] [{i}] Error getting text/href: {e}")
                    continue

                # Skip page furniture — "Add a book", "29 editions",
                # "mark as owned", empty anchors.  Scoring these against the
                # title let them outrank the real result.
                if not is_book_result_link(text, href):
                    continue
                
                if not href or ("/books/" not in href and "book" not in href.lower()):
                    continue

                score = _title_score(text, title)

                # Author check - look for nearby author links
                author_ok = True
                if author_last_names and score >= (TITLE_MATCH_THRESHOLD - 0.1):
                    author_found = False
                    try:
                        # Strategy 1: Check parent container text
                        parent_text = safe_inner_text(link.locator(".."))
                        log(f"    [PICK] [{i}] Parent text: {parent_text[:100]!r}")
                        # Check if ANY author last name appears
                        for author_last in author_last_names:
                            if author_last in fold(parent_text):
                                author_found = True
                                log(f"    [PICK] [{i}] ✓ Author '{author_last}' found in parent")
                                break
                    except Exception as e:
                        log(f"    [PICK] [{i}] Parent check failed: {e}")
                        parent_text = ""
                    
                    # Strategy 2: Look for author links near this book link
                    if not author_found:
                        try:
                            # Find all author links on the page
                            author_links = page.locator('a[href*="/authors/"]')
                            author_count = author_links.count()
                            log(f"    [PICK] [{i}] Found {author_count} author links on page")
                            
                            # Check author links near this book (within reasonable distance)
                            authors_checked = []
                            for j in range(min(author_count, 50)):
                                try:
                                    author_link = author_links.nth(j)
                                    author_text = safe_inner_text(author_link)
                                    authors_checked.append(author_text)
                                    # Check if ANY author last name matches
                                    for author_last in author_last_names:
                                        if author_last in fold(author_text):
                                            author_found = True
                                            log(f"    [PICK] [{i}] ✓ Author '{author_last}' found in link: '{author_text}'")
                                            break
                                    if author_found:
                                        break
                                except Exception:
                                    continue
                            
                            if not author_found and authors_checked:
                                log(f"    [PICK] [{i}] Authors on page: {', '.join(authors_checked[:10])}")
                        except Exception as e:
                            log(f"    [PICK] [{i}] Author link check failed: {e}")
                    
                    # Strategy 3: Check siblings
                    if not author_found:
                        try:
                            siblings_text = ""
                            for sib_selector in ["+ *", "~ *"]:
                                try:
                                    sibling = link.locator(sib_selector).first
                                    siblings_text += " " + safe_inner_text(sibling)
                                except Exception:
                                    pass
                            if siblings_text:
                                log(f"    [PICK] [{i}] Sibling text: {siblings_text[:100]!r}")
                            # Check if ANY author last name appears
                            for author_last in author_last_names:
                                if author_last in fold(siblings_text):
                                    author_found = True
                                    log(f"    [PICK] [{i}] ✓ Author '{author_last}' found in siblings")
                                    break
                        except Exception as e:
                            log(f"    [PICK] [{i}] Sibling check failed: {e}")
                    
                    author_ok = author_found
                    if not author_found:
                        log(f"    [PICK] [{i}] ✗ None of {author_last_names} found anywhere")

                log(f"    [PICK] [{i}] score={score:.3f} author_ok={author_ok} | {text[:80]!r} | {href}")
                all_candidates.append((score, author_ok, text, href))

                if score >= TITLE_MATCH_THRESHOLD and author_ok:
                    if score > best_score:
                        best_score   = score
                        best_locator = link
                        best_text    = text
                        best_href    = href

        except Exception as e:
            log(f"    [PICK] Selector '{selector}' error: {e}")
            continue

    # Log summary of what we found
    if all_candidates:
        log(f"    [PICK] Found {len(all_candidates)} total candidates")
        # Show top 5
        sorted_cands = sorted(all_candidates, key=lambda x: x[0], reverse=True)[:5]
        log(f"    [PICK] Top candidates:")
        for i, (score, author_ok, text, href) in enumerate(sorted_cands, 1):
            log(f"    [PICK]   {i}. score={score:.3f} author_ok={author_ok} | {text[:60]!r}")
    
    # An ASIN/ISBN search that returns exactly one book has already identified
    # it — the title check adds nothing and can only reject a correct result,
    # e.g. when the spreadsheet holds a series name rather than the title.
    if best_locator is None and trust_identifier:
        distinct = {href for _s, _a, _t, href in all_candidates}
        if len(distinct) == 1:
            only_href = next(iter(distinct))
            match = next(c for c in all_candidates if c[3] == only_href)
            log(f"    [PICK] ✓ Single result for an ID search — taking it "
                f"without a title check: {match[2][:80]!r}")
            try:
                page.goto(BASE_URL.rstrip("/") + only_href,
                          wait_until="domcontentloaded", timeout=20000)
                page.wait_for_timeout(2000)
                log(f"    [PICK] ✓ Landed on: {page.url}")
                return True
            except Exception as e:
                log(f"    [PICK] ERROR: Could not open {only_href}: {e}")
                return False

    if best_locator is not None:
        log(f"    [PICK] ✓ Best match (score={best_score:.3f}): {best_text[:80]!r}")
        log(f"    [PICK] Clicking: {best_href}")
        try:
            best_locator.click()
        except Exception as e:
            log(f"    [PICK] ERROR: Click failed: {e}")
            return False
        try:
            page.wait_for_load_state("domcontentloaded", timeout=10000)
        except Exception:
            pass
        page.wait_for_timeout(2000)
        log(f"    [PICK] ✓ Landed on: {page.url}")
        return True

    log(f"    [PICK] ✗ No match >= {TITLE_MATCH_THRESHOLD} found")
    log(f"    [PICK] Threshold: {TITLE_MATCH_THRESHOLD}, Best score found: {best_score:.3f}")
    return False


def book_already_read(page) -> bool:
    """
    Check if the book is already marked as Read on StoryGraph.
    StoryGraph shows a selected "read" status button when the book
    is in the user's Read pile.
    """
    read_selectors = [
        'button.read-status-label[title="Book marked as read"]',
        'button.read-status-label[aria-label="Book marked as read"]',
        'button.read-status-label:has-text("read"):not(:has-text("to read"))',
        '[data-status="read"]',
        # Broader fallback — look for a selected/active read button
        'button[class*="read-status"][class*="selected"]:has-text("read")',
        'button[aria-pressed="true"]:has-text("read")',
    ]
    for selector in read_selectors:
        try:
            locator = page.locator(selector).first
            if locator.count() > 0 and locator.is_visible(timeout=1500):
                log(f"    [READ] Book already marked as Read (found: {selector})")
                return True
        except Exception:
            pass
    return False


def book_already_to_read(page) -> bool:
    """
    Check if book is already marked as to-read.
    IMPORTANT: Only match the SELECTED state button, not the unselected one!
    """
    selected_selectors = [
        'button.read-status-label[title="Book marked as to read"]',
        'button.read-status-label[aria-label="Book marked as to read"]',
        'button.read-status-label:has-text("to read")',
        '[data-status="to-read"]',  # If they use data attributes
        # NOTE: Do NOT use 'button:has-text("to read")' - it matches the unselected button too!
    ]
    for selector in selected_selectors:
        try:
            locator = page.locator(selector).first
            if locator.count() > 0 and locator.is_visible(timeout=2000):
                log(f"    [TO-READ] Already marked (found: {selector})")
                return True
        except Exception:
            pass
    
    log(f"    [TO-READ] Not marked yet (will click button)")
    return False


def click_to_read(page) -> bool:
    unselected_selectors = [
        'button.read-status-button[title="Add to your To-Read Pile"]',
        'button.read-status-button[aria-label="Add to your To-Read Pile"]',
        'button.read-status-button:has-text("to read")',
        'button:has-text("to read"):not(.read-status-label)',  # Button with "to read" but not already selected
        '[aria-label*="To-Read" i]',  # Broader aria-label match
    ]
    for selector in unselected_selectors:
        try:
            locator = page.locator(selector).first
            if locator.count() > 0 and locator.is_visible(timeout=2000):
                log(f"    [TO-READ] Clicking button (found: {selector})")
                locator.click()
                page.wait_for_timeout(1500)
                return True
        except Exception as e:
            log(f"    [TO-READ] Selector '{selector}' failed: {e}")
            pass
    log(f"    [TO-READ] ERROR: Could not find To-Read button")
    return False


def remove_from_to_read(page) -> bool:
    """Take the book off the To-Read pile.

    StoryGraph exposes this as a delete link, much like un-marking owned, and
    sometimes only behind the chevron next to the status pill.  Several shapes
    are tried because the markup has changed before; the log names whichever
    one worked so it can be narrowed later.
    """
    remove_selectors = [
        'a[href*="/remove-from-to-read"]',
        'a[href*="/remove-to-read"]',
        'a[href*="remove"][href*="to-read"]',
        'a.remove-from-to-read-link',
        'a[data-method="delete"]:has-text("to read")',
        'button[title="Remove from your To-Read Pile"]',
        'button[aria-label="Remove from your To-Read Pile"]',
    ]

    def try_selectors(selectors, label) -> bool:
        for selector in selectors:
            try:
                locator = page.locator(selector).first
                if locator.count() > 0 and locator.is_visible(timeout=1500):
                    log(f"    [REMOVE] Clicking {label} (found: {selector})")
                    locator.click()
                    page.wait_for_timeout(2000)
                    return True
            except Exception:
                continue
        return False

    if try_selectors(remove_selectors, "remove link"):
        return not book_already_to_read(page)

    # Not on the page directly — open the status dropdown and look again.
    dropdown_selectors = [
        'button.read-status-label + button',            # chevron beside the pill
        'button[aria-haspopup="true"]:near(.read-status-label)',
        '.read-status-dropdown button',
        'button[aria-expanded="false"]:has-text("to read")',
    ]
    if try_selectors(dropdown_selectors, "status dropdown"):
        menu_selectors = remove_selectors + [
            '*:has-text("remove from to-read")',
            '*:has-text("remove from shelf")',
            'a:has-text("remove")',
            'button:has-text("remove")',
        ]
        if try_selectors(menu_selectors, "remove item"):
            return not book_already_to_read(page)

    log("    [REMOVE] ERROR: Could not find a control to remove it from To-Read")
    return False


def book_already_owned(page) -> bool:
    """
    Return True if the book is already marked as owned on StoryGraph.
    """
    owned_indicators = [
        'a.remove-from-owned-link',  # Actual class when owned
        'a[href*="/remove-owned-book"]',  # Actual href when owned
        'a[data-method="delete"]:has-text("owned")',  # Delete method + "owned" text
        'a:has-text("owned"):not(:has-text("mark as owned"))',  # Just "owned", not "mark as owned"
        # Legacy/fallback selectors (in case StoryGraph changes):
        'a[title="Remove from owned"]',
        'a[aria-label="Remove from owned"]',
        'a.remove-owned-link',
        '*:has-text("owned by you")',
        '.owned-badge',
        '[data-owned="true"]',
    ]
    for selector in owned_indicators:
        try:
            locator = page.locator(selector).first
            if locator.count() > 0 and locator.is_visible(timeout=1500):
                log(f"    [OWNED] Already owned (found: {selector})")
                return True
        except Exception:
            pass
    
    log(f"    [OWNED] Not owned yet (will click mark as owned)")
    return False


def click_mark_as_owned(page) -> bool:
    """
    Click the "Mark as owned" link on the current StoryGraph book page.
    """
    owned_selectors = [
        'a.mark-as-owned-link[title="Mark as owned"]',
        'a.mark-as-owned-link[aria-label="Mark as owned"]',
        'a.mark-as-owned-link',
        'a[href*="mark-as-owned"]',
        'a:has-text("mark as owned")',
        'button:has-text("mark as owned")',  # In case they changed to button
    ]
    for selector in owned_selectors:
        try:
            locator = page.locator(selector).first
            if locator.count() > 0 and locator.is_visible(timeout=2000):
                log(f"    [OWNED] Clicking Mark as Owned (found: {selector})")
                locator.click()
                page.wait_for_timeout(1500)
                return True
        except Exception as e:
            log(f"    [OWNED] Selector '{selector}' failed: {e}")
            pass
    log(f"    [OWNED] ERROR: Could not find Mark as Owned button")
    return False


def process_book(page, row: pd.Series) -> dict:
    title  = clean_text(row.get("Title"))
    author = clean_text(row.get("Author"))
    owned  = is_owned(row)
    queries = build_search_queries(row)

    result = {
        "Title":          title,
        "Author":         author,
        "DuplicateKey":   clean_text(row.get("DuplicateKey")),
        "ASIN":           clean_text(row.get("ASIN")),
        "ISBN_13":        clean_text(row.get("ISBN_13")),
        "ISBN_10":        clean_text(row.get("ISBN_10")),
        STATUS_COL:       "",
        MATCHED_QUERY_COL: "",
        NOTES_COL:        "",
        COMPLETED_COL:    "No",
        OWNED_SG_COL:     "",
        READ_COL:         "",   # will be set to "Yes" if SG shows book as already Read
    }

    if not queries:
        result[STATUS_COL]    = STATUS_FAILED
        result[NOTES_COL]     = "No ASIN/ISBN/title available"
        result[COMPLETED_COL] = "No"   # retry once metadata fills in
        return result

    log(f"    [PROCESS] Trying {len(queries)} queries: {queries}")
    
    for query_num, query in enumerate(queries, 1):
        log(f"    [PROCESS] Query {query_num}/{len(queries)}: {query!r}")
        
        if not submit_search(page, query):
            log(f"    [PROCESS] Search submission failed, trying next query")
            continue

        picked = pick_first_matching_result(
            page, title, author, trust_identifier=is_identifier_query(query)
        )

        # If we didn't pick a result AND we're not already on a book page, try next query
        if not picked:
            if "/books/" in page.url:
                log(f"    [PROCESS] Not picked but already on book page: {page.url}")
            else:
                log(f"    [PROCESS] No result matched, trying next query")
                continue

        result[MATCHED_QUERY_COL] = query
        log(f"    [PROCESS] Processing book page: {page.url}")

        # ── Removed in the selector: take it off the To-Read pile ─────────
        if needs_removal(row):
            if book_already_to_read(page):
                if remove_from_to_read(page):
                    result[STATUS_COL]    = "Removed"
                    result[NOTES_COL]     = "Removed from To-Read pile"
                    result[COMPLETED_COL] = "Yes"
                    log("    [PROCESS] Removed from To-Read")
                else:
                    result[STATUS_COL]    = STATUS_FAILED
                    result[NOTES_COL]     = "On To-Read but remove control not found"
                    result[COMPLETED_COL] = "No"
            else:
                result[STATUS_COL]    = "Removed"
                result[NOTES_COL]     = "Marked Removed; was not on To-Read pile"
                result[COMPLETED_COL] = "Yes"
                log("    [PROCESS] Not on To-Read — nothing to remove")
            return result

        # ── Already Read on StoryGraph? ───────────────────────────────────
        # Check this BEFORE the To-Read logic. If SG shows the book as Read,
        # we skip adding it to To-Read (makes no sense) and flag it in Excel.
        if book_already_read(page):
            result[READ_COL]      = "Yes"
            result[STATUS_COL]    = "Skipped"
            result[NOTES_COL]     = "Already marked Read on StoryGraph — flagged in Excel"
            result[COMPLETED_COL] = "Yes"
            log(f"    [PROCESS] Book is Read on SG — will set Read=Yes in Excel")

        # ── To-Read ──────────────────────────────────────────────────────
        elif book_already_to_read(page):
            result[STATUS_COL]    = "Skipped"
            result[NOTES_COL]     = "Already marked To Read"
            result[COMPLETED_COL] = "Yes"
        elif click_to_read(page):
            result[STATUS_COL]    = "Added"
            result[NOTES_COL]     = "Clicked To Read"
            result[COMPLETED_COL] = "Yes"
        else:
            result[STATUS_COL]    = STATUS_FAILED
            result[NOTES_COL]     = "Book page found, but To Read button not found"
            result[COMPLETED_COL] = "No"   # retryable: often a slow load or markup change

        # ── Owned (only attempted when Owned = Yes in the spreadsheet) ───
        if owned:
            if book_already_owned(page):
                result[OWNED_SG_COL] = "Already Owned"
                result[NOTES_COL] += " | Already marked Owned on SG"
            elif click_mark_as_owned(page):
                result[OWNED_SG_COL] = "Marked Owned"
                result[NOTES_COL]   += " | Clicked Mark as Owned"
            else:
                result[OWNED_SG_COL] = "Owned button not found"
                result[NOTES_COL]   += " | Owned button not found on SG"

        return result

    result[STATUS_COL]    = "Not Found"
    result[NOTES_COL]     = "No matching StoryGraph result after trying all queries"
    result[COMPLETED_COL] = "No"
    log(f"    [PROCESS] ✗ Not found after {len(queries)} queries")
    return result


_OTHER_SHEETS_CACHE = None


def save_excel(df_all: pd.DataFrame):
    """
    Write All Books sheet back while preserving all other sheets.
    Uses atomic write (temp file → rename) so the original is never
    left as 0kb if the write crashes or is interrupted.
    Also keeps a .bak of the last good save.
    """
    global _OTHER_SHEETS_CACHE
    import shutil, tempfile, os

    # The other sheets (Unique Books, Duplicate Summary, Needs Review) are not
    # touched by this script, so read them once instead of re-reading the whole
    # workbook on every save.
    if _OTHER_SHEETS_CACHE is None:
        try:
            sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)
        except Exception:
            sheets = {}
        _OTHER_SHEETS_CACHE = {k: v for k, v in sheets.items() if k != SHEET_NAME}

    all_sheets: dict = dict(_OTHER_SHEETS_CACHE)
    all_sheets[SHEET_NAME] = df_all

    # Write to a temp file in the same directory
    tmp_path = EXCEL_PATH.with_suffix(".tmp.xlsx")
    try:
        with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
            for name, sheet_df in all_sheets.items():
                sheet_df.to_excel(writer, sheet_name=name, index=False)
    except Exception as e:
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
        pass

    try:
        tmp_path.replace(EXCEL_PATH)
    except Exception as e:
        raise RuntimeError(f"Could not replace Excel file: {e}") from e


def main():
    _open_log()
    
    log("="*60)
    log(f"StoryGraph To-Read Processor")
    log(f"TEST_LIMIT={TEST_LIMIT} ({'no limit' if TEST_LIMIT == 0 else f'{TEST_LIMIT} books'})")
    log(f"TITLE_MATCH_THRESHOLD={TITLE_MATCH_THRESHOLD}")
    log(f"Excel: {EXCEL_PATH}")
    log("="*60)
    
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_PATH}")

    df_all = pd.read_excel(EXCEL_PATH, sheet_name="All Books")
    df_all = ensure_storygraph_columns(df_all)

    required_cols = {"Title", "Author"}
    missing = required_cols - set(df_all.columns)
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Fill Short Title for any rows that are missing it
    if SHORT_TITLE_COL not in df_all.columns:
        df_all[SHORT_TITLE_COL] = ""
    blank_st = df_all[SHORT_TITLE_COL].apply(clean_text) == ""
    df_all.loc[blank_st, SHORT_TITLE_COL] = df_all.loc[blank_st, "Title"].apply(short_title)

    df = df_all.copy()
    if "DuplicateKey" in df.columns:
        df = df.drop_duplicates(subset=["DuplicateKey"], keep="first").copy()
    else:
        df = df.drop_duplicates(subset=["Title", "Author"], keep="first").copy()

    df = df[~df.apply(should_skip_storygraph, axis=1)].copy()
    df = df[~df.apply(is_terminal_status, axis=1)].copy()

    # Only process books that came from Amazon (Owned = Yes)
    if OWNED_COL in df.columns:
        owned = df[OWNED_COL].apply(clean_text).str.lower() == "yes"
        # A removal is about your To-Read pile, not about owning the book, so
        # Removed rows come through whether or not Owned is set.
        removals = df.apply(needs_removal, axis=1)
        df = df[owned | removals].copy()
        if removals.any():
            log(f"  [REMOVE] {int(removals.sum())} book(s) marked Removed to take "
                f"off the To-Read pile")

    # Cooldowns: skip anything processed too recently, and hold failed
    # attempts until their retry interval is up.
    df = apply_cooldowns(df)

    if TEST_LIMIT:
        # Random sample rather than the first N rows: a plain head() always
        # works the same end of the sheet, so a small limit would keep
        # revisiting the oldest entries.  sort_index() puts the chosen books
        # back into sheet order so the run log reads top to bottom.
        if TEST_LIMIT < len(df):
            log(f"  [LIMIT] Random sample of {TEST_LIMIT} from {len(df)} eligible book(s)")
            df = df.sample(n=TEST_LIMIT).sort_index().copy()
        else:
            df = df.copy()

    if df.empty:
        log("No rows need StoryGraph processing.")
        return

    results = []
    stopped_early = False

    with sync_playwright() as p:
        context = launch_browser(p)
        page    = context.pages[0] if context.pages else context.new_page()

        wait_for_user_login(page)

        total = len(df)
        for idx, (_, row) in enumerate(df.iterrows(), start=1):
            title         = clean_text(row.get("Title"))
            author        = clean_text(row.get("Author"))
            duplicate_key = clean_text(row.get("DuplicateKey"))
            owned_flag    = "owned=YES" if is_owned(row) else "owned=no"
            log(f"\n{'='*60}")
            log(f"[{idx}/{total}] {title} | {author}  ({owned_flag})")
            log(f"{'='*60}")

            try:
                outcome = process_book(page, row)
            except KeyboardInterrupt:
                log("\n[STOPPED] Keyboard interrupt — saving progress and exiting.")
                stopped_early = True
                break
            except PlaywrightTimeoutError:
                outcome = {
                    "Title": title, "Author": author, "DuplicateKey": duplicate_key,
                    "ASIN": clean_text(row.get("ASIN")),
                    "ISBN_13": clean_text(row.get("ISBN_13")),
                    "ISBN_10": clean_text(row.get("ISBN_10")),
                    STATUS_COL: "Error", MATCHED_QUERY_COL: "",
                    NOTES_COL: "Timeout", COMPLETED_COL: "No",
                    OWNED_SG_COL: "",
                }
            except (PlaywrightError, Exception) as e:
                if _is_browser_closed(e):
                    log(f"\n[STOPPED] Browser was closed — saving progress and exiting.")
                    stopped_early = True
                    break
                outcome = {
                    "Title": title, "Author": author, "DuplicateKey": duplicate_key,
                    "ASIN": clean_text(row.get("ASIN")),
                    "ISBN_13": clean_text(row.get("ISBN_13")),
                    "ISBN_10": clean_text(row.get("ISBN_10")),
                    STATUS_COL: "Error", MATCHED_QUERY_COL: "",
                    NOTES_COL: str(e), COMPLETED_COL: "No",
                    OWNED_SG_COL: "",
                }

            sg_owned_note = f" | SG Owned: {outcome[OWNED_SG_COL]}" if outcome[OWNED_SG_COL] else ""
            log(f" -> {outcome[STATUS_COL]}: {outcome[NOTES_COL]}{sg_owned_note}")
            results.append(outcome)

            # Write back to df_all
            if "DuplicateKey" in df_all.columns and duplicate_key:
                mask = df_all["DuplicateKey"].astype(str).str.strip() == duplicate_key
            else:
                mask = (
                    (df_all["Title"].astype(str).str.strip() == title) &
                    (df_all["Author"].astype(str).str.strip() == author)
                )

            df_all.loc[mask, STATUS_COL]        = outcome[STATUS_COL]
            df_all.loc[mask, MATCHED_QUERY_COL]  = outcome[MATCHED_QUERY_COL]
            df_all.loc[mask, NOTES_COL]          = outcome[NOTES_COL]
            df_all.loc[mask, COMPLETED_COL]      = outcome[COMPLETED_COL]
            if outcome[OWNED_SG_COL]:
                df_all.loc[mask, OWNED_SG_COL]  = outcome[OWNED_SG_COL]
            # If StoryGraph showed the book as already Read, flag it in Excel
            if outcome.get(READ_COL) == "Yes":
                df_all.loc[mask, READ_COL] = "Yes"
                log(f"    [EXCEL] Set Read=Yes for: {outcome['Title']}")
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            df_all.loc[mask, ADDED_DATE_COL] = today

            if idx % SAVE_EVERY_N_BOOKS == 0:
                save_excel(df_all)
                log(f"    [SAVE] Progress written ({idx}/{total})")
            time.sleep(SLEEP_BETWEEN_BOOKS)

        # Flush whatever the last batch left unsaved — this also covers the
        # interrupt and browser-closed breaks above.
        try:
            save_excel(df_all)
            log("    [SAVE] Final progress written")
        except Exception as e:
            log(f"    [SAVE] ERROR: could not write final progress: {e}")

        try:
            context.close()   # persistent context — this also writes the profile out
        except Exception:
            pass

    results_df  = pd.DataFrame(results)
    output_path = EXCEL_PATH.with_name("storygraph_results.xlsx")
    results_df.to_excel(output_path, index=False)
    if stopped_early:
        log(f"\nStopped early — progress saved to: {output_path}")
    else:
        log(f"\nDone. Results saved to: {output_path}")
    
    # Show completion message with log location
    log(f"\n{'='*60}")
    log(f"Complete! Log file: {LOG_FILE}")
    log(f"{'='*60}")
    
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    status_msg = f"Stopped early ({len(results)} books processed)" if stopped_early else f"Completed ({len(results)} books processed)"
    messagebox.showinfo(
        "StoryGraph Complete",
        f"{status_msg}\n\n"
        f"Results: {output_path.name}\n"
        f"Log: {LOG_FILE.name}\n\n"
        f"Use 'View Latest Log' in the launcher to see details.",
        parent=root
    )
    root.destroy()


if __name__ == "__main__":
    main()
