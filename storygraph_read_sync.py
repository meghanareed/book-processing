"""
storygraph_read_sync.py
-----------------------
Walks your StoryGraph "read" shelf and sets Read = Yes in books_output.xlsx
for every book on it.

Why this exists: storygraph_to_read.py only notices that a book is already
read while it is processing that book, and a book stops being processed once
it reaches a terminal status.  So finishing a book on StoryGraph after it was
added never makes it back to the spreadsheet, and the book selector keeps
offering books you have already read.

This reads only — nothing is changed on StoryGraph.
"""

import sys

# Windows consoles and pipes default to cp1252, which cannot encode the ✓ ✗ →
# characters logged below; printing one raises UnicodeEncodeError mid-run.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import re

import pandas as pd
from playwright.sync_api import sync_playwright

# Reuse the login, browser profile and spreadsheet plumbing rather than
# duplicating it.  Importing runs that module's config load, which is harmless.
from storygraph_to_read import (
    BASE_URL,
    DECISION_COL,
    EXCEL_PATH,
    READ_COL,
    SHEET_NAME,
    _SG_CFG,
    _open_log,
    accept_dialogs,
    clean_text,
    ensure_storygraph_columns,
    fold,
    launch_browser,
    log,
    safe_inner_text,
    save_excel,
    short_title,
    wait_for_user_login,
)

# StoryGraph paginates the shelf; stop after this many pages as a runaway guard.
MAX_PAGES = int(_SG_CFG.get("read_sync_max_pages", 200))
USERNAME = clean_text(_SG_CFG.get("username"))


def discover_username(page) -> str:
    """Find the logged-in username from a profile link on the page."""
    for selector in ('a[href^="/profile/"]', 'a[href*="/profile/"]'):
        try:
            locator = page.locator(selector).first
            if locator.count() == 0:
                continue
            href = clean_text(locator.get_attribute("href"))
            match = re.search(r"/profile/([^/?#]+)", href)
            if match:
                return match.group(1)
        except Exception:
            continue
    return ""


def scrape_read_shelf(page, username: str) -> list[tuple[str, str]]:
    """Return (title, author) for every book on the read shelf."""
    books: list[tuple[str, str]] = []
    seen_hrefs: set[str] = set()

    for page_num in range(1, MAX_PAGES + 1):
        url = f"{BASE_URL}/books-read/{username}?page={page_num}"
        log(f"  [READ] Page {page_num}: {url}")
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(1500)
        except Exception as e:
            log(f"  [READ] Could not load page {page_num}: {e}")
            break

        panes = page.locator(".book-pane")
        count = panes.count()
        if count == 0:
            # Fall back to raw book links in case the pane class changes.
            count = page.locator('a[href*="/books/"]').count()
            if count == 0:
                log(f"  [READ] No books on page {page_num} — stopping.")
                break

        found_here = 0
        for i in range(count):
            try:
                pane = panes.nth(i)
                title_link = pane.locator('a[href*="/books/"]').first
                href = clean_text(title_link.get_attribute("href"))
                # Skip the editions/owned helper links that share the prefix.
                if not href or "/editions" in href or href in seen_hrefs:
                    continue
                title = safe_inner_text(title_link)
                if not title:
                    continue
                author = safe_inner_text(pane.locator('a[href*="/authors/"]').first)
                seen_hrefs.add(href)
                books.append((title, author))
                found_here += 1
            except Exception:
                continue

        log(f"  [READ] Page {page_num}: {found_here} book(s)")
        if found_here == 0:
            break

    return books


def apply_to_excel(books: list[tuple[str, str]]) -> None:
    """Set Read = Yes for each shelf book found in the spreadsheet."""
    if not EXCEL_PATH.exists():
        log(f"  ERROR: {EXCEL_PATH} not found.")
        return

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df = ensure_storygraph_columns(df)
    if READ_COL not in df.columns:
        df[READ_COL] = ""

    # Index the spreadsheet by folded title|author, and by title alone as a
    # fallback for rows whose author differs from StoryGraph's spelling.
    by_pair: dict[tuple[str, str], list[int]] = {}
    by_title: dict[str, list[int]] = {}
    for idx, row in df.iterrows():
        title = clean_text(row.get("Title"))
        if not title:
            continue
        author = clean_text(row.get("Author"))
        # Authors are stored as "First Last" and may list several; match on the
        # last name of the first author, as the picker does.
        first_author = author.split(",")[0].strip()
        last_name = fold(first_author.split()[-1]) if first_author else ""
        for key in {fold(title), fold(short_title(title))}:
            by_pair.setdefault((key, last_name), []).append(idx)
            by_title.setdefault(key, []).append(idx)

    already = newly = unmatched = ambiguous = 0
    for title, author in books:
        sg_last = fold(author.split(",")[0].strip().split()[-1]) if author.strip() else ""
        keys = {fold(title), fold(short_title(title))}

        rows: list[int] = []
        for key in keys:
            rows = by_pair.get((key, sg_last), [])
            if rows:
                break
        if not rows:
            # No author agreement — fall back to title, but only when it is
            # unambiguous, so two different books sharing a title aren't both
            # marked read.
            candidates = {i for key in keys for i in by_title.get(key, [])}
            if len(candidates) == 1:
                rows = list(candidates)
            elif len(candidates) > 1:
                ambiguous += 1
                log(f"    [SKIP] {title!r} matches {len(candidates)} rows — left alone")
                continue

        if not rows:
            unmatched += 1
            continue

        for idx in rows:
            if clean_text(df.at[idx, READ_COL]).lower() == "yes":
                already += 1
            else:
                df.at[idx, READ_COL] = "Yes"
                newly += 1
                log(f"    [READ] Marked read: {title!r} by {author!r}")

    log("")
    log(f"  Shelf books:           {len(books)}")
    log(f"  Newly marked Read:     {newly}")
    log(f"  Already marked Read:   {already}")
    log(f"  Not in spreadsheet:    {unmatched}")
    log(f"  Ambiguous title match: {ambiguous}")

    if newly:
        save_excel(df)
        log(f"  [OK] Saved {EXCEL_PATH.name}")
    else:
        log("  Nothing to change — spreadsheet left as is.")


def main() -> None:
    _open_log()
    log("=" * 60)
    log("StoryGraph Read Shelf -> Excel")
    log("=" * 60)

    with sync_playwright() as p:
        context = launch_browser(p)
        page = context.pages[0] if context.pages else context.new_page()
        accept_dialogs(page)

        wait_for_user_login(page)

        username = USERNAME or discover_username(page)
        if not username:
            log("  ERROR: Could not work out your StoryGraph username.")
            log("  Set it in launcher_config.json -> storygraph.username")
            context.close()
            return
        log(f"  Username: {username}")

        books = scrape_read_shelf(page, username)
        log(f"  Found {len(books)} book(s) on the read shelf")

        try:
            context.close()
        except Exception:
            pass

    if books:
        apply_to_excel(books)
    else:
        log("  No books found — nothing to do.")
        log("  If your shelf isn't empty, check the username and the page layout.")


if __name__ == "__main__":
    main()
