# 📚 Book Tools Launcher - Installation Guide

## What You're Getting

A unified launcher that replaces your three .bat files with one clean GUI, plus a book selector that exports decisions to update both spreadsheets automatically.

## Step 1: Install the Launcher

1. **Clone the repo** to `C:\Users\megha\book-tools` (outside OneDrive — see README.md):
   - `book_launcher.py`
   - `apply_reading_log.py`
   - `Launcher.bat`
   - `README.md` (optional, for reference)
   - `book-selector-decisions.example.json` (optional, shows format)

2. **Make sure these files are ALREADY in that folder** (you have them):
   - `books.py`
   - `amazon_owned_books.py`  
   - `storygraph_to_read.py`
   - `books_output.xlsx`
   - `my-reading-log.xlsx`

3. **Create desktop shortcut** (optional but recommended):
   - Right-click `Launcher.bat`
   - Choose "Send to" → "Desktop (create shortcut)"
   - Now you can double-click from your desktop!

4. **Test it**: Double-click `Launcher.bat` - you should see a window with 6 buttons.

---

## Step 2: Update the Book Selector

Your book selector is hosted at https://meghanareed.github.io/my-book-selector/

1. **Download this file**:
   - `my-book-selector-modified.html`

2. **Upload to GitHub**:
   - Go to https://github.com/meghanareed/my-book-selector
   - Click on `my-book-selector.html`
   - Click the pencil icon (Edit)
   - Delete all content
   - Open `my-book-selector-modified.html` in a text editor
   - Copy everything and paste into GitHub
   - Scroll down and click "Commit changes"

3. **Wait 1-2 minutes** for GitHub Pages to rebuild

4. **Test it**: 
   - Go to https://meghanareed.github.io/my-book-selector/
   - You should now see "Export Decisions" in the Update tab

---

## Step 3: First Run

### Workflow:

1. **Launch** - Double-click your desktop shortcut (or Launcher.bat)

2. **Open Book Selector** - Click the button in the launcher
   - Browser opens to https://meghanareed.github.io/my-book-selector/

3. **Upload fresh data** - In the selector's Update tab:
   - Upload your latest `books_output.xlsx`
   - This loads books that aren't marked Read=Yes

4. **Pick books** - Use the Pick tab:
   - Click "Select a Book for Me!"
   - Choose: Read / Ignore / Remove / Again / Close

5. **Export decisions** - When you're done, go to Update tab:
   - Click "Export Decisions JSON"
   - Save to Downloads

6. **Apply decisions** - Back in the launcher:
   - Click "Apply Selector Decisions"
   - Pick the JSON you just downloaded
   - Watch it update both spreadsheets

### What Gets Updated:

✅ `books_output.xlsx` column Y (Read) → set to "Yes" for all decisions
✅ `my-reading-log.xlsx` → new rows added (Title, Author, Genre, Status, Date)

### StoryGraph Skip Rules

Before opening the browser, **Push to StoryGraph** filters out books matching any of these:

| Column | Value | Reason |
|--------|-------|--------|
| `Read` | Yes | Already read — skip entirely |
| `Skip Storygraph` | Yes | Manually excluded |
| `StoryGraph Status` | Added or Skipped | Done in a previous run |
| `StoryGraph Completed` | Yes | Done in a previous run |
| `Owned` | not Yes | Only owned books are pushed |

**Bonus:** If StoryGraph shows a book as already Read (even if your Excel didn't know), the script automatically sets `Read = Yes` in Excel so it gets skipped on all future runs.

---

## What Changed from Before

**OLD WAY:**
- 3 separate .bat files (RunBooks.bat, run_amazon_owned_books.bat, RunStoryGraph.bat)
- Selector deleted books when you clicked "Read it!" - data lost forever
- No way to track what you ignored vs read
- Had to manually update my-reading-log.xlsx
- StoryGraph would try to add already-read books to To-Read pile
- No feedback when StoryGraph found a book you'd already read

**NEW WAY:**
- 1 launcher with buttons for everything
- Selector tracks decisions (Read/Ignored/Removed) without deleting data
- Export/import flow preserves all book metadata
- Automatic updates to both spreadsheets
- Books marked `Read=Yes` are skipped before StoryGraph even opens the browser
- If StoryGraph detects a book as already Read, it sets `Read=Yes` in Excel automatically
- Books marked Read don't show up in selector anymore

---

## Files You Got

| File | What It Does | Where It Goes |
|------|-------------|---------------|
| `book_launcher.py` | The GUI app | book-tools |
| `apply_reading_log.py` | Processes decisions JSON | book-tools |
| `Launcher.bat` | Starts the GUI | book-tools (shortcut to desktop) |
| `my-book-selector-modified.html` | Updated selector | Upload to GitHub |
| `README.md` | Documentation | book-tools (optional) |
| `book-selector-decisions.example.json` | Shows export format | book-tools (optional) |

---

## Troubleshooting

**"Script not found"**
→ Make sure all .py files are in your clone, e.g. `C:\Users\megha\book-tools`

**Python errors on first run**
→ Normal! Each button auto-installs packages first time. Wait for it to finish.

**Playwright takes forever**
→ Also normal first time - it downloads Chromium browser. Takes 2-3 minutes.

**Decisions not exporting**
→ Make sure you uploaded the modified HTML to GitHub and waited for it to rebuild

**Books still showing after marking Read**
→ Re-upload books_output.xlsx in the selector's Update tab to refresh the data

---

## Quick Reference

### Launcher Buttons:

1. **Process Screenshots** → Extracts books from screenshots (books.py)
2. **Update Amazon Owned** → Marks owned books (amazon_owned_books.py)
3. **Push to StoryGraph** → Adds to StoryGraph (storygraph_to_read.py)
4. **Sync Read from StoryGraph** → Marks SG-read books as Read (storygraph_read_sync.py)
5. **Open Book Selector** → Opens https://meghanareed.github.io/my-book-selector/
6. **Apply Selector Decisions** → Imports JSON, updates spreadsheets
7. **Fill Missing Metadata** → Backfills any missing field (reenrich_existing.py)
8. **Backfill Page Counts** → Page counts only, ignoring the skip rules
9. **Open Reading Folder** → Opens folder in Explorer

### Decisions JSON Format:
```json
{
  "exported_at": "2026-04-28T15:30:00",
  "decisions": [
    {
      "asin": "B0CTCQ6MM4",
      "title": "UNBIND",
      "author": "Adam Wright",
      "decision": "Read"
    }
  ]
}
```

Matching priority: ASIN → ISBN-13 → ISBN-10 → DuplicateKey → Title+Author
