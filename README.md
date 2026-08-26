# Book Tools Launcher

Single GUI for all your reading workflow scripts.

## What's Inside

- **book_launcher.py** - Tkinter GUI with buttons for each tool
- **apply_reading_log.py** - Processes decisions from the book selector
- **Launcher.bat** - Double-click this to start the GUI
- **book-selector-decisions.example.json** - Example format for decisions

## Setup

1. Copy all files to `C:\Users\megha\OneDrive\Documents\Reading`
2. Make sure you have these files in the same folder:
   - books.py
   - amazon_owned_books.py
   - storygraph_to_read.py
   - books_output.xlsx
   - my-reading-log.xlsx

3. **Optional:** Right-click `Launcher.bat` → Send to → Desktop (create shortcut)
   This gives you a one-click launcher on your desktop.

## Usage

### The Launcher

Double-click `Launcher.bat` (or your desktop shortcut). You'll see buttons for:

1. **Process Screenshots** - Runs books.py to extract titles from screenshots
2. **Update Amazon Owned** - Marks owned books from Amazon
3. **Push to StoryGraph** - Adds books to your StoryGraph to-read list
4. **Open Book Selector** - Opens https://meghanareed.github.io/my-book-selector/
5. **Apply Selector Decisions** - Imports decisions JSON to update both spreadsheets
6. **Open Reading Folder** - Opens the folder in Explorer

Each script runs in its own console window so you can watch the output.

### The Book Selector Workflow

1. Click "Open Book Selector" in the launcher
2. In your browser, upload a fresh copy of books_output.xlsx (Update tab)
3. Browse and pick books, marking them Read / Ignored / Removed
4. Click "Export Decisions" (in the Update tab)
5. Save the JSON to your Downloads folder
6. Back in the launcher, click "Apply Selector Decisions"
7. Pick the JSON file you just downloaded

The apply script will:
- Set column Y (Read) to "Yes" in books_output.xlsx for each decision
- Append rows to my-reading-log.xlsx

### Push to StoryGraph

The StoryGraph script adds your owned books to your StoryGraph To-Read pile and marks them as owned. Before opening the browser, it skips any book that matches any of these conditions:

| Column | Value | Why skipped |
|--------|-------|-------------|
| `Read` | Yes | Already read — no point adding to To-Read |
| `Skip Storygraph` | Yes | Manually excluded |
| `StoryGraph Status` | Added or Skipped | Already processed in a previous run |
| `StoryGraph Completed` | Yes | Already processed in a previous run |
| `Owned` | anything other than Yes | Only owned books are pushed to StoryGraph |

If StoryGraph shows a book as already **Read** (even if your Excel didn't know), the script sets `Read = Yes` in Excel automatically and skips adding it to To-Read.

### Decisions Format

The selector exports JSON like this:

```json
{
  "exported_at": "2026-04-28T15:30:00",
  "decisions": [
    {
      "asin": "B0CTCQ6MM4",
      "isbn_13": "9781920000000",
      "isbn_10": "1917190077",
      "duplicate_key": "unbind|",
      "title": "UNBIND",
      "author": "Adam Wright",
      "genre": "Romance, Contemporary",
      "decision": "Read"
    }
  ]
}
```

## Troubleshooting

**"Script not found"**: Make sure all your .py files are in the same folder as the launcher.

**Python errors**: Each button automatically installs required packages. If you see errors, try:
1. Open cmd in the Reading folder
2. Run: `py -m pip install --upgrade pip`
3. Run the launcher again

**Playwright errors**: The first time you run Amazon/StoryGraph tools, Playwright will download Chromium. This can take a few minutes.

**Book showing in StoryGraph run even though it's read**: Make sure `Read = Yes` is set in column Y of books_output.xlsx. The script filters these out before opening the browser. If StoryGraph itself shows the book as Read but your Excel doesn't, the script will detect it and set `Read = Yes` automatically for next time.

## Next Steps

Upload the modified `my-book-selector.html` to your GitHub Pages repo to get the export functionality. The modified version:
- Captures ASIN, ISBN, DuplicateKey when importing books
- Filters out books where Read=Yes  
- Tracks decisions (Read/Ignored/Removed)
- Has an "Export Decisions" button
