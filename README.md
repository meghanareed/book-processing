# Book Tools Launcher

Single GUI for all your reading workflow scripts.

## What's Inside

- **book_launcher.py** - Tkinter GUI with buttons for each tool
- **apply_reading_log.py** - Processes decisions from the book selector
- **Launcher.bat** - Double-click this to start the GUI
- **book-selector-decisions.example.json** - Example format for decisions

## Setup

The code lives in a git clone; the data stays in OneDrive. Keep them apart — OneDrive
syncing a `.git` folder corrupts the repository. See `REPO_SETUP.md` for the full rationale.

```
C:\Users\megha\book-tools\                     <- this repo (code)
C:\Users\megha\OneDrive\Documents\Reading\     <- spreadsheets and logs
C:\Users\megha\OneDrive\Pictures\...\DCIM\Books  <- incoming screenshots
```

```powershell
cd C:\Users\megha
git clone https://github.com/meghanareed/book-processing.git book-tools
cd book-tools

py -m venv venv                 # name it "venv" — Launcher.bat looks for that
.\venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m playwright install chromium

copy launcher_config.example.json launcher_config.json
```

Then start the launcher and open **Settings** to set your OpenAI API key and confirm the
three folder paths on the Books.py tab.

**Optional:** Right-click `Launcher.bat` → Send to → Desktop (create shortcut)
for a one-click launcher.

## Usage

### The Launcher

Double-click `Launcher.bat` (or your desktop shortcut). You'll see buttons for:

1. **Process Screenshots** - Runs books.py to extract titles from screenshots
2. **Update Amazon Owned** - Marks owned books from Amazon
3. **Push to StoryGraph** - Adds books to your StoryGraph to-read list
4. **Open Book Selector** - Opens https://meghanareed.github.io/my-book-selector/
5. **Apply Selector Decisions** - Imports decisions JSON to update both spreadsheets
6. **Fill Missing Metadata** - Runs reenrich_existing.py to backfill PageCount, Genre, Tropes
7. **Open Reading Folder** - Opens the data folder in Explorer
8. **View Latest Log** - Opens the newest log file from the data folder's `logs\`

Output streams into the launcher's Output Log panel as each script runs. The panel clears
at the start of every run; the full history is in `logs\` in the data folder.

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

**Python errors**: Each button installs its required packages before running. If you see errors, try:
1. Open PowerShell in your `book-tools` folder
2. Run: `py -m pip install -r requirements.txt`
3. Run the launcher again

**Playwright errors**: The first time you run Amazon/StoryGraph tools, Playwright will download Chromium. This can take a few minutes.

**Book showing in StoryGraph run even though it's read**: Make sure `Read = Yes` is set in column Y of books_output.xlsx. The script filters these out before opening the browser. If StoryGraph itself shows the book as Read but your Excel doesn't, the script will detect it and set `Read = Yes` automatically for next time.

## Next Steps

Upload the modified `my-book-selector.html` to your GitHub Pages repo to get the export functionality. The modified version:
- Captures ASIN, ISBN, DuplicateKey when importing books
- Filters out books where Read=Yes  
- Tracks decisions (Read/Ignored/Removed)
- Has an "Export Decisions" button
