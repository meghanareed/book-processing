import pandas as pd
from pathlib import Path
import sys

# Windows consoles and pipes default to cp1252, which cannot encode the ✓ ✗ →
# characters logged below; printing one raises UnicodeEncodeError mid-run.
# Force UTF-8 so this holds however the script is started.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

EXCEL_PATH = Path(r"C:\Users\megha\OneDrive\Documents\Reading\books_output.xlsx")

# Column names
OWNED_COL = "Owned"
READ_COL = "Read"
STATUS_COL = "StoryGraph Status"
COMPLETED_COL = "StoryGraph Completed"

def clean_text(val) -> str:
    if pd.isna(val):
        return ""
    return str(val).strip()

print("="*60)
print("StoryGraph Processing Check")
print("="*60)

# Read Excel
df_all = pd.read_excel(EXCEL_PATH, sheet_name="All Books")
print(f"\nTotal books in Excel: {len(df_all)}")

# Owned = in the Amazon library, read or not.
if OWNED_COL not in df_all.columns:
    print(f"\n  !! No '{OWNED_COL}' column in the sheet at all.")
    print("     Run Update Amazon Owned to rebuild it.")
    input("\nPress Enter to close...")
    raise SystemExit(1)

owned_books = df_all[df_all[OWNED_COL].apply(clean_text).str.lower() == "yes"]
print(f"Books with Owned=Yes (your Amazon library): {len(owned_books)}")

# Read is what separates the library from the pile still to read.  Compare the
# unread number against the Kindle app -- that is the one that should match.
if READ_COL in df_all.columns:
    read_flag = owned_books[READ_COL].apply(clean_text).str.lower() == "yes"
    print(f"  of those, already read  : {int(read_flag.sum())}")
    print(f"  of those, still to read : {int((~read_flag).sum())}")
    owned_books = owned_books[~read_flag]
else:
    print(f"  (no '{READ_COL}' column yet -- nothing has been marked read)")

# Check already completed
if COMPLETED_COL in df_all.columns:
    not_completed = owned_books[owned_books[COMPLETED_COL].apply(clean_text).str.lower() != "yes"]
    print(f"Books NOT already completed: {len(not_completed)}")
else:
    not_completed = owned_books
    print(f"No Completed column found - all owned books will be processed")

# Check terminal status
if STATUS_COL in df_all.columns:
    terminal_statuses = {"added", "skipped"}
    pending = not_completed[~not_completed[STATUS_COL].apply(clean_text).str.lower().isin(terminal_statuses)]
    print(f"Books needing processing: {len(pending)}")
    
    if len(pending) > 0:
        print("\nFirst 5 books that will be processed:")
        for i, row in pending.head(5).iterrows():
            title = clean_text(row.get("Title"))
            author = clean_text(row.get("Author"))
            print(f"  - {title} by {author}")
    else:
        print("\n⚠️  NO BOOKS WILL BE PROCESSED!")
        print("\nReasons a book might be skipped:")
        print("  - Owned ≠ Yes")
        print("  - Read = Yes (already finished)")
        print("  - StoryGraph Completed = Yes")
        print("  - StoryGraph Status = Added or Skipped")
else:
    print(f"\nBooks to process: {len(not_completed)}")

print("\n" + "="*60)
input("\nPress Enter to close...")
