import pandas as pd
from pathlib import Path

EXCEL_PATH = Path(r"C:\Users\megha\OneDrive\Documents\Reading\books_output.xlsx")

# Column names
OWNED_COL = "Owned"
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

# Check Owned column
owned_books = df_all[df_all[OWNED_COL].apply(clean_text).str.lower() == "yes"]
print(f"Books with Owned=Yes: {len(owned_books)}")

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
        print("  - StoryGraph Completed = Yes")
        print("  - StoryGraph Status = Added or Skipped")
else:
    print(f"\nBooks to process: {len(not_completed)}")

print("\n" + "="*60)
input("\nPress Enter to close...")
