"""
StoryGraph Diagnostic - Shows exactly why books are being filtered out
Run this to see how many books qualify for StoryGraph processing
"""
import pandas as pd
from pathlib import Path
import json

EXCEL_PATH = Path(r"C:\Users\megha\OneDrive\Documents\Reading\books_output.xlsx")
CONFIG_PATH = EXCEL_PATH.parent / "launcher_config.json"

print("="*60)
print("StoryGraph Processing Diagnostic")
print("="*60)

# Check launcher config
if CONFIG_PATH.exists():
    try:
        config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        max_books = config.get("storygraph", {}).get("max_books", 0)
        print(f"✓ Launcher config found: max_books = {max_books}")
    except Exception as e:
        print(f"✗ Error reading launcher config: {e}")
        max_books = 0
else:
    print(f"✗ No launcher config at: {CONFIG_PATH}")
    max_books = 0

print(f"\n")

# Load Excel
df = pd.read_excel(EXCEL_PATH, sheet_name="All Books")
print(f"Total rows in Excel: {len(df)}")

# Check for required columns
cols = {
    "Owned": "Owned" in df.columns,
    "StoryGraph Status": "StoryGraph Status" in df.columns,
    "StoryGraph Completed": "StoryGraph Completed" in df.columns,
    "Skip Storygraph": "Skip Storygraph" in df.columns,
    "DuplicateKey": "DuplicateKey" in df.columns,
}

print(f"\nColumns present:")
for col, present in cols.items():
    print(f"  {col}: {'✓' if present else '✗ MISSING'}")

print(f"\n" + "="*60)
print("Filtering Analysis")
print("="*60)

def clean(v):
    return str(v or '').strip()

# Start with full dataset
current = df.copy()
print(f"Starting with: {len(current)} rows")

# Step 1: Remove duplicates
if "DuplicateKey" in current.columns:
    before = len(current)
    current = current.drop_duplicates(subset=["DuplicateKey"], keep="first")
    removed = before - len(current)
    print(f"After removing duplicates: {len(current)} rows (-{removed})")

# Step 2: Skip Storygraph filter
if "Skip Storygraph" in current.columns:
    before = len(current)
    current = current[current["Skip Storygraph"].apply(lambda x: clean(x).lower() != "yes")]
    removed = before - len(current)
    print(f"After Skip Storygraph filter: {len(current)} rows (-{removed})")

# Step 3: Terminal status filter
if "StoryGraph Completed" in current.columns and "StoryGraph Status" in current.columns:
    before = len(current)
    current = current[
        (current["StoryGraph Completed"].apply(lambda x: clean(x).lower() != "yes")) &
        (~current["StoryGraph Status"].isin(["Added", "Skipped"]))
    ]
    removed = before - len(current)
    print(f"After terminal status filter: {len(current)} rows (-{removed})")

# Step 4: Owned = Yes filter
if "Owned" in current.columns:
    before = len(current)
    current = current[current["Owned"].apply(lambda x: clean(x).lower() == "yes")]
    removed = before - len(current)
    print(f"After Owned=Yes filter: {len(current)} rows (-{removed})")

# Step 5: TEST_LIMIT
if max_books and max_books > 0:
    before = len(current)
    current = current.head(max_books)
    removed = before - len(current)
    print(f"After max_books limit ({max_books}): {len(current)} rows (-{removed})")

print(f"\n" + "="*60)
print(f"FINAL: {len(current)} books will be processed")
print("="*60)

if len(current) > 0:
    print(f"\nFirst 5 books that will be processed:")
    for i, (idx, row) in enumerate(current.head(5).iterrows(), 1):
        title = row.get("Title", "")
        author = row.get("Author", "")
        print(f"{i}. {title} | {author}")
else:
    print("\n⚠️ NO BOOKS WILL BE PROCESSED")
    print("\nPossible reasons:")
    print("  1. All books already have StoryGraph Completed=Yes")
    print("  2. All books have Skip Storygraph=Yes")
    print("  3. No books have Owned=Yes")
    print("  4. max_books is set to a very low number")

print(f"\nTo process more books:")
print(f"  • Set launcher Settings > StoryGraph > Max Books = 0")
print(f"  • Check books_output.xlsx for Owned=Yes books")
print(f"  • Look for books with StoryGraph Completed≠Yes")

input("\nPress Enter to close...")
