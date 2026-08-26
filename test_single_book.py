"""
Single Book Test for StoryGraph
Tests the complete flow with one book to verify everything works.
"""
import sys
from pathlib import Path

# Add the parent directory to path so we can import from the main script
sys.path.insert(0, str(Path(__file__).parent))

from playwright.sync_api import sync_playwright
import pandas as pd

# ============================================================
# TEST CONFIGURATION
# ============================================================
TEST_ASIN = "B0F7SN9KZV"  # The book you mentioned
EXCEL_PATH = Path(r"C:\Users\megha\OneDrive\Documents\Reading\books_output.xlsx")

# Import functions from main script
from storygraph_to_read import (
    wait_for_user_login,
    process_book,
    log,
    _open_log,
    HEADLESS
)

def main():
    print("="*70)
    print(f"SINGLE BOOK TEST - ASIN: {TEST_ASIN}")
    print("="*70)
    
    _open_log()
    
    # Read Excel to find the book
    df_all = pd.read_excel(EXCEL_PATH, sheet_name="All Books")
    
    # Find the book with this ASIN
    book_row = df_all[df_all["ASIN"].astype(str).str.strip() == TEST_ASIN]
    
    if book_row.empty:
        print(f"\n❌ ERROR: ASIN {TEST_ASIN} not found in Excel!")
        print(f"\nSearched in: {EXCEL_PATH}")
        print(f"Sheet: All Books")
        input("\nPress Enter to exit...")
        return
    
    book_row = book_row.iloc[0]
    
    print(f"\n✓ Found book in Excel:")
    print(f"  Title:  {book_row.get('Title')}")
    print(f"  Author: {book_row.get('Author')}")
    print(f"  ASIN:   {book_row.get('ASIN')}")
    print(f"  Owned:  {book_row.get('Owned')}")
    print(f"\nStarting browser test...\n")
    
    # Launch browser and test
    with sync_playwright() as p:
        print("🌐 Opening browser...")
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context()
        page = context.new_page()
        
        print("🔐 Waiting for login...")
        wait_for_user_login(page)
        
        print("\n" + "="*70)
        print("PROCESSING BOOK")
        print("="*70)
        
        result = process_book(page, book_row)
        
        print("\n" + "="*70)
        print("RESULT")
        print("="*70)
        print(f"Status:         {result.get('StoryGraph Status')}")
        print(f"Notes:          {result.get('StoryGraph Notes')}")
        print(f"Matched Query:  {result.get('Matched Query')}")
        print(f"Completed:      {result.get('StoryGraph Completed')}")
        print(f"Owned on SG:    {result.get('Owned on SG')}")
        
        print("\n" + "="*70)
        
        if result.get('StoryGraph Status') == 'Added':
            print("✅ SUCCESS! Book was added to StoryGraph!")
        elif result.get('StoryGraph Status') == 'Skipped':
            if 'Already marked' in result.get('StoryGraph Notes', ''):
                print("✅ SUCCESS! Book was already on StoryGraph (as expected)")
            else:
                print(f"⚠️  SKIPPED: {result.get('StoryGraph Notes')}")
        else:
            print(f"❌ FAILED: {result.get('StoryGraph Notes')}")
        
        print("\nBrowser will stay open for 10 seconds so you can verify...")
        page.wait_for_timeout(10000)
        
        print("\nClosing browser...")
        browser.close()
    
    print("\n" + "="*70)
    print("Test complete!")
    print("="*70)
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
