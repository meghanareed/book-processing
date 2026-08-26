# StoryGraph Matching - Troubleshooting Guide

## 🔧 What I Fixed

### 1. **Lowered Match Threshold**
- **Old:** 0.65 (65% similarity required)
- **New:** 0.60 (60% similarity required)
- **Why:** StoryGraph may format titles differently (capitalization, punctuation, subtitles)

### 2. **Enhanced Debugging Logs**
Every run now creates `storygraph_YYYYMMDD_HHMMSS.log` with:
- Every search query attempted
- All candidates found for each search
- Title similarity scores (0.0-1.0)
- Author verification results
- Top 5 matches ranked by score
- Exactly why each match failed

### 3. **More Flexible Selectors**
Added fallback selectors for:
- Search results (5 different selector strategies)
- To-Read button (5 variations)
- Mark as Owned link (6 variations)

### 4. **Better Author Checking**
- Checks parent container first (fast)
- Falls back to sibling elements
- Less strict - allows near-matches to proceed

### 5. **Launcher Integration**
Reads max_books setting from launcher config

---

## 📊 Reading the Logs

The log file shows **exactly** what's happening. Here's what to look for:

### Example: Successful Match
```
[21:30:15] [1/5] Dark Desires | Jane Doe (owned=YES)
[21:30:15]     [PROCESS] Trying 3 queries: ['B0ABC123', '9781234567890', 'Dark Desires Jane Doe']
[21:30:15]     [SEARCH] Searching for: 'Dark Desires Jane Doe'
[21:30:16]     [SEARCH] Search submitted, landed on: https://app.thestorygraph.com/search?query=...
[21:30:16]     [PICK] Looking for: title='Dark Desires', author='Jane Doe' (last name: 'doe')
[21:30:16]     [PICK] Selector 'a[href*="/books/"]': 12 candidates found
[21:30:16]     [PICK] [0] score=0.950 author_ok=True | "Dark Desires" | /books/abc123
[21:30:16]     [PICK] ✓ Best match (score=0.950): "Dark Desires"
[21:30:17]     [PICK] ✓ Landed on: https://app.thestorygraph.com/books/abc123
[21:30:17]     [TO-READ] Clicking button (found: button.read-status-button)
[21:30:18]     [OWNED] Clicking Mark as Owned (found: a.mark-as-owned-link)
[21:30:19]  -> Added: Clicked To Read | Clicked Mark as Owned
```

### Example: No Match Found
```
[21:30:20] [2/5] God of Ruin | Rina Kent (owned=YES)
[21:30:20]     [PROCESS] Trying 4 queries: ['B0XYZ789', 'God of Ruin Rina Kent', 'God of Ruin']
[21:30:20]     [SEARCH] Searching for: 'God of Ruin Rina Kent'
[21:30:21]     [PICK] Looking for: title='God of Ruin', author='Rina Kent' (last name: 'kent')
[21:30:21]     [PICK] Selector 'a[href*="/books/"]': 8 candidates found
[21:30:21]     [PICK] Top candidates:
[21:30:21]     [PICK]   1. score=0.450 author_ok=True | "Gods of Ruin (Series)" 
[21:30:21]     [PICK]   2. score=0.380 author_ok=False | "Ruin and Rising"
[21:30:21]     [PICK] ✗ No match >= 0.60 found
[21:30:21]     [PICK] Threshold: 0.60, Best score found: 0.450
[21:30:21]     [PROCESS] No result matched, trying next query
[21:30:21]     [SEARCH] Searching for: 'God of Ruin'
...
```

---

## 🔍 Debugging Steps

### Step 1: Check the Log File
1. Run "Push to StoryGraph" from launcher
2. Open the newest `storygraph_YYYYMMDD_HHMMSS.log` file
3. Find the book that's failing
4. Look for the `[PICK]` section

### Step 2: Interpret the Results

**If you see candidates but no matches:**
```
[PICK] Selector 'a[href*="/books/"]': 12 candidates found
[PICK] Top candidates:
  1. score=0.550 author_ok=True | "Book Title Here"
  2. score=0.450 author_ok=False | "Another Book"
[PICK] ✗ No match >= 0.60 found
```
**Problem:** Best score (0.550) is below threshold (0.60)
**Solution:** Lower `TITLE_MATCH_THRESHOLD` in the script (try 0.50)

**If you see "0 candidates found" for ALL selectors:**
```
[PICK] Selector 'a[href*="/books/"]': 0 candidates found
[PICK] Selector 'main a[href*="/books/"]': 0 candidates found
```
**Problem:** StoryGraph changed their HTML structure
**Solution:** Inspect the search results page HTML to find new selectors

**If author_ok=False is blocking good matches:**
```
[PICK] [0] score=0.850 author_ok=False | "Perfect Title Match"
```
**Problem:** Author verification is too strict
**Solution:** Check if author name format changed (first+last vs last, first)

### Step 3: Quick Fixes

**Lower the threshold:**
Edit `storygraph_to_read.py`, line 30:
```python
TITLE_MATCH_THRESHOLD = 0.50  # Was 0.60
```

**Disable author check temporarily (for testing):**
Find line ~220, change:
```python
author_ok = True  # Force True to bypass author check
```

### Step 4: HTML Structure Changed?

If StoryGraph completely redesigned their site:

1. Open the log, find: `[SEARCH] Search submitted, landed on: [URL]`
2. Copy that URL and open it in a browser
3. Right-click a book result → Inspect
4. Look for the link element's selector
5. Update the `selectors` list in `pick_first_matching_result()`

---

## 🎯 Common Issues

### Issue: "No candidates found"
**Cause:** StoryGraph changed HTML structure
**Fix:** Inspect page, update selectors

### Issue: "score=0.45 but threshold is 0.60"
**Cause:** Title has subtitle, punctuation differences, etc.
**Fix:** Lower threshold or improve `_title_score()` function

### Issue: "author_ok=False blocking matches"
**Cause:** Author name format mismatch or not visible near result
**Fix:** Make author check more lenient or disable it

### Issue: "Already marked To Read" for books not in your list
**Cause:** Button selector changed
**Fix:** Check `book_already_to_read()` selectors

---

## 📝 Next Steps If Still Failing

1. **Share the log file** - The timestamped log shows exactly what's happening

2. **Test with one book manually:**
   - Set `TEST_LIMIT = 1` in Settings
   - Run the script
   - Share the log section for that one book

3. **Check if StoryGraph changed:**
   - Log in to StoryGraph manually
   - Search for a known book
   - Right-click the result → Inspect
   - Check if the HTML structure matches our selectors

4. **Provide HTML sample:**
   - If structure changed, share the HTML of a search result
   - I can update selectors to match

---

## 🛠️ Files Modified

1. **storygraph_to_read.py** - Enhanced with debugging & logging
2. **Launcher v2.1** - Now supports max_books setting

**All console output is logged** - nothing gets lost anymore!
