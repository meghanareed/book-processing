@echo off
:: Run from this file's own folder so the clone location does not matter
cd /d "%~dp0"
set PYTHONUNBUFFERED=1

echo ============================================================
echo SINGLE BOOK TEST - B0F7SN9KZV
echo ============================================================
echo.
echo This will test the complete flow with just one book:
echo   1. Open browser
echo   2. Wait for login
echo   3. Search for book
echo   4. Match author
echo   5. Click book link
echo   6. Click "to read"
echo   7. Click "mark as owned"
echo.
echo Browser will stay open for 10 seconds so you can verify!
echo.
pause

python -u test_single_book.py

echo.
pause
