@echo off
:: Run from this file's own folder so the clone location does not matter
cd /d "%~dp0"
python check_books_to_process.py
