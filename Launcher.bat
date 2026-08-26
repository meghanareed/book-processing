@echo off
:: ============================================================
::  Launcher.bat
::  Double-click to start the Book Tools GUI
:: ============================================================
cd /d "%~dp0"

:: Try venv first, then pythonw (GUI mode - no console), then python
if exist "venv\Scripts\pythonw.exe" (
    start "" "venv\Scripts\pythonw.exe" book_launcher.py
) else (
    where pythonw >nul 2>&1
    if %errorlevel%==0 (
        start "" pythonw book_launcher.py
    ) else (
        where py >nul 2>&1
        if %errorlevel%==0 (
            start "" pyw book_launcher.py
        ) else (
            echo ERROR: Python not found.
            echo Install Python from https://www.python.org/downloads/
            pause
            exit /b
        )
    )
)
