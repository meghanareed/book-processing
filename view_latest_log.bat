@echo off
:: Logs are written next to the spreadsheets, not next to the code
set "LOGDIR=C:\Users\megha\OneDrive\Documents\Reading\logs"
cd /d "%LOGDIR%"

echo Finding latest StoryGraph log...
echo.

REM Find the most recent storygraph log file
for /f "delims=" %%f in ('dir /b /o-d storygraph_*.log 2^>nul') do (
    set LATEST=%%f
    goto :found
)

:found
if not defined LATEST (
    echo No log files found in logs folder!
    pause
    exit /b
)

echo Latest log: %LATEST%
echo Location: %LOGDIR%\%LATEST%
echo.
echo ============================================================
type "%LATEST%"
echo ============================================================
echo.
pause
