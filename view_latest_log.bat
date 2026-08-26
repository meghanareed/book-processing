@echo off
cd /d "C:\Users\megha\OneDrive\Documents\Reading\logs"

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
echo Location: C:\Users\megha\OneDrive\Documents\Reading\logs\%LATEST%
echo.
echo ============================================================
type "%LATEST%"
echo ============================================================
echo.
pause
