@echo off
REM ====================================================
REM  TW Relative Rotation - Daily Update Pipeline
REM    1. Fetch data + calc metrics -> Excel
REM    2. Generate website -> docs/index.html
REM    3. Diff signal state -> Email notify
REM    4. Commit + push -> GitHub Pages
REM ====================================================
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

echo.
echo [1/4] Fetching data and computing metrics...
python "scripts\update_data.py" --days 400
if errorlevel 1 goto :fail

echo.
echo [2/4] Generating website...
python "scripts\generate_html.py"
if errorlevel 1 goto :fail

echo.
echo [3/4] Checking signal changes and sending email...
python "scripts\notify_email.py"

echo.
echo [4/4] Deploying to GitHub Pages...
git add -A docs/ scripts/ README.md .gitignore .env.example update.bat
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "daily update %date% %time%"
    git push
) else (
    echo   ^(no changes, skip push^)
)

echo.
echo Done. Closing in 5 seconds...
timeout /t 5 >nul
exit /b 0

:fail
echo.
echo [ERROR] Pipeline failed. Check error messages above.
echo (Tip: If Excel file is open, close it and re-run.)
pause
exit /b 1
