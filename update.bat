@echo off
chcp 65001 >nul
REM ====================================================
REM  台股相對輪動模型 - 每日完整更新 pipeline
REM    1. 抓資料 + 計算指標 -> Excel
REM    2. 生成網頁 -> site/index.html
REM    3. 比對信號變化 -> Email 通知
REM    4. 提交 + 推送 -> GitHub Pages
REM ====================================================
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set FAILED=0

echo.
echo [1/4] 抓取資料與計算指標...
python "scripts\update_data.py" --days 400
if errorlevel 1 ( set FAILED=1 & goto :done )

echo.
echo [2/4] 生成網頁...
python "scripts\generate_html.py"
if errorlevel 1 ( set FAILED=1 & goto :done )

echo.
echo [3/4] 比對信號變化、寄送通知...
python "scripts\notify_email.py"

echo.
echo [4/4] 部署到 GitHub Pages...
git add -A docs/ scripts/ README.md .gitignore .env.example update.bat
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "daily update %date% %time%"
    git push
) else (
    echo   (無變動，略過 push)
)

:done
echo.
if %FAILED%==1 (
    echo [錯誤] pipeline 失敗，按任意鍵關閉...
    pause >nul
    exit /b 1
)
echo 全部完成，5 秒後關閉...
timeout /t 5 >nul
