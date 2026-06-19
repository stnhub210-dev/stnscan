@echo off
chcp 65001 >nul
title STN Scan - Web Link Sync
cd /d "%~dp0"

echo ============================================
echo    STN Scan : Web Link Sync (scan.html)
echo ============================================
echo.

echo [1/6] Cleaning desktop.ini / locks inside .git ...
del /a /s /q ".git\desktop.ini" >nul 2>&1
del /f /q ".git\*.lock" >nul 2>&1
echo     done.

echo [2/6] Regenerating scan.html with live drive links ...
python generate_html.py once
if exist "_web_dash_tmp.html" del /f /q "_web_dash_tmp.html" >nul 2>&1

echo [3/6] Staging changes ...
git add -A

echo [4/6] Commit ...
git commit -m "sync web links"

echo [5/6] Push to main (triggers auto-deploy) ...
git push origin main

echo [6/6] Force deploy to gh-pages (instant) ...
git push --force origin main:gh-pages

echo.
echo ============================================
echo    DONE. Refresh https://stnscan.co.kr after 1-2 min.
echo    Top file links will be clickable.
echo ============================================
echo.
pause
