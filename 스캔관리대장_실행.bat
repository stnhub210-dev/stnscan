@echo off
chcp 65001 > nul
title 스캔관리대장 서버
cd /d C:\Users\User\Desktop\스캔
echo.
echo ========================================
echo   스캔관리대장 서버 시작 중...
echo   포트: 8765
echo   종료하려면 이 창을 닫으세요
echo ========================================
echo.
python generate_html.py
pause
