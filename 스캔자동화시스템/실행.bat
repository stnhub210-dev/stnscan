@echo off
chcp 65001 >nul
title 스캔 자동화 시스템
cd /d %~dp0
echo ==========================================
echo   스캔 자동화 통합 관리 시스템 v1.0
echo ==========================================
echo.
python auto_manager.py
pause
