@echo off
chcp 65001 >nul
title 시작프로그램 등록
echo PC 켤때마다 자동 실행 등록 중...

set SCRIPT_PATH=%~dp0auto_manager.py
set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

:: pythonw로 실행 (CMD창 없이 백그라운드)
echo Set oShell = CreateObject("WScript.Shell") > "%TEMP%\create_shortcut.vbs"
echo Set oShortcut = oShell.CreateShortcut("%STARTUP%\스캔자동화.lnk") >> "%TEMP%\create_shortcut.vbs"
echo oShortcut.TargetPath = "pythonw" >> "%TEMP%\create_shortcut.vbs"
echo oShortcut.Arguments = """%SCRIPT_PATH%""" >> "%TEMP%\create_shortcut.vbs"
echo oShortcut.WorkingDirectory = "%~dp0" >> "%TEMP%\create_shortcut.vbs"
echo oShortcut.Description = "스캔 자동화 시스템" >> "%TEMP%\create_shortcut.vbs"
echo oShortcut.Save >> "%TEMP%\create_shortcut.vbs"
cscript //nologo "%TEMP%\create_shortcut.vbs"

echo.
echo ✅ 시작프로그램 등록 완료!
echo    PC 켤때마다 백그라운드에서 자동 실행됩니다.
echo.
echo 지금 바로 시작하려면 실행.bat 을 클릭하세요.
pause
