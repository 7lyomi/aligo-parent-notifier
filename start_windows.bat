@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

if not exist ".venv\Scripts\python.exe" goto NO_INSTALL
if not exist ".env" goto NO_ENV

echo Starting Parent Alimtalk Tool v7...
echo Keep this window open while using the program.
echo Press Ctrl+C here to stop it.
echo.

start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://127.0.0.1:5050'"
".venv\Scripts\python.exe" app.py

echo.
echo The program has stopped.
pause
exit /b 0

:NO_INSTALL
echo [ERROR] Installation is not complete.
echo Run install_windows.bat first.
pause
exit /b 1

:NO_ENV
echo [ERROR] The .env file is missing.
echo Run install_windows.bat again.
pause
exit /b 1
