@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo Parent Alimtalk Tool v7.1.0 - Setup
echo ==========================================
echo.

where py >nul 2>nul
if not errorlevel 1 goto USE_PY

where python >nul 2>nul
if not errorlevel 1 goto USE_PYTHON

echo [ERROR] Python was not found.
echo Install Python first and enable "Add python.exe to PATH".
echo Then close this window and run this file again.
echo.
pause
exit /b 1

:USE_PY
set "PY_CMD=py -3"
goto SETUP

:USE_PYTHON
set "PY_CMD=python"

:SETUP
echo [1/4] Checking Python...
%PY_CMD% --version
if errorlevel 1 goto ERROR

echo.
echo [2/4] Creating virtual environment...
%PY_CMD% -m venv .venv
if errorlevel 1 goto ERROR

echo.
echo [3/4] Installing packages...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto ERROR
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto ERROR

echo.
echo [4/4] Preparing .env...
if exist ".env" goto ENV_EXISTS
copy /y ".env.example" ".env" >nul
echo A new .env file was created.
start "" notepad ".env"
goto DONE

:ENV_EXISTS
echo Existing .env file was kept.

:DONE
echo.
echo Setup completed.
echo Save the .env file, then run start_windows.bat.
echo.
pause
exit /b 0

:ERROR
echo.
echo [ERROR] Setup failed.
echo Take a screenshot of the lines above.
echo.
pause
exit /b 1
