@echo off
setlocal
echo Checking Python...
echo.
where py
py -3 --version
echo.
where python
python --version
echo.
echo If at least one version number appears, Python is available.
pause
