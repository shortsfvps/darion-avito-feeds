@echo off
chcp 65001 > nul
pushd "%~dp0"

where python >nul 2>nul
if errorlevel 1 goto NOPYTHON

python scripts\generate_feeds.py %*
if errorlevel 1 goto FAIL

if "%~1"=="--scheduled" goto EXIT
pause
goto EXIT

:NOPYTHON
echo [ERROR] Python is not installed or not added to PATH!
echo Please install Python 3.10+ from https://www.python.org/
pause
goto EXIT

:FAIL
echo.
echo [ERROR] Feed generation failed with error code %errorlevel%!
pause

:EXIT
popd
