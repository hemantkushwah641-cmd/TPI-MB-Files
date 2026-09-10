@echo off
setlocal
cd /d "%~dp0"
title TPI MB Downloader - Install
echo.
echo  TPI Measurement Book  -  first-time install
echo.

set "PYLAUNCH="
where py >nul 2>nul
if %errorlevel%==0 (
  set "PYLAUNCH=py -3"
) else (
  where python >nul 2>nul
  if %errorlevel%==0 (
    set "PYLAUNCH=python"
  )
)
if not defined PYLAUNCH (
  echo ERROR: Python 3.12+ not found.
  echo Install Python from python.org and tick "Add python.exe to PATH".
  pause
  exit /b 1
)

echo [1/4] Creating virtual environment...
%PYLAUNCH% -m venv .venv
if errorlevel 1 (
  echo ERROR: venv failed.
  pause
  exit /b 1
)

echo [2/4] Installing Python packages...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo ERROR: pip install failed. Check internet.
  pause
  exit /b 1
)

echo [3/4] Installing Chromium for Playwright...
".venv\Scripts\python.exe" -m playwright install chromium

echo [4/4] Desktop shortcut...
cscript //nologo CREATE_SHORTCUT.vbs

echo.
echo  Install complete.
echo  Use the desktop icon  "TPI MB Downloader"  or START_TPI.bat
echo  First run: create a master password (8+ characters).
echo  Every 15 days the app asks for it again.
echo  Portal passwords stay encrypted. Hide modules from Modules / Lock.
echo  Use UNINSTALL.bat if you want to remove the app later.
echo.
pause
