@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title TPI MB Downloader - Uninstall
echo.
echo  This will remove TPI Measurement Book from this PC.
echo  Desktop shortcut and Python packages (.venv) will be deleted.
echo.
set /p OK=Type YES to uninstall: 
if /I not "%OK%"=="YES" (
  echo Cancelled.
  pause
  exit /b 0
)

echo.
echo Closing the app if it is open...
taskkill /F /IM pythonw.exe >nul 2>nul
timeout /t 1 /nobreak >nul

echo Removing desktop shortcut...
del /F /Q "%USERPROFILE%\Desktop\TPI MB Downloader.lnk" >nul 2>nul
del /F /Q "%PUBLIC%\Desktop\TPI MB Downloader.lnk" >nul 2>nul

echo Removing virtual environment...
if exist ".venv\" (
  rmdir /S /Q ".venv" >nul 2>nul
)

echo.
set /p DATA=Also delete saved IDs and settings (.tpidata)? Type YES or NO: 
if /I "%DATA%"=="YES" (
  if exist ".tpidata\" attrib -h -s ".tpidata" >nul 2>nul
  rmdir /S /Q ".tpidata" >nul 2>nul
  echo Saved IDs removed.
) else (
  echo Saved IDs kept.
)

echo.
echo Uninstall finished.
echo You can now delete this folder if you want a full remove:
echo   %cd%
echo.
pause
