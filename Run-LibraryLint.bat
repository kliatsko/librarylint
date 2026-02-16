@echo off
:: LibraryLint Launcher
:: Double-click this file to run LibraryLint without needing to open PowerShell manually

title LibraryLint

:: Change to the script directory
cd /d "%~dp0"

:: Set UTF-8 codepage for box-drawing characters
chcp 65001 >nul 2>&1

:: Display splash
echo.
echo   +------------------------------------------+
echo   :         L I B R A R Y L I N T            :
echo   :    Media Library Organization Tool        :
echo   +------------------------------------------+
echo.

:: Run LibraryLint with execution policy bypass
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0LibraryLint.ps1" %*

:: Keep window open if there was an error
if errorlevel 1 (
    echo.
    echo Press any key to exit...
    pause >nul
)
