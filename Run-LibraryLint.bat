@echo off
:: LibraryLint Launcher
:: Double-click this file to run LibraryLint without needing to open PowerShell manually

title LibraryLint

:: Change to the script directory
cd /d "%~dp0"

:: Run LibraryLint with execution policy bypass
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0LibraryLint.ps1" %*

:: Keep window open if there was an error
if errorlevel 1 (
    echo.
    echo Press any key to exit...
    pause >nul
)
