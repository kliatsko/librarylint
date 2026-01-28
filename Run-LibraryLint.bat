@echo off
:: LibraryLint Launcher
:: Double-click this file to run LibraryLint without needing to open PowerShell manually

title LibraryLint

:: Change to the script directory
cd /d "%~dp0"

:: Display splash
echo.
echo   [96m╔══════════════════════════════════════════╗[0m
echo   [96m║[0m         [93mL I B R A R Y L I N T[0m           [96m║[0m
echo   [96m║[0m    [37mMedia Library Organization Tool[0m      [96m║[0m
echo   [96m╚══════════════════════════════════════════╝[0m
echo.

:: Run LibraryLint with execution policy bypass
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0LibraryLint.ps1" %*

:: Keep window open if there was an error
if errorlevel 1 (
    echo.
    echo Press any key to exit...
    pause >nul
)
