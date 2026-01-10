@echo off
title LibraryLint - Media Library Organization Tool
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0LibraryLint.ps1" %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Press any key to close...
    pause >nul
)
