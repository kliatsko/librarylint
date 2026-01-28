<#
.SYNOPSIS
    Builds LibraryLint release packages (installer and portable ZIP)
.DESCRIPTION
    Creates both an Inno Setup installer and a portable ZIP file for distribution.
    Requires Inno Setup to be installed for the installer build.
.PARAMETER SkipInstaller
    Skip building the installer (useful if Inno Setup is not installed)
.PARAMETER SkipPortable
    Skip building the portable ZIP
.EXAMPLE
    .\Build-Release.ps1
.EXAMPLE
    .\Build-Release.ps1 -SkipInstaller
#>

param(
    [switch]$SkipInstaller,
    [switch]$SkipPortable
)

$ErrorActionPreference = "Stop"

# Get version from main script
$scriptContent = Get-Content -Path ".\LibraryLint.ps1" -Raw
if ($scriptContent -match '\$script:AppVersion\s*=\s*"([^"]+)"') {
    $version = $matches[1]
} else {
    Write-Error "Could not find version in LibraryLint.ps1"
    exit 1
}

Write-Host "Building LibraryLint v$version" -ForegroundColor Cyan
Write-Host ""

# Create dist folder
if (-not (Test-Path ".\dist")) {
    New-Item -ItemType Directory -Path ".\dist" | Out-Null
}

# Build Installer
if (-not $SkipInstaller) {
    Write-Host "=== Building Installer ===" -ForegroundColor Yellow

    # Find Inno Setup
    $isccPaths = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )

    $iscc = $isccPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $iscc) {
        Write-Host "Inno Setup not found. Install with: winget install JRSoftware.InnoSetup" -ForegroundColor Red
        Write-Host "Skipping installer build..." -ForegroundColor Yellow
    } else {
        Write-Host "Using Inno Setup: $iscc" -ForegroundColor Gray

        # Update version in .iss file
        $issPath = ".\installer\LibraryLint.iss"
        $issContent = Get-Content -Path $issPath -Raw
        $issContent = $issContent -replace '#define MyAppVersion "[^"]+"', "#define MyAppVersion `"$version`""
        Set-Content -Path $issPath -Value $issContent -NoNewline

        # Compile
        & $iscc ".\installer\LibraryLint.iss"

        if ($LASTEXITCODE -eq 0) {
            Write-Host "Installer created: dist\LibraryLint-$version-Setup.exe" -ForegroundColor Green
        } else {
            Write-Host "Installer build failed!" -ForegroundColor Red
        }
    }
    Write-Host ""
}

# Build Portable ZIP
if (-not $SkipPortable) {
    Write-Host "=== Building Portable ZIP ===" -ForegroundColor Yellow

    $zipPath = ".\dist\LibraryLint-$version-Portable.zip"

    # Remove old ZIP if exists
    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force
    }

    # Files to include
    $files = @(
        "LibraryLint.ps1",
        "Run-LibraryLint.bat",
        "LibraryLint.psd1",
        "README.md",
        "CHANGELOG.md",
        "GETTING_STARTED.md",
        "LICENSE"
    )

    # Create temp folder for packaging
    $tempDir = ".\dist\LibraryLint-$version"
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    # Copy files
    foreach ($file in $files) {
        if (Test-Path $file) {
            Copy-Item $file -Destination $tempDir
        }
    }

    # Copy folders
    Copy-Item ".\modules" -Destination "$tempDir\modules" -Recurse
    Copy-Item ".\config" -Destination "$tempDir\config" -Recurse

    # Copy assets if exists
    if (Test-Path ".\assets") {
        Copy-Item ".\assets" -Destination "$tempDir\assets" -Recurse
    }

    # Create ZIP
    Compress-Archive -Path $tempDir -DestinationPath $zipPath -Force

    # Cleanup temp folder
    Remove-Item $tempDir -Recurse -Force

    Write-Host "Portable ZIP created: dist\LibraryLint-$version-Portable.zip" -ForegroundColor Green
    Write-Host ""
}

# Summary
Write-Host "=== Build Complete ===" -ForegroundColor Cyan
Write-Host ""
Get-ChildItem ".\dist" | ForEach-Object {
    $size = "{0:N2} MB" -f ($_.Length / 1MB)
    Write-Host "  $($_.Name) ($size)" -ForegroundColor White
}
Write-Host ""
Write-Host "Upload these files to GitHub Releases for v$version" -ForegroundColor Gray
