# Building the LibraryLint Installer

## Prerequisites

1. **Install Inno Setup**
   ```powershell
   winget install JRSoftware.InnoSetup
   ```

2. **Create icon file** (optional but recommended)
   - Place a 256x256 .ico file at `assets/icon.ico`
   - If no icon exists, the installer will skip it

## Building the Installer

### Option 1: Using Inno Setup GUI
1. Open Inno Setup Compiler
2. File → Open → select `installer/LibraryLint.iss`
3. Build → Compile
4. Output: `dist/LibraryLint-5.2.3-Setup.exe`

### Option 2: Command Line
```powershell
# From the LibraryLint root directory
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" installer\LibraryLint.iss
```

## Output Files

After building:
- `dist/LibraryLint-5.2.3-Setup.exe` - Installer executable

## Creating a Portable ZIP

For users who prefer portable installations:

```powershell
# From the LibraryLint root directory
$version = "5.2.3"
$files = @(
    "LibraryLint.ps1",
    "Run-LibraryLint.bat",
    "LibraryLint.psd1",
    "README.md",
    "CHANGELOG.md",
    "GETTING_STARTED.md",
    "LICENSE",
    "modules",
    "config"
)

# Create dist folder if needed
New-Item -ItemType Directory -Force -Path "dist" | Out-Null

# Create ZIP
Compress-Archive -Path $files -DestinationPath "dist\LibraryLint-$version-Portable.zip" -Force
```

## Updating Version

When releasing a new version:

1. Update `$script:AppVersion` in `LibraryLint.ps1`
2. Update `ModuleVersion` in `LibraryLint.psd1`
3. Update `#define MyAppVersion` in `installer/LibraryLint.iss`
4. Update `CHANGELOG.md`
5. Rebuild installer

## What the Installer Does

1. Copies LibraryLint files to `C:\Program Files\LibraryLint` (or user choice)
2. Creates Start Menu shortcuts
3. Optionally creates Desktop shortcut
4. Optionally installs dependencies via winget:
   - 7-Zip
   - FFmpeg
   - yt-dlp
   - MediaInfo
5. Registers uninstaller in Windows Add/Remove Programs
