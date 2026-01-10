<#
.SYNOPSIS
    LibraryLint - Automated media library organization and cleanup tool

.DESCRIPTION
    This PowerShell script automates the cleanup and organization of downloaded media files (movies and TV shows).
    It performs the following operations:

    FOR MOVIES:
    - Extracts archives (.rar, .zip, .7z, .tar, .gz, .bz2)
    - Removes unnecessary files (samples, proofs, screenshots)
    - Processes trailers: Move to _Trailers folder or delete
    - Processes subtitles: Keep preferred language (English), delete others
    - Creates individual folders for loose video files
    - Cleans folder names by removing quality/codec/release tags
    - Formats movie years with parentheses (e.g., "Movie 2024" -> "Movie (2024)")
    - Replaces dots with spaces in folder names
    - Generates Kodi-compatible NFO files (optional)
    - Parses existing NFO files for metadata display

    FOR TV SHOWS:
    - Extracts all archives
    - Removes unnecessary files (samples, proofs, screenshots)
    - Processes subtitles: Keep preferred language
    - Parses episode info (S01E01, 1x01, multi-episode S01E01-E03)
    - Organizes episodes into Season folders
    - Renames episodes to standard format (optional)
    - Shows episode summary with gap detection
    - Removes empty folders

    FEATURES:
    - Dry-run mode: Preview all changes before applying them
    - Comprehensive logging: All operations logged with timestamps
    - Error handling: Graceful handling of locked files and permission issues
    - Progress tracking: Visual progress indicators for long operations
    - Multi-format support: Handles mp4, mkv, avi, mov, wmv, flv, m4v
    - Subtitle handling: Keep English subtitles (.srt, .sub, .idx, .ass, .ssa, .vtt)
    - Trailer management: Move trailers to _Trailers folder instead of deleting
    - NFO support: Read existing and generate new Kodi NFO files
    - Duplicate detection: Find and remove duplicate movies with quality scoring
    - Quality scoring: Score files by resolution, codec, source, audio, HDR
    - Health check: Validate library for issues (empty folders, missing files, etc.)
    - Codec analysis: Analyze video codecs and generate transcoding queue
    - TMDB integration: Fetch movie/show metadata from The Movie Database
    - Automatic 7-Zip installation if not present
    - Statistics summary: Detailed report of all operations performed

    NEW IN v5.0:
    - Rebranded from MediaCleaner to LibraryLint
    - Dry run report: Auto-generated HTML report showing proposed changes
    - Safe filename sanitization for Windows compatibility
    - Configuration file: Save/load settings to JSON file
    - MediaInfo integration: Accurate codec detection from file headers
    - Undo/rollback support: Manifest-based rollback of changes
    - Enhanced duplicate detection: File hashing for exact duplicates
    - Export reports: CSV, HTML, and JSON library exports
    - Retry logic: Automatic retry of failed operations with backoff
    - Verbose mode: Optional debug output with -Verbose flag

    NEW IN v5.1:
    - Modular architecture: Sync and Mirror modules
    - SFTP Sync: Download files from seedbox/remote server with WinSCP
    - Mirror Backup: Robocopy-based mirroring to external drives
    - First-run setup wizard: Guided configuration for new users
    - Library transfer: Move processed movies to main collection
    - Bracket path handling: Fixed issues with brackets in filenames
    - Improved bitrate warnings: Better detection of encodes vs source files

.PARAMETER None
    This script is interactive and prompts for all necessary inputs

.EXAMPLE
    .\LibraryLint.ps1
    Runs the script interactively, prompting for dry-run mode and folder selection

.NOTES
    Created By: Nick Kliatsko
    Last Updated: 01/10/2026
    Version: 5.1

    Requirements:
    - Windows 10 or later
    - PowerShell 5.1 or later
    - 7-Zip (will be installed automatically if not present)

    Supported Video Formats:
    - .mp4, .mkv, .avi, .mov, .wmv, .flv, .m4v

    Supported Archive Formats:
    - .rar, .zip, .7z, .tar, .gz, .bz2

    Log File Location:
    - %LOCALAPPDATA%\LibraryLint\Logs\LibraryLint_YYYYMMDD_HHMMSS.log

.LINK
    https://github.com/kliatsko/librarylint
#>

#============================================
# INITIALIZATION & CONFIGURATION
#============================================

# Script parameters for command-line usage
param(
    [string]$ConfigFile = $null,
    [switch]$Verbose,
    [switch]$NoAutoOpen  # Disable auto-opening reports in dry run mode
)

# Verbose mode flag
$script:VerboseMode = $Verbose.IsPresent

function Write-Verbose-Message {
    param([string]$Message, [string]$Color = "Magenta")
    if ($script:VerboseMode) {
        Write-Host "[DEBUG] $Message" -ForegroundColor $Color
    }
}

Write-Verbose-Message "Script starting..."
Write-Verbose-Message "Loading Windows Forms assembly..."
Add-Type -AssemblyName System.Windows.Forms
Write-Verbose-Message "Windows Forms loaded successfully"

# AppData folder for logs, config, and undo manifests (won't be uploaded to GitHub)
$script:AppDataFolder = Join-Path $env:LOCALAPPDATA "LibraryLint"
$script:LogsFolder = Join-Path $script:AppDataFolder "Logs"
$script:ReportsFolder = Join-Path $script:AppDataFolder "Reports"
$script:UndoFolder = Join-Path $script:AppDataFolder "Undo"

# Create folders if they don't exist
@($script:AppDataFolder, $script:LogsFolder, $script:ReportsFolder, $script:UndoFolder) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -Path $_ -ItemType Directory -Force | Out-Null
    }
}

# Load modules (if available)
$script:ModulesPath = Join-Path $PSScriptRoot "modules"
$script:MirrorModuleLoaded = $false
$script:SyncModuleLoaded = $false

if (Test-Path $script:ModulesPath) {
    $mirrorModule = Join-Path $script:ModulesPath "Mirror.psm1"
    $syncModule = Join-Path $script:ModulesPath "Sync.psm1"

    if (Test-Path $mirrorModule) {
        try {
            Import-Module $mirrorModule -Force -ErrorAction Stop
            $script:MirrorModuleLoaded = $true
            Write-Verbose-Message "Mirror module loaded"
        } catch {
            Write-Verbose-Message "Failed to load Mirror module: $_"
        }
    }

    if (Test-Path $syncModule) {
        try {
            Import-Module $syncModule -Force -ErrorAction Stop
            $script:SyncModuleLoaded = $true
            Write-Verbose-Message "Sync module loaded"
        } catch {
            Write-Verbose-Message "Failed to load Sync module: $_"
        }
    }
}

# Configuration file path
$script:ConfigFilePath = if ($ConfigFile) { $ConfigFile } else { Join-Path $script:AppDataFolder "LibraryLint.config.json" }

# Default configuration
$script:DefaultConfig = @{
    LogFile = Join-Path $script:LogsFolder "LibraryLint_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    DryRun = $false
    SevenZipPath = "C:\Program Files\7-Zip\7z.exe"
    MediaInfoPath = "C:\Program Files\MediaInfo\MediaInfo.exe"
    FFmpegPath = "C:\Program Files\ffmpeg\bin\ffmpeg.exe"
    VideoExtensions = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.m4v')
    SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt')
    ArchiveExtensions = @('*.rar', '*.zip', '*.7z', '*.tar', '*.gz', '*.bz2')
    ArchiveCleanupPatterns = @('*.r??', '*.zip', '*.7z', '*.tar', '*.gz', '*.bz2')
    UnnecessaryPatterns = @('*Sample*', '*Preview*', '*Proof*', '*Screens*')
    TrailerPatterns = @('*Trailer*', '*trailer*', '*TRAILER*', '*Teaser*', '*teaser*')
    PreferredSubtitleLanguages = @('eng', 'en', 'english')
    KeepSubtitles = $true
    KeepTrailers = $true
    GenerateNFO = $false
    OrganizeSeasons = $true
    RenameEpisodes = $false
    CheckDuplicates = $false
    TMDBApiKey = $null
    FanartTVApiKey = $null
    MoviesInboxPath = $null
    MoviesLibraryPath = $null  # Main movie collection path (e.g., G:\Movies)
    TVShowsInboxPath = $null
    TVShowsLibraryPath = $null  # Main TV shows collection path (e.g., G:\Shows)

    # Mirror/Backup settings
    MirrorSourceDrive = "G:"
    MirrorDestDrive = "F:"
    MirrorFolders = @("Movies", "Shows")

    # SFTP Sync settings
    SFTPHost = $null
    SFTPPort = 22
    SFTPUsername = $null
    SFTPPassword = $null
    SFTPPrivateKeyPath = $null
    SFTPRemotePaths = @("/downloads")
    SFTPDeleteAfterDownload = $false

    DownloadTrailers = $false
    TrailerQuality = "720p"  # 1080p, 720p, or 480p
    YtDlpPath = $null
    YtDlpCookieBrowser = $null  # Browser to get cookies from for age-restricted videos (chrome, firefox, edge, brave)
    DownloadSubtitles = $false
    SubtitleLanguage = "en"  # Language code for subtitles (en, es, fr, de, etc.)
    SubdlApiKey = $null
    EnableParallelProcessing = $true
    MaxParallelJobs = 4
    EnableUndo = $true
    RetryCount = 3
    RetryDelaySeconds = 2
    Tags = @(
        # Resolution tags
        '1080p', '2160p', '720p', '480p', '4K', '2K', 'UHD',
        # Video quality/source tags
        'HDRip', 'DVDRip', 'BRRip', 'BR-Rip', 'BDRip', 'BD-Rip', 'WEB-DL', 'WEBRip', 'BluRay', 'DVDR', 'DVDScr',
        # Codec tags
        'x264', 'x265', 'X265', 'H264', 'H265', 'HEVC', 'hevc-d3g', 'XviD', 'DivX', 'AVC',
        # Audio tags
        'AAC', 'AC3', 'DTS', 'Atmos', 'TrueHD', 'DD5.1', '5.1', '7.1', 'DTS-HD',
        # HDR/Color tags
        'HDR', 'HDR10', 'HDR10+', 'DolbyVision', 'SDR', '10bit', '8bit',
        # Release type tags
        'Extended', 'Unrated', 'Remastered', 'REPACK', 'PROPER', 'iNTERNAL', 'LiMiTED', 'REAL', 'HC', 'ExtCut',
        'Anniversary Edition', 'Restored', "Director's Cut", 'Theatrical', 'DUBBED', 'SUBBED',
        # Language/Region tags
        'MULTi', 'DUAL', 'ENG', 'MULTI.VFF', 'TRUEFRENCH', 'FRENCH',
        # Release group examples
        'YIFY', 'YTS', 'RARBG', 'SPARKS', 'AMRAP', 'CMRG', 'FGT', 'EVO', 'STUTTERSHIT', 'FLEET', 'ION10',
        # Misc tags
        '1080-hd4u', 'NF', 'AMZN', 'HULU', 'WEB'
    )
}

# Initialize Config from defaults
$script:Config = $script:DefaultConfig.Clone()

#============================================
# CONFIGURATION FILE FUNCTIONS
#============================================

<#
.SYNOPSIS
    Loads configuration from JSON file
.PARAMETER Path
    Path to the configuration file
.OUTPUTS
    Boolean indicating success
#>
function Import-Configuration {
    param(
        [string]$Path = $script:ConfigFilePath
    )

    if (Test-Path $Path) {
        try {
            Write-Verbose-Message "Loading configuration from: $Path"
            $jsonConfig = Get-Content -Path $Path -Raw | ConvertFrom-Json

            # Merge loaded config with defaults (loaded values override defaults)
            foreach ($property in $jsonConfig.PSObject.Properties) {
                if ($script:Config.ContainsKey($property.Name)) {
                    # Handle arrays specially
                    if ($property.Value -is [System.Array] -or $property.Value -is [System.Collections.ArrayList]) {
                        $script:Config[$property.Name] = @($property.Value)
                    } else {
                        $script:Config[$property.Name] = $property.Value
                    }
                }
            }

            # Ensure LogFile is unique for each session
            $script:Config.LogFile = Join-Path $script:LogsFolder "LibraryLint_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

            Write-Host "Configuration loaded from: $Path" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "Warning: Could not load config file: $_" -ForegroundColor Yellow
            Write-Host "Using default configuration" -ForegroundColor Cyan
            return $false
        }
    } else {
        Write-Verbose-Message "No configuration file found at: $Path"
        return $false
    }
}

<#
.SYNOPSIS
    Saves current configuration to JSON file
.PARAMETER Path
    Path to save the configuration file
#>
function Export-Configuration {
    param(
        [string]$Path = $script:ConfigFilePath
    )

    try {
        # Create a clean config object for export (exclude session-specific values)
        $exportConfig = @{}
        foreach ($key in $script:Config.Keys) {
            if ($key -ne 'LogFile') {  # Don't save session-specific log file
                $exportConfig[$key] = $script:Config[$key]
            }
        }

        $exportConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
        Write-Host "Configuration saved to: $Path" -ForegroundColor Green
        Write-Log "Configuration saved to: $Path" "INFO"
        return $true
    }
    catch {
        Write-Host "Error saving configuration: $_" -ForegroundColor Red
        Write-Log "Error saving configuration: $_" "ERROR"
        return $false
    }
}

<#
.SYNOPSIS
    Prompts user to save current configuration
#>
function Invoke-ConfigurationSavePrompt {
    $saveConfig = Read-Host "`nSave current settings as defaults? (Y/N) [N]"
    if ($saveConfig -eq 'Y' -or $saveConfig -eq 'y') {
        Export-Configuration
    }
}

# Load configuration file if it exists
Import-Configuration | Out-Null

# Global statistics tracking
$script:Stats = @{
    StartTime = $null
    EndTime = $null
    FilesDeleted = 0
    BytesDeleted = 0
    ArchivesExtracted = 0
    ArchivesFailed = 0
    FoldersCreated = 0
    FoldersRenamed = 0
    FilesMoved = 0
    EmptyFoldersRemoved = 0
    SubtitlesProcessed = 0
    SubtitlesDeleted = 0
    TrailersMoved = 0
    NFOFilesCreated = 0
    NFOFilesRead = 0
    Errors = 0
    Warnings = 0
    OperationsRetried = 0
}

# Undo manifest for rollback support
$script:UndoManifest = @{
    SessionId = [guid]::NewGuid().ToString()
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Operations = @()
}

# Failed operations queue for retry
$script:FailedOperations = @()

# Track processed path for dry run report
$script:ProcessedPath = $null
$script:ProcessedMediaType = $null
$script:NoAutoOpen = $NoAutoOpen.IsPresent

#============================================
# HELPER FUNCTIONS
#============================================

<#
.SYNOPSIS
    Writes a message to the log file with timestamp and severity level
.PARAMETER Message
    The message to log
.PARAMETER Level
    The severity level (INFO, WARNING, ERROR, DRY-RUN)
#>
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:Config.LogFile -Value $logMessage

    # Track errors and warnings in stats
    if ($Level -eq "ERROR") { $script:Stats.Errors++ }
    if ($Level -eq "WARNING") { $script:Stats.Warnings++ }
}

<#
.SYNOPSIS
    Sanitizes a string for use as a Windows filename/folder name
.DESCRIPTION
    Replaces characters that are illegal in Windows filenames with a hyphen.
    Illegal characters: \ / : * ? " < > |
.PARAMETER Name
    The string to sanitize
.OUTPUTS
    A string safe for use as a Windows filename
#>
function Get-SafeFileName {
    param([string]$Name)

    # Replace illegal Windows filename characters with hyphen
    $safe = $Name -replace '[\\/:*?"<>|]', '-'

    # Collapse multiple consecutive hyphens into one
    $safe = $safe -replace '-+', '-'

    # Remove leading/trailing hyphens and spaces
    $safe = $safe.Trim('-', ' ')

    return $safe
}

#============================================
# UNDO/ROLLBACK FUNCTIONS
#============================================

<#
.SYNOPSIS
    Records an operation to the undo manifest for potential rollback
.PARAMETER OperationType
    Type of operation (Move, Rename, Delete, Create)
.PARAMETER SourcePath
    Original path/name
.PARAMETER DestinationPath
    New path/name (if applicable)
.PARAMETER FileSize
    Size of the file (for deleted files tracking)
#>
function Add-UndoOperation {
    param(
        [string]$OperationType,
        [string]$SourcePath,
        [string]$DestinationPath = $null,
        [long]$FileSize = 0
    )

    if (-not $script:Config.EnableUndo) { return }

    $operation = @{
        Type = $OperationType
        Source = $SourcePath
        Destination = $DestinationPath
        Size = $FileSize
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    $script:UndoManifest.Operations += $operation
    Write-Verbose-Message "Recorded undo operation: $OperationType - $SourcePath"
}

<#
.SYNOPSIS
    Saves the undo manifest to a JSON file
.PARAMETER Path
    Optional path for the manifest file
#>
function Save-UndoManifest {
    param(
        [string]$Path = $null
    )

    if (-not $script:Config.EnableUndo) { return }
    if ($script:UndoManifest.Operations.Count -eq 0) {
        Write-Verbose-Message "No operations to save in undo manifest"
        return
    }

    if (-not $Path) {
        $Path = Join-Path $script:UndoFolder "LibraryLint_Undo_$($script:UndoManifest.SessionId.Substring(0,8)).json"
    }

    try {
        $script:UndoManifest | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
        Write-Host "Undo manifest saved to: $Path" -ForegroundColor Cyan
        Write-Log "Undo manifest saved with $($script:UndoManifest.Operations.Count) operations" "INFO"
    }
    catch {
        Write-Host "Warning: Could not save undo manifest: $_" -ForegroundColor Yellow
        Write-Log "Error saving undo manifest: $_" "WARNING"
    }
}

<#
.SYNOPSIS
    Loads and executes an undo manifest to rollback changes
.PARAMETER ManifestPath
    Path to the undo manifest JSON file
#>
function Invoke-UndoOperations {
    param(
        [string]$ManifestPath
    )

    if (-not (Test-Path $ManifestPath)) {
        Write-Host "Undo manifest not found: $ManifestPath" -ForegroundColor Red
        return
    }

    try {
        $manifest = Get-Content -Path $ManifestPath -Raw | ConvertFrom-Json
        Write-Host "`nUndo Manifest: $($manifest.SessionId)" -ForegroundColor Cyan
        Write-Host "Created: $($manifest.Timestamp)" -ForegroundColor Gray
        Write-Host "Operations: $($manifest.Operations.Count)" -ForegroundColor Gray

        $confirm = Read-Host "`nUndo all operations? (Y/N) [N]"
        if ($confirm -ne 'Y' -and $confirm -ne 'y') {
            Write-Host "Undo cancelled" -ForegroundColor Yellow
            return
        }

        # Process operations in reverse order
        $reversed = $manifest.Operations | Sort-Object { $_.Timestamp } -Descending
        $undone = 0
        $failed = 0

        foreach ($op in $reversed) {
            try {
                switch ($op.Type) {
                    "Move" {
                        if (Test-Path $op.Destination) {
                            Move-Item -Path $op.Destination -Destination $op.Source -Force -ErrorAction Stop
                            Write-Host "Restored: $($op.Source)" -ForegroundColor Green
                            $undone++
                        } else {
                            Write-Host "Cannot restore (destination missing): $($op.Destination)" -ForegroundColor Yellow
                            $failed++
                        }
                    }
                    "Rename" {
                        if (Test-Path $op.Destination) {
                            Rename-Item -Path $op.Destination -NewName (Split-Path $op.Source -Leaf) -Force -ErrorAction Stop
                            Write-Host "Renamed back: $(Split-Path $op.Source -Leaf)" -ForegroundColor Green
                            $undone++
                        } else {
                            Write-Host "Cannot restore (file missing): $($op.Destination)" -ForegroundColor Yellow
                            $failed++
                        }
                    }
                    "Create" {
                        if (Test-Path $op.Source) {
                            Remove-Item -Path $op.Source -Recurse -Force -ErrorAction Stop
                            Write-Host "Removed created item: $($op.Source)" -ForegroundColor Green
                            $undone++
                        }
                    }
                    "Delete" {
                        Write-Host "Cannot restore deleted file: $($op.Source)" -ForegroundColor Yellow
                        $failed++
                    }
                }
            }
            catch {
                Write-Host "Failed to undo: $($op.Source) - $_" -ForegroundColor Red
                $failed++
            }
        }

        Write-Host "`nUndo complete: $undone restored, $failed failed" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Error processing undo manifest: $_" -ForegroundColor Red
    }
}

#============================================
# RETRY LOGIC FUNCTIONS
#============================================

<#
.SYNOPSIS
    Adds a failed operation to the retry queue
.PARAMETER OperationType
    Type of operation that failed
.PARAMETER Parameters
    Hashtable of parameters for the operation
.PARAMETER ErrorMessage
    The error message from the failure
#>
function Add-FailedOperation {
    param(
        [string]$OperationType,
        [hashtable]$Parameters,
        [string]$ErrorMessage
    )

    $script:FailedOperations += @{
        Type = $OperationType
        Parameters = $Parameters
        Error = $ErrorMessage
        RetryCount = 0
        Timestamp = Get-Date
    }
}

<#
.SYNOPSIS
    Retries all failed operations
.DESCRIPTION
    Attempts to retry failed operations with exponential backoff
#>
function Invoke-RetryFailedOperations {
    if ($script:FailedOperations.Count -eq 0) {
        return
    }

    Write-Host "`nRetrying $($script:FailedOperations.Count) failed operation(s)..." -ForegroundColor Yellow
    Write-Log "Starting retry of $($script:FailedOperations.Count) failed operations" "INFO"

    $stillFailed = @()

    foreach ($op in $script:FailedOperations) {
        if ($op.RetryCount -ge $script:Config.RetryCount) {
            Write-Host "Max retries reached for: $($op.Type) - $($op.Parameters.Path)" -ForegroundColor Red
            $stillFailed += $op
            continue
        }

        $op.RetryCount++
        $delay = $script:Config.RetryDelaySeconds * [math]::Pow(2, $op.RetryCount - 1)
        Write-Host "Retry $($op.RetryCount)/$($script:Config.RetryCount) (waiting ${delay}s): $($op.Type)" -ForegroundColor Cyan

        Start-Sleep -Seconds $delay

        try {
            switch ($op.Type) {
                "Move" {
                    Move-Item -Path $op.Parameters.Source -Destination $op.Parameters.Destination -Force -ErrorAction Stop
                    Write-Host "Retry successful: Move $($op.Parameters.Source)" -ForegroundColor Green
                    $script:Stats.OperationsRetried++
                }
                "Delete" {
                    Remove-Item -Path $op.Parameters.Path -Recurse -Force -ErrorAction Stop
                    Write-Host "Retry successful: Delete $($op.Parameters.Path)" -ForegroundColor Green
                    $script:Stats.OperationsRetried++
                }
                "Rename" {
                    Rename-Item -Path $op.Parameters.Path -NewName $op.Parameters.NewName -Force -ErrorAction Stop
                    Write-Host "Retry successful: Rename $($op.Parameters.Path)" -ForegroundColor Green
                    $script:Stats.OperationsRetried++
                }
                "Extract" {
                    # Re-attempt extraction
                    $process = Start-Process -FilePath $script:Config.SevenZipPath `
                        -ArgumentList "x", "-o`"$($op.Parameters.ExtractPath)`"", "`"$($op.Parameters.ArchivePath)`"", "-r", "-y" `
                        -NoNewWindow -PassThru -Wait
                    if ($process.ExitCode -eq 0) {
                        Write-Host "Retry successful: Extract $($op.Parameters.ArchivePath)" -ForegroundColor Green
                        $script:Stats.OperationsRetried++
                    } else {
                        throw "Extraction failed with exit code: $($process.ExitCode)"
                    }
                }
            }
        }
        catch {
            Write-Host "Retry failed: $_" -ForegroundColor Yellow
            $op.Error = $_.ToString()
            $stillFailed += $op
        }
    }

    $script:FailedOperations = $stillFailed

    if ($stillFailed.Count -gt 0) {
        Write-Host "$($stillFailed.Count) operation(s) still failed after retries" -ForegroundColor Yellow
        Write-Log "$($stillFailed.Count) operations failed after all retries" "WARNING"
    } else {
        Write-Host "All retry operations completed successfully" -ForegroundColor Green
    }
}

#============================================
# MEDIAINFO INTEGRATION
#============================================

<#
.SYNOPSIS
    Checks if MediaInfo CLI is installed
.OUTPUTS
    Boolean indicating if MediaInfo is available
#>
function Test-MediaInfoInstallation {
    if (Test-Path $script:Config.MediaInfoPath) {
        Write-Verbose-Message "MediaInfo found at: $($script:Config.MediaInfoPath)"
        return $true
    }

    # Try to find in PATH
    $mediaInfoInPath = Get-Command "mediainfo" -ErrorAction SilentlyContinue
    if ($mediaInfoInPath) {
        $script:Config.MediaInfoPath = $mediaInfoInPath.Source
        Write-Verbose-Message "MediaInfo found in PATH: $($script:Config.MediaInfoPath)"
        return $true
    }

    Write-Verbose-Message "MediaInfo not found - using filename parsing for codec detection"
    return $false
}

<#
.SYNOPSIS
    Checks if FFmpeg is installed and available
.OUTPUTS
    Boolean indicating if FFmpeg is available
#>
function Test-FFmpegInstallation {
    if (Test-Path $script:Config.FFmpegPath) {
        Write-Verbose-Message "FFmpeg found at: $($script:Config.FFmpegPath)"
        return $true
    }

    # Try to find in PATH
    $ffmpegInPath = Get-Command "ffmpeg" -ErrorAction SilentlyContinue
    if ($ffmpegInPath) {
        $script:Config.FFmpegPath = $ffmpegInPath.Source
        Write-Verbose-Message "FFmpeg found in PATH: $($script:Config.FFmpegPath)"
        return $true
    }

    # Check common installation locations
    $commonPaths = @(
        "C:\ffmpeg\bin\ffmpeg.exe",
        "C:\Program Files\ffmpeg\bin\ffmpeg.exe",
        "C:\Program Files (x86)\ffmpeg\bin\ffmpeg.exe",
        "$env:USERPROFILE\ffmpeg\bin\ffmpeg.exe",
        "$env:LOCALAPPDATA\ffmpeg\bin\ffmpeg.exe"
    )

    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            $script:Config.FFmpegPath = $path
            Write-Verbose-Message "FFmpeg found at: $path"
            return $true
        }
    }

    Write-Verbose-Message "FFmpeg not found"
    return $false
}

<#
.SYNOPSIS
    Checks if yt-dlp is installed and accessible
.OUTPUTS
    Boolean indicating if yt-dlp is available
#>
function Test-YtDlpInstallation {
    if ($script:Config.YtDlpPath -and (Test-Path $script:Config.YtDlpPath)) {
        Write-Verbose-Message "yt-dlp found at: $($script:Config.YtDlpPath)"
        return $true
    }

    # Try to find in PATH
    $ytdlpInPath = Get-Command "yt-dlp" -ErrorAction SilentlyContinue
    if ($ytdlpInPath) {
        $script:Config.YtDlpPath = $ytdlpInPath.Source
        Write-Verbose-Message "yt-dlp found in PATH: $($script:Config.YtDlpPath)"
        return $true
    }

    # Check common installation locations
    $commonPaths = @(
        "C:\yt-dlp\yt-dlp.exe",
        "C:\Program Files\yt-dlp\yt-dlp.exe",
        "C:\Program Files (x86)\yt-dlp\yt-dlp.exe",
        "$env:USERPROFILE\yt-dlp\yt-dlp.exe",
        "$env:LOCALAPPDATA\yt-dlp\yt-dlp.exe",
        "$env:LOCALAPPDATA\Programs\yt-dlp\yt-dlp.exe",
        "$env:APPDATA\yt-dlp\yt-dlp.exe"
    )

    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            $script:Config.YtDlpPath = $path
            Write-Verbose-Message "yt-dlp found at: $path"
            return $true
        }
    }

    Write-Verbose-Message "yt-dlp not found"
    return $false
}

<#
.SYNOPSIS
    Gets detailed video information using MediaInfo CLI
.PARAMETER FilePath
    Path to the video file
.OUTPUTS
    Hashtable with detailed video properties
#>
function Get-MediaInfoDetails {
    param(
        [string]$FilePath
    )

    $info = @{
        VideoCodec = "Unknown"
        AudioCodec = "Unknown"
        Resolution = "Unknown"
        Width = 0
        Height = 0
        Duration = 0
        Bitrate = 0
        FrameRate = 0
        HDR = $false
        AudioChannels = 0
        Container = "Unknown"
    }

    if (-not (Test-MediaInfoInstallation)) {
        return $null
    }

    try {
        # Use JSON output for more reliable parsing - single call gets everything
        $jsonOutput = & $script:Config.MediaInfoPath --Output=JSON "$FilePath" 2>$null

        if (-not $jsonOutput) {
            Write-Log "MediaInfo returned no output for: $FilePath" "WARNING"
            return $null
        }

        $mediaData = $jsonOutput | ConvertFrom-Json

        if (-not $mediaData.media -or -not $mediaData.media.track) {
            Write-Log "MediaInfo JSON has no media/track data for: $FilePath" "WARNING"
            return $null
        }

        # Get General track info
        $generalTrack = $mediaData.media.track | Where-Object { $_.'@type' -eq 'General' } | Select-Object -First 1
        if ($generalTrack) {
            if ($generalTrack.Duration) { $info.Duration = [long]([double]$generalTrack.Duration * 1000) }
            if ($generalTrack.OverallBitRate) { $info.Bitrate = [long]$generalTrack.OverallBitRate }
            if ($generalTrack.Format) { $info.Container = $generalTrack.Format }
        }

        # Get Video track info
        $videoTrack = $mediaData.media.track | Where-Object { $_.'@type' -eq 'Video' } | Select-Object -First 1
        if (-not $videoTrack) {
            Write-Log "MediaInfo found no video track in: $FilePath" "WARNING"
        } else {
            if ($videoTrack.Format) { $info.VideoCodec = $videoTrack.Format }
            if ($videoTrack.Width) { $info.Width = [int]$videoTrack.Width }
            if ($videoTrack.Height) { $info.Height = [int]$videoTrack.Height }
            if ($videoTrack.FrameRate) { $info.FrameRate = [double]$videoTrack.FrameRate }

            # Log if we got a video track but no dimensions
            if ($info.Width -eq 0 -or $info.Height -eq 0) {
                Write-Log "MediaInfo video track missing dimensions for: $FilePath (Width=$($videoTrack.Width), Height=$($videoTrack.Height))" "WARNING"
            }

            # Check for HDR
            $hdrFormat = $videoTrack.HDR_Format
            $colorSpace = $videoTrack.colour_primaries
            if ($hdrFormat -or $colorSpace -match 'BT.2020') {
                $info.HDR = $true
            }
        }

        # Get Audio track info
        $audioTrack = $mediaData.media.track | Where-Object { $_.'@type' -eq 'Audio' } | Select-Object -First 1
        if ($audioTrack) {
            if ($audioTrack.Format) { $info.AudioCodec = $audioTrack.Format }
            if ($audioTrack.Channels) { $info.AudioChannels = [int]$audioTrack.Channels }
        }

        # Determine resolution label
        if ($info.Height -ge 2160) { $info.Resolution = "2160p" }
        elseif ($info.Height -ge 1080) { $info.Resolution = "1080p" }
        elseif ($info.Height -ge 720) { $info.Resolution = "720p" }
        elseif ($info.Height -ge 480) { $info.Resolution = "480p" }
        elseif ($info.Height -gt 0) { $info.Resolution = "$($info.Height)p" }
        else {
            Write-Log "MediaInfo could not determine resolution for: $FilePath" "WARNING"
        }

        Write-Verbose-Message "MediaInfo: $($info.Resolution) $($info.VideoCodec) $($info.AudioCodec)"
        return $info
    }
    catch {
        Write-Log "Error getting MediaInfo for $FilePath : $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Gets detailed HDR format information using MediaInfo CLI
.PARAMETER FilePath
    Path to the video file
.OUTPUTS
    Hashtable with HDR format details (Format, Profile, Compatibility)
.DESCRIPTION
    Extracts detailed HDR metadata including Dolby Vision profile,
    HDR10+ dynamic metadata, HLG, and compatibility layers
#>
function Get-MediaInfoHDRFormat {
    param(
        [string]$FilePath
    )

    if (-not (Test-MediaInfoInstallation)) {
        return $null
    }

    try {
        $hdrFormat = & $script:Config.MediaInfoPath --Inform="Video;%HDR_Format%" "$FilePath" 2>$null
        $hdrFormatCompat = & $script:Config.MediaInfoPath --Inform="Video;%HDR_Format_Compatibility%" "$FilePath" 2>$null
        $hdrFormatProfile = & $script:Config.MediaInfoPath --Inform="Video;%HDR_Format_Profile%" "$FilePath" 2>$null
        $colorPrimaries = & $script:Config.MediaInfoPath --Inform="Video;%colour_primaries%" "$FilePath" 2>$null
        $transferChar = & $script:Config.MediaInfoPath --Inform="Video;%transfer_characteristics%" "$FilePath" 2>$null

        $result = @{
            Format = $null
            Profile = $null
            Compatibility = $null
            ColorPrimaries = $null
            TransferCharacteristics = $null
        }

        if ($hdrFormat) { $result.Format = $hdrFormat.Trim() }
        if ($hdrFormatProfile) { $result.Profile = $hdrFormatProfile.Trim() }
        if ($hdrFormatCompat) { $result.Compatibility = $hdrFormatCompat.Trim() }
        if ($colorPrimaries) { $result.ColorPrimaries = $colorPrimaries.Trim() }
        if ($transferChar) { $result.TransferCharacteristics = $transferChar.Trim() }

        # Normalize HDR format names
        if ($result.Format) {
            if ($result.Format -match 'Dolby Vision|DOVI') {
                $result.Format = "Dolby Vision"
            }
            elseif ($result.Format -match 'HDR10\+|SMPTE ST 2094') {
                $result.Format = "HDR10+"
            }
            elseif ($result.Format -match 'SMPTE ST 2086|HDR10') {
                $result.Format = "HDR10"
            }
            elseif ($result.Format -match 'HLG|ARIB STD-B67') {
                $result.Format = "HLG"
            }
        }
        # Detect HDR from transfer characteristics if HDR_Format not present
        elseif ($transferChar -match 'PQ|SMPTE ST 2084') {
            $result.Format = "HDR10"
        }
        elseif ($transferChar -match 'HLG') {
            $result.Format = "HLG"
        }
        # Check for BT.2020 color space (wide color gamut, often indicates HDR)
        elseif ($colorPrimaries -match 'BT\.2020') {
            $result.Format = "HDR"
        }

        if ($result.Format) {
            return $result
        }
        return $null
    }
    catch {
        Write-Verbose-Message "Error getting HDR format: $_" "Yellow"
        return $null
    }
}

#============================================
# IMPROVED DUPLICATE DETECTION WITH HASHING
#============================================

<#
.SYNOPSIS
    Calculates a partial hash of a file for duplicate detection
.PARAMETER FilePath
    Path to the file
.PARAMETER SampleSize
    Number of bytes to sample from start and end (default 1MB each)
.OUTPUTS
    String hash value
.DESCRIPTION
    Instead of hashing the entire file (slow for large videos),
    this samples the first and last portions for a quick fingerprint
#>
function Get-FilePartialHash {
    param(
        [string]$FilePath,
        [int]$SampleSize = 1MB
    )

    try {
        $file = [System.IO.File]::OpenRead($FilePath)
        $fileLength = $file.Length

        # For small files, hash the entire file
        if ($fileLength -le ($SampleSize * 2)) {
            $file.Close()
            return (Get-FileHash -Path $FilePath -Algorithm MD5).Hash
        }

        $hasher = [System.Security.Cryptography.MD5]::Create()

        # Read first chunk
        $buffer = New-Object byte[] $SampleSize
        $file.Read($buffer, 0, $SampleSize) | Out-Null

        # Seek to end and read last chunk
        $file.Seek(-$SampleSize, [System.IO.SeekOrigin]::End) | Out-Null
        $endBuffer = New-Object byte[] $SampleSize
        $file.Read($endBuffer, 0, $SampleSize) | Out-Null

        $file.Close()

        # Combine buffers and hash
        $combined = $buffer + $endBuffer + [System.BitConverter]::GetBytes($fileLength)
        $hashBytes = $hasher.ComputeHash($combined)
        $hash = [System.BitConverter]::ToString($hashBytes) -replace '-', ''

        return $hash
    }
    catch {
        Write-Verbose-Message "Error calculating hash for $FilePath : $_" "Yellow"
        return $null
    }
}

<#
.SYNOPSIS
    Enhanced duplicate detection using NFO metadata, file hashing, and title matching
.PARAMETER Path
    Root path of the media library
.OUTPUTS
    Array of duplicate groups with match information and confidence levels
.DESCRIPTION
    Detects duplicates using multiple strategies:
    1. IMDB/TMDB ID matching (highest confidence - same movie guaranteed)
    2. File hash matching (exact same file content)
    3. Title + Year matching (high confidence)
    4. Title-only matching with similar file sizes (lower confidence)
#>
function Find-DuplicateMoviesEnhanced {
    param(
        [string]$Path
    )

    Write-Host "`nScanning for duplicate movies (enhanced)..." -ForegroundColor Yellow
    Write-Host "This scan checks: NFO metadata (IMDB/TMDB IDs), file hashes, and titles" -ForegroundColor Gray
    Write-Log "Starting enhanced duplicate scan in: $Path" "INFO"

    try {
        $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_Trailers' }

        if ($movieFolders.Count -eq 0) {
            Write-Host "No movie folders found" -ForegroundColor Cyan
            return @()
        }

        # Build lookups for different matching strategies
        $imdbLookup = @{}      # Key: IMDB ID
        $tmdbLookup = @{}      # Key: TMDB ID
        $hashLookup = @{}      # Key: File hash
        $titleYearLookup = @{} # Key: normalized title|year
        $titleOnlyLookup = @{} # Key: normalized title (for movies without years)

        $totalFolders = $movieFolders.Count
        $currentIndex = 0
        $nfoCount = 0

        foreach ($folder in $movieFolders) {
            $currentIndex++
            $percentComplete = [math]::Round(($currentIndex / $totalFolders) * 100)
            Write-Progress -Activity "Scanning for duplicates" -Status "Processing $currentIndex of $totalFolders - $($folder.Name)" -PercentComplete $percentComplete

            # Find the main video file
            $videoFile = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                Sort-Object Length -Descending |
                Select-Object -First 1

            if (-not $videoFile) { continue }

            # Get title info from folder name
            $titleInfo = Get-NormalizedTitle -Name $folder.Name
            $quality = Get-QualityScore -FileName $folder.Name -FilePath $videoFile.FullName

            # Try to read NFO file for IMDB/TMDB IDs
            $nfoFile = Get-ChildItem -LiteralPath $folder.FullName -Filter "*.nfo" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            $nfoMetadata = $null
            $imdbId = $null
            $tmdbId = $null
            $nfoYear = $null
            $nfoTitle = $null

            if ($nfoFile) {
                $nfoMetadata = Read-NFOFile -NfoPath $nfoFile.FullName
                if ($nfoMetadata) {
                    $nfoCount++
                    $imdbId = $nfoMetadata.IMDBID
                    $tmdbId = $nfoMetadata.TMDBID
                    $nfoYear = $nfoMetadata.Year
                    $nfoTitle = $nfoMetadata.Title
                }
            }

            # Use NFO year if folder name doesn't have one
            $finalYear = if ($titleInfo.Year) { $titleInfo.Year } elseif ($nfoYear) { $nfoYear } else { $null }

            # Calculate partial hash for exact duplicate detection
            $fileHash = Get-FilePartialHash -FilePath $videoFile.FullName

            $entry = @{
                Path = $folder.FullName
                OriginalName = $folder.Name
                Year = $finalYear
                Quality = $quality
                FileSize = $videoFile.Length
                FilePath = $videoFile.FullName
                FileHash = $fileHash
                IMDBID = $imdbId
                TMDBID = $tmdbId
                NFOTitle = $nfoTitle
                MatchType = @()      # Will track how this was matched
                Confidence = $null   # High, Medium, Low
            }

            # Add to IMDB lookup (most reliable)
            if ($imdbId) {
                if (-not $imdbLookup.ContainsKey($imdbId)) {
                    $imdbLookup[$imdbId] = @()
                }
                $imdbLookup[$imdbId] += $entry
            }

            # Add to TMDB lookup
            if ($tmdbId) {
                if (-not $tmdbLookup.ContainsKey($tmdbId)) {
                    $tmdbLookup[$tmdbId] = @()
                }
                $tmdbLookup[$tmdbId] += $entry
            }

            # Add to hash lookup
            if ($fileHash) {
                if (-not $hashLookup.ContainsKey($fileHash)) {
                    $hashLookup[$fileHash] = @()
                }
                $hashLookup[$fileHash] += $entry
            }

            # Add to title+year lookup (use NFO title if available for better matching)
            $normalizedTitle = $titleInfo.NormalizedTitle
            if ($nfoTitle) {
                $nfoTitleInfo = Get-NormalizedTitle -Name $nfoTitle
                $normalizedTitle = $nfoTitleInfo.NormalizedTitle
            }

            if ($finalYear) {
                $titleYearKey = "$normalizedTitle|$finalYear"
                if (-not $titleYearLookup.ContainsKey($titleYearKey)) {
                    $titleYearLookup[$titleYearKey] = @()
                }
                $titleYearLookup[$titleYearKey] += $entry
            } else {
                # No year - add to title-only lookup
                if (-not $titleOnlyLookup.ContainsKey($normalizedTitle)) {
                    $titleOnlyLookup[$normalizedTitle] = @()
                }
                $titleOnlyLookup[$normalizedTitle] += $entry
            }
        }

        Write-Progress -Activity "Scanning for duplicates" -Completed
        Write-Host "Scanned $totalFolders folders ($nfoCount with NFO files)" -ForegroundColor Gray

        # Find duplicates with priority ordering
        $duplicates = @()
        $processedPaths = @{}

        # 1. IMDB ID matches (highest confidence)
        foreach ($id in $imdbLookup.Keys) {
            if ($imdbLookup[$id].Count -gt 1) {
                $group = $imdbLookup[$id]
                foreach ($entry in $group) {
                    $entry.MatchType += "IMDB"
                    $entry.Confidence = "High"
                    $processedPaths[$entry.Path] = $true
                }
                $duplicates += ,@($group)
                Write-Log "IMDB duplicate group: $id - $($group.Count) entries: $(($group | ForEach-Object { $_.OriginalName }) -join ', ')" "INFO"
            }
        }

        # 2. TMDB ID matches (high confidence, skip if already matched by IMDB)
        foreach ($id in $tmdbLookup.Keys) {
            $unprocessed = @($tmdbLookup[$id] | Where-Object { -not $processedPaths.ContainsKey($_.Path) })
            if ($unprocessed.Count -gt 1) {
                foreach ($entry in $unprocessed) {
                    $entry.MatchType += "TMDB"
                    $entry.Confidence = "High"
                    $processedPaths[$entry.Path] = $true
                }
                $duplicates += ,@($unprocessed)
                Write-Log "TMDB duplicate group: $id - $($unprocessed.Count) entries: $(($unprocessed | ForEach-Object { $_.OriginalName }) -join ', ')" "INFO"
            }
        }

        # 3. Exact file hash matches (exact duplicate files)
        foreach ($hash in $hashLookup.Keys) {
            $unprocessed = @($hashLookup[$hash] | Where-Object { -not $processedPaths.ContainsKey($_.Path) })
            if ($unprocessed.Count -gt 1) {
                foreach ($entry in $unprocessed) {
                    $entry.MatchType += "ExactHash"
                    $entry.Confidence = "High"
                    $processedPaths[$entry.Path] = $true
                }
                $duplicates += ,@($unprocessed)
                Write-Log "Hash duplicate group: $($hash.Substring(0,8))... - $($unprocessed.Count) entries: $(($unprocessed | ForEach-Object { $_.OriginalName }) -join ', ')" "INFO"
            }
        }

        # 4. Title + Year matches (medium confidence)
        foreach ($key in $titleYearLookup.Keys) {
            $unprocessed = @($titleYearLookup[$key] | Where-Object { -not $processedPaths.ContainsKey($_.Path) })
            if ($unprocessed.Count -gt 1) {
                foreach ($entry in $unprocessed) {
                    $entry.MatchType += "TitleYear"
                    $entry.Confidence = "Medium"
                    $processedPaths[$entry.Path] = $true
                }
                $duplicates += ,@($unprocessed)
                Write-Log "Title+Year duplicate group: $key - $($unprocessed.Count) entries: $(($unprocessed | ForEach-Object { $_.OriginalName }) -join ', ')" "INFO"
            }
        }

        # 5. Title-only matches with similar file sizes (low confidence - needs review)
        foreach ($title in $titleOnlyLookup.Keys) {
            $unprocessed = @($titleOnlyLookup[$title] | Where-Object { -not $processedPaths.ContainsKey($_.Path) })
            if ($unprocessed.Count -gt 1) {
                # Only flag as duplicates if file sizes are within 20% of each other
                $avgSize = ($unprocessed | Measure-Object -Property FileSize -Average).Average
                $similarSizeEntries = @($unprocessed | Where-Object {
                    $avgSize -eq 0 -or ([math]::Abs($_.FileSize - $avgSize) / $avgSize) -lt 0.2
                })

                if ($similarSizeEntries.Count -gt 1) {
                    foreach ($entry in $similarSizeEntries) {
                        $entry.MatchType += "TitleOnly"
                        $entry.Confidence = "Low"
                        $processedPaths[$entry.Path] = $true
                    }
                    $duplicates += ,@($similarSizeEntries)
                    Write-Log "Title-only duplicate group: $title - $($similarSizeEntries.Count) entries: $(($similarSizeEntries | ForEach-Object { $_.OriginalName }) -join ', ')" "INFO"
                }
            }
        }

        Write-Log "Enhanced duplicate scan found $($duplicates.Count) duplicate groups" "INFO"

        # Summary by confidence
        $highConf = ($duplicates | Where-Object { $_[0].Confidence -eq "High" }).Count
        $medConf = ($duplicates | Where-Object { $_[0].Confidence -eq "Medium" }).Count
        $lowConf = ($duplicates | Where-Object { $_[0].Confidence -eq "Low" }).Count
        Write-Host "`nDuplicate summary: $highConf high-confidence, $medConf medium, $lowConf low (needs review)" -ForegroundColor Yellow

        return $duplicates
    }
    catch {
        Write-Host "Error scanning for duplicates: $_" -ForegroundColor Red
        Write-Log "Error in enhanced duplicate scan: $_" "ERROR"
        return @()
    }
}

<#
.SYNOPSIS
    Enhanced duplicate detection with interactive cleanup options
.PARAMETER Path
    Root path of the media library
.DESCRIPTION
    Scans for duplicates using NFO metadata (IMDB/TMDB IDs), file hashes, and title matching.
    Displays a detailed report with confidence levels, then prompts the user for action:
    - Move duplicates to review folder (safe - nothing deleted)
    - Automatically delete duplicates (keeps best quality)
    - Export report to CSV
    - Do nothing (report only)
#>
function Invoke-EnhancedDuplicateDetection {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    # Find duplicates
    $duplicates = Find-DuplicateMoviesEnhanced -Path $Path

    if ($duplicates.Count -eq 0) {
        Write-Host "`nNo duplicate movies found - your library is clean!" -ForegroundColor Green
        Write-Log "No duplicates found" "INFO"
        return
    }

    # Sort by confidence (High first, then Medium, then Low)
    $sortedDuplicates = $duplicates | Sort-Object {
        switch ($_[0].Confidence) {
            "High" { 0 }
            "Medium" { 1 }
            "Low" { 2 }
            default { 3 }
        }
    }

    # Calculate summary stats
    $highConf = @($duplicates | Where-Object { $_[0].Confidence -eq "High" }).Count
    $medConf = @($duplicates | Where-Object { $_[0].Confidence -eq "Medium" }).Count
    $lowConf = @($duplicates | Where-Object { $_[0].Confidence -eq "Low" }).Count

    # Calculate potential space savings
    $totalSavings = 0
    foreach ($group in $duplicates) {
        $sorted = $group | Sort-Object {
            $hasYear = if ($_.OriginalName -match '\(\d{4}\)') { 1 } else { 0 }
            ($_.Quality.Score * 1000) + ($hasYear * 100) + ($_.FileSize / 1MB)
        } -Descending
        # Sum up sizes of all but the best quality
        $sorted | Select-Object -Skip 1 | ForEach-Object { $totalSavings += $_.FileSize }
    }

    # Display report header
    Write-Host "`n" -NoNewline
    Write-Host "╔══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║              ENHANCED DUPLICATE DETECTION REPORT                     ║" -ForegroundColor Yellow
    Write-Host "╠══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
    Write-Host "║  Total duplicate groups: $($duplicates.Count.ToString().PadRight(44))║" -ForegroundColor White
    Write-Host "║  - High confidence (IMDB/TMDB/Hash): $($highConf.ToString().PadRight(32))║" -ForegroundColor Green
    Write-Host "║  - Medium confidence (Title+Year):  $($medConf.ToString().PadRight(32))║" -ForegroundColor Cyan
    Write-Host "║  - Low confidence (Title only):     $($lowConf.ToString().PadRight(32))║" -ForegroundColor DarkYellow
    Write-Host "╠══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
    Write-Host "║  Potential space savings: $((Format-FileSize $totalSavings).PadRight(43))║" -ForegroundColor Magenta
    Write-Host "╚══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow

    # Display duplicate groups
    $groupNum = 1
    $currentConfidence = $null

    foreach ($group in $sortedDuplicates) {
        $confidence = $group[0].Confidence
        $matchTypes = ($group | ForEach-Object { $_.MatchType } | Sort-Object -Unique) -join ", "

        # Print section header when confidence level changes
        if ($confidence -ne $currentConfidence) {
            $currentConfidence = $confidence
            $confColor = switch ($confidence) {
                "High" { "Green" }
                "Medium" { "Cyan" }
                "Low" { "DarkYellow" }
                default { "Gray" }
            }
            Write-Host "`n=== $confidence Confidence Duplicates ===" -ForegroundColor $confColor
            switch ($confidence) {
                "High" { Write-Host "    (Same IMDB/TMDB ID or identical file - definitely duplicates)" -ForegroundColor DarkGray }
                "Medium" { Write-Host "    (Same title and year - very likely duplicates)" -ForegroundColor DarkGray }
                "Low" { Write-Host "    (Similar title, no year - review manually)" -ForegroundColor DarkGray }
            }
        }

        $confColor = switch ($confidence) {
            "High" { "Green" }
            "Medium" { "Cyan" }
            "Low" { "DarkYellow" }
            default { "Gray" }
        }

        Write-Host "`n[$groupNum] " -NoNewline -ForegroundColor White
        Write-Host "[$confidence] " -NoNewline -ForegroundColor $confColor
        Write-Host "Match: $matchTypes" -ForegroundColor Gray

        # Show IMDB/TMDB ID if available
        $sampleEntry = $group[0]
        if ($sampleEntry.IMDBID) {
            Write-Host "    IMDB: $($sampleEntry.IMDBID)" -ForegroundColor DarkCyan
        }

        # Sort by quality - best first
        $sorted = $group | Sort-Object {
            $hasYear = if ($_.OriginalName -match '\(\d{4}\)') { 1 } else { 0 }
            ($_.Quality.Score * 1000) + ($hasYear * 100) + ($_.FileSize / 1MB)
        } -Descending

        $first = $true
        foreach ($movie in $sorted) {
            $sizeStr = Format-FileSize $movie.FileSize
            $scoreStr = "Score: $($movie.Quality.Score)"
            $hashStr = if ($movie.FileHash) { $movie.FileHash.Substring(0, 8) + "..." } else { "N/A" }

            if ($first) {
                Write-Host "  [KEEP] " -ForegroundColor Green -NoNewline
                $first = $false
            } else {
                Write-Host "  [DEL?] " -ForegroundColor Red -NoNewline
            }

            Write-Host "$($movie.OriginalName)" -ForegroundColor White
            Write-Host "         $scoreStr | $sizeStr | Hash: $hashStr" -ForegroundColor Gray

            # Show quality details
            $qualityDetails = @()
            if ($movie.Quality.Resolution) { $qualityDetails += $movie.Quality.Resolution }
            if ($movie.Quality.Source) { $qualityDetails += $movie.Quality.Source }
            if ($movie.Quality.Codec) { $qualityDetails += $movie.Quality.Codec }
            if ($qualityDetails.Count -gt 0) {
                Write-Host "         $($qualityDetails -join ' | ')" -ForegroundColor DarkGray
            }
        }

        $groupNum++
    }

    Write-Log "Enhanced duplicate report: $($duplicates.Count) groups found (High: $highConf, Medium: $medConf, Low: $lowConf)" "INFO"

    # Prompt for action
    Write-Host "`n" -NoNewline
    Write-Host "╔══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                        WHAT WOULD YOU LIKE TO DO?                    ║" -ForegroundColor Cyan
    Write-Host "╠══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
    Write-Host "║  1. Move duplicates to review folder (safe - nothing deleted)        ║" -ForegroundColor White
    Write-Host "║  2. Auto-delete duplicates (HIGH confidence only)                    ║" -ForegroundColor Yellow
    Write-Host "║  3. Auto-delete ALL duplicates (use with caution!)                   ║" -ForegroundColor Red
    Write-Host "║  4. Export report to CSV                                             ║" -ForegroundColor White
    Write-Host "║  5. Do nothing (report only)                                         ║" -ForegroundColor Gray
    Write-Host "╚══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    $action = Read-Host "`nSelect action (1-5)"

    switch ($action) {
        "1" {
            # Move to review folder
            Write-Host "`n--- Moving Duplicates to Review Folder ---" -ForegroundColor Yellow
            Write-Host "Lower-quality duplicates will be moved to: $Path\_Duplicates_Review" -ForegroundColor Gray
            Write-Host "You can review them there and delete manually when ready." -ForegroundColor Gray

            $confirm = Read-Host "`nProceed? (Y/N)"
            if ($confirm -match '^[Yy]') {
                Invoke-DuplicateMove -Path $Path -Duplicates $duplicates -ReviewFolderName "_Duplicates_Review"
            } else {
                Write-Host "Cancelled." -ForegroundColor Yellow
            }
        }
        "2" {
            # Auto-delete HIGH confidence only
            $highConfDupes = @($duplicates | Where-Object { $_[0].Confidence -eq "High" })
            if ($highConfDupes.Count -eq 0) {
                Write-Host "`nNo high-confidence duplicates to delete." -ForegroundColor Yellow
                return
            }

            Write-Host "`n--- Auto-Delete HIGH Confidence Duplicates ---" -ForegroundColor Yellow
            Write-Host "This will PERMANENTLY DELETE $($highConfDupes.Count) duplicate groups." -ForegroundColor Red
            Write-Host "Only high-confidence duplicates (same IMDB/TMDB ID or file hash) will be deleted." -ForegroundColor Gray
            Write-Host "The BEST quality version of each movie will be kept." -ForegroundColor Green

            $confirm = Read-Host "`nType 'DELETE' to confirm (or anything else to cancel)"
            if ($confirm -eq 'DELETE') {
                Invoke-DuplicateDelete -Path $Path -Duplicates $highConfDupes
            } else {
                Write-Host "Cancelled." -ForegroundColor Yellow
            }
        }
        "3" {
            # Auto-delete ALL
            Write-Host "`n--- Auto-Delete ALL Duplicates ---" -ForegroundColor Red
            Write-Host "WARNING: This will PERMANENTLY DELETE duplicates from ALL confidence levels!" -ForegroundColor Red
            Write-Host "This includes LOW confidence matches that may not be true duplicates." -ForegroundColor Yellow
            Write-Host "Total: $($duplicates.Count) groups | Space savings: $(Format-FileSize $totalSavings)" -ForegroundColor Gray

            $confirm = Read-Host "`nType 'DELETE ALL' to confirm (or anything else to cancel)"
            if ($confirm -eq 'DELETE ALL') {
                Invoke-DuplicateDelete -Path $Path -Duplicates $duplicates
            } else {
                Write-Host "Cancelled." -ForegroundColor Yellow
            }
        }
        "4" {
            # Export to CSV
            $csvPath = Join-Path $Path "DuplicateReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            Export-DuplicateReport -Duplicates $duplicates -OutputPath $csvPath
            Write-Host "`nReport exported to: $csvPath" -ForegroundColor Green
        }
        "5" {
            Write-Host "`nNo action taken. Report complete." -ForegroundColor Gray
        }
        default {
            Write-Host "`nInvalid selection. No action taken." -ForegroundColor Yellow
        }
    }
}

<#
.SYNOPSIS
    Moves duplicate movies to a review folder
.PARAMETER Path
    Root path of the movie library
.PARAMETER Duplicates
    Array of duplicate groups from Find-DuplicateMoviesEnhanced
.PARAMETER ReviewFolderName
    Name of the review folder
#>
function Invoke-DuplicateMove {
    param(
        [string]$Path,
        [array]$Duplicates,
        [string]$ReviewFolderName = "_Duplicates_Review"
    )

    $reviewPath = Join-Path $Path $ReviewFolderName
    if (-not (Test-Path $reviewPath)) {
        New-Item -Path $reviewPath -ItemType Directory -Force | Out-Null
    }

    # Create log file
    $logFile = Join-Path $reviewPath "DUPLICATE_CLEANUP_LOG.txt"
    $logContent = @()
    $logContent += "=" * 70
    $logContent += "DUPLICATE CLEANUP LOG - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $logContent += "=" * 70
    $logContent += ""
    $logContent += "This folder contains lower-quality duplicate movies."
    $logContent += "The BETTER quality version was kept in your main library."
    $logContent += ""
    $logContent += "Review each folder, then delete this entire folder to reclaim space."
    $logContent += ""
    $logContent += "=" * 70
    $logContent += ""

    $movedCount = 0
    $totalSize = 0
    $groupNum = 1

    foreach ($group in $Duplicates) {
        $sorted = $group | Sort-Object {
            $hasYear = if ($_.OriginalName -match '\(\d{4}\)') { 1 } else { 0 }
            ($_.Quality.Score * 1000) + ($hasYear * 100) + ($_.FileSize / 1MB)
        } -Descending

        $keep = $sorted[0]
        $toMove = $sorted | Select-Object -Skip 1

        $logContent += "-" * 70
        $logContent += "GROUP $groupNum - Confidence: $($group[0].Confidence)"
        $logContent += "KEPT: $($keep.OriginalName) (Score: $($keep.Quality.Score), Size: $(Format-FileSize $keep.FileSize))"
        $logContent += "MOVED:"

        foreach ($dupe in $toMove) {
            Write-Host "Moving: $($dupe.OriginalName)" -ForegroundColor Gray

            $destPath = Join-Path $reviewPath $dupe.OriginalName
            try {
                if (Test-Path $destPath) {
                    $suffix = 1
                    while (Test-Path "$destPath`_$suffix") { $suffix++ }
                    $destPath = "$destPath`_$suffix"
                }

                Move-Item -Path $dupe.Path -Destination $destPath -Force -ErrorAction Stop
                $movedCount++
                $totalSize += $dupe.FileSize
                $logContent += "  - $($dupe.OriginalName) (Score: $($dupe.Quality.Score), Size: $(Format-FileSize $dupe.FileSize))"
                Write-Log "Moved duplicate: $($dupe.Path) -> $destPath" "INFO"
            }
            catch {
                Write-Host "  ERROR: $_" -ForegroundColor Red
                $logContent += "  - ERROR moving $($dupe.OriginalName): $_"
            }
        }

        $logContent += ""
        $groupNum++
    }

    $logContent += "=" * 70
    $logContent += "SUMMARY: Moved $movedCount folders | Space: $(Format-FileSize $totalSize)"
    $logContent | Out-File -FilePath $logFile -Encoding UTF8

    Write-Host "`n" -NoNewline
    Write-Host "╔══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                         MOVE COMPLETE                                ║" -ForegroundColor Green
    Write-Host "╠══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Green
    Write-Host "║  Folders moved: $($movedCount.ToString().PadRight(53))║" -ForegroundColor White
    Write-Host "║  Space to reclaim: $((Format-FileSize $totalSize).PadRight(50))║" -ForegroundColor Magenta
    Write-Host "║  Review folder: $($reviewPath.Substring(0, [Math]::Min(53, $reviewPath.Length)).PadRight(53))║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Green

    Write-Host "`nNext: Review the folder, then delete it to reclaim space." -ForegroundColor Yellow
    Write-Log "Duplicate move complete: $movedCount moved, $(Format-FileSize $totalSize) to reclaim" "INFO"
}

<#
.SYNOPSIS
    Permanently deletes duplicate movies (keeps best quality)
.PARAMETER Path
    Root path of the movie library
.PARAMETER Duplicates
    Array of duplicate groups to process
#>
function Invoke-DuplicateDelete {
    param(
        [string]$Path,
        [array]$Duplicates
    )

    $deletedCount = 0
    $totalSize = 0
    $errors = @()

    foreach ($group in $Duplicates) {
        $sorted = $group | Sort-Object {
            $hasYear = if ($_.OriginalName -match '\(\d{4}\)') { 1 } else { 0 }
            ($_.Quality.Score * 1000) + ($hasYear * 100) + ($_.FileSize / 1MB)
        } -Descending

        $keep = $sorted[0]
        $toDelete = $sorted | Select-Object -Skip 1

        Write-Host "Keeping: $($keep.OriginalName)" -ForegroundColor Green

        foreach ($dupe in $toDelete) {
            Write-Host "  Deleting: $($dupe.OriginalName)" -ForegroundColor Red
            try {
                Remove-Item -Path $dupe.Path -Recurse -Force -ErrorAction Stop
                $deletedCount++
                $totalSize += $dupe.FileSize
                Write-Log "Deleted duplicate: $($dupe.Path)" "INFO"
            }
            catch {
                Write-Host "    ERROR: $_" -ForegroundColor Red
                $errors += "Failed to delete $($dupe.OriginalName): $_"
                Write-Log "Failed to delete duplicate $($dupe.Path): $_" "ERROR"
            }
        }
    }

    Write-Host "`n" -NoNewline
    Write-Host "╔══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                        DELETION COMPLETE                             ║" -ForegroundColor Green
    Write-Host "╠══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Green
    Write-Host "║  Folders deleted: $($deletedCount.ToString().PadRight(51))║" -ForegroundColor White
    Write-Host "║  Space reclaimed: $((Format-FileSize $totalSize).PadRight(51))║" -ForegroundColor Magenta
    if ($errors.Count -gt 0) {
        Write-Host "║  Errors: $($errors.Count.ToString().PadRight(60))║" -ForegroundColor Red
    }
    Write-Host "╚══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Green

    Write-Log "Duplicate deletion complete: $deletedCount deleted, $(Format-FileSize $totalSize) reclaimed" "INFO"
}

<#
.SYNOPSIS
    Exports duplicate report to CSV
.PARAMETER Duplicates
    Array of duplicate groups
.PARAMETER OutputPath
    Path for the CSV file
#>
function Export-DuplicateReport {
    param(
        [array]$Duplicates,
        [string]$OutputPath
    )

    $rows = @()
    $groupNum = 1

    foreach ($group in $Duplicates) {
        $sorted = $group | Sort-Object {
            $hasYear = if ($_.OriginalName -match '\(\d{4}\)') { 1 } else { 0 }
            ($_.Quality.Score * 1000) + ($hasYear * 100) + ($_.FileSize / 1MB)
        } -Descending

        $first = $true
        foreach ($movie in $sorted) {
            $rows += [PSCustomObject]@{
                Group = $groupNum
                Confidence = $movie.Confidence
                MatchType = ($movie.MatchType -join ", ")
                Action = if ($first) { "KEEP" } else { "DELETE" }
                FolderName = $movie.OriginalName
                Path = $movie.Path
                QualityScore = $movie.Quality.Score
                FileSize = $movie.FileSize
                FileSizeFormatted = Format-FileSize $movie.FileSize
                IMDBID = $movie.IMDBID
                TMDBID = $movie.TMDBID
                Resolution = $movie.Quality.Resolution
                Source = $movie.Quality.Source
                Codec = $movie.Quality.Codec
            }
            $first = $false
        }
        $groupNum++
    }

    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Duplicate report exported to: $OutputPath" "INFO"
}

# Keep the old function name as an alias for backwards compatibility
function Show-EnhancedDuplicateReport {
    param([string]$Path)
    Invoke-EnhancedDuplicateDetection -Path $Path
}

#============================================
# EXPORT/REPORT FUNCTIONS
#============================================

<#
.SYNOPSIS
    Exports library data to CSV format
.PARAMETER Path
    Root path of the media library
.PARAMETER OutputPath
    Path for the output CSV file
.PARAMETER MediaType
    Type of media (Movies or TVShows)
#>
function Export-LibraryToCSV {
    param(
        [string]$Path,
        [string]$OutputPath = $null,
        [string]$MediaType = "Movies"
    )

    if (-not $OutputPath) {
        $OutputPath = Join-Path $PSScriptRoot "MediaLibrary_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    }

    Write-Host "`nExporting library to CSV..." -ForegroundColor Yellow
    Write-Log "Starting CSV export for: $Path" "INFO"

    try {
        $items = @()

        if ($MediaType -eq "Movies") {
            $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne '_Trailers' }

            foreach ($folder in $folders) {
                $videoFile = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                    Sort-Object Length -Descending |
                    Select-Object -First 1

                if (-not $videoFile) { continue }

                $titleInfo = Get-NormalizedTitle -Name $folder.Name
                $quality = Get-QualityScore -FileName $folder.Name -FilePath $videoFile.FullName

                $items += [PSCustomObject]@{
                    Title = $titleInfo.NormalizedTitle
                    Year = $titleInfo.Year
                    FolderName = $folder.Name
                    FileName = $videoFile.Name
                    FileSizeMB = [math]::Round($videoFile.Length / 1MB, 2)
                    Resolution = $quality.Resolution
                    VideoCodec = $quality.Codec
                    AudioCodec = $quality.Audio
                    Source = $quality.Source
                    QualityScore = $quality.Score
                    HDR = $quality.HDR
                    HDRFormat = $quality.HDRFormat
                    Bitrate = if ($quality.Bitrate -gt 0) { [math]::Round($quality.Bitrate / 1000000, 1) } else { $null }
                    Container = $videoFile.Extension.TrimStart('.')
                    Path = $folder.FullName
                    DataSource = $quality.DataSource
                }
            }
        }
        else {
            # TV Shows
            $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

            foreach ($file in $videoFiles) {
                $epInfo = Get-EpisodeInfo -FileName $file.Name
                $quality = Get-QualityScore -FileName $file.Name -FilePath $file.FullName

                $items += [PSCustomObject]@{
                    ShowTitle = $epInfo.ShowTitle
                    Season = $epInfo.Season
                    Episode = $epInfo.Episode
                    EpisodeTitle = $epInfo.EpisodeTitle
                    FileName = $file.Name
                    FileSizeMB = [math]::Round($file.Length / 1MB, 2)
                    Resolution = $quality.Resolution
                    VideoCodec = $quality.Codec
                    AudioCodec = $quality.Audio
                    QualityScore = $quality.Score
                    HDR = $quality.HDR
                    HDRFormat = $quality.HDRFormat
                    Bitrate = if ($quality.Bitrate -gt 0) { [math]::Round($quality.Bitrate / 1000000, 1) } else { $null }
                    Container = $file.Extension.TrimStart('.')
                    Path = $file.FullName
                    DataSource = $quality.DataSource
                }
            }
        }

        if ($items.Count -gt 0) {
            $items | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-Host "Exported $($items.Count) items to: $OutputPath" -ForegroundColor Green
            Write-Log "CSV export completed: $($items.Count) items to $OutputPath" "INFO"
        } else {
            Write-Host "No items found to export" -ForegroundColor Yellow
        }

        return $OutputPath
    }
    catch {
        Write-Host "Error exporting to CSV: $_" -ForegroundColor Red
        Write-Log "Error exporting to CSV: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Generates an HTML report of the media library
.PARAMETER Path
    Root path of the media library
.PARAMETER OutputPath
    Path for the output HTML file
.PARAMETER MediaType
    Type of media (Movies or TVShows)
#>
function Export-LibraryToHTML {
    param(
        [string]$Path,
        [string]$OutputPath = $null,
        [string]$MediaType = "Movies"
    )

    if (-not $OutputPath) {
        $OutputPath = Join-Path $script:ReportsFolder "MediaLibrary_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    }

    Write-Host "`nGenerating HTML report..." -ForegroundColor Yellow
    Write-Log "Starting HTML report for: $Path" "INFO"

    try {
        $items = @()
        $stats = @{
            TotalItems = 0
            TotalSizeGB = 0
            ByResolution = @{}
            ByCodec = @{}
            ByYear = @{}
        }

        if ($MediaType -eq "Movies") {
            $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne '_Trailers' }

            $totalFolders = @($folders).Count
            $currentIndex = 0

            foreach ($folder in $folders) {
                $currentIndex++
                Write-Host "`rScanning: $currentIndex / $totalFolders - $($folder.Name)".PadRight(80) -NoNewline

                $videoFile = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                    Sort-Object Length -Descending |
                    Select-Object -First 1

                if (-not $videoFile) { continue }

                $titleInfo = Get-NormalizedTitle -Name $folder.Name
                $quality = Get-QualityScore -FileName $folder.Name -FilePath $videoFile.FullName

                $items += @{
                    Title = $titleInfo.NormalizedTitle
                    Year = $titleInfo.Year
                    Resolution = $quality.Resolution
                    Codec = $quality.Codec
                    Size = $videoFile.Length
                    Score = $quality.Score
                    HDR = $quality.HDR
                    HDRFormat = $quality.HDRFormat
                }

                # Update stats
                $stats.TotalItems++
                $stats.TotalSizeGB += $videoFile.Length / 1GB

                if (-not $stats.ByResolution.ContainsKey($quality.Resolution)) { $stats.ByResolution[$quality.Resolution] = 0 }
                $stats.ByResolution[$quality.Resolution]++

                if (-not $stats.ByCodec.ContainsKey($quality.Codec)) { $stats.ByCodec[$quality.Codec] = 0 }
                $stats.ByCodec[$quality.Codec]++

                if ($titleInfo.Year) {
                    if (-not $stats.ByYear.ContainsKey($titleInfo.Year)) { $stats.ByYear[$titleInfo.Year] = 0 }
                    $stats.ByYear[$titleInfo.Year]++
                }
            }
            Write-Host "`rScanning complete: $totalFolders items processed".PadRight(80)
        }

        # Generate HTML
        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LibraryLint Library Report</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #1a1a2e; color: #eee; padding: 20px; }
        h1 { color: #00d4ff; margin-bottom: 20px; text-align: center; }
        h2 { color: #00d4ff; margin: 20px 0 10px; border-bottom: 2px solid #00d4ff; padding-bottom: 5px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: #16213e; padding: 20px; border-radius: 10px; text-align: center; }
        .stat-value { font-size: 2.5em; color: #00d4ff; font-weight: bold; }
        .stat-label { color: #888; margin-top: 5px; }
        .chart-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .chart { background: #16213e; padding: 20px; border-radius: 10px; }
        .bar-chart { margin-top: 10px; }
        .bar-item { display: flex; align-items: center; margin: 8px 0; }
        .bar-label { width: 100px; font-size: 0.9em; }
        .bar-container { flex: 1; background: #0f3460; height: 25px; border-radius: 5px; overflow: hidden; }
        .bar { height: 100%; background: linear-gradient(90deg, #00d4ff, #0097b2); transition: width 0.5s; }
        .bar-value { margin-left: 10px; font-size: 0.9em; color: #888; }
        table { width: 100%; border-collapse: collapse; background: #16213e; border-radius: 10px; overflow: hidden; }
        th { background: #0f3460; color: #00d4ff; padding: 15px; text-align: left; }
        td { padding: 12px 15px; border-bottom: 1px solid #0f3460; }
        tr:hover { background: #1f4068; }
        .quality-high { color: #00ff88; }
        .quality-medium { color: #ffaa00; }
        .quality-low { color: #ff4444; }
        .search-box { width: 100%; padding: 12px; margin-bottom: 20px; background: #16213e; border: 2px solid #0f3460; border-radius: 8px; color: #eee; font-size: 1em; }
        .search-box:focus { outline: none; border-color: #00d4ff; }
        .footer { text-align: center; margin-top: 30px; color: #666; font-size: 0.9em; }
    </style>
</head>
<body>
    <h1>LibraryLint Library Report</h1>
    <p style="text-align: center; color: #888; margin-bottom: 30px;">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>

    <div class="stats-grid">
        <div class="stat-card">
            <div class="stat-value">$($stats.TotalItems)</div>
            <div class="stat-label">Total $MediaType</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$([math]::Round($stats.TotalSizeGB, 1)) GB</div>
            <div class="stat-label">Total Size</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$($stats.ByResolution.Keys.Count)</div>
            <div class="stat-label">Resolution Types</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$($stats.ByCodec.Keys.Count)</div>
            <div class="stat-label">Codec Types</div>
        </div>
    </div>

    <div class="chart-container">
        <div class="chart">
            <h2>By Resolution</h2>
            <div class="bar-chart">
"@

        # Add resolution bars
        $maxRes = ($stats.ByResolution.Values | Measure-Object -Maximum).Maximum
        foreach ($res in ($stats.ByResolution.GetEnumerator() | Sort-Object Value -Descending)) {
            $pct = if ($maxRes -gt 0) { [math]::Round(($res.Value / $maxRes) * 100) } else { 0 }
            $html += @"
                <div class="bar-item">
                    <span class="bar-label">$($res.Key)</span>
                    <div class="bar-container"><div class="bar" style="width: $pct%"></div></div>
                    <span class="bar-value">$($res.Value)</span>
                </div>
"@
        }

        $html += @"
            </div>
        </div>
        <div class="chart">
            <h2>By Codec</h2>
            <div class="bar-chart">
"@

        # Add codec bars
        $maxCodec = ($stats.ByCodec.Values | Measure-Object -Maximum).Maximum
        foreach ($codec in ($stats.ByCodec.GetEnumerator() | Sort-Object Value -Descending)) {
            $pct = if ($maxCodec -gt 0) { [math]::Round(($codec.Value / $maxCodec) * 100) } else { 0 }
            $html += @"
                <div class="bar-item">
                    <span class="bar-label">$($codec.Key)</span>
                    <div class="bar-container"><div class="bar" style="width: $pct%"></div></div>
                    <span class="bar-value">$($codec.Value)</span>
                </div>
"@
        }

        $html += @"
            </div>
        </div>
    </div>

    <h2>Library Contents</h2>
    <input type="text" class="search-box" placeholder="Search movies..." onkeyup="filterTable(this.value)">
    <table id="libraryTable">
        <thead>
            <tr>
                <th>Title</th>
                <th>Year</th>
                <th>Resolution</th>
                <th>Codec</th>
                <th>Size</th>
                <th>Score</th>
            </tr>
        </thead>
        <tbody>
"@

        # Add table rows
        foreach ($item in ($items | Sort-Object { $_.Title })) {
            $sizeStr = Format-FileSize $item.Size
            $scoreClass = if ($item.Score -ge 100) { "quality-high" } elseif ($item.Score -ge 50) { "quality-medium" } else { "quality-low" }
            $html += @"
            <tr>
                <td>$([System.Web.HttpUtility]::HtmlEncode($item.Title))</td>
                <td>$($item.Year)</td>
                <td>$($item.Resolution)</td>
                <td>$($item.Codec)</td>
                <td>$sizeStr</td>
                <td class="$scoreClass">$($item.Score)</td>
            </tr>
"@
        }

        $html += @"
        </tbody>
    </table>

    <div class="footer">
        <p>Generated by LibraryLint v5.0</p>
    </div>

    <script>
        function filterTable(query) {
            const table = document.getElementById('libraryTable');
            const rows = table.getElementsByTagName('tr');
            query = query.toLowerCase();
            for (let i = 1; i < rows.length; i++) {
                const cells = rows[i].getElementsByTagName('td');
                let found = false;
                for (let j = 0; j < cells.length; j++) {
                    if (cells[j].textContent.toLowerCase().includes(query)) {
                        found = true;
                        break;
                    }
                }
                rows[i].style.display = found ? '' : 'none';
            }
        }
    </script>
</body>
</html>
"@

        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "HTML report generated: $OutputPath" -ForegroundColor Green
        Write-Log "HTML report generated: $OutputPath" "INFO"

        # Optionally open in browser
        $openReport = Read-Host "Open report in browser? (Y/N) [Y]"
        if ($openReport -ne 'N' -and $openReport -ne 'n') {
            Start-Process $OutputPath
        }

        return $OutputPath
    }
    catch {
        Write-Host "Error generating HTML report: $_" -ForegroundColor Red
        Write-Log "Error generating HTML report: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Exports library data to JSON format for backup or integration
.PARAMETER Path
    Root path of the media library
.PARAMETER OutputPath
    Path for the output JSON file
#>
function Export-LibraryToJSON {
    param(
        [string]$Path,
        [string]$OutputPath = $null,
        [string]$MediaType = "Movies"
    )

    if (-not $OutputPath) {
        $OutputPath = Join-Path $PSScriptRoot "MediaLibrary_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    }

    Write-Host "`nExporting library to JSON..." -ForegroundColor Yellow

    try {
        $library = @{
            ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            MediaType = $MediaType
            SourcePath = $Path
            Items = @()
        }

        if ($MediaType -eq "Movies") {
            $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne '_Trailers' }

            foreach ($folder in $folders) {
                $videoFile = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                    Sort-Object Length -Descending |
                    Select-Object -First 1

                if (-not $videoFile) { continue }

                $titleInfo = Get-NormalizedTitle -Name $folder.Name
                $quality = Get-QualityScore -FileName $folder.Name -FilePath $videoFile.FullName

                $library.Items += @{
                    Title = $titleInfo.NormalizedTitle
                    Year = $titleInfo.Year
                    FolderName = $folder.Name
                    FileName = $videoFile.Name
                    FileSize = $videoFile.Length
                    Quality = $quality
                    Path = $folder.FullName
                }
            }
        }

        $library | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "Exported $($library.Items.Count) items to: $OutputPath" -ForegroundColor Green
        Write-Log "JSON export completed: $($library.Items.Count) items" "INFO"

        return $OutputPath
    }
    catch {
        Write-Host "Error exporting to JSON: $_" -ForegroundColor Red
        Write-Log "Error exporting to JSON: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Interactive export menu
.PARAMETER Path
    Root path of the media library
#>
function Invoke-ExportMenu {
    param(
        [string]$Path,
        [string]$MediaType = "Movies"
    )

    Write-Host "`n=== Export Library ===" -ForegroundColor Cyan
    Write-Host "1. Export to CSV (spreadsheet-compatible)"
    Write-Host "2. Export to HTML (visual report)"
    Write-Host "3. Export to JSON (backup/integration)"
    Write-Host "4. Export All Formats"
    Write-Host "5. Cancel"

    $choice = Read-Host "`nSelect export format"

    switch ($choice) {
        "1" { Export-LibraryToCSV -Path $Path -MediaType $MediaType }
        "2" { Export-LibraryToHTML -Path $Path -MediaType $MediaType }
        "3" { Export-LibraryToJSON -Path $Path -MediaType $MediaType }
        "4" {
            Export-LibraryToCSV -Path $Path -MediaType $MediaType
            Export-LibraryToHTML -Path $Path -MediaType $MediaType
            Export-LibraryToJSON -Path $Path -MediaType $MediaType
        }
        default { Write-Host "Export cancelled" -ForegroundColor Yellow }
    }
}

<#
.SYNOPSIS
    Generates and optionally auto-opens a dry run summary report
.DESCRIPTION
    Creates an HTML report summarizing all changes that would be made in dry run mode.
    Automatically opens in the default browser unless -NoAutoOpen was specified.
.PARAMETER Path
    The path that was processed
.PARAMETER MediaType
    Type of media (Movies or TVShows)
#>
function Export-DryRunReport {
    param(
        [string]$Path,
        [string]$MediaType = "Movies"
    )

    if (-not $Path) {
        Write-Verbose-Message "No path provided for dry run report"
        return $null
    }

    $reportPath = Join-Path $script:ReportsFolder "DryRun_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

    Write-Host "`nGenerating dry run report..." -ForegroundColor Yellow
    Write-Log "Generating dry run report for: $Path" "INFO"

    try {
        # Read the log file to extract DRY-RUN entries
        $logContent = @()
        if (Test-Path $script:Config.LogFile) {
            $logContent = Get-Content -Path $script:Config.LogFile -ErrorAction SilentlyContinue |
                Where-Object { $_ -match '\[DRY-RUN\]' }
        }

        $duration = if ($script:Stats.EndTime -and $script:Stats.StartTime) {
            $script:Stats.EndTime - $script:Stats.StartTime
        } else {
            (Get-Date) - $script:Stats.StartTime
        }

        # Build action summary from log entries
        $actions = @{
            Deletions = @($logContent | Where-Object { $_ -match 'Would delete' })
            Moves = @($logContent | Where-Object { $_ -match 'Would move' })
            Renames = @($logContent | Where-Object { $_ -match 'Would rename' })
            Creates = @($logContent | Where-Object { $_ -match 'Would create' })
            Extractions = @($logContent | Where-Object { $_ -match 'Would extract' })
        }

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LibraryLint Dry Run Report</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #1a1a2e; color: #eee; padding: 20px; }
        h1 { color: #ffaa00; margin-bottom: 10px; text-align: center; }
        h2 { color: #00d4ff; margin: 20px 0 10px; border-bottom: 2px solid #00d4ff; padding-bottom: 5px; }
        .warning-banner { background: #664400; border: 2px solid #ffaa00; border-radius: 10px; padding: 15px; margin-bottom: 20px; text-align: center; }
        .warning-banner h2 { color: #ffaa00; border: none; margin: 0; }
        .warning-banner p { color: #ffcc66; margin-top: 5px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin-bottom: 30px; }
        .stat-card { background: #16213e; padding: 15px; border-radius: 10px; text-align: center; }
        .stat-value { font-size: 2em; color: #00d4ff; font-weight: bold; }
        .stat-label { color: #888; margin-top: 5px; font-size: 0.9em; }
        .action-section { background: #16213e; padding: 15px; border-radius: 10px; margin-bottom: 15px; }
        .action-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
        .action-count { background: #0f3460; padding: 5px 15px; border-radius: 20px; font-size: 0.9em; }
        .action-list { max-height: 300px; overflow-y: auto; }
        .action-item { padding: 8px 12px; border-bottom: 1px solid #0f3460; font-family: 'Consolas', monospace; font-size: 0.85em; word-break: break-all; }
        .action-item:last-child { border-bottom: none; }
        .action-item:hover { background: #1f4068; }
        .delete { border-left: 3px solid #ff4444; }
        .move { border-left: 3px solid #44aaff; }
        .rename { border-left: 3px solid #44ff88; }
        .create { border-left: 3px solid #aa44ff; }
        .extract { border-left: 3px solid #ffaa44; }
        .empty-message { color: #666; font-style: italic; padding: 10px; }
        .footer { text-align: center; margin-top: 30px; color: #666; font-size: 0.9em; }
        .path-info { background: #0f3460; padding: 10px 15px; border-radius: 8px; margin-bottom: 20px; font-family: 'Consolas', monospace; word-break: break-all; }
    </style>
</head>
<body>
    <h1>🔍 Dry Run Report</h1>
    <div class="warning-banner">
        <h2>⚠️ PREVIEW ONLY - NO CHANGES MADE</h2>
        <p>This report shows what would happen if you run LibraryLint in live mode.</p>
    </div>

    <div class="path-info">
        <strong>Scanned Path:</strong> $([System.Web.HttpUtility]::HtmlEncode($Path))<br>
        <strong>Media Type:</strong> $MediaType<br>
        <strong>Scan Duration:</strong> $("{0:hh\:mm\:ss}" -f $duration)
    </div>

    <div class="stats-grid">
        <div class="stat-card">
            <div class="stat-value">$($actions.Deletions.Count)</div>
            <div class="stat-label">Files to Delete</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$($actions.Moves.Count)</div>
            <div class="stat-label">Files to Move</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$($actions.Renames.Count)</div>
            <div class="stat-label">Items to Rename</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$($actions.Creates.Count)</div>
            <div class="stat-label">Items to Create</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$($actions.Extractions.Count)</div>
            <div class="stat-label">Archives to Extract</div>
        </div>
    </div>
"@

        # Add action sections
        $sectionData = @(
            @{ Title = "Files to Delete"; Items = $actions.Deletions; Class = "delete"; Icon = "🗑️" }
            @{ Title = "Files to Move"; Items = $actions.Moves; Class = "move"; Icon = "📦" }
            @{ Title = "Items to Rename"; Items = $actions.Renames; Class = "rename"; Icon = "✏️" }
            @{ Title = "Items to Create"; Items = $actions.Creates; Class = "create"; Icon = "📁" }
            @{ Title = "Archives to Extract"; Items = $actions.Extractions; Class = "extract"; Icon = "📂" }
        )

        foreach ($section in $sectionData) {
            $html += @"
    <div class="action-section">
        <div class="action-header">
            <h2>$($section.Icon) $($section.Title)</h2>
            <span class="action-count">$($section.Items.Count) items</span>
        </div>
        <div class="action-list">
"@
            if ($section.Items.Count -eq 0) {
                $html += "            <div class='empty-message'>No actions in this category</div>`n"
            } else {
                foreach ($item in $section.Items) {
                    # Extract the relevant part of the log entry
                    $cleanItem = $item -replace '^\[.*?\] \[DRY-RUN\] ', ''
                    $html += "            <div class='action-item $($section.Class)'>$([System.Web.HttpUtility]::HtmlEncode($cleanItem))</div>`n"
                }
            }
            $html += @"
        </div>
    </div>
"@
        }

        $html += @"
    <div class="footer">
        <p>Generated by LibraryLint v5.0 | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p style="margin-top: 10px;">To apply these changes, run LibraryLint again and select <strong>No</strong> for dry-run mode.</p>
    </div>
</body>
</html>
"@

        $html | Out-File -FilePath $reportPath -Encoding UTF8
        Write-Host "Dry run report generated: $reportPath" -ForegroundColor Green
        Write-Log "Dry run report generated: $reportPath" "INFO"

        # Auto-open unless disabled
        if (-not $script:NoAutoOpen) {
            Write-Host "Opening report in browser..." -ForegroundColor Cyan
            Start-Process $reportPath
        }

        return $reportPath
    }
    catch {
        Write-Host "Error generating dry run report: $_" -ForegroundColor Red
        Write-Log "Error generating dry run report: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Displays a folder browser dialog for user to select a directory
.OUTPUTS
    String - The full path of the selected folder
#>
function Select-FolderDialog {
    param(
        [string]$Description = "Select the folder to clean up"
    )
    Write-Host "Opening folder dialog... (check taskbar if it doesn't appear)" -ForegroundColor Yellow
    Write-Verbose-Message "Opening folder browser dialog..."
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.ShowNewFolderButton = $true

    $result = $folderBrowser.ShowDialog()
    Write-Verbose-Message "Dialog closed with result: $result"

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Verbose-Message "Selected path: $($folderBrowser.SelectedPath)"
        return $folderBrowser.SelectedPath
    } else {
        Write-Host "No folder selected" -ForegroundColor Red
        Write-Log "No folder selected by user" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Formats bytes into human-readable format
.PARAMETER Bytes
    The number of bytes to format
#>
function Format-FileSize {
    param([long]$Bytes)

    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

<#
.SYNOPSIS
    Displays the statistics summary at the end of processing
#>
function Show-Statistics {
    $script:Stats.EndTime = Get-Date
    $duration = $script:Stats.EndTime - $script:Stats.StartTime

    Write-Host "`n" -NoNewline
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                    PROCESSING SUMMARY                        ║" -ForegroundColor Cyan
    Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan

    # Time stats
    Write-Host "║  " -ForegroundColor Cyan -NoNewline
    Write-Host "Duration:              " -NoNewline
    Write-Host ("{0:hh\:mm\:ss}" -f $duration).PadRight(39) -ForegroundColor White -NoNewline
    Write-Host "║" -ForegroundColor Cyan

    Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan

    # File operations
    if ($script:Stats.FilesDeleted -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Files Deleted:         " -NoNewline
        Write-Host "$($script:Stats.FilesDeleted)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.BytesDeleted -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Space Reclaimed:       " -NoNewline
        Write-Host (Format-FileSize $script:Stats.BytesDeleted).PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.ArchivesExtracted -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Archives Extracted:    " -NoNewline
        Write-Host "$($script:Stats.ArchivesExtracted)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.ArchivesFailed -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Archives Failed:       " -NoNewline
        Write-Host "$($script:Stats.ArchivesFailed)".PadRight(39) -ForegroundColor Yellow -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.FoldersCreated -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Folders Created:       " -NoNewline
        Write-Host "$($script:Stats.FoldersCreated)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.FoldersRenamed -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Folders Renamed:       " -NoNewline
        Write-Host "$($script:Stats.FoldersRenamed)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.FilesMoved -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Files Moved:           " -NoNewline
        Write-Host "$($script:Stats.FilesMoved)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.EmptyFoldersRemoved -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Empty Folders Removed: " -NoNewline
        Write-Host "$($script:Stats.EmptyFoldersRemoved)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.SubtitlesProcessed -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Subtitles Processed:   " -NoNewline
        Write-Host "$($script:Stats.SubtitlesProcessed)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.SubtitlesDeleted -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Subtitles Deleted:     " -NoNewline
        Write-Host "$($script:Stats.SubtitlesDeleted)".PadRight(39) -ForegroundColor Yellow -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.TrailersMoved -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Trailers Moved:        " -NoNewline
        Write-Host "$($script:Stats.TrailersMoved)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.OperationsRetried -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Operations Retried:    " -NoNewline
        Write-Host "$($script:Stats.OperationsRetried)".PadRight(39) -ForegroundColor Yellow -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.NFOFilesCreated -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "NFO Files Created:     " -NoNewline
        Write-Host "$($script:Stats.NFOFilesCreated)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.NFOFilesRead -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "NFO Files Parsed:      " -NoNewline
        Write-Host "$($script:Stats.NFOFilesRead)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan

    # Errors and warnings
    Write-Host "║  " -ForegroundColor Cyan -NoNewline
    Write-Host "Errors:                " -NoNewline
    $errorColor = if ($script:Stats.Errors -gt 0) { "Red" } else { "Green" }
    Write-Host "$($script:Stats.Errors)".PadRight(39) -ForegroundColor $errorColor -NoNewline
    Write-Host "║" -ForegroundColor Cyan

    Write-Host "║  " -ForegroundColor Cyan -NoNewline
    Write-Host "Warnings:              " -NoNewline
    $warningColor = if ($script:Stats.Warnings -gt 0) { "Yellow" } else { "Green" }
    Write-Host "$($script:Stats.Warnings)".PadRight(39) -ForegroundColor $warningColor -NoNewline
    Write-Host "║" -ForegroundColor Cyan

    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    # Log the summary
    Write-Log "Processing completed - Duration: $($duration.ToString('hh\:mm\:ss')), Files Deleted: $($script:Stats.FilesDeleted), Space Reclaimed: $(Format-FileSize $script:Stats.BytesDeleted), Archives Extracted: $($script:Stats.ArchivesExtracted), Errors: $($script:Stats.Errors)" "INFO"
}

#============================================
# 7-ZIP FUNCTIONS
#============================================

<#
.SYNOPSIS
    Verifies 7-Zip is installed, installs it if not present
.OUTPUTS
    Boolean - True if 7-Zip is available, False otherwise
#>
function Test-SevenZipInstallation {
    Write-Host "Checking 7-Zip installation..." -ForegroundColor Yellow

    if (Test-Path -Path $script:Config.SevenZipPath) {
        Write-Host "7-Zip installed" -ForegroundColor Green
        Write-Log "7-Zip found at $($script:Config.SevenZipPath)" "INFO"
        return $true
    }

    Write-Host "7-Zip not installed. Installing now..." -ForegroundColor Yellow
    Write-Log "7-Zip not found, attempting installation" "INFO"

    $installerPath = "$env:TEMP\7z-installer.exe"
    $7zipUrl = "https://www.7-zip.org/a/7z2408-x64.exe"

    try {
        Write-Host "Downloading 7-Zip..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $7zipUrl -OutFile $installerPath -UseBasicParsing

        Write-Host "Installing 7-Zip..." -ForegroundColor Yellow
        Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait -NoNewWindow

        Remove-Item $installerPath -ErrorAction SilentlyContinue

        if (Test-Path $script:Config.SevenZipPath) {
            Write-Host "7-Zip installed successfully" -ForegroundColor Green
            Write-Log "7-Zip installed successfully" "INFO"
            return $true
        } else {
            Write-Host "7-Zip installation failed" -ForegroundColor Red
            Write-Log "7-Zip installation failed" "ERROR"
            return $false
        }
    }
    catch {
        Write-Host "Error installing 7-Zip: $_" -ForegroundColor Red
        Write-Host "Please install 7-Zip manually from https://www.7-zip.org/" -ForegroundColor Yellow
        Write-Log "Error installing 7-Zip: $_" "ERROR"
        return $false
    }
}

#============================================
# FILE CLEANUP FUNCTIONS
#============================================

<#
.SYNOPSIS
    Removes unnecessary files matching specified patterns
.PARAMETER Path
    The root path to search for unnecessary files
#>
function Remove-UnnecessaryFiles {
    param(
        [string]$Path
    )

    Write-Host "Cleaning unnecessary files..." -ForegroundColor Yellow
    Write-Log "Starting unnecessary file cleanup in: $Path" "INFO"

    try {
        $filesToDelete = @()

        foreach ($pattern in $script:Config.UnnecessaryPatterns) {
            $files = Get-ChildItem -Path $Path -Filter $pattern -Recurse -ErrorAction SilentlyContinue
            if ($files) {
                $filesToDelete += $files
            }
        }

        if ($filesToDelete.Count -gt 0) {
            Write-Host "Found $($filesToDelete.Count) unnecessary file(s)/folder(s)" -ForegroundColor Cyan

            foreach ($item in $filesToDelete) {
                $itemSize = if ($item.PSIsContainer) {
                    (Get-ChildItem -Path $item.FullName -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                } else {
                    $item.Length
                }

                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would delete: $($item.FullName)" -ForegroundColor Yellow
                    Write-Log "Would delete: $($item.FullName)" "DRY-RUN"
                } else {
                    Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "Deleted: $($item.Name)" -ForegroundColor Gray
                    Write-Log "Deleted: $($item.FullName)" "INFO"
                    $script:Stats.FilesDeleted++
                    $script:Stats.BytesDeleted += $itemSize
                }
            }
            Write-Host "Unnecessary files cleaned" -ForegroundColor Green
        } else {
            Write-Host "No unnecessary files found" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Warning: Some files could not be deleted: $_" -ForegroundColor Yellow
        Write-Log "Error during file cleanup: $_" "ERROR"
    }
}

#============================================
# SUBTITLE FUNCTIONS
#============================================

<#
.SYNOPSIS
    Processes subtitle files - keeps preferred languages, removes others
.PARAMETER Path
    The root path to search for subtitle files
.DESCRIPTION
    Scans for subtitle files (.srt, .sub, .idx, .ass, .ssa, .vtt) and either:
    - Keeps them if they match preferred languages or if KeepSubtitles is true
    - Deletes non-preferred language subtitles
    - Moves subtitle files to be alongside their video files
#>
function Invoke-SubtitleProcessing {
    param(
        [string]$Path
    )

    Write-Host "Processing subtitle files..." -ForegroundColor Yellow
    Write-Log "Starting subtitle processing in: $Path" "INFO"

    try {
        # Find all subtitle files
        $subtitleFiles = @()
        foreach ($ext in $script:Config.SubtitleExtensions) {
            $files = Get-ChildItem -Path $Path -Filter "*$ext" -Recurse -ErrorAction SilentlyContinue
            if ($files) {
                $subtitleFiles += $files
            }
        }

        if ($subtitleFiles.Count -eq 0) {
            Write-Host "No subtitle files found" -ForegroundColor Cyan
            return
        }

        Write-Host "Found $($subtitleFiles.Count) subtitle file(s)" -ForegroundColor Cyan

        foreach ($subtitle in $subtitleFiles) {
            $script:Stats.SubtitlesProcessed++

            # Check if subtitle matches preferred language
            $isPreferred = $false
            $subtitleNameLower = $subtitle.BaseName.ToLower()

            foreach ($lang in $script:Config.PreferredSubtitleLanguages) {
                if ($subtitleNameLower -match "\.$lang$" -or $subtitleNameLower -match "\.$lang\." -or $subtitleNameLower -match "_$lang$" -or $subtitleNameLower -match "_$lang[_\.]") {
                    $isPreferred = $true
                    break
                }
            }

            # If no language tag found, assume it's the default/preferred language
            $hasLanguageTag = $subtitleNameLower -match '\.(eng|en|english|spa|es|spanish|fre|fr|french|ger|de|german|ita|it|italian|por|pt|portuguese|rus|ru|russian|jpn|ja|japanese|chi|zh|chinese|kor|ko|korean|ara|ar|arabic|hin|hi|hindi|dut|nl|dutch|pol|pl|polish|swe|sv|swedish|nor|no|norwegian|dan|da|danish|fin|fi|finnish)(\.|$|_)'
            if (-not $hasLanguageTag) {
                $isPreferred = $true  # No language tag = default language, keep it
            }

            if ($script:Config.KeepSubtitles -and $isPreferred) {
                Write-Host "Keeping subtitle: $($subtitle.Name)" -ForegroundColor Green
                Write-Log "Keeping subtitle: $($subtitle.FullName)" "INFO"
            } else {
                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would delete subtitle: $($subtitle.Name)" -ForegroundColor Yellow
                    Write-Log "Would delete subtitle: $($subtitle.FullName)" "DRY-RUN"
                } else {
                    $subtitleSize = $subtitle.Length
                    Remove-Item -Path $subtitle.FullName -Force -ErrorAction SilentlyContinue
                    Write-Host "Deleted subtitle: $($subtitle.Name)" -ForegroundColor Gray
                    Write-Log "Deleted subtitle: $($subtitle.FullName)" "INFO"
                    $script:Stats.SubtitlesDeleted++
                    $script:Stats.BytesDeleted += $subtitleSize
                }
            }
        }

        Write-Host "Subtitle processing completed" -ForegroundColor Green
    }
    catch {
        Write-Host "Error processing subtitles: $_" -ForegroundColor Red
        Write-Log "Error processing subtitles: $_" "ERROR"
    }
}

#============================================
# TRAILER FUNCTIONS
#============================================

<#
.SYNOPSIS
    Moves trailer files to a _Trailers subfolder instead of deleting them
.PARAMETER Path
    The root path to search for trailer files
.DESCRIPTION
    Finds trailer/teaser files and moves them to a _Trailers folder,
    preserving them for later viewing while keeping the main library clean
#>
function Move-TrailersToFolder {
    param(
        [string]$Path
    )

    Write-Host "Processing trailer files..." -ForegroundColor Yellow
    Write-Log "Starting trailer processing in: $Path" "INFO"

    try {
        $trailerFiles = @()

        foreach ($pattern in $script:Config.TrailerPatterns) {
            # Search for trailer video files
            foreach ($ext in $script:Config.VideoExtensions) {
                $files = Get-ChildItem -Path $Path -Filter "$pattern$ext" -Recurse -ErrorAction SilentlyContinue
                if ($files) {
                    $trailerFiles += $files
                }
            }
            # Also search for trailer folders
            $trailerFolders = Get-ChildItem -Path $Path -Directory -Filter $pattern -Recurse -ErrorAction SilentlyContinue
            if ($trailerFolders) {
                foreach ($folder in $trailerFolders) {
                    $videosInFolder = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                        Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }
                    if ($videosInFolder) {
                        $trailerFiles += $videosInFolder
                    }
                }
            }
        }

        # Remove duplicates
        $trailerFiles = $trailerFiles | Select-Object -Unique

        if ($trailerFiles.Count -eq 0) {
            Write-Host "No trailer files found" -ForegroundColor Cyan
            return
        }

        Write-Host "Found $($trailerFiles.Count) trailer file(s)" -ForegroundColor Cyan

        # Create _Trailers folder at root
        $trailersFolder = Join-Path $Path "_Trailers"

        if (-not $script:Config.DryRun) {
            if (-not (Test-Path $trailersFolder)) {
                New-Item -Path $trailersFolder -ItemType Directory -Force | Out-Null
                Write-Host "Created _Trailers folder" -ForegroundColor Green
                Write-Log "Created _Trailers folder at: $trailersFolder" "INFO"
                $script:Stats.FoldersCreated++
            }
        }

        foreach ($trailer in $trailerFiles) {
            $destPath = Join-Path $trailersFolder $trailer.Name

            # Handle duplicate names by adding a number
            $counter = 1
            while (Test-Path $destPath) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($trailer.Name)
                $extension = $trailer.Extension
                $destPath = Join-Path $trailersFolder "$baseName`_$counter$extension"
                $counter++
            }

            if ($script:Config.DryRun) {
                Write-Host "[DRY-RUN] Would move trailer: $($trailer.Name) -> _Trailers/" -ForegroundColor Yellow
                Write-Log "Would move trailer: $($trailer.FullName) to $destPath" "DRY-RUN"
            } else {
                try {
                    Move-Item -Path $trailer.FullName -Destination $destPath -Force -ErrorAction Stop
                    Write-Host "Moved trailer: $($trailer.Name)" -ForegroundColor Green
                    Write-Log "Moved trailer: $($trailer.FullName) to $destPath" "INFO"
                    $script:Stats.TrailersMoved++
                }
                catch {
                    Write-Host "Warning: Could not move trailer $($trailer.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error moving trailer $($trailer.FullName): $_" "WARNING"
                }
            }
        }

        Write-Host "Trailer processing completed" -ForegroundColor Green
    }
    catch {
        Write-Host "Error processing trailers: $_" -ForegroundColor Red
        Write-Log "Error processing trailers: $_" "ERROR"
    }
}

#============================================
# NFO FILE FUNCTIONS
#============================================

<#
.SYNOPSIS
    Parses an existing NFO file and extracts metadata
.PARAMETER NfoPath
    The path to the NFO file to parse
.OUTPUTS
    Hashtable containing extracted metadata (title, year, plot, etc.)
.DESCRIPTION
    Reads Kodi-compatible NFO files and extracts metadata including:
    - Title, original title, sort title
    - Year, release date
    - Plot, tagline
    - Rating, votes
    - IMDB/TMDB IDs
    - Genre, studio, director, actors
#>
function Read-NFOFile {
    param(
        [string]$NfoPath
    )

    $metadata = @{
        Title = $null
        OriginalTitle = $null
        SortTitle = $null
        Year = $null
        Plot = $null
        Tagline = $null
        Rating = $null
        Votes = $null
        IMDBID = $null
        TMDBID = $null
        Genres = @()
        Studios = @()
        Directors = @()
        Actors = @()
        Runtime = $null
    }

    try {
        if (-not (Test-Path $NfoPath)) {
            Write-Log "NFO file not found: $NfoPath" "WARNING"
            return $null
        }

        # Read raw content first to check if it's XML
        $rawContent = Get-Content -Path $NfoPath -Raw -Encoding UTF8 -ErrorAction Stop

        # Check if this looks like a Kodi XML NFO (not a scene release NFO)
        # Scene NFOs typically contain BBCode [img], ASCII art, or plain text
        if ($rawContent -notmatch '^\s*(<\?xml|<movie|<tvshow|<episodedetails)') {
            # Not a Kodi XML NFO - skip silently
            Write-Log "Skipping non-XML NFO file: $NfoPath" "DEBUG"
            return $null
        }

        [xml]$nfoContent = $rawContent
        $script:Stats.NFOFilesRead++

        # Movie NFO
        if ($nfoContent.movie) {
            $movie = $nfoContent.movie
            $metadata.Title = $movie.title
            $metadata.OriginalTitle = $movie.originaltitle
            $metadata.SortTitle = $movie.sorttitle
            $metadata.Year = $movie.year
            $metadata.Plot = $movie.plot
            $metadata.Tagline = $movie.tagline
            $metadata.Rating = $movie.rating
            $metadata.Votes = $movie.votes
            $metadata.Runtime = $movie.runtime

            # Extract IDs
            if ($movie.uniqueid) {
                foreach ($id in $movie.uniqueid) {
                    if ($id.type -eq 'imdb') { $metadata.IMDBID = $id.'#text' }
                    if ($id.type -eq 'tmdb') { $metadata.TMDBID = $id.'#text' }
                }
            }
            # Fallback for older NFO format
            if (-not $metadata.IMDBID -and $movie.imdbid) { $metadata.IMDBID = $movie.imdbid }
            if (-not $metadata.IMDBID -and $movie.id) { $metadata.IMDBID = $movie.id }

            # Extract genres
            if ($movie.genre) {
                $metadata.Genres = @($movie.genre)
            }

            # Extract studios
            if ($movie.studio) {
                $metadata.Studios = @($movie.studio)
            }

            # Extract directors
            if ($movie.director) {
                $metadata.Directors = @($movie.director)
            }

            # Extract actors
            if ($movie.actor) {
                foreach ($actor in $movie.actor) {
                    $metadata.Actors += @{
                        Name = $actor.name
                        Role = $actor.role
                        Thumb = $actor.thumb
                    }
                }
            }

            Write-Log "Parsed NFO file: $NfoPath - Title: $($metadata.Title)" "INFO"
        }
        # TV Show NFO
        elseif ($nfoContent.tvshow) {
            $show = $nfoContent.tvshow
            $metadata.Title = $show.title
            $metadata.OriginalTitle = $show.originaltitle
            $metadata.Year = $show.year
            $metadata.Plot = $show.plot
            $metadata.Rating = $show.rating

            if ($show.uniqueid) {
                foreach ($id in $show.uniqueid) {
                    if ($id.type -eq 'imdb') { $metadata.IMDBID = $id.'#text' }
                    if ($id.type -eq 'tmdb') { $metadata.TMDBID = $id.'#text' }
                }
            }

            if ($show.genre) {
                $metadata.Genres = @($show.genre)
            }

            Write-Log "Parsed TV Show NFO file: $NfoPath - Title: $($metadata.Title)" "INFO"
        }
        # Episode NFO
        elseif ($nfoContent.episodedetails) {
            $episode = $nfoContent.episodedetails
            $metadata.Title = $episode.title
            $metadata.Plot = $episode.plot
            $metadata.Rating = $episode.rating
            $metadata.Season = $episode.season
            $metadata.Episode = $episode.episode
            $metadata.Aired = $episode.aired

            Write-Log "Parsed Episode NFO file: $NfoPath - Title: $($metadata.Title)" "INFO"
        }

        return $metadata
    }
    catch {
        Write-Log "Error parsing NFO file $NfoPath : $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Generates a Kodi-compatible NFO file for a movie
.PARAMETER VideoPath
    The path to the video file
.PARAMETER Title
    The movie title (if not provided, extracted from folder/file name)
.PARAMETER Year
    The movie year (if not provided, extracted from folder/file name)
.DESCRIPTION
    Creates a basic NFO file that Kodi can use to identify the movie.
    The NFO file is named the same as the video file with .nfo extension.
#>
function New-MovieNFO {
    param(
        [string]$VideoPath,
        [string]$Title = $null,
        [string]$Year = $null
    )

    try {
        $videoFile = Get-Item $VideoPath -ErrorAction Stop
        $nfoPath = Join-Path $videoFile.DirectoryName "$([System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)).nfo"

        # Skip if NFO already exists and has content
        if (Test-Path $nfoPath) {
            $existingNfo = Get-Content $nfoPath -Raw -ErrorAction SilentlyContinue
            # Check if NFO is empty or just has basic structure without real data
            if ($existingNfo -and $existingNfo.Length -gt 500 -and $existingNfo -match '<plot>.+</plot>') {
                Write-Host "NFO already exists: $($videoFile.Name)" -ForegroundColor Cyan
                Write-Log "NFO already exists, skipping: $nfoPath" "INFO"
                return
            }
            # NFO exists but is empty or minimal - will regenerate
            Write-Host "Regenerating incomplete NFO: $($videoFile.Name)" -ForegroundColor Yellow
            Write-Log "NFO exists but is empty/minimal, regenerating: $nfoPath" "INFO"
        }

        # Extract title and year from folder name if not provided
        if (-not $Title -or -not $Year) {
            $folderName = $videoFile.Directory.Name

            # Try to extract year in parentheses format: "Movie Name (2024)" or "Movie Name (2024_1080p_x265)"
            # The year must be at the start of the parentheses, but may have other content after it
            if ($folderName -match '^(.+?)\s*\((\d{4})\)') {
                # Clean format: Movie Name (2024)
                if (-not $Title) { $Title = $Matches[1].Trim() }
                if (-not $Year) { $Year = $Matches[2] }
            }
            elseif ($folderName -match '^(.+?)\s*\((\d{4})[_\s]') {
                # Year followed by underscore or space inside parens: Movie Name (2024_1080p_x265)
                if (-not $Title) { $Title = $Matches[1].Trim() }
                if (-not $Year) { $Year = $Matches[2] }
            }
            # Try to extract year without parentheses: "Movie Name 2024"
            elseif ($folderName -match '^(.+?)\s+(\d{4})$') {
                if (-not $Title) { $Title = $Matches[1].Trim() }
                if (-not $Year) { $Year = $Matches[2] }
            }
            # Just use folder name as title
            else {
                if (-not $Title) { $Title = $folderName }
            }
        }

        # Clean up title
        $Title = $Title -replace '\.', ' '
        $Title = $Title -replace '\s+', ' '
        $Title = $Title.Trim()

        if ($script:Config.DryRun) {
            Write-Host "[DRY-RUN] Would create NFO for: $Title $(if($Year){"($Year)"})" -ForegroundColor Yellow
            Write-Log "Would create NFO for: $Title ($Year) at $nfoPath" "DRY-RUN"
            return
        }

        # Try to fetch metadata from TMDB if API key is available
        if ($script:Config.TMDBApiKey) {
            Write-Log "Searching TMDB for: '$Title' (Year: $Year)" "DEBUG"
            $searchResult = Search-TMDBMovie -Title $Title -Year $Year -ApiKey $script:Config.TMDBApiKey

            # If search with year fails, try without year (folder year might be wrong)
            if (-not $searchResult -and $Year) {
                Write-Log "TMDB search with year failed, retrying without year filter" "DEBUG"
                $searchResult = Search-TMDBMovie -Title $Title -ApiKey $script:Config.TMDBApiKey
            }

            if ($searchResult -and $searchResult.Id) {
                Write-Log "TMDB found: $($searchResult.Title) (ID: $($searchResult.Id))" "DEBUG"
                # Get full movie details from TMDB
                $tmdbMetadata = Get-TMDBMovieDetails -MovieId $searchResult.Id -ApiKey $script:Config.TMDBApiKey

                if ($tmdbMetadata) {
                    $null = New-MovieNFOFromTMDB -Metadata $tmdbMetadata -NFOPath $nfoPath
                    Write-Host "Created NFO from TMDB: $($tmdbMetadata.Title) ($($tmdbMetadata.Year))" -ForegroundColor Green
                    Write-Log "Created NFO from TMDB for: $Title" "INFO"

                    # Download artwork (poster, fanart, actor images) from TMDB
                    $artworkCount = Save-MovieArtwork -Metadata $tmdbMetadata -MovieFolder $videoFile.DirectoryName

                    # Download additional artwork (clearlogo, banner, etc.) from fanart.tv
                    if ($script:Config.FanartTVApiKey -and $tmdbMetadata.TMDBID) {
                        $fanartCount = Save-FanartTVArtwork -TMDBID $tmdbMetadata.TMDBID -MovieFolder $videoFile.DirectoryName -ApiKey $script:Config.FanartTVApiKey
                        $artworkCount += $fanartCount
                    }

                    if ($artworkCount -gt 0) {
                        Write-Host "    Downloaded $artworkCount artwork file(s)" -ForegroundColor Cyan
                    }

                    # Download trailer if enabled and available
                    if ($script:Config.DownloadTrailers -and $tmdbMetadata.TrailerKey) {
                        Save-MovieTrailer -TrailerKey $tmdbMetadata.TrailerKey -MovieFolder $videoFile.DirectoryName -MovieTitle $tmdbMetadata.Title -Quality $script:Config.TrailerQuality
                    }

                    # Download subtitle if enabled
                    if ($script:Config.DownloadSubtitles) {
                        $null = Save-MovieSubtitle -IMDBID $tmdbMetadata.IMDBID -Title $tmdbMetadata.Title -Year $tmdbMetadata.Year -MovieFolder $videoFile.DirectoryName -MovieTitle $tmdbMetadata.Title -Language $script:Config.SubtitleLanguage
                    }

                    return
                }
            }

            Write-Host "TMDB lookup failed for: $Title - creating basic NFO" -ForegroundColor Yellow
            Write-Log "TMDB lookup failed for: $Title, creating basic NFO" "WARNING"
        } else {
            Write-Log "No TMDB API key configured - creating basic NFO for: $Title" "INFO"
        }

        # Fallback: Create basic NFO without TMDB data
        $nfoContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<movie>
    <title>$([System.Security.SecurityElement]::Escape($Title))</title>
    $(if($Year){"<year>$Year</year>"})
    <plot></plot>
    <outline></outline>
    <tagline></tagline>
    <runtime></runtime>
    <thumb></thumb>
    <fanart></fanart>
    <mpaa></mpaa>
    <genre></genre>
    <studio></studio>
    <director></director>
</movie>
"@

        $nfoContent | Out-File -FilePath $nfoPath -Encoding UTF8 -Force
        Write-Host "Created basic NFO: $Title $(if($Year){"($Year)"})" -ForegroundColor Green
        Write-Log "Created basic NFO file: $nfoPath" "INFO"
        $script:Stats.NFOFilesCreated++
    }
    catch {
        Write-Host "Error creating NFO for $VideoPath : $_" -ForegroundColor Red
        Write-Log "Error creating NFO for $VideoPath : $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Generates NFO files for all movies in a folder that don't have them
.PARAMETER Path
    The root path of the movie library
#>
function Invoke-NFOGeneration {
    param(
        [string]$Path
    )

    if (-not $script:Config.GenerateNFO) {
        return
    }

    Write-Host "Generating NFO files for movies..." -ForegroundColor Yellow
    Write-Log "Starting NFO generation in: $Path" "INFO"

    try {
        # Get all movie folders (folders containing video files)
        $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue

        foreach ($folder in $movieFolders) {
            # Skip the _Trailers folder
            if ($folder.Name -eq '_Trailers') {
                continue
            }

            # Find video files in this folder
            $videoFiles = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

            foreach ($video in $videoFiles) {
                New-MovieNFO -VideoPath $video.FullName
            }
        }

        # Also check for video files directly in root
        $rootVideos = Get-ChildItem -Path $Path -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        foreach ($video in $rootVideos) {
            New-MovieNFO -VideoPath $video.FullName
        }

        Write-Host "NFO generation completed" -ForegroundColor Green
    }
    catch {
        Write-Host "Error during NFO generation: $_" -ForegroundColor Red
        Write-Log "Error during NFO generation: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Scans and displays metadata from existing NFO files
.PARAMETER Path
    The root path to scan for NFO files
#>
function Show-NFOMetadata {
    param(
        [string]$Path
    )

    Write-Host "`nScanning for existing NFO files..." -ForegroundColor Yellow
    Write-Log "Scanning for NFO files in: $Path" "INFO"

    try {
        $nfoFiles = Get-ChildItem -Path $Path -Filter "*.nfo" -Recurse -ErrorAction SilentlyContinue

        if ($nfoFiles.Count -eq 0) {
            Write-Host "No NFO files found" -ForegroundColor Cyan
            return
        }

        Write-Host "Found $($nfoFiles.Count) NFO file(s)`n" -ForegroundColor Cyan

        foreach ($nfo in $nfoFiles) {
            $metadata = Read-NFOFile -NfoPath $nfo.FullName

            if ($metadata -and $metadata.Title) {
                $displayTitle = $metadata.Title
                if ($metadata.Year) { $displayTitle += " ($($metadata.Year))" }

                Write-Host "  $displayTitle" -ForegroundColor White
                if ($metadata.IMDBID) {
                    Write-Host "    IMDB: $($metadata.IMDBID)" -ForegroundColor Gray
                }
                if ($metadata.Genres -and $metadata.Genres.Count -gt 0) {
                    Write-Host "    Genres: $($metadata.Genres -join ', ')" -ForegroundColor Gray
                }
            }
        }
    }
    catch {
        Write-Host "Error scanning NFO files: $_" -ForegroundColor Red
        Write-Log "Error scanning NFO files: $_" "ERROR"
    }
}

#============================================
# FOLDER YEAR FIX FUNCTIONS
#============================================

<#
.SYNOPSIS
    Fixes missing years in movie folder names
.PARAMETER Path
    The root path to scan for movie folders
.PARAMETER WhatIf
    If true, only shows what would be changed without making changes
.DESCRIPTION
    Scans movie folders for missing years, looks up the year from:
    1. Existing NFO file in the folder
    2. TMDB API (confirms NFO year or finds missing year)
    Then renames folders to "Movie Title (Year)" format and updates/creates NFO files.
#>
function Repair-MovieFolderYears {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [switch]$WhatIf = $false
    )

    Write-Host "`n=== Fix Missing Years in Movie Folders ===" -ForegroundColor Cyan
    if ($WhatIf) {
        Write-Host "[DRY-RUN MODE - No changes will be made]" -ForegroundColor Yellow
    }
    Write-Log "Starting folder year fix in: $Path (WhatIf: $WhatIf)" "INFO"

    if (-not $script:Config.TMDBApiKey) {
        Write-Host "`nWarning: TMDB API key not configured. Only NFO files will be used for year lookup." -ForegroundColor Yellow
        Write-Host "To enable TMDB lookups, add your API key to the config file." -ForegroundColor Yellow
        $response = Read-Host "Continue anyway? (Y/N)"
        if ($response -notmatch '^[Yy]') {
            return
        }
    }

    $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue
    $foldersToFix = @()
    $fixed = 0
    $skipped = 0
    $failed = 0
    $noYearFound = @()

    Write-Host "`nScanning folders..." -ForegroundColor Yellow

    foreach ($folder in $folders) {
        # Skip folders that already have a year
        if ($folder.Name -match '\((19|20)\d{2}\)\s*$') {
            continue
        }

        # Get the title from folder name (clean it up)
        $titleInfo = Get-NormalizedTitle -Name $folder.Name
        $cleanTitle = $titleInfo.NormalizedTitle

        # Check for existing NFO file
        $nfoFile = Get-ChildItem -LiteralPath $folder.FullName -Filter "*.nfo" -File -ErrorAction SilentlyContinue | Select-Object -First 1
        $nfoYear = $null
        $nfoMetadata = $null

        if ($nfoFile) {
            $nfoMetadata = Read-NFOFile -NfoPath $nfoFile.FullName
            if ($nfoMetadata -and $nfoMetadata.Year) {
                $nfoYear = $nfoMetadata.Year
            }
        }

        # Search TMDB to confirm or find year
        $tmdbResult = $null
        $tmdbYear = $null

        if ($script:Config.TMDBApiKey) {
            # If we have NFO year, search with it for better accuracy
            $tmdbResult = Search-TMDBMovie -Title $cleanTitle -Year $nfoYear -ApiKey $script:Config.TMDBApiKey
            if ($tmdbResult) {
                $tmdbYear = $tmdbResult.Year
            }
        }

        # Determine final year to use
        $finalYear = $null
        $yearSource = $null

        if ($nfoYear -and $tmdbYear) {
            if ($nfoYear -eq $tmdbYear) {
                $finalYear = $nfoYear
                $yearSource = "NFO + TMDB (confirmed)"
            } else {
                # Prefer TMDB as it's authoritative, but note the discrepancy
                $finalYear = $tmdbYear
                $yearSource = "TMDB (NFO had $nfoYear)"
            }
        } elseif ($tmdbYear) {
            $finalYear = $tmdbYear
            $yearSource = "TMDB"
        } elseif ($nfoYear) {
            $finalYear = $nfoYear
            $yearSource = "NFO only"
        }

        if ($finalYear) {
            # Use TMDB title if available (preserves proper casing), otherwise use cleaned folder name
            $displayTitle = if ($tmdbResult -and $tmdbResult.Title) { $tmdbResult.Title } else { $cleanTitle }

            $foldersToFix += @{
                Folder = $folder
                CleanTitle = $displayTitle
                Year = $finalYear
                YearSource = $yearSource
                NFOFile = $nfoFile
                NFOMetadata = $nfoMetadata
                TMDBResult = $tmdbResult
            }
        } else {
            $noYearFound += @{
                Folder = $folder
                CleanTitle = $cleanTitle
            }
        }
    }

    # Display summary
    Write-Host "`n=== Scan Results ===" -ForegroundColor Cyan
    Write-Host "Folders to fix: $($foldersToFix.Count)" -ForegroundColor Green
    Write-Host "No year found: $($noYearFound.Count)" -ForegroundColor Yellow

    if ($noYearFound.Count -gt 0) {
        Write-Host "`nFolders where no year could be determined:" -ForegroundColor Yellow
        $noYearFound | Select-Object -First 10 | ForEach-Object {
            Write-Host "  - $($_.Folder.Name)" -ForegroundColor Gray
        }
        if ($noYearFound.Count -gt 10) {
            Write-Host "  ... and $($noYearFound.Count - 10) more" -ForegroundColor Gray
        }
    }

    if ($foldersToFix.Count -eq 0) {
        Write-Host "`nNo folders need fixing!" -ForegroundColor Green
        return
    }

    # Show what will be changed
    Write-Host "`n=== Proposed Changes ===" -ForegroundColor Cyan
    foreach ($item in $foldersToFix) {
        $safeTitle = Get-SafeFileName -Name $item.CleanTitle
        $newName = "$safeTitle ($($item.Year))"
        Write-Host "`n  Current: " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Folder.Name)" -ForegroundColor White
        Write-Host "  New:     " -NoNewline -ForegroundColor Gray
        Write-Host "$newName" -ForegroundColor Green
        Write-Host "  Source:  " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.YearSource)" -ForegroundColor Cyan
    }

    if ($WhatIf) {
        Write-Host "`n[DRY-RUN] No changes made. Run without -WhatIf to apply changes." -ForegroundColor Yellow
        Write-Log "Dry-run completed: $($foldersToFix.Count) folders would be fixed" "INFO"
        return
    }

    # Confirm before proceeding
    Write-Host ""
    $confirm = Read-Host "Proceed with renaming $($foldersToFix.Count) folder(s)? (Y/N)"
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }

    # Apply changes
    Write-Host "`nApplying changes..." -ForegroundColor Yellow

    foreach ($item in $foldersToFix) {
        # Sanitize title for Windows filename (replace illegal characters like : / \ etc.)
        $safeTitle = Get-SafeFileName -Name $item.CleanTitle
        $newName = "$safeTitle ($($item.Year))"
        $newPath = Join-Path (Split-Path $item.Folder.FullName -Parent) $newName

        try {
            # Check if target already exists
            if (Test-Path $newPath) {
                Write-Host "  Skipped (target exists): $($item.Folder.Name)" -ForegroundColor Yellow
                $skipped++
                continue
            }

            # Rename the folder
            Rename-Item -Path $item.Folder.FullName -NewName $newName -ErrorAction Stop
            Write-Host "  Renamed: $($item.Folder.Name) -> $newName" -ForegroundColor Green
            Write-Log "Renamed folder: $($item.Folder.FullName) -> $newPath" "INFO"
            $fixed++

            # Update or create NFO file
            $videoFile = Get-ChildItem -Path $newPath -File -ErrorAction SilentlyContinue |
                         Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                         Select-Object -First 1

            if ($videoFile) {
                # Get full TMDB details if we have a result
                $fullMetadata = $null
                if ($item.TMDBResult -and $item.TMDBResult.TMDBID) {
                    $fullMetadata = Get-TMDBMovieDetails -MovieId $item.TMDBResult.TMDBID -ApiKey $script:Config.TMDBApiKey
                }

                # Create/update NFO with best available data
                $nfoPath = Join-Path $newPath "$([System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)).nfo"

                if ($fullMetadata) {
                    $null = New-MovieNFOFromTMDB -Metadata $fullMetadata -NFOPath $nfoPath
                } elseif (-not $item.NFOFile) {
                    # Create basic NFO if none exists
                    New-MovieNFO -VideoPath $videoFile.FullName -Title $item.CleanTitle -Year $item.Year
                }
            }
        }
        catch {
            Write-Host "  Failed: $($item.Folder.Name) - $_" -ForegroundColor Red
            Write-Log "Failed to rename folder $($item.Folder.FullName): $_" "ERROR"
            $failed++
        }
    }

    # Final summary
    Write-Host "`n=== Summary ===" -ForegroundColor Cyan
    Write-Host "Fixed: $fixed" -ForegroundColor Green
    Write-Host "Skipped: $skipped" -ForegroundColor Yellow
    Write-Host "Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Green' })
    Write-Host "No year found: $($noYearFound.Count)" -ForegroundColor Yellow

    Write-Log "Folder year fix completed: Fixed=$fixed, Skipped=$skipped, Failed=$failed, NoYear=$($noYearFound.Count)" "INFO"
}

<#
.SYNOPSIS
    Creates a Kodi-compatible NFO file from TMDB metadata
.PARAMETER Metadata
    Hashtable containing TMDB movie metadata
.PARAMETER NFOPath
    The path where the NFO file should be created
#>
function New-MovieNFOFromTMDB {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Metadata,
        [Parameter(Mandatory=$true)]
        [string]$NFOPath,
        [string]$ErrorReasonVar = $null
    )

    try {
        # Verify the parent directory exists
        $parentDir = Split-Path $NFOPath -Parent
        if (-not (Test-Path $parentDir)) {
            throw "Parent directory does not exist: $parentDir"
        }

        # Escape XML special characters in text content
        $title = [System.Security.SecurityElement]::Escape($Metadata.Title)
        $originalTitle = [System.Security.SecurityElement]::Escape($Metadata.OriginalTitle)
        $plot = [System.Security.SecurityElement]::Escape($Metadata.Overview)
        $tagline = [System.Security.SecurityElement]::Escape($Metadata.Tagline)

        $genresXml = ""
        if ($Metadata.Genres) {
            $genresXml = ($Metadata.Genres | ForEach-Object {
                $escaped = [System.Security.SecurityElement]::Escape($_)
                "    <genre>$escaped</genre>"
            }) -join "`n"
        }

        $studiosXml = ""
        if ($Metadata.Studios) {
            $studiosXml = ($Metadata.Studios | Select-Object -First 3 | ForEach-Object {
                $escaped = [System.Security.SecurityElement]::Escape($_)
                "    <studio>$escaped</studio>"
            }) -join "`n"
        }

        $directorsXml = ""
        if ($Metadata.Directors) {
            $directorsXml = ($Metadata.Directors | ForEach-Object {
                $escaped = [System.Security.SecurityElement]::Escape($_)
                "    <director>$escaped</director>"
            }) -join "`n"
        }

        $actorsXml = ""
        if ($Metadata.Actors) {
            $actorsXml = ($Metadata.Actors | Select-Object -First 5 | ForEach-Object {
                $name = [System.Security.SecurityElement]::Escape($_.Name)
                $role = [System.Security.SecurityElement]::Escape($_.Role)
                "    <actor>`n        <name>$name</name>`n        <role>$role</role>`n    </actor>"
            }) -join "`n"
        }

        $nfoContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<movie>
    <title>$title</title>
    <originaltitle>$originalTitle</originaltitle>
    <year>$($Metadata.Year)</year>
    <plot>$plot</plot>
    <tagline>$tagline</tagline>
    <runtime>$($Metadata.Runtime)</runtime>
    <rating>$($Metadata.VoteAverage)</rating>
    <votes>$($Metadata.VoteCount)</votes>
    <uniqueid type="tmdb" default="true">$($Metadata.TMDBID)</uniqueid>
    <uniqueid type="imdb">$($Metadata.IMDBID)</uniqueid>
$genresXml
$studiosXml
$directorsXml
$actorsXml
</movie>
"@

        $nfoContent | Out-File -LiteralPath $NFOPath -Encoding UTF8 -Force
        Write-Log "Created NFO from TMDB: $NFOPath" "INFO"
        $script:Stats.NFOFilesCreated++
        return $true
    }
    catch {
        Write-Host "  Error creating NFO: $_" -ForegroundColor Red
        Write-Log "Error creating NFO from TMDB: $_" "ERROR"
        if ($ErrorReasonVar) {
            Set-Variable -Name $ErrorReasonVar -Value $_.Exception.Message -Scope 1
        }
        return $false
    }
}

#============================================
# DUPLICATE DETECTION & QUALITY SCORING
#============================================

<#
.SYNOPSIS
    Detects quality concerns based on bitrate, resolution, and file size
.DESCRIPTION
    Flags files that may have quality issues:
    - Low bitrate for resolution (1080p under 2.5 Mbps, 4K under 8 Mbps)
    - Suspiciously small file size for runtime (< 1GB for 90+ min movie)
    - Source/bitrate mismatch (BluRay source with WEB-level bitrate)
.PARAMETER Quality
    The quality hashtable from Get-QualityScore
.PARAMETER FileName
    The original filename for source detection
.OUTPUTS
    Array of concern strings (empty if no concerns)
#>
function Get-QualityConcerns {
    param(
        [hashtable]$Quality,
        [string]$FileName
    )

    $concerns = @()
    $fileNameLower = $FileName.ToLower()

    # Calculate bitrate in Mbps
    $bitrateMbps = if ($Quality.Bitrate -gt 0) { [math]::Round($Quality.Bitrate / 1000000, 2) } else { 0 }

    # Resolution-based bitrate thresholds
    if ($bitrateMbps -gt 0) {
        switch ($Quality.Resolution) {
            "2160p" {
                if ($bitrateMbps -lt 8) {
                    $concerns += "Very low bitrate for 4K (${bitrateMbps} Mbps, expect 15+ Mbps)"
                } elseif ($bitrateMbps -lt 12) {
                    $concerns += "Low bitrate for 4K (${bitrateMbps} Mbps, expect 15+ Mbps)"
                }
            }
            "1080p" {
                if ($bitrateMbps -lt 2) {
                    $concerns += "Very low bitrate for 1080p (${bitrateMbps} Mbps, expect 5+ Mbps)"
                } elseif ($bitrateMbps -lt 3) {
                    $concerns += "Low bitrate for 1080p (${bitrateMbps} Mbps, expect 5+ Mbps)"
                }
            }
            "720p" {
                if ($bitrateMbps -lt 1) {
                    $concerns += "Very low bitrate for 720p (${bitrateMbps} Mbps, expect 2+ Mbps)"
                }
            }
        }
    }

    # File size vs duration check (if we have both)
    if ($Quality.FileSize -gt 0 -and $Quality.Duration -gt 0) {
        $durationMinutes = $Quality.Duration / 60000  # Duration is in milliseconds
        $fileSizeGB = $Quality.FileSize / 1GB

        # Flag movies over 80 minutes that are under 1GB
        if ($durationMinutes -ge 80 -and $fileSizeGB -lt 1) {
            $concerns += "Small file for runtime ($([math]::Round($fileSizeGB, 2)) GB for $([math]::Round($durationMinutes, 0)) min)"
        }

        # Flag 4K movies over 80 minutes that are under 5GB
        if ($Quality.Resolution -eq "2160p" -and $durationMinutes -ge 80 -and $fileSizeGB -lt 5) {
            $concerns += "Small 4K file ($([math]::Round($fileSizeGB, 2)) GB for $([math]::Round($durationMinutes, 0)) min, expect 10+ GB)"
        }
    }

    # Source/bitrate mismatch - only flag if it appears to be a full rip, not an encode
    if ($bitrateMbps -gt 0) {
        # Check if this looks like an encode (has encoder/quality tags indicating re-encoding)
        $isEncode = $fileNameLower -match 'x264|x265|hevc|h\.?264|h\.?265|avc|web-?dl|web-?rip|hdrip|dvdrip|brrip|yify|yts|rarbg|\d{3,4}p'

        # Remux should always have high bitrate (these are direct disc copies)
        if ($fileNameLower -match 'remux' -and $bitrateMbps -lt 15) {
            $concerns += "Low bitrate for Remux (${bitrateMbps} Mbps, expect 20+ Mbps)"
        }
        # Only flag BluRay source mismatch if it's NOT an encode
        # An encode with "BluRay" just means BluRay was the source - lower bitrate is expected
        elseif (-not $isEncode -and $fileNameLower -match 'bluray|blu-ray' -and $bitrateMbps -lt 10) {
            $concerns += "Low bitrate for apparent BluRay rip (${bitrateMbps} Mbps) - if this is an encode, ignore"
        }
    }

    # File size only check (when we don't have MediaInfo duration)
    if ($Quality.FileSize -gt 0 -and $Quality.Duration -eq 0) {
        $fileSizeGB = $Quality.FileSize / 1GB

        # Very small files are suspicious for any resolution
        if ($Quality.Resolution -eq "1080p" -and $fileSizeGB -lt 0.7) {
            $concerns += "Very small 1080p file ($([math]::Round($fileSizeGB, 2)) GB)"
        }
        elseif ($Quality.Resolution -eq "2160p" -and $fileSizeGB -lt 2) {
            $concerns += "Very small 4K file ($([math]::Round($fileSizeGB, 2)) GB)"
        }
    }

    return $concerns
}

<#
.SYNOPSIS
    Calculates a quality score for a video file based on its properties
.PARAMETER FileName
    The filename to analyze
.OUTPUTS
    Hashtable with Score, Resolution, Codec, Source, and details
.DESCRIPTION
    Analyzes filename for quality indicators and assigns a score:
    - Resolution: 2160p=100, 1080p=80, 720p=60, 480p=40
    - Source: BluRay=30, WEB-DL=25, WEBRip=20, HDTV=15, DVDRip=10
    - Codec: x265/HEVC=20, x264=15, XviD=5
    - Audio: Atmos=15, TrueHD=12, DTS-HD=10, DTS=8, AC3=5, AAC=3
    - HDR: HDR10+=15, HDR10=12, HDR=10, DolbyVision=15
#>
function Get-QualityScore {
    param(
        [string]$FileName,
        [string]$FilePath = $null
    )

    $quality = @{
        Score = 0
        Resolution = "Unknown"
        Codec = "Unknown"
        Source = "Unknown"
        Audio = "Unknown"
        HDR = $false
        HDRFormat = $null
        Bitrate = 0
        Width = 0
        Height = 0
        AudioChannels = 0
        Details = @()
        DataSource = "Filename"
        QualityConcerns = @()
        Duration = 0
        FileSize = 0
    }

    # Try MediaInfo first if FilePath is provided
    $mediaInfo = $null
    if ($FilePath -and (Test-Path $FilePath -PathType Leaf)) {
        $mediaInfo = Get-MediaInfoDetails -FilePath $FilePath
    }

    if ($mediaInfo) {
        $quality.DataSource = "MediaInfo"
        $quality.Width = $mediaInfo.Width
        $quality.Height = $mediaInfo.Height
        $quality.Bitrate = $mediaInfo.Bitrate
        $quality.AudioChannels = $mediaInfo.AudioChannels
        $quality.Duration = $mediaInfo.Duration

        # Get file size
        if ($FilePath -and (Test-Path $FilePath)) {
            $quality.FileSize = (Get-Item $FilePath).Length
        }

        # Resolution from actual dimensions
        if ($mediaInfo.Height -ge 2160) {
            $quality.Resolution = "2160p"
            $quality.Score += 100
            $quality.Details += "4K/2160p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+100)"
        }
        elseif ($mediaInfo.Height -ge 1080) {
            $quality.Resolution = "1080p"
            $quality.Score += 80
            $quality.Details += "1080p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+80)"
        }
        elseif ($mediaInfo.Height -ge 720) {
            $quality.Resolution = "720p"
            $quality.Score += 60
            $quality.Details += "720p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+60)"
        }
        elseif ($mediaInfo.Height -ge 480) {
            $quality.Resolution = "480p"
            $quality.Score += 40
            $quality.Details += "480p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+40)"
        }
        elseif ($mediaInfo.Height -gt 0) {
            $quality.Resolution = "$($mediaInfo.Height)p"
            $quality.Score += 20
            $quality.Details += "$($mediaInfo.Height)p [MediaInfo] (+20)"
        }

        # Video codec from MediaInfo
        $videoCodec = $mediaInfo.VideoCodec
        if ($videoCodec) {
            switch -Regex ($videoCodec) {
                'HEVC|H\.?265|V_MPEGH' {
                    $quality.Codec = "HEVC/x265"
                    $quality.Score += 20
                    $quality.Details += "HEVC/x265 [MediaInfo: $videoCodec] (+20)"
                }
                'AVC|H\.?264|V_MPEG4/ISO/AVC' {
                    $quality.Codec = "x264"
                    $quality.Score += 15
                    $quality.Details += "x264 [MediaInfo: $videoCodec] (+15)"
                }
                'AV1' {
                    $quality.Codec = "AV1"
                    $quality.Score += 25
                    $quality.Details += "AV1 [MediaInfo] (+25)"
                }
                'VP9' {
                    $quality.Codec = "VP9"
                    $quality.Score += 18
                    $quality.Details += "VP9 [MediaInfo] (+18)"
                }
                'MPEG-4|DivX|XviD' {
                    $quality.Codec = "XviD"
                    $quality.Score += 5
                    $quality.Details += "MPEG-4/XviD [MediaInfo: $videoCodec] (+5)"
                }
                'VC-1|WMV' {
                    $quality.Codec = "VC-1"
                    $quality.Score += 8
                    $quality.Details += "VC-1 [MediaInfo] (+8)"
                }
                default {
                    $quality.Codec = $videoCodec
                    $quality.Score += 10
                    $quality.Details += "$videoCodec [MediaInfo] (+10)"
                }
            }
        }

        # Audio codec from MediaInfo
        $audioCodec = $mediaInfo.AudioCodec
        $channels = $mediaInfo.AudioChannels
        if ($audioCodec) {
            $channelInfo = if ($channels -gt 0) { " ${channels}ch" } else { "" }
            switch -Regex ($audioCodec) {
                'Atmos|E-AC-3.*Atmos|TrueHD.*Atmos' {
                    $quality.Audio = "Atmos"
                    $quality.Score += 15
                    $quality.Details += "Atmos [MediaInfo$channelInfo] (+15)"
                }
                'TrueHD' {
                    $quality.Audio = "TrueHD"
                    $quality.Score += 12
                    $quality.Details += "TrueHD [MediaInfo$channelInfo] (+12)"
                }
                'DTS-HD|DTS.*HD' {
                    $quality.Audio = "DTS-HD"
                    $quality.Score += 10
                    $quality.Details += "DTS-HD [MediaInfo$channelInfo] (+10)"
                }
                'DTS.*X|DTS:X' {
                    $quality.Audio = "DTS:X"
                    $quality.Score += 14
                    $quality.Details += "DTS:X [MediaInfo$channelInfo] (+14)"
                }
                '^DTS$|^DTS\s' {
                    $quality.Audio = "DTS"
                    $quality.Score += 8
                    $quality.Details += "DTS [MediaInfo$channelInfo] (+8)"
                }
                'E-AC-3|EAC3|DD\+|Dolby Digital Plus' {
                    $quality.Audio = "EAC3"
                    $quality.Score += 7
                    $quality.Details += "EAC3/DD+ [MediaInfo$channelInfo] (+7)"
                }
                'AC-3|AC3|Dolby Digital' {
                    $quality.Audio = "AC3"
                    $quality.Score += 5
                    $quality.Details += "AC3 [MediaInfo$channelInfo] (+5)"
                }
                'AAC' {
                    $quality.Audio = "AAC"
                    $quality.Score += 3
                    $quality.Details += "AAC [MediaInfo$channelInfo] (+3)"
                }
                'FLAC' {
                    $quality.Audio = "FLAC"
                    $quality.Score += 6
                    $quality.Details += "FLAC [MediaInfo$channelInfo] (+6)"
                }
                'PCM|LPCM' {
                    $quality.Audio = "PCM"
                    $quality.Score += 4
                    $quality.Details += "PCM [MediaInfo$channelInfo] (+4)"
                }
                'Opus' {
                    $quality.Audio = "Opus"
                    $quality.Score += 4
                    $quality.Details += "Opus [MediaInfo$channelInfo] (+4)"
                }
                'Vorbis' {
                    $quality.Audio = "Vorbis"
                    $quality.Score += 2
                    $quality.Details += "Vorbis [MediaInfo$channelInfo] (+2)"
                }
                'MP3|MPEG Audio' {
                    $quality.Audio = "MP3"
                    $quality.Score += 1
                    $quality.Details += "MP3 [MediaInfo$channelInfo] (+1)"
                }
                default {
                    $quality.Audio = $audioCodec
                    $quality.Score += 2
                    $quality.Details += "$audioCodec [MediaInfo$channelInfo] (+2)"
                }
            }
        }

        # HDR detection from MediaInfo
        if ($mediaInfo.HDR) {
            $quality.HDR = $true
            # Get detailed HDR format if available
            $hdrFormat = Get-MediaInfoHDRFormat -FilePath $FilePath
            if ($hdrFormat) {
                $quality.HDRFormat = $hdrFormat.Format
                switch ($hdrFormat.Format) {
                    "Dolby Vision" {
                        $quality.Score += 18
                        $quality.Details += "Dolby Vision [MediaInfo] (+18)"
                    }
                    "HDR10+" {
                        $quality.Score += 16
                        $quality.Details += "HDR10+ [MediaInfo] (+16)"
                    }
                    "HDR10" {
                        $quality.Score += 12
                        $quality.Details += "HDR10 [MediaInfo] (+12)"
                    }
                    "HLG" {
                        $quality.Score += 10
                        $quality.Details += "HLG [MediaInfo] (+10)"
                    }
                    default {
                        $quality.Score += 10
                        $quality.Details += "HDR [MediaInfo] (+10)"
                    }
                }
            }
            else {
                $quality.Score += 10
                $quality.Details += "HDR [MediaInfo] (+10)"
            }
        }

        # Source detection still from filename (MediaInfo can't detect source)
        $fileNameLower = $FileName.ToLower()
        if ($fileNameLower -match 'bluray|blu-ray|bdrip|brrip') {
            $quality.Source = "BluRay"
            $quality.Score += 30
            $quality.Details += "BluRay (+30)"
        }
        elseif ($fileNameLower -match 'remux') {
            $quality.Source = "Remux"
            $quality.Score += 35
            $quality.Details += "Remux (+35)"
        }
        elseif ($fileNameLower -match 'web-dl|webdl') {
            $quality.Source = "WEB-DL"
            $quality.Score += 25
            $quality.Details += "WEB-DL (+25)"
        }
        elseif ($fileNameLower -match 'webrip') {
            $quality.Source = "WEBRip"
            $quality.Score += 20
            $quality.Details += "WEBRip (+20)"
        }
        elseif ($fileNameLower -match 'hdtv') {
            $quality.Source = "HDTV"
            $quality.Score += 15
            $quality.Details += "HDTV (+15)"
        }
        elseif ($fileNameLower -match 'dvdrip') {
            $quality.Source = "DVDRip"
            $quality.Score += 10
            $quality.Details += "DVDRip (+10)"
        }

        # Bitrate bonus (higher bitrate = better quality)
        if ($quality.Bitrate -gt 0) {
            $bitrateMbps = [math]::Round($quality.Bitrate / 1000000, 1)
            if ($bitrateMbps -ge 40) {
                $quality.Score += 20
                $quality.Details += "High Bitrate [${bitrateMbps} Mbps] (+20)"
            }
            elseif ($bitrateMbps -ge 20) {
                $quality.Score += 15
                $quality.Details += "Good Bitrate [${bitrateMbps} Mbps] (+15)"
            }
            elseif ($bitrateMbps -ge 10) {
                $quality.Score += 10
                $quality.Details += "Moderate Bitrate [${bitrateMbps} Mbps] (+10)"
            }
            elseif ($bitrateMbps -ge 5) {
                $quality.Score += 5
                $quality.Details += "Low Bitrate [${bitrateMbps} Mbps] (+5)"
            }
        }

        # Quality Concerns detection (MediaInfo path)
        $quality.QualityConcerns = @(Get-QualityConcerns -Quality $quality -FileName $FileName)

        return $quality
    }

    # Fallback to filename parsing if MediaInfo not available
    $quality.DataSource = "Filename"
    $fileNameLower = $FileName.ToLower()

    # Resolution scoring
    if ($fileNameLower -match '2160p|4k|uhd') {
        $quality.Resolution = "2160p"
        $quality.Score += 100
        $quality.Details += "4K/2160p (+100)"
    }
    elseif ($fileNameLower -match '1080p') {
        $quality.Resolution = "1080p"
        $quality.Score += 80
        $quality.Details += "1080p (+80)"
    }
    elseif ($fileNameLower -match '720p') {
        $quality.Resolution = "720p"
        $quality.Score += 60
        $quality.Details += "720p (+60)"
    }
    elseif ($fileNameLower -match '480p|dvd') {
        $quality.Resolution = "480p"
        $quality.Score += 40
        $quality.Details += "480p (+40)"
    }

    # Source scoring
    if ($fileNameLower -match 'remux') {
        $quality.Source = "Remux"
        $quality.Score += 35
        $quality.Details += "Remux (+35)"
    }
    elseif ($fileNameLower -match 'bluray|blu-ray|bdrip|brrip') {
        $quality.Source = "BluRay"
        $quality.Score += 30
        $quality.Details += "BluRay (+30)"
    }
    elseif ($fileNameLower -match 'web-dl|webdl') {
        $quality.Source = "WEB-DL"
        $quality.Score += 25
        $quality.Details += "WEB-DL (+25)"
    }
    elseif ($fileNameLower -match 'webrip') {
        $quality.Source = "WEBRip"
        $quality.Score += 20
        $quality.Details += "WEBRip (+20)"
    }
    elseif ($fileNameLower -match 'hdtv') {
        $quality.Source = "HDTV"
        $quality.Score += 15
        $quality.Details += "HDTV (+15)"
    }
    elseif ($fileNameLower -match 'dvdrip') {
        $quality.Source = "DVDRip"
        $quality.Score += 10
        $quality.Details += "DVDRip (+10)"
    }

    # Codec scoring
    if ($fileNameLower -match 'av1') {
        $quality.Codec = "AV1"
        $quality.Score += 25
        $quality.Details += "AV1 (+25)"
    }
    elseif ($fileNameLower -match 'x265|h\.?265|hevc') {
        $quality.Codec = "HEVC/x265"
        $quality.Score += 20
        $quality.Details += "HEVC/x265 (+20)"
    }
    elseif ($fileNameLower -match 'vp9') {
        $quality.Codec = "VP9"
        $quality.Score += 18
        $quality.Details += "VP9 (+18)"
    }
    elseif ($fileNameLower -match 'x264|h\.?264|avc') {
        $quality.Codec = "x264"
        $quality.Score += 15
        $quality.Details += "x264 (+15)"
    }
    elseif ($fileNameLower -match 'xvid|divx') {
        $quality.Codec = "XviD"
        $quality.Score += 5
        $quality.Details += "XviD (+5)"
    }

    # Audio scoring
    if ($fileNameLower -match 'atmos') {
        $quality.Audio = "Atmos"
        $quality.Score += 15
        $quality.Details += "Atmos (+15)"
    }
    elseif ($fileNameLower -match 'dts[\s\.\-]?x|dtsx') {
        $quality.Audio = "DTS:X"
        $quality.Score += 14
        $quality.Details += "DTS:X (+14)"
    }
    elseif ($fileNameLower -match 'truehd') {
        $quality.Audio = "TrueHD"
        $quality.Score += 12
        $quality.Details += "TrueHD (+12)"
    }
    elseif ($fileNameLower -match 'dts-hd|dtshd|dts[\s\.\-]?hd[\s\.\-]?ma') {
        $quality.Audio = "DTS-HD"
        $quality.Score += 10
        $quality.Details += "DTS-HD (+10)"
    }
    elseif ($fileNameLower -match 'dts') {
        $quality.Audio = "DTS"
        $quality.Score += 8
        $quality.Details += "DTS (+8)"
    }
    elseif ($fileNameLower -match 'eac3|ddp|dd\+|dolby\s*digital\s*plus') {
        $quality.Audio = "EAC3"
        $quality.Score += 7
        $quality.Details += "EAC3/DD+ (+7)"
    }
    elseif ($fileNameLower -match 'ac3|dd5\.?1') {
        $quality.Audio = "AC3"
        $quality.Score += 5
        $quality.Details += "AC3 (+5)"
    }
    elseif ($fileNameLower -match 'flac') {
        $quality.Audio = "FLAC"
        $quality.Score += 6
        $quality.Details += "FLAC (+6)"
    }
    elseif ($fileNameLower -match 'aac') {
        $quality.Audio = "AAC"
        $quality.Score += 3
        $quality.Details += "AAC (+3)"
    }
    elseif ($fileNameLower -match 'opus') {
        $quality.Audio = "Opus"
        $quality.Score += 4
        $quality.Details += "Opus (+4)"
    }

    # HDR scoring
    if ($fileNameLower -match 'dolby[\s\.\-]?vision|dovi|dv[\s\.\-]hdr|\.dv\.') {
        $quality.HDR = $true
        $quality.HDRFormat = "Dolby Vision"
        $quality.Score += 18
        $quality.Details += "Dolby Vision (+18)"
    }
    elseif ($fileNameLower -match 'hdr10\+|hdr10plus') {
        $quality.HDR = $true
        $quality.HDRFormat = "HDR10+"
        $quality.Score += 16
        $quality.Details += "HDR10+ (+16)"
    }
    elseif ($fileNameLower -match 'hdr10') {
        $quality.HDR = $true
        $quality.HDRFormat = "HDR10"
        $quality.Score += 12
        $quality.Details += "HDR10 (+12)"
    }
    elseif ($fileNameLower -match 'hlg') {
        $quality.HDR = $true
        $quality.HDRFormat = "HLG"
        $quality.Score += 10
        $quality.Details += "HLG (+10)"
    }
    elseif ($fileNameLower -match 'hdr') {
        $quality.HDR = $true
        $quality.HDRFormat = "HDR"
        $quality.Score += 10
        $quality.Details += "HDR (+10)"
    }

    # Get file size for filename-based analysis
    if ($FilePath -and (Test-Path $FilePath)) {
        $quality.FileSize = (Get-Item $FilePath).Length
    }

    # Quality Concerns detection (Filename path - limited without MediaInfo)
    $quality.QualityConcerns = @(Get-QualityConcerns -Quality $quality -FileName $FileName)

    return $quality
}

<#
.SYNOPSIS
    Extracts a normalized title from a movie folder or filename for duplicate matching
.PARAMETER Name
    The folder or file name to normalize
.OUTPUTS
    Hashtable with NormalizedTitle and Year
#>
function Get-NormalizedTitle {
    param(
        [string]$Name,
        [switch]$Strict  # If set, does minimal normalization for stricter matching
    )

    $result = @{
        NormalizedTitle = $null
        Year = $null
    }

    # Remove extension if present
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Name)

    # Try to extract year (must be in parentheses, brackets, or preceded by space for strict matching)
    if ($name -match '[\(\[\s](19|20)\d{2}[\)\]\s]?') {
        $yearMatch = [regex]::Match($name, '[\(\[\s](19|20\d{2})[\)\]\s]?')
        if ($yearMatch.Success) {
            $result.Year = [regex]::Match($yearMatch.Value, '(19|20)\d{2}').Value
        }
    }

    # Remove everything after year or quality tags
    $title = $name -replace '[\(\[]?(19|20)\d{2}[\)\]]?.*$', ''
    $title = $title -replace '\s*(720p|1080p|2160p|4K|HDRip|DVDRip|BRRip|BluRay|WEB-DL|WEBRip|x264|x265|HEVC|AAC|AC3|DTS|10bit|H\.?264|H\.?265).*$', ''

    # Normalize the title
    $title = $title -replace '\.', ' '
    $title = $title -replace '[_]', ' '
    # Only replace hyphens surrounded by spaces (not in titles like "Spider-Man")
    $title = $title -replace '\s+-\s+', ' '
    $title = $title -replace '\s+', ' '
    $title = $title.Trim().ToLower()

    if (-not $Strict) {
        # For non-strict matching, remove leading articles
        $title = $title -replace '^(the|a|an)\s+', ''
    }

    $result.NormalizedTitle = $title

    return $result
}

<#
.SYNOPSIS
    Finds potential duplicate movies in the library (basic title matching)
.PARAMETER Path
    The root path of the movie library
.OUTPUTS
    Array of duplicate groups, each containing matching movies
.DESCRIPTION
    Basic duplicate detection using title normalization only.
    For more accurate detection using NFO metadata (IMDB/TMDB IDs) and file hashing,
    use Find-DuplicateMoviesEnhanced instead.
#>
function Find-DuplicateMovies {
    param(
        [string]$Path
    )

    Write-Host "`nScanning for duplicate movies..." -ForegroundColor Yellow
    Write-Log "Starting duplicate scan in: $Path" "INFO"

    try {
        $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_Trailers' }

        if ($movieFolders.Count -eq 0) {
            Write-Host "No movie folders found" -ForegroundColor Cyan
            return @()
        }

        # Build a lookup of normalized titles
        $titleLookup = @{}
        $totalFolders = $movieFolders.Count
        $currentIndex = 0

        foreach ($folder in $movieFolders) {
            $currentIndex++
            $percentComplete = [math]::Round(($currentIndex / $totalFolders) * 100)
            Write-Progress -Activity "Scanning for duplicates" -Status "Processing $currentIndex of $totalFolders - $($folder.Name)" -PercentComplete $percentComplete

            $titleInfo = Get-NormalizedTitle -Name $folder.Name

            # Find the main video file for size info
            $videoFile = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                Sort-Object Length -Descending |
                Select-Object -First 1

            $quality = Get-QualityScore -FileName $folder.Name -FilePath $(if ($videoFile) { $videoFile.FullName } else { $null })

            $entry = @{
                Path = $folder.FullName
                OriginalName = $folder.Name
                Year = $titleInfo.Year
                Quality = $quality
                FileSize = if ($videoFile) { $videoFile.Length } else { 0 }
            }

            $key = $titleInfo.NormalizedTitle
            if ($titleInfo.Year) {
                $key = "$($titleInfo.NormalizedTitle)|$($titleInfo.Year)"
            }

            if (-not $titleLookup.ContainsKey($key)) {
                $titleLookup[$key] = @()
            }
            $titleLookup[$key] += $entry
        }

        Write-Progress -Activity "Scanning for duplicates" -Completed

        # Find duplicates (groups with more than 1 entry)
        $duplicates = @()
        foreach ($key in $titleLookup.Keys) {
            if ($titleLookup[$key].Count -gt 1) {
                $duplicates += ,@($titleLookup[$key])
            }
        }

        return $duplicates
    }
    catch {
        Write-Host "Error scanning for duplicates: $_" -ForegroundColor Red
        Write-Log "Error scanning for duplicates: $_" "ERROR"
        return @()
    }
}

<#
.SYNOPSIS
    Displays duplicate movies and suggests which to keep based on quality
.PARAMETER Path
    The root path of the movie library
#>
function Show-DuplicateReport {
    param(
        [string]$Path
    )

    $duplicates = Find-DuplicateMovies -Path $Path

    if ($duplicates.Count -eq 0) {
        Write-Host "No duplicate movies found!" -ForegroundColor Green
        Write-Log "No duplicates found" "INFO"
        return
    }

    Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║                    DUPLICATE REPORT                          ║" -ForegroundColor Yellow
    Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
    Write-Host "║  Found $($duplicates.Count) potential duplicate group(s)".PadRight(63) + "║" -ForegroundColor Yellow
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow

    $groupNum = 1
    foreach ($group in $duplicates) {
        Write-Host "`n[$groupNum] Potential Duplicates:" -ForegroundColor Cyan

        # Sort by quality score descending
        $sorted = $group | Sort-Object { $_.Quality.Score } -Descending

        $first = $true
        foreach ($movie in $sorted) {
            $sizeStr = Format-FileSize $movie.FileSize
            $scoreStr = "Score: $($movie.Quality.Score)"

            if ($first) {
                Write-Host "  [KEEP] " -ForegroundColor Green -NoNewline
                $first = $false
            } else {
                Write-Host "  [DEL?] " -ForegroundColor Red -NoNewline
            }

            Write-Host "$($movie.OriginalName)" -ForegroundColor White
            Write-Host "         $scoreStr | $sizeStr | $($movie.Quality.Resolution) $($movie.Quality.Source)" -ForegroundColor Gray
        }

        $groupNum++
    }

    Write-Host "`n" -NoNewline
    Write-Log "Found $($duplicates.Count) duplicate groups" "INFO"
}

<#
.SYNOPSIS
    Interactively removes duplicate movies, keeping highest quality versions
.PARAMETER Path
    The root path of the movie library
#>
function Remove-DuplicateMovies {
    param(
        [string]$Path
    )

    $duplicates = Find-DuplicateMovies -Path $Path

    if ($duplicates.Count -eq 0) {
        Write-Host "No duplicate movies found!" -ForegroundColor Green
        return
    }

    Write-Host "`nFound $($duplicates.Count) duplicate group(s)" -ForegroundColor Yellow

    $confirmAll = Read-Host "Delete lower quality duplicates? (Y/N/Review each) [N]"

    if ($confirmAll -eq 'Y' -or $confirmAll -eq 'y') {
        foreach ($group in $duplicates) {
            $sorted = $group | Sort-Object { $_.Quality.Score } -Descending
            # Keep first (highest quality), delete the rest
            $toDelete = $sorted | Select-Object -Skip 1

            foreach ($movie in $toDelete) {
                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would delete: $($movie.OriginalName)" -ForegroundColor Yellow
                    Write-Log "Would delete duplicate: $($movie.Path)" "DRY-RUN"
                } else {
                    try {
                        $size = (Get-ChildItem -Path $movie.Path -Recurse | Measure-Object -Property Length -Sum).Sum
                        Remove-Item -Path $movie.Path -Recurse -Force -ErrorAction Stop
                        Write-Host "Deleted: $($movie.OriginalName)" -ForegroundColor Red
                        Write-Log "Deleted duplicate: $($movie.Path)" "INFO"
                        $script:Stats.FilesDeleted++
                        $script:Stats.BytesDeleted += $size
                    }
                    catch {
                        Write-Host "Error deleting $($movie.OriginalName): $_" -ForegroundColor Red
                        Write-Log "Error deleting duplicate $($movie.Path): $_" "ERROR"
                    }
                }
            }
        }
    }
    elseif ($confirmAll -eq 'R' -or $confirmAll -eq 'r') {
        # Review each duplicate group
        foreach ($group in $duplicates) {
            $sorted = $group | Sort-Object { $_.Quality.Score } -Descending

            Write-Host "`nDuplicate Group:" -ForegroundColor Cyan
            $i = 1
            foreach ($movie in $sorted) {
                Write-Host "  [$i] $($movie.OriginalName) (Score: $($movie.Quality.Score))" -ForegroundColor White
                $i++
            }

            $choice = Read-Host "Enter number to KEEP (others deleted), or S to skip"
            if ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $sorted.Count) {
                $keepIndex = [int]$choice - 1
                for ($j = 0; $j -lt $sorted.Count; $j++) {
                    if ($j -ne $keepIndex) {
                        $movie = $sorted[$j]
                        if ($script:Config.DryRun) {
                            Write-Host "[DRY-RUN] Would delete: $($movie.OriginalName)" -ForegroundColor Yellow
                        } else {
                            try {
                                $size = (Get-ChildItem -Path $movie.Path -Recurse | Measure-Object -Property Length -Sum).Sum
                                Remove-Item -Path $movie.Path -Recurse -Force -ErrorAction Stop
                                Write-Host "Deleted: $($movie.OriginalName)" -ForegroundColor Red
                                $script:Stats.FilesDeleted++
                                $script:Stats.BytesDeleted += $size
                            }
                            catch {
                                Write-Host "Error: $_" -ForegroundColor Red
                            }
                        }
                    }
                }
            } else {
                Write-Host "Skipped" -ForegroundColor Cyan
            }
        }
    } else {
        Write-Host "No duplicates removed" -ForegroundColor Cyan
    }
}

<#
.SYNOPSIS
    Transfers processed movies from inbox to main library
.PARAMETER InboxPath
    The inbox path containing processed movies
.PARAMETER LibraryPath
    The main library path to move movies to
.DESCRIPTION
    Moves all movie folders from the inbox to the main library.
    Checks for existing folders and prompts for conflict resolution.
#>
function Move-MoviesToLibrary {
    param(
        [string]$InboxPath,
        [string]$LibraryPath
    )

    Write-Host "`n--- Transfer to Library ---" -ForegroundColor Cyan
    Write-Log "Starting transfer from $InboxPath to $LibraryPath" "INFO"

    # Get all movie folders in inbox
    $movieFolders = Get-ChildItem -LiteralPath $InboxPath -Directory -ErrorAction SilentlyContinue

    if ($movieFolders.Count -eq 0) {
        Write-Host "No movies to transfer" -ForegroundColor Gray
        return @{ Moved = 0; Skipped = 0; Upgraded = 0; Failed = 0 }
    }

    Write-Host "Found $($movieFolders.Count) movie(s) to transfer" -ForegroundColor White

    $stats = @{
        Moved = 0
        Skipped = 0
        Upgraded = 0
        Failed = 0
        BytesMoved = 0
    }

    # Collect potential upgrades for batch prompt
    $potentialUpgrades = @()

    foreach ($folder in $movieFolders) {
        $destPath = Join-Path $LibraryPath $folder.Name

        Write-Host "  $($folder.Name)" -ForegroundColor White -NoNewline

        # Check if destination already exists
        if (Test-Path -LiteralPath $destPath) {
            # Compare quality scores to see if this is an upgrade
            $inboxVideo = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '\.(mkv|mp4|avi|mov|wmv|m4v)$' } |
                Select-Object -First 1

            $libraryVideo = Get-ChildItem -LiteralPath $destPath -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '\.(mkv|mp4|avi|mov|wmv|m4v)$' } |
                Select-Object -First 1

            if ($inboxVideo -and $libraryVideo) {
                $inboxQuality = Get-QualityScore -FileName $inboxVideo.Name -FilePath $inboxVideo.FullName
                $libraryQuality = Get-QualityScore -FileName $libraryVideo.Name -FilePath $libraryVideo.FullName

                if ($inboxQuality.Score -gt $libraryQuality.Score) {
                    # This is a quality upgrade
                    $inboxSize = (Get-ChildItem -LiteralPath $folder.FullName -Recurse -File -ErrorAction SilentlyContinue |
                        Measure-Object -Property Length -Sum).Sum
                    $librarySize = (Get-ChildItem -LiteralPath $destPath -Recurse -File -ErrorAction SilentlyContinue |
                        Measure-Object -Property Length -Sum).Sum

                    Write-Host " [UPGRADE AVAILABLE]" -ForegroundColor Magenta
                    Write-Host "      Library: $($libraryQuality.Resolution) $($libraryQuality.Codec) (Score: $($libraryQuality.Score), $(Format-FileSize $librarySize))" -ForegroundColor DarkGray
                    Write-Host "      Inbox:   $($inboxQuality.Resolution) $($inboxQuality.Codec) (Score: $($inboxQuality.Score), $(Format-FileSize $inboxSize))" -ForegroundColor DarkCyan

                    $potentialUpgrades += @{
                        Folder = $folder
                        DestPath = $destPath
                        InboxQuality = $inboxQuality
                        LibraryQuality = $libraryQuality
                        InboxSize = $inboxSize
                        LibrarySize = $librarySize
                    }
                    continue
                }
            }

            Write-Host " [EXISTS - skipped]" -ForegroundColor Yellow
            Write-Log "Skipped $($folder.Name) - already exists in library" "WARNING"
            $stats.Skipped++
            continue
        }

        if ($script:Config.DryRun) {
            Write-Host " [DRY-RUN - would move]" -ForegroundColor Yellow
            Write-Log "Would move $($folder.Name) to library" "DRY-RUN"
            $stats.Moved++
            continue
        }

        try {
            # Calculate size before moving
            $folderSize = (Get-ChildItem -LiteralPath $folder.FullName -Recurse -File -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum

            # Move the folder
            Move-Item -LiteralPath $folder.FullName -Destination $destPath -Force -ErrorAction Stop

            Write-Host " [MOVED]" -ForegroundColor Green
            Write-Log "Moved $($folder.Name) to $destPath" "INFO"
            $stats.Moved++
            $stats.BytesMoved += $folderSize
        }
        catch {
            Write-Host " [FAILED: $_]" -ForegroundColor Red
            Write-Log "Failed to move $($folder.Name): $_" "ERROR"
            $stats.Failed++
        }
    }

    # Handle potential upgrades
    if ($potentialUpgrades.Count -gt 0 -and -not $script:Config.DryRun) {
        Write-Host "`n--- Quality Upgrades Available ---" -ForegroundColor Magenta
        Write-Host "Found $($potentialUpgrades.Count) movie(s) that would upgrade existing library versions." -ForegroundColor White
        Write-Host "Upgrading will replace the library version with the higher quality inbox version." -ForegroundColor Gray

        $upgradeChoice = Read-Host "`nApply quality upgrades? (Y/N/Review) [N]"

        if ($upgradeChoice -eq 'Y' -or $upgradeChoice -eq 'y') {
            foreach ($upgrade in $potentialUpgrades) {
                Write-Host "  Upgrading $($upgrade.Folder.Name)..." -ForegroundColor White -NoNewline
                try {
                    # Remove old library version
                    Remove-Item -LiteralPath $upgrade.DestPath -Recurse -Force -ErrorAction Stop
                    # Move new version
                    Move-Item -LiteralPath $upgrade.Folder.FullName -Destination $upgrade.DestPath -Force -ErrorAction Stop
                    Write-Host " [UPGRADED]" -ForegroundColor Green
                    Write-Log "Upgraded $($upgrade.Folder.Name): Score $($upgrade.LibraryQuality.Score) -> $($upgrade.InboxQuality.Score)" "INFO"
                    $stats.Upgraded++
                    $stats.BytesMoved += $upgrade.InboxSize
                }
                catch {
                    Write-Host " [FAILED: $_]" -ForegroundColor Red
                    Write-Log "Failed to upgrade $($upgrade.Folder.Name): $_" "ERROR"
                    $stats.Failed++
                }
            }
        }
        elseif ($upgradeChoice -eq 'R' -or $upgradeChoice -eq 'r' -or $upgradeChoice -eq 'Review' -or $upgradeChoice -eq 'review') {
            foreach ($upgrade in $potentialUpgrades) {
                Write-Host "`n  $($upgrade.Folder.Name)" -ForegroundColor Cyan
                Write-Host "    Library: $($upgrade.LibraryQuality.Resolution) $($upgrade.LibraryQuality.Codec) (Score: $($upgrade.LibraryQuality.Score), $(Format-FileSize $upgrade.LibrarySize))" -ForegroundColor Yellow
                Write-Host "    Inbox:   $($upgrade.InboxQuality.Resolution) $($upgrade.InboxQuality.Codec) (Score: $($upgrade.InboxQuality.Score), $(Format-FileSize $upgrade.InboxSize))" -ForegroundColor Green

                $singleChoice = Read-Host "    Upgrade this movie? (Y/N) [N]"
                if ($singleChoice -eq 'Y' -or $singleChoice -eq 'y') {
                    try {
                        Remove-Item -LiteralPath $upgrade.DestPath -Recurse -Force -ErrorAction Stop
                        Move-Item -LiteralPath $upgrade.Folder.FullName -Destination $upgrade.DestPath -Force -ErrorAction Stop
                        Write-Host "    [UPGRADED]" -ForegroundColor Green
                        Write-Log "Upgraded $($upgrade.Folder.Name): Score $($upgrade.LibraryQuality.Score) -> $($upgrade.InboxQuality.Score)" "INFO"
                        $stats.Upgraded++
                        $stats.BytesMoved += $upgrade.InboxSize
                    }
                    catch {
                        Write-Host "    [FAILED: $_]" -ForegroundColor Red
                        Write-Log "Failed to upgrade $($upgrade.Folder.Name): $_" "ERROR"
                        $stats.Failed++
                    }
                } else {
                    Write-Host "    [SKIPPED]" -ForegroundColor Gray
                    $stats.Skipped++
                }
            }
        } else {
            Write-Host "Upgrades skipped" -ForegroundColor Gray
            $stats.Skipped += $potentialUpgrades.Count
        }
    }

    # Summary
    Write-Host "`nTransfer complete:" -ForegroundColor Cyan
    Write-Host "  Moved:    $($stats.Moved) movies ($(Format-FileSize $stats.BytesMoved))" -ForegroundColor Green
    if ($stats.Upgraded -gt 0) {
        Write-Host "  Upgraded: $($stats.Upgraded) movies" -ForegroundColor Magenta
    }
    if ($stats.Skipped -gt 0) {
        Write-Host "  Skipped:  $($stats.Skipped) movies (already in library)" -ForegroundColor Yellow
    }
    if ($stats.Failed -gt 0) {
        Write-Host "  Failed:   $($stats.Failed) movies" -ForegroundColor Red
    }

    return $stats
}

#============================================
# HEALTH CHECK & CODEC ANALYSIS
#============================================

<#
.SYNOPSIS
    Performs a health check on the media library
.PARAMETER Path
    The root path of the media library
.DESCRIPTION
    Validates the media library by checking for:
    - Empty folders
    - Missing video files in movie folders
    - Corrupted/zero-byte files
    - Orphaned subtitle files (no matching video)
    - Missing NFO files
    - Naming issues
    - Very small video files (likely samples)
#>
function Invoke-LibraryHealthCheck {
    param(
        [string]$Path,
        [string]$MediaType = "Movies"  # Movies or TVShows
    )

    Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                  LIBRARY HEALTH CHECK                        ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    Write-Log "Starting health check for: $Path" "INFO"

    $issues = @{
        EmptyFolders = @()
        NoVideoFiles = @()
        ZeroByteFiles = @()
        OrphanedSubtitles = @()
        MissingNFO = @()
        SmallVideos = @()
        NamingIssues = @()
    }

    try {
        # Check for empty folders
        Write-Host "`nChecking for empty folders..." -ForegroundColor Yellow
        $emptyFolders = Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue |
            Where-Object { (Get-ChildItem -Path $_.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0 }
        $issues.EmptyFolders = $emptyFolders

        # Check movie folders for missing videos
        Write-Host "Checking for folders without video files..." -ForegroundColor Yellow
        $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_Trailers' }

        foreach ($folder in $folders) {
            $hasVideo = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                Select-Object -First 1

            if (-not $hasVideo) {
                $issues.NoVideoFiles += $folder
            }
        }

        # Check for zero-byte files
        Write-Host "Checking for corrupted/zero-byte files..." -ForegroundColor Yellow
        $zeroByteFiles = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -eq 0 }
        $issues.ZeroByteFiles = $zeroByteFiles

        # Check for very small video files (likely samples, under 50MB)
        Write-Host "Checking for suspiciously small video files..." -ForegroundColor Yellow
        $smallVideos = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object {
                $script:Config.VideoExtensions -contains $_.Extension.ToLower() -and
                $_.Length -lt 50MB -and
                $_.Name -notmatch 'sample|preview|trailer|teaser'
            }
        $issues.SmallVideos = $smallVideos

        # Check for orphaned subtitle files
        Write-Host "Checking for orphaned subtitle files..." -ForegroundColor Yellow
        $subtitleFiles = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.SubtitleExtensions -contains $_.Extension.ToLower() }

        foreach ($sub in $subtitleFiles) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sub.Name)
            # Remove language suffix if present
            $baseName = $baseName -replace '\.(eng|en|english|spa|es|fre|fr|ger|de)$', ''

            $hasMatchingVideo = Get-ChildItem -Path $sub.DirectoryName -File -ErrorAction SilentlyContinue |
                Where-Object {
                    $script:Config.VideoExtensions -contains $_.Extension.ToLower() -and
                    [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -eq $baseName
                } | Select-Object -First 1

            if (-not $hasMatchingVideo) {
                $issues.OrphanedSubtitles += $sub
            }
        }

        # Check for missing NFO files (movies only)
        if ($MediaType -eq "Movies") {
            Write-Host "Checking for missing NFO files..." -ForegroundColor Yellow
            foreach ($folder in $folders) {
                $hasNFO = Get-ChildItem -LiteralPath $folder.FullName -Filter "*.nfo" -File -ErrorAction SilentlyContinue |
                    Select-Object -First 1
                $hasVideo = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                    Select-Object -First 1

                if ($hasVideo -and -not $hasNFO) {
                    $issues.MissingNFO += $folder
                }
            }
        }

        # Check for naming issues
        Write-Host "Checking for naming issues..." -ForegroundColor Yellow
        foreach ($folder in $folders) {
            # Check for dots in folder names (should be spaces)
            if ($folder.Name -match '\..*\.' -and $folder.Name -notmatch '\(\d{4}\)') {
                $issues.NamingIssues += @{
                    Path = $folder.FullName
                    Issue = "Contains dots (should be spaces)"
                }
            }
            # Check for missing year
            if ($MediaType -eq "Movies" -and $folder.Name -notmatch '(19|20)\d{2}') {
                $issues.NamingIssues += @{
                    Path = $folder.FullName
                    Issue = "Missing year"
                }
            }
        }

        # Display results
        Write-Host "`n=== Health Check Results ===" -ForegroundColor Cyan

        $totalIssues = 0

        if ($issues.EmptyFolders.Count -gt 0) {
            Write-Host "`nEmpty Folders ($($issues.EmptyFolders.Count)):" -ForegroundColor Yellow
            $issues.EmptyFolders | ForEach-Object {
                Write-Host "  - $($_.FullName)" -ForegroundColor Gray
            }
            $totalIssues += $issues.EmptyFolders.Count
        }

        if ($issues.NoVideoFiles.Count -gt 0) {
            Write-Host "`nFolders Without Video Files ($($issues.NoVideoFiles.Count)):" -ForegroundColor Yellow
            $issues.NoVideoFiles | ForEach-Object {
                Write-Host "  - $($_.Name)" -ForegroundColor Gray
            }
            $totalIssues += $issues.NoVideoFiles.Count
        }

        if ($issues.ZeroByteFiles.Count -gt 0) {
            Write-Host "`nZero-Byte Files ($($issues.ZeroByteFiles.Count)):" -ForegroundColor Red
            $issues.ZeroByteFiles | ForEach-Object {
                Write-Host "  - $($_.FullName)" -ForegroundColor Gray
            }
            $totalIssues += $issues.ZeroByteFiles.Count
        }

        if ($issues.SmallVideos.Count -gt 0) {
            Write-Host "`nSuspiciously Small Videos ($($issues.SmallVideos.Count)):" -ForegroundColor Yellow
            $issues.SmallVideos | ForEach-Object {
                Write-Host "  - $($_.Name) ($(Format-FileSize $_.Length))" -ForegroundColor Gray
            }
            $totalIssues += $issues.SmallVideos.Count
        }

        if ($issues.OrphanedSubtitles.Count -gt 0) {
            Write-Host "`nOrphaned Subtitle Files ($($issues.OrphanedSubtitles.Count)):" -ForegroundColor Yellow
            $issues.OrphanedSubtitles | ForEach-Object {
                Write-Host "  - $($_.Name)" -ForegroundColor Gray
            }
            $totalIssues += $issues.OrphanedSubtitles.Count
        }

        if ($issues.MissingNFO.Count -gt 0) {
            Write-Host "`nMovies Missing NFO Files ($($issues.MissingNFO.Count)):" -ForegroundColor Yellow
            $issues.MissingNFO | Select-Object -First 10 | ForEach-Object {
                Write-Host "  - $($_.Name)" -ForegroundColor Gray
            }
            if ($issues.MissingNFO.Count -gt 10) {
                Write-Host "  ... and $($issues.MissingNFO.Count - 10) more" -ForegroundColor Gray
            }
            $totalIssues += $issues.MissingNFO.Count
        }

        if ($issues.NamingIssues.Count -gt 0) {
            Write-Host "`nNaming Issues ($($issues.NamingIssues.Count)):" -ForegroundColor Yellow
            $issues.NamingIssues | Select-Object -First 10 | ForEach-Object {
                Write-Host "  - $($_.Path): $($_.Issue)" -ForegroundColor Gray
            }
            if ($issues.NamingIssues.Count -gt 10) {
                Write-Host "  ... and $($issues.NamingIssues.Count - 10) more" -ForegroundColor Gray
            }
            $totalIssues += $issues.NamingIssues.Count
        }

        # Summary
        Write-Host "`n" -NoNewline
        if ($totalIssues -eq 0) {
            Write-Host "Library is healthy! No issues found." -ForegroundColor Green
        } else {
            Write-Host "Found $totalIssues issue(s) in the library." -ForegroundColor Yellow
        }

        Write-Log "Health check completed: $totalIssues issues found" "INFO"
        return $issues
    }
    catch {
        Write-Host "Error during health check: $_" -ForegroundColor Red
        Write-Log "Error during health check: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Analyzes video files using MediaInfo or file properties
.PARAMETER Path
    The path to the video file or folder
.DESCRIPTION
    Extracts codec information from video files and generates reports.
    Uses file naming patterns if MediaInfo is not available.
#>
function Get-VideoCodecInfo {
    param(
        [string]$FilePath
    )

    $info = @{
        FileName = [System.IO.Path]::GetFileName($FilePath)
        FileSize = 0
        VideoCodec = "Unknown"
        AudioCodec = "Unknown"
        Resolution = "Unknown"
        HDR = $false
        Container = "Unknown"
        NeedsTranscode = $false
        TranscodeReason = @()
    }

    try {
        $file = Get-Item $FilePath -ErrorAction Stop
        $info.FileSize = $file.Length
        $info.Container = $file.Extension.TrimStart('.').ToUpper()

        # Get quality info using MediaInfo when available
        $quality = Get-QualityScore -FileName $file.Name -FilePath $file.FullName
        $info.Resolution = $quality.Resolution
        $info.VideoCodec = $quality.Codec
        $info.AudioCodec = $quality.Audio
        $info.HDR = $quality.HDR

        # Determine processing needed based on codec and container
        # TranscodeMode: "none" = keep as-is, "remux" = copy streams to MKV, "transcode" = re-encode
        $info.TranscodeMode = "none"

        # XviD/DivX are legacy codecs - need full transcode to H.264
        if ($info.VideoCodec -eq "XviD" -or $info.VideoCodec -eq "DivX" -or $info.VideoCodec -eq "MPEG-4") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "transcode"
            $info.TranscodeReason += "XviD/DivX/MPEG-4 is a legacy codec - will transcode to H.264"
        }
        # H.264 in AVI container - remux only (no quality loss)
        elseif (($info.VideoCodec -eq "H264" -or $info.VideoCodec -eq "AVC" -or $info.VideoCodec -eq "H.264") -and $info.Container -eq "AVI") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "remux"
            $info.TranscodeReason += "H.264 in AVI container - will remux to MKV (no re-encoding)"
        }
        # AVI with unknown codec - likely legacy, transcode it
        elseif ($info.Container -eq "AVI") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "transcode"
            $info.TranscodeReason += "AVI with legacy codec - will transcode to H.264"
        }
        # HEVC/x265 - keep as-is (it's a modern efficient codec)
        # H.264 in modern containers (MKV, MP4) - keep as-is
        # Note: HDR and 4K are NOT flagged for transcode - they're features, not problems

        return $info
    }
    catch {
        Write-Log "Error analyzing file $FilePath : $_" "ERROR"
        return $info
    }
}

<#
.SYNOPSIS
    Generates a codec analysis report and transcoding queue
.PARAMETER Path
    The root path of the media library
.PARAMETER ExportPath
    Optional path to export the transcoding queue CSV
#>
function Invoke-CodecAnalysis {
    param(
        [string]$Path,
        [string]$ExportPath = $null
    )

    Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                    CODEC ANALYSIS                            ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    Write-Log "Starting codec analysis for: $Path" "INFO"

    try {
        # Find all video files
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($videoFiles.Count -eq 0) {
            Write-Host "No video files found" -ForegroundColor Cyan
            return
        }

        Write-Host "`nAnalyzing $($videoFiles.Count) video file(s)..." -ForegroundColor Yellow

        $analysis = @{
            TotalFiles = $videoFiles.Count
            TotalSize = 0
            ByResolution = @{}
            ByCodec = @{}
            ByContainer = @{}
            NeedTranscode = @()
            AllFiles = @()
            FilesWithConcerns = 0
        }

        $current = 1
        foreach ($file in $videoFiles) {
            $percentComplete = ($current / $videoFiles.Count) * 100
            Write-Progress -Activity "Analyzing codecs" -Status "[$current/$($videoFiles.Count)] $($file.Name)" -PercentComplete $percentComplete

            $info = Get-VideoCodecInfo -FilePath $file.FullName
            $analysis.TotalSize += $info.FileSize

            # Get quality score for sorting
            $quality = Get-QualityScore -FileName $file.Name -FilePath $file.FullName

            # Count by resolution
            if (-not $analysis.ByResolution.ContainsKey($info.Resolution)) {
                $analysis.ByResolution[$info.Resolution] = 0
            }
            $analysis.ByResolution[$info.Resolution]++

            # Count by codec
            if (-not $analysis.ByCodec.ContainsKey($info.VideoCodec)) {
                $analysis.ByCodec[$info.VideoCodec] = 0
            }
            $analysis.ByCodec[$info.VideoCodec]++

            # Count by container
            if (-not $analysis.ByContainer.ContainsKey($info.Container)) {
                $analysis.ByContainer[$info.Container] = 0
            }
            $analysis.ByContainer[$info.Container]++

            # Store all file info for quality report
            $fileInfo = @{
                Path = $file.FullName
                FileName = $info.FileName
                FolderName = $file.Directory.Name
                Size = $info.FileSize
                Resolution = $info.Resolution
                Codec = $info.VideoCodec
                AudioCodec = $info.AudioCodec
                Container = $info.Container
                HDR = $info.HDR
                QualityScore = $quality.Score
                Bitrate = $quality.Bitrate
                QualityConcerns = $quality.QualityConcerns
                HasConcerns = ($quality.QualityConcerns.Count -gt 0)
                NeedsTranscode = $info.NeedsTranscode
                TranscodeMode = $info.TranscodeMode
                TranscodeReason = $info.TranscodeReason -join "; "
            }
            $analysis.AllFiles += $fileInfo

            # Track files with quality concerns
            if ($quality.QualityConcerns.Count -gt 0) {
                $analysis.FilesWithConcerns++
            }

            # Add to transcode queue if needed
            if ($info.NeedsTranscode) {
                $analysis.NeedTranscode += $fileInfo
            }

            $current++
        }
        Write-Progress -Activity "Analyzing codecs" -Completed

        # Display results
        Write-Host "`n=== Library Statistics ===" -ForegroundColor Cyan
        Write-Host "Total Files: $($analysis.TotalFiles)" -ForegroundColor White
        Write-Host "Total Size: $(Format-FileSize $analysis.TotalSize)" -ForegroundColor White

        Write-Host "`n=== By Resolution ===" -ForegroundColor Cyan
        $analysis.ByResolution.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            $pct = [math]::Round(($_.Value / $analysis.TotalFiles) * 100, 1)
            Write-Host "  $($_.Key): $($_.Value) ($pct%)" -ForegroundColor White
        }

        Write-Host "`n=== By Video Codec ===" -ForegroundColor Cyan
        $analysis.ByCodec.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            $pct = [math]::Round(($_.Value / $analysis.TotalFiles) * 100, 1)
            Write-Host "  $($_.Key): $($_.Value) ($pct%)" -ForegroundColor White
        }

        Write-Host "`n=== By Container ===" -ForegroundColor Cyan
        $analysis.ByContainer.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            $pct = [math]::Round(($_.Value / $analysis.TotalFiles) * 100, 1)
            Write-Host "  $($_.Key): $($_.Value) ($pct%)" -ForegroundColor White
        }

        # Quality score range
        $scores = $analysis.AllFiles | ForEach-Object { $_.QualityScore }
        $minScore = ($scores | Measure-Object -Minimum).Minimum
        $maxScore = ($scores | Measure-Object -Maximum).Maximum
        $avgScore = [math]::Round(($scores | Measure-Object -Average).Average, 1)
        Write-Host "`n=== Quality Scores ===" -ForegroundColor Cyan
        Write-Host "  Range: $minScore - $maxScore (Average: $avgScore)" -ForegroundColor White

        # Show quality concerns summary
        if ($analysis.FilesWithConcerns -gt 0) {
            Write-Host "`n=== Quality Concerns ===" -ForegroundColor Yellow
            Write-Host "  Files with concerns: $($analysis.FilesWithConcerns)" -ForegroundColor $(if ($analysis.FilesWithConcerns -gt 10) { 'Red' } else { 'Yellow' })
        }

        # Interactive quality report menu - loop until user exits
        $continueReports = $true
        while ($continueReports) {
            Write-Host "`n=== Quality Reports ===" -ForegroundColor Yellow
            Write-Host "1. Show lowest quality files (re-acquisition candidates)"
            Write-Host "2. Show highest quality files"
            Write-Host "3. Show files by resolution (ascending)"
            Write-Host "4. Show legacy codec files (XviD, DivX, etc.)"
            Write-Host "5. Show files with quality concerns"
            Write-Host "6. Export full quality report to CSV"
            Write-Host "7. Done with reports"

            $qualityChoice = Read-Host "`nSelect option [7]"

            switch ($qualityChoice) {
            "1" {
                # Lowest quality files
                Write-Host "`n=== Lowest Quality Files (Re-acquisition Candidates) ===" -ForegroundColor Red
                $lowestQuality = $analysis.AllFiles | Sort-Object QualityScore | Select-Object -First 20
                $rank = 1
                foreach ($file in $lowestQuality) {
                    $hdrTag = if ($file.HDR) { " [HDR]" } else { "" }
                    $color = if ($file.QualityScore -lt 50) { "Red" } elseif ($file.QualityScore -lt 80) { "Yellow" } else { "White" }
                    Write-Host "`n  $rank. $($file.FolderName)" -ForegroundColor $color
                    Write-Host "     Score: $($file.QualityScore) | $($file.Resolution) | $($file.Codec)$hdrTag | $(Format-FileSize $file.Size)" -ForegroundColor Gray
                    $rank++
                }
            }
            "2" {
                # Highest quality files
                Write-Host "`n=== Highest Quality Files ===" -ForegroundColor Green
                $highestQuality = $analysis.AllFiles | Sort-Object QualityScore -Descending | Select-Object -First 20
                $rank = 1
                foreach ($file in $highestQuality) {
                    $hdrTag = if ($file.HDR) { " [HDR]" } else { "" }
                    Write-Host "`n  $rank. $($file.FolderName)" -ForegroundColor Green
                    Write-Host "     Score: $($file.QualityScore) | $($file.Resolution) | $($file.Codec)$hdrTag | $(Format-FileSize $file.Size)" -ForegroundColor Gray
                    $rank++
                }
            }
            "3" {
                # By resolution ascending (worst first)
                Write-Host "`n=== Files by Resolution (Lowest First) ===" -ForegroundColor Yellow
                $resolutionOrder = @{
                    "Unknown" = 0; "360p" = 1; "480p" = 2; "576p" = 3; "720p" = 4; "1080p" = 5; "2160p" = 6
                }
                $byResolution = $analysis.AllFiles | Sort-Object {
                    $res = $_.Resolution
                    if ($resolutionOrder.ContainsKey($res)) { $resolutionOrder[$res] }
                    else {
                        # Extract numeric value for non-standard resolutions
                        if ($res -match '(\d+)p') { [int]$matches[1] / 1000 }
                        else { 0 }
                    }
                }, QualityScore | Select-Object -First 30

                $currentRes = ""
                foreach ($file in $byResolution) {
                    if ($file.Resolution -ne $currentRes) {
                        $currentRes = $file.Resolution
                        Write-Host "`n  --- $currentRes ---" -ForegroundColor Cyan
                    }
                    $hdrTag = if ($file.HDR) { " [HDR]" } else { "" }
                    Write-Host "    $($file.FolderName) | $($file.Codec)$hdrTag | Score: $($file.QualityScore)" -ForegroundColor Gray
                }
            }
            "4" {
                # Legacy codec files - use regex matching for flexibility
                Write-Host "`n=== Legacy Codec Files ===" -ForegroundColor Yellow
                $legacyFiles = $analysis.AllFiles | Where-Object {
                    $_.Codec -match 'XviD|DivX|MPEG-4|VC-1|WMV|MPEG4' -or
                    $_.Container -eq "AVI"
                } | Sort-Object QualityScore

                if ($legacyFiles.Count -eq 0) {
                    Write-Host "  No legacy codec files found!" -ForegroundColor Green
                } else {
                    Write-Host "  Found $($legacyFiles.Count) file(s) with legacy codecs:" -ForegroundColor White
                    foreach ($file in $legacyFiles) {
                        Write-Host "`n    $($file.FolderName)" -ForegroundColor Yellow
                        Write-Host "      $($file.Resolution) | $($file.Codec) | $($file.Container) | Score: $($file.QualityScore)" -ForegroundColor Gray
                    }
                }
            }
            "5" {
                # Files with quality concerns
                Write-Host "`n=== Files with Quality Concerns ===" -ForegroundColor Yellow
                $concernFiles = $analysis.AllFiles | Where-Object { $_.HasConcerns } | Sort-Object QualityScore

                if ($concernFiles.Count -eq 0) {
                    Write-Host "  No quality concerns found!" -ForegroundColor Green
                } else {
                    Write-Host "  Found $($concernFiles.Count) file(s) with potential quality issues:" -ForegroundColor White
                    foreach ($file in $concernFiles) {
                        $bitrateMbps = if ($file.Bitrate -gt 0) { "$([math]::Round($file.Bitrate / 1000000, 1)) Mbps" } else { "N/A" }
                        Write-Host "`n    $($file.FolderName)" -ForegroundColor Yellow
                        Write-Host "      $($file.Resolution) | $($file.Codec) | $bitrateMbps | $(Format-FileSize $file.Size)" -ForegroundColor Gray
                        foreach ($concern in $file.QualityConcerns) {
                            Write-Host "      ! $concern" -ForegroundColor Red
                        }
                    }
                }
            }
            "6" {
                # Export full quality report
                $qualityExportPath = Join-Path $script:ReportsFolder "QualityReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                $analysis.AllFiles | Sort-Object QualityScore |
                    Select-Object FolderName, FileName, QualityScore, Resolution, Codec, AudioCodec, Container, HDR, @{N='SizeMB';E={[math]::Round($_.Size/1MB,2)}}, @{N='BitrateMbps';E={if($_.Bitrate -gt 0){[math]::Round($_.Bitrate/1000000,1)}else{'N/A'}}}, @{N='QualityConcerns';E={$_.QualityConcerns -join '; '}}, NeedsTranscode, TranscodeReason, Path |
                    Export-Csv -Path $qualityExportPath -NoTypeInformation -Encoding UTF8
                Write-Host "`nQuality report exported to: $qualityExportPath" -ForegroundColor Green
                Write-Log "Quality report exported to: $qualityExportPath" "INFO"
            }
            default {
                $continueReports = $false
            }
        }
        } # End while loop

        # Transcode queue
        if ($analysis.NeedTranscode.Count -gt 0) {
            Write-Host "`n=== Potential Transcoding Queue ===" -ForegroundColor Yellow
            Write-Host "Found $($analysis.NeedTranscode.Count) file(s) that may benefit from transcoding:" -ForegroundColor White

            $analysis.NeedTranscode | Sort-Object QualityScore | Select-Object -First 10 | ForEach-Object {
                Write-Host "`n  $($_.FolderName)" -ForegroundColor White
                Write-Host "    Score: $($_.QualityScore) | $(Format-FileSize $_.Size) | $($_.Resolution) | $($_.Codec)" -ForegroundColor Gray
                Write-Host "    Reason: $($_.TranscodeReason)" -ForegroundColor Yellow
            }

            if ($analysis.NeedTranscode.Count -gt 10) {
                Write-Host "`n  ... and $($analysis.NeedTranscode.Count - 10) more files" -ForegroundColor Gray
            }

            # Export option
            if (-not $ExportPath) {
                $exportInput = Read-Host "`nExport transcoding queue to CSV? (Y/N) [N]"
                if ($exportInput -eq 'Y' -or $exportInput -eq 'y') {
                    $ExportPath = Join-Path $script:ReportsFolder "TranscodeQueue_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                }
            }

            if ($ExportPath) {
                $analysis.NeedTranscode | Sort-Object QualityScore |
                    Select-Object FolderName, FileName, QualityScore, @{N='SizeMB';E={[math]::Round($_.Size/1MB,2)}}, Resolution, Codec, Container, TranscodeReason, Path |
                    Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
                Write-Host "`nTranscoding queue exported to: $ExportPath" -ForegroundColor Green
                Write-Log "Transcoding queue exported to: $ExportPath" "INFO"
            }
        } else {
            Write-Host "`nNo files require transcoding for compatibility." -ForegroundColor Green
        }

        Write-Log "Codec analysis completed: $($analysis.TotalFiles) files analyzed" "INFO"
        return $analysis
    }
    catch {
        Write-Host "Error during codec analysis: $_" -ForegroundColor Red
        Write-Log "Error during codec analysis: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Generates FFmpeg commands for transcoding files in the queue
.PARAMETER TranscodeQueue
    Array of files needing transcoding (from Invoke-CodecAnalysis)
.PARAMETER TargetCodec
    Target video codec (default: libx264)
.DESCRIPTION
    Creates a PowerShell script that processes files in-place:
    - Remux: Copies streams to MKV container (fast, no quality loss)
    - Transcode: Re-encodes legacy codecs to H.264
    Output files replace originals in their original folders.
    Original files are deleted only after successful processing.
#>
<#
.SYNOPSIS
    Runs transcoding inline for a queue of files
.PARAMETER TranscodeQueue
    Array of file info objects to transcode
.PARAMETER TargetCodec
    FFmpeg codec to use (default: libx264)
#>
function Invoke-Transcode {
    param(
        [array]$TranscodeQueue,
        [string]$TargetCodec = "libx264"
    )

    if ($TranscodeQueue.Count -eq 0) {
        Write-Host "No files to transcode" -ForegroundColor Cyan
        return
    }

    # Check FFmpeg installation
    if (-not (Test-FFmpegInstallation)) {
        Write-Host "`nFFmpeg not found!" -ForegroundColor Red
        Write-Host "Please install FFmpeg to use transcoding features." -ForegroundColor Yellow
        Write-Host "Download from: https://ffmpeg.org/download.html" -ForegroundColor Cyan
        Write-Host "Or install via: winget install ffmpeg" -ForegroundColor Cyan
        Write-Log "FFmpeg not found - transcoding aborted" "ERROR"
        return
    }

    $total = $TranscodeQueue.Count
    $current = 0
    $remuxed = 0
    $transcoded = 0
    $failed = 0

    Write-Host "`n=== Starting Transcode ===" -ForegroundColor Cyan
    Write-Host "Files to process: $total" -ForegroundColor White
    Write-Host "Press Ctrl+C to abort at any time`n" -ForegroundColor Gray

    foreach ($item in $TranscodeQueue) {
        $current++
        $inputPath = $item.Path
        $inputDir = [System.IO.Path]::GetDirectoryName($inputPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.FileName)
        $outputPath = Join-Path $inputDir ($baseName + ".mkv")
        $mode = if ($item.TranscodeMode) { $item.TranscodeMode } else { "transcode" }

        # Skip remux if input and output are the same (already .mkv)
        if ($mode -eq "remux" -and $inputPath -eq $outputPath) {
            Write-Host "[$current/$total] Skipping remux (already .mkv): $($item.FileName)" -ForegroundColor Gray
            continue
        }

        # Use temp file during processing
        $tempOutput = $outputPath + ".tmp.mkv"

        if ($mode -eq "remux") {
            Write-Host "[$current/$total] Remuxing: $($item.FileName)" -ForegroundColor Cyan
            $ffmpegArgs = "-i `"$inputPath`" -c:v copy -c:a copy -c:s copy `"$tempOutput`" -y"
        } else {
            Write-Host "[$current/$total] Transcoding: $($item.FileName)" -ForegroundColor Yellow
            Write-Host "  Reason: $($item.TranscodeReason -join '; ')" -ForegroundColor Gray
            $ffmpegArgs = "-i `"$inputPath`" -map 0:v -map 0:a -map 0:s? -c:v $TargetCodec -crf 23 -preset medium -c:a aac -b:a 192k -c:s copy `"$tempOutput`" -y"
        }

        # Run FFmpeg
        $process = Start-Process -FilePath "ffmpeg" -ArgumentList $ffmpegArgs -NoNewWindow -Wait -PassThru

        if ($process.ExitCode -eq 0 -and (Test-Path $tempOutput)) {
            $outSize = (Get-Item $tempOutput).Length
            if ($outSize -gt 0) {
                # Delete original and rename temp to final
                Remove-Item -Path $inputPath -Force
                Rename-Item -Path $tempOutput -NewName ([System.IO.Path]::GetFileName($outputPath))

                if ($mode -eq "remux") {
                    Write-Host "  Remuxed successfully" -ForegroundColor Green
                    $remuxed++
                } else {
                    Write-Host "  Transcoded successfully" -ForegroundColor Green
                    $transcoded++
                }
            } else {
                Write-Host "  Failed (output empty)" -ForegroundColor Red
                Remove-Item -Path $tempOutput -Force -ErrorAction SilentlyContinue
                $failed++
            }
        } else {
            Write-Host "  Failed" -ForegroundColor Red
            Remove-Item -Path $tempOutput -Force -ErrorAction SilentlyContinue
            $failed++
        }
    }

    Write-Host "`n=== Transcode Complete ===" -ForegroundColor Cyan
    Write-Host "  Remuxed: $remuxed" -ForegroundColor Cyan
    Write-Host "  Transcoded: $transcoded" -ForegroundColor Yellow
    Write-Host "  Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
    Write-Log "Transcode complete - Remuxed: $remuxed, Transcoded: $transcoded, Failed: $failed" "INFO"
}

function New-TranscodeScript {
    param(
        [array]$TranscodeQueue,
        [string]$TargetCodec = "libx264",
        [string]$TargetResolution = $null  # e.g., "1920:1080" for 1080p
    )

    if ($TranscodeQueue.Count -eq 0) {
        Write-Host "No files in transcode queue" -ForegroundColor Cyan
        return
    }

    # Check FFmpeg installation
    if (-not (Test-FFmpegInstallation)) {
        Write-Host "`nFFmpeg not found!" -ForegroundColor Red
        Write-Host "Please install FFmpeg to use transcoding features." -ForegroundColor Yellow
        Write-Host "Download from: https://ffmpeg.org/download.html" -ForegroundColor Cyan
        Write-Host "Or install via: winget install ffmpeg" -ForegroundColor Cyan
        Write-Log "FFmpeg not found - transcode script generation aborted" "ERROR"
        return $null
    }

    # Count files by mode
    $remuxCount = ($TranscodeQueue | Where-Object { $_.TranscodeMode -eq "remux" }).Count
    $transcodeCount = ($TranscodeQueue | Where-Object { $_.TranscodeMode -eq "transcode" -or $_.TranscodeMode -eq $null }).Count

    $ffmpegPath = $script:Config.FFmpegPath
    $scriptPath = Join-Path $script:AppDataFolder "transcode_$(Get-Date -Format 'yyyyMMdd_HHmmss').ps1"

    $scriptContent = @"
# LibraryLint Transcode Script
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# Target Codec: $TargetCodec
# Files to process: $($TranscodeQueue.Count)
#   - Remux only (no re-encoding): $remuxCount
#   - Full transcode: $transcodeCount
#
# Output files are saved in the same folder as the original.
# Original files are deleted after successful processing.

`$ffmpegPath = "ffmpeg"  # Update this path if ffmpeg is not in PATH

`$files = @(
"@

    $itemCount = $TranscodeQueue.Count
    $currentIndex = 0
    foreach ($item in $TranscodeQueue) {
        $currentIndex++
        # Output goes to same folder as input, with .mkv extension
        $inputDir = [System.IO.Path]::GetDirectoryName($item.Path)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.FileName)
        $mode = if ($item.TranscodeMode) { $item.TranscodeMode } else { "transcode" }

        # Output is always .mkv (same name as input, different extension if needed)
        # If input is already .mkv, the script uses a temp file during processing
        $outputFile = Join-Path $inputDir ($baseName + ".mkv")
        # Escape single quotes by doubling them for PowerShell string literals
        $escapedInput = $item.Path -replace "'", "''"
        $escapedOutput = $outputFile -replace "'", "''"
        # Only add comma if not the last item
        $comma = if ($currentIndex -lt $itemCount) { "," } else { "" }
        $scriptContent += "`n    @{ Input = '$escapedInput'; Output = '$escapedOutput'; Mode = '$mode' }$comma"
    }

    $resolutionParam = if ($TargetResolution) { "-vf scale=$TargetResolution" } else { "" }

    $scriptContent += @"

)

`$total = `$files.Count
`$current = 1
`$remuxed = 0
`$transcoded = 0
`$failed = 0
`$deleted = 0

foreach (`$file in `$files) {
    # Skip remux if input and output are the same file (already .mkv container)
    if (`$file.Mode -eq "remux" -and `$file.Input -eq `$file.Output) {
        Write-Host "[`$current/`$total] Skipping remux (already .mkv): `$(`$file.Input)" -ForegroundColor Gray
        `$current++
        continue
    }

    # Use temp file to avoid issues when input/output have same name
    `$tempOutput = `$file.Output + ".tmp.mkv"

    if (`$file.Mode -eq "remux") {
        # REMUX: Copy streams without re-encoding (fast, no quality loss)
        Write-Host "[`$current/`$total] Remuxing (no re-encode): `$(`$file.Input)" -ForegroundColor Cyan
        & `$ffmpegPath -i "`$(`$file.Input)" -c:v copy -c:a copy -c:s copy "`$tempOutput" -y

        if (`$LASTEXITCODE -eq 0 -and (Test-Path `$tempOutput)) {
            # Verify output file is valid (has size > 0)
            `$outSize = (Get-Item `$tempOutput).Length
            if (`$outSize -gt 0) {
                # Delete original and rename temp to final
                Remove-Item -Path `$file.Input -Force
                Rename-Item -Path `$tempOutput -NewName ([System.IO.Path]::GetFileName(`$file.Output))
                Write-Host "  -> Remuxed & replaced: `$(`$file.Output)" -ForegroundColor Green
                `$remuxed++
                `$deleted++
            } else {
                Write-Host "  -> Failed (output empty): `$(`$file.Input)" -ForegroundColor Red
                Remove-Item -Path `$tempOutput -Force -ErrorAction SilentlyContinue
                `$failed++
            }
        } else {
            Write-Host "  -> Failed: `$(`$file.Input)" -ForegroundColor Red
            Remove-Item -Path `$tempOutput -Force -ErrorAction SilentlyContinue
            `$failed++
        }
    } else {
        # TRANSCODE: Re-encode video to H.264
        Write-Host "[`$current/`$total] Transcoding: `$(`$file.Input)" -ForegroundColor Yellow
        & `$ffmpegPath -i "`$(`$file.Input)" -map 0:v -map 0:a -map 0:s? -c:v $TargetCodec -crf 23 -preset medium $resolutionParam -c:a aac -b:a 192k -c:s copy "`$tempOutput" -y

        if (`$LASTEXITCODE -eq 0 -and (Test-Path `$tempOutput)) {
            # Verify output file is valid (has size > 0)
            `$outSize = (Get-Item `$tempOutput).Length
            if (`$outSize -gt 0) {
                # Delete original and rename temp to final
                Remove-Item -Path `$file.Input -Force
                Rename-Item -Path `$tempOutput -NewName ([System.IO.Path]::GetFileName(`$file.Output))
                Write-Host "  -> Transcoded & replaced: `$(`$file.Output)" -ForegroundColor Green
                `$transcoded++
                `$deleted++

                # 5-second pause to allow user to stop
                Write-Host "  Press any key within 5 seconds to stop, or wait to continue..." -ForegroundColor Magenta
                `$timeout = 5
                `$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                while (`$stopwatch.Elapsed.TotalSeconds -lt `$timeout) {
                    if ([Console]::KeyAvailable) {
                        `$null = [Console]::ReadKey(`$true)
                        Write-Host "``n  User requested stop. Exiting..." -ForegroundColor Yellow
                        Write-Host "``n========== Summary (Stopped Early) ==========" -ForegroundColor Cyan
                        Write-Host "Remuxed (fast, no quality loss): `$remuxed" -ForegroundColor Cyan
                        Write-Host "Transcoded (re-encoded): `$transcoded" -ForegroundColor Yellow
                        Write-Host "Original files deleted: `$deleted" -ForegroundColor Green
                        Write-Host "Failed: `$failed" -ForegroundColor Red
                        Write-Host "==============================================" -ForegroundColor Cyan
                        exit
                    }
                    Start-Sleep -Milliseconds 100
                }
            } else {
                Write-Host "  -> Failed (output empty): `$(`$file.Input)" -ForegroundColor Red
                Remove-Item -Path `$tempOutput -Force -ErrorAction SilentlyContinue
                `$failed++
            }
        } else {
            Write-Host "  -> Failed: `$(`$file.Input)" -ForegroundColor Red
            Remove-Item -Path `$tempOutput -Force -ErrorAction SilentlyContinue
            `$failed++
        }
    }

    `$current++
}

Write-Host "`n========== Summary ==========" -ForegroundColor Cyan
Write-Host "Remuxed (fast, no quality loss): `$remuxed" -ForegroundColor Cyan
Write-Host "Transcoded (re-encoded): `$transcoded" -ForegroundColor Yellow
Write-Host "Original files deleted: `$deleted" -ForegroundColor Green
Write-Host "Failed: `$failed" -ForegroundColor $(if (`$failed -gt 0) { 'Red' } else { 'Green' })
Write-Host "==============================" -ForegroundColor Cyan
"@

    $scriptContent | Out-File -FilePath $scriptPath -Encoding UTF8
    Write-Host "`nTranscode script created: $scriptPath" -ForegroundColor Green
    Write-Host "  - Remux only (fast, no quality loss): $remuxCount files" -ForegroundColor Cyan
    Write-Host "  - Full transcode (re-encode): $transcodeCount files" -ForegroundColor Yellow
    Write-Host "`nRun the script to start processing, or edit it to customize parameters." -ForegroundColor Cyan
    Write-Log "Transcode script created: $scriptPath - Remux: $remuxCount, Transcode: $transcodeCount" "INFO"

    return $scriptPath
}

#============================================
# TMDB API INTEGRATION
#============================================

<#
.SYNOPSIS
    Tests if a TMDB API key is valid
.PARAMETER ApiKey
    TMDB API key to validate
.OUTPUTS
    Boolean indicating if the key is valid
#>
function Test-TMDBApiKey {
    param(
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $false
    }

    try {
        # Use the configuration endpoint which requires a valid API key
        $url = "https://api.themoviedb.org/3/configuration?api_key=$ApiKey"
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        # If we get here, the key is valid
        if ($response.images) {
            return $true
        }
        return $false
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 401) {
            Write-Host "Invalid API key" -ForegroundColor Red
        } else {
            Write-Host "Error validating API key: $_" -ForegroundColor Red
        }
        return $false
    }
}

<#
.SYNOPSIS
    Searches TMDB for a movie by title and year
.PARAMETER Title
    The movie title to search for
.PARAMETER Year
    The movie year (optional but recommended)
.PARAMETER ApiKey
    TMDB API key (get one free at https://www.themoviedb.org/settings/api)
.OUTPUTS
    Hashtable with movie metadata from TMDB
#>
function Search-TMDBMovie {
    param(
        [string]$Title,
        [string]$Year = $null,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        Write-Log "TMDB API key not provided" "WARNING"
        return $null
    }

    try {
        $encodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)
        $url = "https://api.themoviedb.org/3/search/movie?api_key=$ApiKey&query=$encodedTitle"

        if ($Year) {
            $url += "&year=$Year"
        }

        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        if ($response.results -and $response.results.Count -gt 0) {
            # If year was provided, try to find an exact year match first
            $movie = $null
            if ($Year) {
                $movie = $response.results | Where-Object {
                    $_.release_date -and $_.release_date.StartsWith($Year)
                } | Select-Object -First 1
            }

            # Fall back to first result if no exact year match
            if (-not $movie) {
                $movie = $response.results[0]
            }

            return @{
                Id = $movie.id
                Title = $movie.title
                OriginalTitle = $movie.original_title
                Year = if ($movie.release_date) { $movie.release_date.Substring(0,4) } else { $null }
                Overview = $movie.overview
                Rating = $movie.vote_average
                Votes = $movie.vote_count
                PosterPath = if ($movie.poster_path) { "https://image.tmdb.org/t/p/w500$($movie.poster_path)" } else { $null }
                BackdropPath = if ($movie.backdrop_path) { "https://image.tmdb.org/t/p/original$($movie.backdrop_path)" } else { $null }
            }
        }

        return $null
    }
    catch {
        Write-Log "Error searching TMDB: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Gets detailed movie information from TMDB by ID
.PARAMETER MovieId
    The TMDB movie ID
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with detailed movie metadata
#>
function Get-TMDBMovieDetails {
    param(
        [int]$MovieId,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    try {
        $url = "https://api.themoviedb.org/3/movie/$MovieId`?api_key=$ApiKey&append_to_response=credits,external_ids,videos"
        $movie = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        $directors = @()
        $cast = @()

        if ($movie.credits) {
            $directors = $movie.credits.crew | Where-Object { $_.job -eq 'Director' } | Select-Object -ExpandProperty name
            $cast = $movie.credits.cast | Select-Object -First 10 | ForEach-Object {
                @{
                    Name = $_.name
                    Role = $_.character
                    Thumb = if ($_.profile_path) { "https://image.tmdb.org/t/p/w185$($_.profile_path)" } else { $null }
                }
            }
        }

        # Get YouTube trailer key (prefer official trailers, then teasers)
        $trailerKey = $null
        if ($movie.videos -and $movie.videos.results) {
            $trailer = $movie.videos.results | Where-Object {
                $_.site -eq 'YouTube' -and $_.type -eq 'Trailer' -and $_.official -eq $true
            } | Select-Object -First 1

            if (-not $trailer) {
                $trailer = $movie.videos.results | Where-Object {
                    $_.site -eq 'YouTube' -and $_.type -eq 'Trailer'
                } | Select-Object -First 1
            }

            if (-not $trailer) {
                $trailer = $movie.videos.results | Where-Object {
                    $_.site -eq 'YouTube' -and $_.type -eq 'Teaser'
                } | Select-Object -First 1
            }

            if ($trailer) {
                $trailerKey = $trailer.key
            }
        }

        return @{
            Id = $movie.id
            TMDBID = $movie.id
            Title = $movie.title
            OriginalTitle = $movie.original_title
            Tagline = $movie.tagline
            Year = if ($movie.release_date) { $movie.release_date.Substring(0,4) } else { $null }
            ReleaseDate = $movie.release_date
            Overview = $movie.overview
            VoteAverage = $movie.vote_average
            VoteCount = $movie.vote_count
            Runtime = $movie.runtime
            Genres = $movie.genres | Select-Object -ExpandProperty name
            Studios = $movie.production_companies | Select-Object -ExpandProperty name
            Directors = $directors
            Actors = $cast
            IMDBID = $movie.external_ids.imdb_id
            PosterPath = if ($movie.poster_path) { "https://image.tmdb.org/t/p/w500$($movie.poster_path)" } else { $null }
            BackdropPath = if ($movie.backdrop_path) { "https://image.tmdb.org/t/p/original$($movie.backdrop_path)" } else { $null }
            TrailerKey = $trailerKey
        }
    }
    catch {
        Write-Log "Error getting TMDB movie details: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Downloads an image from a URL to a local path
.PARAMETER Url
    The URL of the image to download
.PARAMETER DestinationPath
    The local path to save the image
.PARAMETER Description
    Description for logging purposes
#>
function Save-ImageFromUrl {
    param(
        [string]$Url,
        [string]$DestinationPath,
        [string]$Description = "image"
    )

    if (-not $Url) {
        return $false
    }

    # Skip if file already exists
    if (Test-Path $DestinationPath) {
        Write-Log "$Description already exists: $DestinationPath" "INFO"
        return $true
    }

    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($Url, $DestinationPath)
        Write-Log "Downloaded $Description to: $DestinationPath" "INFO"
        return $true
    }
    catch {
        Write-Log "Error downloading $Description from $Url : $_" "ERROR"
        return $false
    }
    finally {
        if ($webClient) {
            $webClient.Dispose()
        }
    }
}

<#
.SYNOPSIS
    Downloads all movie artwork from TMDB metadata
.PARAMETER Metadata
    Hashtable containing TMDB movie metadata (from Get-TMDBMovieDetails)
.PARAMETER MovieFolder
    The folder where the movie is located
#>
function Save-MovieArtwork {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Metadata,
        [Parameter(Mandatory=$true)]
        [string]$MovieFolder
    )

    $downloadedCount = 0

    # Download poster
    if ($Metadata.PosterPath) {
        $posterPath = Join-Path $MovieFolder "poster.jpg"
        if (-not (Test-Path $posterPath)) {
            Write-Host "    Downloading poster..." -ForegroundColor Gray
            if (Save-ImageFromUrl -Url $Metadata.PosterPath -DestinationPath $posterPath -Description "poster") {
                $downloadedCount++
            }
        }
    }

    # Download fanart/backdrop
    if ($Metadata.BackdropPath) {
        $fanartPath = Join-Path $MovieFolder "fanart.jpg"
        if (-not (Test-Path $fanartPath)) {
            Write-Host "    Downloading fanart..." -ForegroundColor Gray
            if (Save-ImageFromUrl -Url $Metadata.BackdropPath -DestinationPath $fanartPath -Description "fanart") {
                $downloadedCount++
            }
        }
    }

    # Download actor images
    if ($Metadata.Actors -and $Metadata.Actors.Count -gt 0) {
        $actorsFolder = Join-Path $MovieFolder ".actors"
        $actorsWithThumbs = $Metadata.Actors | Where-Object { $_.Thumb }

        if ($actorsWithThumbs.Count -gt 0) {
            # Create .actors folder if it doesn't exist
            if (-not (Test-Path $actorsFolder)) {
                New-Item -Path $actorsFolder -ItemType Directory -Force | Out-Null
            }

            $actorCount = 0
            foreach ($actor in $actorsWithThumbs) {
                # Sanitize actor name for filename
                $safeActorName = $actor.Name -replace '[\\/:*?"<>|]', '_'
                $actorImagePath = Join-Path $actorsFolder "$safeActorName.jpg"

                if (-not (Test-Path $actorImagePath)) {
                    if ($actorCount -eq 0) {
                        Write-Host "    Downloading actor images..." -ForegroundColor Gray
                    }
                    if (Save-ImageFromUrl -Url $actor.Thumb -DestinationPath $actorImagePath -Description "actor image for $($actor.Name)") {
                        $downloadedCount++
                        $actorCount++
                    }
                }
            }
        }
    }

    return $downloadedCount
}

<#
.SYNOPSIS
    Tests if a fanart.tv API key is valid
.PARAMETER ApiKey
    Fanart.tv API key to validate
.OUTPUTS
    Boolean indicating if the key is valid
#>
function Test-FanartTVApiKey {
    param(
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $false
    }

    try {
        # Use a known movie (The Matrix - TMDB ID 603) to test the API
        $url = "https://webservice.fanart.tv/v3/movies/603?api_key=$ApiKey"
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        if ($response.tmdb_id) {
            return $true
        }
        return $false
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 401) {
            Write-Host "Invalid fanart.tv API key" -ForegroundColor Red
        } else {
            Write-Host "Error validating fanart.tv API key: $_" -ForegroundColor Red
        }
        return $false
    }
}

<#
.SYNOPSIS
    Gets movie artwork from fanart.tv
.PARAMETER TMDBID
    The TMDB movie ID
.PARAMETER ApiKey
    Fanart.tv API key
.OUTPUTS
    Hashtable with artwork URLs (clearlogo, hdmovielogo, movieposter, moviebackground, etc.)
#>
function Get-FanartTVMovieArt {
    param(
        [int]$TMDBID,
        [string]$ApiKey
    )

    if (-not $ApiKey -or -not $TMDBID) {
        return $null
    }

    try {
        $url = "https://webservice.fanart.tv/v3/movies/$TMDBID`?api_key=$ApiKey"
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        # Extract best artwork URLs (first English or first available)
        $artwork = @{
            ClearLogo = $null
            HDClearLogo = $null
            Banner = $null
            Thumb = $null
            Disc = $null
        }

        # HD Movie Logo (clearlogo) - prefer English
        if ($response.hdmovielogo) {
            $englishLogo = $response.hdmovielogo | Where-Object { $_.lang -eq 'en' } | Select-Object -First 1
            if ($englishLogo) {
                $artwork.HDClearLogo = $englishLogo.url
            } else {
                $artwork.HDClearLogo = $response.hdmovielogo[0].url
            }
        }

        # Regular Movie Logo (fallback)
        if (-not $artwork.HDClearLogo -and $response.movielogo) {
            $englishLogo = $response.movielogo | Where-Object { $_.lang -eq 'en' } | Select-Object -First 1
            if ($englishLogo) {
                $artwork.ClearLogo = $englishLogo.url
            } else {
                $artwork.ClearLogo = $response.movielogo[0].url
            }
        }

        # Movie Banner
        if ($response.moviebanner) {
            $englishBanner = $response.moviebanner | Where-Object { $_.lang -eq 'en' } | Select-Object -First 1
            if ($englishBanner) {
                $artwork.Banner = $englishBanner.url
            } else {
                $artwork.Banner = $response.moviebanner[0].url
            }
        }

        # Movie Thumb (landscape)
        if ($response.moviethumb) {
            $englishThumb = $response.moviethumb | Where-Object { $_.lang -eq 'en' } | Select-Object -First 1
            if ($englishThumb) {
                $artwork.Thumb = $englishThumb.url
            } else {
                $artwork.Thumb = $response.moviethumb[0].url
            }
        }

        # Movie Disc
        if ($response.moviedisc) {
            $englishDisc = $response.moviedisc | Where-Object { $_.lang -eq 'en' } | Select-Object -First 1
            if ($englishDisc) {
                $artwork.Disc = $englishDisc.url
            } else {
                $artwork.Disc = $response.moviedisc[0].url
            }
        }

        return $artwork
    }
    catch {
        # 404 is expected for movies not in fanart.tv database
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -ne 404) {
            Write-Log "Error getting fanart.tv artwork: $_" "WARNING"
        }
        return $null
    }
}

<#
.SYNOPSIS
    Downloads movie artwork from fanart.tv
.PARAMETER TMDBID
    The TMDB movie ID
.PARAMETER MovieFolder
    The folder where the movie is located
.PARAMETER ApiKey
    Fanart.tv API key
.OUTPUTS
    Number of artwork files downloaded
#>
function Save-FanartTVArtwork {
    param(
        [int]$TMDBID,
        [string]$MovieFolder,
        [string]$ApiKey
    )

    if (-not $ApiKey -or -not $TMDBID) {
        return 0
    }

    $fanartData = Get-FanartTVMovieArt -TMDBID $TMDBID -ApiKey $ApiKey
    if (-not $fanartData) {
        return 0
    }

    $downloadedCount = 0

    # Download clearlogo (HD preferred, regular as fallback)
    $clearlogoUrl = if ($fanartData.HDClearLogo) { $fanartData.HDClearLogo } else { $fanartData.ClearLogo }
    if ($clearlogoUrl) {
        $clearlogoPath = Join-Path $MovieFolder "clearlogo.png"
        if (-not (Test-Path $clearlogoPath)) {
            Write-Host "    Downloading clearlogo..." -ForegroundColor Gray
            if (Save-ImageFromUrl -Url $clearlogoUrl -DestinationPath $clearlogoPath -Description "clearlogo") {
                $downloadedCount++
            }
        }
    }

    # Download banner
    if ($fanartData.Banner) {
        $bannerPath = Join-Path $MovieFolder "banner.jpg"
        if (-not (Test-Path $bannerPath)) {
            Write-Host "    Downloading banner..." -ForegroundColor Gray
            if (Save-ImageFromUrl -Url $fanartData.Banner -DestinationPath $bannerPath -Description "banner") {
                $downloadedCount++
            }
        }
    }

    # Download landscape/thumb
    if ($fanartData.Thumb) {
        $landscapePath = Join-Path $MovieFolder "landscape.jpg"
        if (-not (Test-Path $landscapePath)) {
            Write-Host "    Downloading landscape..." -ForegroundColor Gray
            if (Save-ImageFromUrl -Url $fanartData.Thumb -DestinationPath $landscapePath -Description "landscape") {
                $downloadedCount++
            }
        }
    }

    # Download disc art
    if ($fanartData.Disc) {
        $discPath = Join-Path $MovieFolder "disc.png"
        if (-not (Test-Path $discPath)) {
            Write-Host "    Downloading disc art..." -ForegroundColor Gray
            if (Save-ImageFromUrl -Url $fanartData.Disc -DestinationPath $discPath -Description "disc art") {
                $downloadedCount++
            }
        }
    }

    return $downloadedCount
}

<#
.SYNOPSIS
    Downloads a movie trailer from YouTube using yt-dlp
.PARAMETER TrailerKey
    The YouTube video ID
.PARAMETER MovieFolder
    The folder to save the trailer to
.PARAMETER MovieTitle
    The movie title (used for naming the file)
.PARAMETER Quality
    Video quality: 1080p, 720p, or 480p
.OUTPUTS
    Boolean indicating success
#>
function Save-MovieTrailer {
    param(
        [string]$TrailerKey,
        [string]$MovieFolder,
        [string]$MovieTitle,
        [string]$Quality = "720p"
    )

    if (-not $TrailerKey) {
        Write-Log "No trailer key provided" "DEBUG"
        return $false
    }

    if (-not (Test-YtDlpInstallation)) {
        Write-Log "yt-dlp not found - cannot download trailer" "WARNING"
        return $false
    }

    # Check if trailer already exists
    $existingTrailers = Get-ChildItem -Path $MovieFolder -Filter "*-trailer.*" -ErrorAction SilentlyContinue
    if ($existingTrailers) {
        Write-Log "Trailer already exists in $MovieFolder" "DEBUG"
        return $true
    }

    # Clean movie title for filename
    $cleanTitle = $MovieTitle -replace '[\\/:*?"<>|]', ''
    $trailerPath = Join-Path $MovieFolder "$cleanTitle-trailer.mp4"

    # Map quality to yt-dlp format
    $formatSpec = switch ($Quality) {
        "1080p" { "bestvideo[height<=1080][ext=mp4]+bestaudio[ext=m4a]/best[height<=1080][ext=mp4]/best[height<=1080]" }
        "720p"  { "bestvideo[height<=720][ext=mp4]+bestaudio[ext=m4a]/best[height<=720][ext=mp4]/best[height<=720]" }
        "480p"  { "bestvideo[height<=480][ext=mp4]+bestaudio[ext=m4a]/best[height<=480][ext=mp4]/best[height<=480]" }
        default { "bestvideo[height<=720][ext=mp4]+bestaudio[ext=m4a]/best[height<=720][ext=mp4]/best[height<=720]" }
    }

    $youtubeUrl = "https://www.youtube.com/watch?v=$TrailerKey"

    try {
        Write-Host "    Downloading trailer ($Quality)..." -ForegroundColor Gray

        # Quote the trailer path to handle spaces in folder/file names
        $ytdlpArgs = @(
            "--no-playlist",
            "-f", "`"$formatSpec`"",
            "--merge-output-format", "mp4",
            "-o", "`"$trailerPath`"",
            "--no-warnings",
            "--quiet"
        )

        # Add cookies from browser to handle age-restricted videos
        if ($script:Config.YtDlpCookieBrowser) {
            $ytdlpArgs += "--cookies-from-browser"
            $ytdlpArgs += $script:Config.YtDlpCookieBrowser
        }

        $ytdlpArgs += "`"$youtubeUrl`""

        $process = Start-Process -FilePath $script:Config.YtDlpPath -ArgumentList $ytdlpArgs -NoNewWindow -Wait -PassThru

        if ($process.ExitCode -eq 0 -and (Test-Path $trailerPath)) {
            $fileSize = (Get-Item $trailerPath).Length
            Write-Host "    Trailer downloaded ($(Format-FileSize $fileSize))" -ForegroundColor Green
            Write-Log "Downloaded trailer for $MovieTitle" "INFO"
            return $true
        } else {
            Write-Log "yt-dlp failed to download trailer for $MovieTitle (exit code: $($process.ExitCode))" "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "Error downloading trailer for $MovieTitle : $_" "ERROR"
        return $false
    }
}

<#
.SYNOPSIS
    Downloads missing trailers for all movies in a library folder
.DESCRIPTION
    Scans through movie folders, checks for existing trailers, and downloads
    missing ones using TMDB metadata and yt-dlp
.PARAMETER Path
    The root path of the movie library
.PARAMETER Quality
    Trailer quality: 1080p, 720p, or 480p (default: from config or 720p)
#>
function Invoke-TrailerDownload {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [string]$Quality
    )

    if (-not $Quality) {
        $Quality = if ($script:Config.TrailerQuality) { $script:Config.TrailerQuality } else { "720p" }
    }

    Write-Host "`n--- Download Missing Trailers ---" -ForegroundColor Yellow
    Write-Log "Starting trailer download scan in: $Path" "INFO"

    # Check prerequisites
    if (-not (Test-YtDlpInstallation)) {
        Write-Host "yt-dlp is required for trailer downloads" -ForegroundColor Red
        Write-Host "Install with: winget install yt-dlp" -ForegroundColor Yellow
        return
    }

    if (-not $script:Config.TMDBApiKey) {
        Write-Host "TMDB API key is required to find trailer URLs" -ForegroundColor Red
        return
    }

    # Get all movie folders
    $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue

    if (-not $movieFolders -or $movieFolders.Count -eq 0) {
        Write-Host "No movie folders found in $Path" -ForegroundColor Yellow
        return
    }

    Write-Host "Scanning $($movieFolders.Count) folders for missing trailers..." -ForegroundColor Cyan
    Write-Host "Quality: $Quality" -ForegroundColor Gray

    $stats = @{
        Total = $movieFolders.Count
        AlreadyHave = 0
        Downloaded = 0
        Failed = 0
        NoTrailerAvailable = 0
    }

    foreach ($folder in $movieFolders) {
        # Check if trailer already exists
        $existingTrailers = Get-ChildItem -LiteralPath $folder.FullName -Filter "*-trailer.*" -ErrorAction SilentlyContinue
        if ($existingTrailers) {
            $stats.AlreadyHave++
            continue
        }

        # Parse folder name for movie info
        $folderName = $folder.Name
        $titleMatch = [regex]::Match($folderName, '^(.+?)\s*\((\d{4})\)$')

        if (-not $titleMatch.Success) {
            # Try without year
            $movieTitle = $folderName
            $movieYear = $null
        } else {
            $movieTitle = $titleMatch.Groups[1].Value.Trim()
            $movieYear = $titleMatch.Groups[2].Value
        }

        Write-Host "`nProcessing: $folderName" -ForegroundColor White

        # Try to get trailer from existing NFO first
        $nfoFile = Get-ChildItem -LiteralPath $folder.FullName -Filter "*.nfo" -ErrorAction SilentlyContinue | Select-Object -First 1
        $trailerKey = $null

        if ($nfoFile) {
            try {
                $nfoContent = Get-Content -Path $nfoFile.FullName -Raw -ErrorAction SilentlyContinue
                if ($nfoContent -match '<trailer>.*youtube.*watch\?v=([^<&"]+)') {
                    $trailerKey = $matches[1]
                    Write-Host "  Found trailer URL in NFO" -ForegroundColor Gray
                }
            } catch {}
        }

        # If no trailer in NFO, search TMDB
        if (-not $trailerKey) {
            Write-Host "  Searching TMDB..." -ForegroundColor Gray

            $searchResult = Search-TMDBMovie -Title $movieTitle -Year $movieYear
            if ($searchResult) {
                $tmdbMetadata = Get-TMDBMovieMetadata -MovieId $searchResult.id
                if ($tmdbMetadata -and $tmdbMetadata.TrailerKey) {
                    $trailerKey = $tmdbMetadata.TrailerKey
                }
            }
        }

        if ($trailerKey) {
            $success = Save-MovieTrailer -TrailerKey $trailerKey -MovieFolder $folder.FullName -MovieTitle $movieTitle -Quality $Quality
            if ($success) {
                $stats.Downloaded++
            } else {
                $stats.Failed++
            }
        } else {
            Write-Host "  No trailer available" -ForegroundColor Yellow
            $stats.NoTrailerAvailable++
        }

        # Rate limiting to avoid hammering TMDB
        Start-Sleep -Milliseconds 250
    }

    # Summary
    Write-Host "`n--- Trailer Download Summary ---" -ForegroundColor Cyan
    Write-Host "Total folders scanned: $($stats.Total)" -ForegroundColor White
    Write-Host "Already had trailers:  $($stats.AlreadyHave)" -ForegroundColor Green
    Write-Host "Downloaded:            $($stats.Downloaded)" -ForegroundColor Green
    Write-Host "No trailer available:  $($stats.NoTrailerAvailable)" -ForegroundColor Yellow
    Write-Host "Failed:                $($stats.Failed)" -ForegroundColor $(if ($stats.Failed -gt 0) { 'Red' } else { 'White' })

    Write-Log "Trailer download completed: Downloaded=$($stats.Downloaded), Failed=$($stats.Failed), NoTrailer=$($stats.NoTrailerAvailable), AlreadyHad=$($stats.AlreadyHave)" "INFO"
}

<#
.SYNOPSIS
    Searches Subdl.com for subtitles by IMDB ID or movie title
.PARAMETER IMDBID
    The IMDB ID of the movie (e.g., tt0107290)
.PARAMETER Title
    The movie title (fallback if no IMDB ID)
.PARAMETER Year
    The movie year (used with title search)
.PARAMETER Language
    Language code (e.g., en, es, fr, de)
.OUTPUTS
    Array of subtitle results or $null
#>
function Search-SubdlSubtitle {
    param(
        [string]$IMDBID,
        [string]$Title,
        [string]$Year,
        [string]$Language = "en",
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        Write-Log "Subdl API key not provided" "WARNING"
        return $null
    }

    try {
        $baseUrl = "https://api.subdl.com/api/v1/subtitles"
        $results = $null

        # Try IMDB ID first (most accurate)
        if ($IMDBID) {
            $url = "$baseUrl`?api_key=$ApiKey&imdb_id=$IMDBID&languages=$Language&subs_per_page=10"
            Write-Log "Searching Subdl by IMDB ID: $IMDBID" "DEBUG"

            $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

            if ($response.status -and $response.subtitles -and $response.subtitles.Count -gt 0) {
                $results = $response.subtitles
            }
        }

        # Fallback to title search
        if (-not $results -and $Title) {
            $encodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)
            $url = "$baseUrl`?api_key=$ApiKey&film_name=$encodedTitle&languages=$Language&subs_per_page=10"
            if ($Year) {
                $url += "&year=$Year"
            }
            Write-Log "Searching Subdl by title: $Title ($Year)" "DEBUG"

            $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

            if ($response.status -and $response.subtitles -and $response.subtitles.Count -gt 0) {
                $results = $response.subtitles
            }
        }

        if ($results) {
            Write-Log "Found $($results.Count) subtitle(s) on Subdl" "DEBUG"
            return $results
        }

        Write-Log "No subtitles found on Subdl for IMDB: $IMDBID, Title: $Title" "DEBUG"
        return $null
    }
    catch {
        Write-Log "Error searching Subdl: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Downloads and saves a subtitle from Subdl.com
.PARAMETER SubtitleUrl
    The URL to download the subtitle from
.PARAMETER MovieFolder
    The folder to save the subtitle to
.PARAMETER MovieTitle
    The movie title (used for naming the file)
.PARAMETER Language
    Language code for the subtitle filename
.OUTPUTS
    Boolean indicating success
#>
function Save-MovieSubtitle {
    param(
        [string]$IMDBID,
        [string]$Title,
        [string]$Year,
        [string]$MovieFolder,
        [string]$MovieTitle,
        [string]$Language = "en",
        [string]$ApiKey
    )

    # Check if API key is available
    if (-not $ApiKey) {
        $ApiKey = $script:Config.SubdlApiKey
    }
    if (-not $ApiKey) {
        Write-Log "Subdl API key not configured" "WARNING"
        return $false
    }

    # Check if subtitle already exists
    $existingSubtitles = Get-ChildItem -Path $MovieFolder -Filter "*.$Language.srt" -ErrorAction SilentlyContinue
    if ($existingSubtitles) {
        Write-Log "Subtitle already exists in $MovieFolder" "DEBUG"
        return $true
    }

    # Also check for generic .srt files
    $genericSrt = Get-ChildItem -Path $MovieFolder -Filter "*.srt" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '\.(en|es|fr|de|it|pt|ru|ja|ko|zh)\.' }
    if ($genericSrt) {
        Write-Log "Generic subtitle file already exists in $MovieFolder" "DEBUG"
        return $true
    }

    # Rate limiting: 1 request per second for Subdl
    Start-Sleep -Seconds 1

    # Search for subtitles
    $subtitles = Search-SubdlSubtitle -IMDBID $IMDBID -Title $Title -Year $Year -Language $Language -ApiKey $ApiKey

    if (-not $subtitles) {
        Write-Log "No subtitles found for: $MovieTitle" "DEBUG"
        return $false
    }

    # Get the best subtitle (first result, sorted by Subdl's ranking)
    $subtitle = $subtitles | Select-Object -First 1

    if (-not $subtitle.url) {
        Write-Log "Subtitle has no download URL" "WARNING"
        return $false
    }

    try {
        Write-Host "    Downloading subtitle ($Language)..." -ForegroundColor Gray

        # Create temp folder for extraction
        $tempFolder = Join-Path $env:TEMP "LibraryLint_Sub_$(Get-Random)"
        New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null

        # Download the subtitle zip
        $zipPath = Join-Path $tempFolder "subtitle.zip"
        $downloadUrl = $subtitle.url

        # Subdl returns a path, construct full URL
        if (-not $downloadUrl.StartsWith("http")) {
            $downloadUrl = "https://dl.subdl.com$downloadUrl"
        }

        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -ErrorAction Stop

        # Extract the zip
        Expand-Archive -Path $zipPath -DestinationPath $tempFolder -Force

        # Find SRT file in extracted content
        $srtFiles = Get-ChildItem -Path $tempFolder -Filter "*.srt" -Recurse -ErrorAction SilentlyContinue

        if ($srtFiles) {
            # Use the first/largest SRT file
            $srtFile = $srtFiles | Sort-Object Length -Descending | Select-Object -First 1

            # Clean movie title for filename
            $cleanTitle = $MovieTitle -replace '[\\/:*?"<>|]', ''
            $destPath = Join-Path $MovieFolder "$cleanTitle.$Language.srt"

            Copy-Item -Path $srtFile.FullName -Destination $destPath -Force

            Write-Host "    Subtitle downloaded ($Language)" -ForegroundColor Green
            Write-Log "Downloaded subtitle for $MovieTitle" "INFO"

            # Cleanup
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue

            return $true
        } else {
            Write-Log "No SRT file found in downloaded archive" "WARNING"
        }

        # Cleanup
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        return $false
    }
    catch {
        Write-Log "Error downloading subtitle for $MovieTitle : $_" "ERROR"
        # Cleanup on error
        if (Test-Path $tempFolder) {
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        return $false
    }
}

<#
.SYNOPSIS
    Downloads missing subtitles for an existing movie library
.PARAMETER Path
    The root path of the movie library
.PARAMETER Language
    Language code for subtitles (default: en)
.DESCRIPTION
    Scans the library for movie folders with NFO files, extracts IMDB IDs,
    and downloads subtitles from Subdl.com for movies that don't have them.
#>
function Invoke-SubtitleDownload {
    param(
        [string]$Path,
        [string]$Language = "en"
    )

    Write-Host "`nScanning library for movies without subtitles..." -ForegroundColor Cyan
    Write-Log "Starting subtitle download scan for: $Path" "INFO"

    # Get all movie folders (folders containing video files)
    $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue

    $totalFolders = $movieFolders.Count
    $processed = 0
    $downloaded = 0
    $skipped = 0
    $failed = 0
    $noNfo = 0

    Write-Host "Found $totalFolders movie folders to scan" -ForegroundColor White

    foreach ($folder in $movieFolders) {
        $processed++
        $percentComplete = [math]::Round(($processed / $totalFolders) * 100)
        Write-Progress -Activity "Downloading Subtitles" -Status "$processed of $totalFolders folders" -PercentComplete $percentComplete

        # Check if subtitle already exists
        $existingSubs = Get-ChildItem -LiteralPath $folder.FullName -Filter "*.srt" -ErrorAction SilentlyContinue
        if ($existingSubs) {
            $skipped++
            Write-Log "Skipping $($folder.Name) - subtitle exists" "DEBUG"
            continue
        }

        # Find NFO file
        $nfoFile = Get-ChildItem -LiteralPath $folder.FullName -Filter "*.nfo" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $nfoFile) {
            $noNfo++
            Write-Log "Skipping $($folder.Name) - no NFO file" "DEBUG"
            continue
        }

        # Parse NFO for IMDB ID and title
        $metadata = Read-NFOFile -NfoPath $nfoFile.FullName
        if (-not $metadata) {
            $noNfo++
            continue
        }

        $movieTitle = $metadata.Title
        if (-not $movieTitle) {
            # Fallback to folder name - handle both clean (2024) and dirty (2024_1080p) formats
            $movieTitle = $folder.Name -replace '\s*\(\d{4}[^)]*\)\s*$', ''
        }

        Write-Host "  [$processed/$totalFolders] $movieTitle..." -ForegroundColor Gray -NoNewline

        # Try to download subtitle
        $result = Save-MovieSubtitle -IMDBID $metadata.IMDBID -Title $movieTitle -Year $metadata.Year -MovieFolder $folder.FullName -MovieTitle $movieTitle -Language $Language

        if ($result) {
            $downloaded++
            Write-Host " downloaded" -ForegroundColor Green
        } else {
            $failed++
            Write-Host " not found" -ForegroundColor Yellow
        }
    }

    Write-Progress -Activity "Downloading Subtitles" -Completed

    # Summary
    Write-Host "`n=== Subtitle Download Summary ===" -ForegroundColor Cyan
    Write-Host "Total folders scanned: $totalFolders" -ForegroundColor White
    Write-Host "Subtitles downloaded:  $downloaded" -ForegroundColor Green
    Write-Host "Already had subtitles: $skipped" -ForegroundColor Gray
    Write-Host "No NFO file:           $noNfo" -ForegroundColor Yellow
    Write-Host "Not found on Subdl:    $failed" -ForegroundColor Yellow

    Write-Log "Subtitle download complete - Downloaded: $downloaded, Skipped: $skipped, Failed: $failed, No NFO: $noNfo" "INFO"

    return @{
        Total = $totalFolders
        Downloaded = $downloaded
        Skipped = $skipped
        Failed = $failed
        NoNfo = $noNfo
    }
}

<#
.SYNOPSIS
    Fixes orphaned subtitle files by renaming them to match the video file
.DESCRIPTION
    Scans for subtitle files that don't match any video filename in their folder.
    If exactly one video exists in the folder, renames the subtitle to match it.
    Handles multiple orphaned subs per folder and preserves language suffixes.
.PARAMETER Path
    The root path of the movie library
.PARAMETER WhatIf
    If specified, shows what would be renamed without making changes
#>
function Repair-OrphanedSubtitles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [switch]$WhatIf
    )

    Write-Host "`n--- Fix Orphaned Subtitles ---" -ForegroundColor Yellow
    if ($WhatIf) {
        Write-Host "(Dry run - no changes will be made)" -ForegroundColor Cyan
    }
    Write-Log "Starting orphaned subtitle repair in: $Path (WhatIf: $WhatIf)" "INFO"

    $stats = @{
        Total = 0
        Renamed = 0
        Deleted = 0
        NoVideo = 0
        MultipleVideos = 0
        AlreadyHasMatch = 0
        Errors = 0
    }
    $errorFolders = @()

    # Get all subtitle files
    $subtitleFiles = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $script:Config.SubtitleExtensions -contains $_.Extension.ToLower() }

    Write-Host "Scanning $($subtitleFiles.Count) subtitle files..." -ForegroundColor Cyan

    foreach ($sub in $subtitleFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sub.Name)
        # Extract language suffix if present
        $langSuffix = ""
        if ($baseName -match '\.(eng|en|english|spa|es|spanish|fre|fr|french|ger|de|german|por|pt|portuguese|ita|it|italian|rus|ru|russian|jpn|ja|japanese|chi|zh|chinese|kor|ko|korean|ara|ar|arabic|hin|hi|hindi|dut|nl|dutch|swe|sv|swedish|nor|no|norwegian|dan|da|danish|fin|fi|finnish|pol|pl|polish|tur|tr|turkish|heb|he|hebrew|tha|th|thai|vie|vi|vietnamese|ind|id|indonesian|msa|ms|malay|forced)$') {
            $langSuffix = $matches[0]
            $baseName = $baseName -replace '\.(eng|en|english|spa|es|spanish|fre|fr|french|ger|de|german|por|pt|portuguese|ita|it|italian|rus|ru|russian|jpn|ja|japanese|chi|zh|chinese|kor|ko|korean|ara|ar|arabic|hin|hi|hindi|dut|nl|dutch|swe|sv|swedish|nor|no|norwegian|dan|da|danish|fin|fi|finnish|pol|pl|polish|tur|tr|turkish|heb|he|hebrew|tha|th|thai|vie|vi|vietnamese|ind|id|indonesian|msa|ms|malay|forced)$', ''
        }

        # Check if this subtitle matches a video file
        $hasMatchingVideo = Get-ChildItem -Path $sub.DirectoryName -File -ErrorAction SilentlyContinue |
            Where-Object {
                $script:Config.VideoExtensions -contains $_.Extension.ToLower() -and
                [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -eq $baseName
            } | Select-Object -First 1

        if ($hasMatchingVideo) {
            # Not orphaned, skip
            continue
        }

        $stats.Total++

        # Find all video files in this folder
        $videoFiles = @(Get-ChildItem -Path $sub.DirectoryName -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
            Where-Object { $_.Name -notmatch 'sample|preview|trailer|teaser' })

        if ($videoFiles.Count -eq 0) {
            # No video file - delete the orphaned subtitle
            Write-Host "  No video: $($sub.Name)" -ForegroundColor DarkGray
            if (-not $WhatIf) {
                try {
                    Remove-Item -Path $sub.FullName -Force
                    Write-Log "Deleted orphaned subtitle (no video): $($sub.FullName)" "INFO"
                } catch {
                    Write-Log "Error deleting orphaned subtitle: $_" "ERROR"
                    $stats.Errors++
                }
            }
            $stats.Deleted++
            continue
        }

        if ($videoFiles.Count -gt 1) {
            # Multiple videos - can't determine which one to match
            Write-Host "  Multiple videos in folder: $($sub.DirectoryName)" -ForegroundColor Yellow
            $stats.MultipleVideos++
            continue
        }

        # Exactly one video - rename subtitle to match
        $video = $videoFiles[0]
        $videoBaseName = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)

        # Check if a matching subtitle already exists for this video
        $existingMatch = Get-ChildItem -Path $sub.DirectoryName -File -ErrorAction SilentlyContinue |
            Where-Object {
                $script:Config.SubtitleExtensions -contains $_.Extension.ToLower() -and
                $_.FullName -ne $sub.FullName -and
                [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -match "^$([regex]::Escape($videoBaseName))"
            } | Select-Object -First 1

        if ($existingMatch) {
            # Video already has a matching subtitle - delete this orphaned one
            Write-Host "  Already has match, deleting: $($sub.Name)" -ForegroundColor DarkGray
            if (-not $WhatIf) {
                try {
                    Remove-Item -Path $sub.FullName -Force
                    Write-Log "Deleted redundant orphaned subtitle: $($sub.FullName)" "INFO"
                } catch {
                    Write-Log "Error deleting redundant subtitle: $_" "ERROR"
                    $stats.Errors++
                }
            }
            $stats.AlreadyHasMatch++
            continue
        }

        # Build new filename
        $newName = $videoBaseName + $langSuffix + $sub.Extension

        Write-Host "  Rename: $($sub.Name) -> $newName" -ForegroundColor Green

        if (-not $WhatIf) {
            try {
                Rename-Item -Path $sub.FullName -NewName $newName -Force
                Write-Log "Renamed orphaned subtitle: $($sub.Name) -> $newName" "INFO"
                $stats.Renamed++
            } catch {
                Write-Host "    Error: $_" -ForegroundColor Red
                Write-Log "Error renaming subtitle: $_" "ERROR"
                $errorFolders += @{ Folder = $sub.DirectoryName; Reason = "Rename failed: $_" }
                $stats.Errors++
            }
        } else {
            $stats.Renamed++
        }
    }

    # Summary
    Write-Host "`n--- Orphaned Subtitle Repair Summary ---" -ForegroundColor Cyan
    Write-Host "Total orphaned found:     $($stats.Total)" -ForegroundColor White
    Write-Host "Renamed to match video:   $($stats.Renamed)" -ForegroundColor Green
    Write-Host "Deleted (no video):       $($stats.Deleted)" -ForegroundColor Gray
    Write-Host "Deleted (had match):      $($stats.AlreadyHasMatch)" -ForegroundColor Gray
    Write-Host "Skipped (multiple videos):$($stats.MultipleVideos)" -ForegroundColor Yellow
    Write-Host "Errors:                   $($stats.Errors)" -ForegroundColor $(if ($stats.Errors -gt 0) { 'Red' } else { 'Gray' })

    if ($errorFolders.Count -gt 0) {
        Write-Host "`n--- Folders With Errors ---" -ForegroundColor Yellow
        foreach ($err in $errorFolders) {
            Write-Host "  $($err.Folder)" -ForegroundColor White
            Write-Host "    $($err.Reason)" -ForegroundColor Gray
        }
    }

    Write-Log "Orphaned subtitle repair complete - Renamed: $($stats.Renamed), Deleted: $($stats.Deleted + $stats.AlreadyHasMatch), Errors: $($stats.Errors)" "INFO"

    return $stats
}

<#
.SYNOPSIS
    Removes empty folders from a media library
.DESCRIPTION
    Recursively scans for and removes empty folders, starting from the deepest
    level and working up. This handles nested empty folders correctly.
.PARAMETER Path
    The root path of the media library
.PARAMETER WhatIf
    If specified, shows what would be deleted without making changes
#>
function Remove-EmptyFolders {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [switch]$WhatIf
    )

    Write-Host "`n--- Remove Empty Folders ---" -ForegroundColor Yellow
    if ($WhatIf) {
        Write-Host "(Dry run - no changes will be made)" -ForegroundColor Cyan
    }
    Write-Log "Starting empty folder cleanup in: $Path (WhatIf: $WhatIf)" "INFO"

    $stats = @{
        Found = 0
        Deleted = 0
        Errors = 0
    }
    $deletedFolders = @()
    $errorFolders = @()

    # Get all directories, sorted by depth (deepest first) to handle nested empties
    $allFolders = Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue |
        Sort-Object { $_.FullName.Split([IO.Path]::DirectorySeparatorChar).Count } -Descending

    Write-Host "Scanning $($allFolders.Count) folders..." -ForegroundColor Cyan

    foreach ($folder in $allFolders) {
        # Check if folder is empty (no files and no non-empty subdirectories)
        $contents = Get-ChildItem -LiteralPath $folder.FullName -Force -ErrorAction SilentlyContinue

        if ($contents.Count -eq 0) {
            $stats.Found++
            $deletedFolders += $folder.FullName

            Write-Host "  Empty: $($folder.FullName)" -ForegroundColor Gray

            if (-not $WhatIf) {
                try {
                    Remove-Item -Path $folder.FullName -Force -ErrorAction Stop
                    Write-Log "Deleted empty folder: $($folder.FullName)" "INFO"
                    $stats.Deleted++
                } catch {
                    Write-Host "    Error: $_" -ForegroundColor Red
                    Write-Log "Error deleting empty folder: $_" "ERROR"
                    $errorFolders += @{ Folder = $folder.FullName; Reason = "$_" }
                    $stats.Errors++
                }
            } else {
                $stats.Deleted++
            }
        }
    }

    # Summary
    Write-Host "`n--- Empty Folder Cleanup Summary ---" -ForegroundColor Cyan
    Write-Host "Empty folders found:   $($stats.Found)" -ForegroundColor White
    Write-Host "Deleted:               $($stats.Deleted)" -ForegroundColor Green
    Write-Host "Errors:                $($stats.Errors)" -ForegroundColor $(if ($stats.Errors -gt 0) { 'Red' } else { 'Gray' })

    if ($stats.Found -eq 0) {
        Write-Host "`nNo empty folders found." -ForegroundColor Green
    }

    if ($errorFolders.Count -gt 0) {
        Write-Host "`n--- Folders With Errors ---" -ForegroundColor Yellow
        foreach ($err in $errorFolders) {
            Write-Host "  $($err.Folder)" -ForegroundColor White
            Write-Host "    $($err.Reason)" -ForegroundColor Gray
        }
    }

    Write-Log "Empty folder cleanup complete - Found: $($stats.Found), Deleted: $($stats.Deleted), Errors: $($stats.Errors)" "INFO"

    return $stats
}

<#
.SYNOPSIS
    Refreshes/repairs metadata for movies in a library
.PARAMETER Path
    The root path of the movie library
.PARAMETER MediaType
    Movies or TVShows
.PARAMETER Overwrite
    If true, overwrites existing NFO files. If false, only creates missing ones.
#>
function Invoke-MetadataRefresh {
    param(
        [string]$Path,
        [string]$MediaType = "Movies",
        [bool]$Overwrite = $false
    )

    Write-Host "`nRefreshing metadata for $MediaType in: $Path" -ForegroundColor Cyan
    Write-Log "Starting metadata refresh for $MediaType in: $Path" "INFO"

    if ($MediaType -eq "Movies") {
        $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_Trailers' }

        $totalFolders = $folders.Count
        $processedCount = 0
        $updatedCount = 0
        $skippedCount = 0
        $errorCount = 0
        $errorFolders = @()

        foreach ($folder in $folders) {
            $processedCount++
            Write-Host "`n[$processedCount/$totalFolders] $($folder.Name)" -ForegroundColor White

            # Find video file in folder
            $videoFile = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                Select-Object -First 1

            if (-not $videoFile) {
                Write-Host "  No video file found - skipping" -ForegroundColor Yellow
                $skippedCount++
                continue
            }

            $nfoPath = Join-Path $folder.FullName "$([System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)).nfo"
            $hasNfo = Test-Path $nfoPath

            # Skip if NFO exists and we're not overwriting
            if ($hasNfo -and -not $Overwrite) {
                Write-Host "  NFO exists - skipping" -ForegroundColor Gray
                $skippedCount++
                continue
            }

            # Extract title and year from folder name
            $title = $null
            $year = $null

            if ($folder.Name -match '^(.+?)\s*\((\d{4})\)') {
                $title = $Matches[1].Trim()
                $year = $Matches[2]
            } elseif ($folder.Name -match '^(.+?)\s+(\d{4})$') {
                $title = $Matches[1].Trim()
                $year = $Matches[2]
            } else {
                $title = $folder.Name
            }

            # Clean title
            $title = $title -replace '\.', ' '
            $title = $title -replace '\s+', ' '
            $title = $title.Trim()

            # Search TMDB
            $searchResult = Search-TMDBMovie -Title $title -Year $year -ApiKey $script:Config.TMDBApiKey

            if ($searchResult -and $searchResult.Id) {
                $tmdbMetadata = Get-TMDBMovieDetails -MovieId $searchResult.Id -ApiKey $script:Config.TMDBApiKey

                if ($tmdbMetadata) {
                    # Create/update NFO
                    $nfoErrorReason = ""
                    $nfoSuccess = New-MovieNFOFromTMDB -Metadata $tmdbMetadata -NFOPath $nfoPath -ErrorReasonVar "nfoErrorReason"

                    if ($nfoSuccess) {
                        Write-Host "  Created NFO: $($tmdbMetadata.Title) ($($tmdbMetadata.Year))" -ForegroundColor Green

                        # Download artwork from TMDB
                        $artworkCount = Save-MovieArtwork -Metadata $tmdbMetadata -MovieFolder $folder.FullName

                        # Download artwork from fanart.tv
                        if ($script:Config.FanartTVApiKey -and $tmdbMetadata.TMDBID) {
                            $fanartCount = Save-FanartTVArtwork -TMDBID $tmdbMetadata.TMDBID -MovieFolder $folder.FullName -ApiKey $script:Config.FanartTVApiKey
                            $artworkCount += $fanartCount
                        }

                        if ($artworkCount -gt 0) {
                            Write-Host "  Downloaded $artworkCount artwork file(s)" -ForegroundColor Cyan
                        }

                        $updatedCount++
                    } else {
                        $reason = if ($nfoErrorReason) { "NFO creation failed: $nfoErrorReason" } else { "Failed to create NFO file" }
                        $errorFolders += @{ Folder = $folder.FullName; Reason = $reason; NFOPath = $nfoPath }
                        $errorCount++
                    }
                } else {
                    Write-Host "  Could not get TMDB details" -ForegroundColor Yellow
                    $errorFolders += @{ Folder = $folder.FullName; Reason = "Could not get TMDB details" }
                    $errorCount++
                }
            } else {
                Write-Host "  Not found on TMDB" -ForegroundColor Yellow
                $errorFolders += @{ Folder = $folder.FullName; Reason = "Not found on TMDB" }
                $errorCount++
            }
        }

        Write-Host "`n--- Metadata Refresh Complete ---" -ForegroundColor Green
        Write-Host "  Processed: $processedCount folders" -ForegroundColor White
        Write-Host "  Updated: $updatedCount" -ForegroundColor Green
        Write-Host "  Skipped: $skippedCount" -ForegroundColor Gray
        Write-Host "  Errors: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { 'Yellow' } else { 'Gray' })

        # Display error locations for easy review
        if ($errorFolders.Count -gt 0) {
            Write-Host "`n--- Folders With Errors ---" -ForegroundColor Yellow
            foreach ($err in $errorFolders) {
                Write-Host "  $($err.Folder)" -ForegroundColor White
                Write-Host "    Reason: $($err.Reason)" -ForegroundColor Gray
                if ($err.NFOPath) {
                    Write-Host "    NFO Path: $($err.NFOPath)" -ForegroundColor DarkGray
                }
            }
        }
    } else {
        Write-Host "TV Show metadata refresh not yet implemented" -ForegroundColor Yellow
        Write-Host "TV Shows will be supported when TVDB integration is added" -ForegroundColor Gray
    }
}

<#
.SYNOPSIS
    Searches TMDB for a TV show by title
.PARAMETER Title
    The TV show title to search for
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with TV show metadata from TMDB
#>
function Search-TMDBTVShow {
    param(
        [string]$Title,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    try {
        $encodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)
        $url = "https://api.themoviedb.org/3/search/tv?api_key=$ApiKey&query=$encodedTitle"

        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        if ($response.results -and $response.results.Count -gt 0) {
            $show = $response.results[0]

            return @{
                Id = $show.id
                Title = $show.name
                OriginalTitle = $show.original_name
                FirstAirDate = $show.first_air_date
                Year = if ($show.first_air_date) { $show.first_air_date.Substring(0,4) } else { $null }
                Overview = $show.overview
                Rating = $show.vote_average
                PosterPath = if ($show.poster_path) { "https://image.tmdb.org/t/p/w500$($show.poster_path)" } else { $null }
            }
        }

        return $null
    }
    catch {
        Write-Log "Error searching TMDB TV: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Gets episode details from TMDB
.PARAMETER ShowId
    The TMDB TV show ID
.PARAMETER Season
    The season number
.PARAMETER Episode
    The episode number
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with episode metadata
#>
function Get-TMDBEpisode {
    param(
        [int]$ShowId,
        [int]$Season,
        [int]$Episode,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    try {
        $url = "https://api.themoviedb.org/3/tv/$ShowId/season/$Season/episode/$Episode`?api_key=$ApiKey"
        $ep = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        return @{
            Title = $ep.name
            Overview = $ep.overview
            AirDate = $ep.air_date
            Season = $ep.season_number
            Episode = $ep.episode_number
            Rating = $ep.vote_average
            StillPath = if ($ep.still_path) { "https://image.tmdb.org/t/p/w300$($ep.still_path)" } else { $null }
        }
    }
    catch {
        Write-Log "Error getting TMDB episode: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Fetches metadata from TMDB for all movies in a library
.PARAMETER Path
    The root path of the movie library
.PARAMETER ApiKey
    TMDB API key
#>
function Invoke-TMDBMetadataFetch {
    param(
        [string]$Path,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        Write-Host "TMDB API key required. Get one free at https://www.themoviedb.org/settings/api" -ForegroundColor Yellow
        return
    }

    Write-Host "`nFetching metadata from TMDB..." -ForegroundColor Yellow
    Write-Log "Starting TMDB metadata fetch for: $Path" "INFO"

    try {
        $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_Trailers' }

        if ($movieFolders.Count -eq 0) {
            Write-Host "No movie folders found" -ForegroundColor Cyan
            return
        }

        $processed = 0
        $found = 0
        $total = $movieFolders.Count

        foreach ($folder in $movieFolders) {
            $processed++
            Write-Progress -Activity "Fetching TMDB metadata" -Status "[$processed/$total] $($folder.Name)" -PercentComplete (($processed / $total) * 100)

            # Check if NFO already exists with TMDB data
            $existingNfo = Get-ChildItem -LiteralPath $folder.FullName -Filter "*.nfo" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($existingNfo) {
                $nfoContent = Get-Content $existingNfo.FullName -Raw -ErrorAction SilentlyContinue
                if ($nfoContent -match 'uniqueid type="tmdb"') {
                    Write-Host "Skipping (has TMDB data): $($folder.Name)" -ForegroundColor Gray
                    continue
                }
            }

            # Get video file
            $videoFile = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                Select-Object -First 1

            if (-not $videoFile) {
                continue
            }

            # Extract title and year from folder name
            $titleInfo = Get-NormalizedTitle -Name $folder.Name
            $searchTitle = $titleInfo.NormalizedTitle -replace '\s+', ' '

            # Search TMDB
            $searchResult = Search-TMDBMovie -Title $searchTitle -Year $titleInfo.Year -ApiKey $ApiKey

            if ($searchResult) {
                # Get detailed info
                $details = Get-TMDBMovieDetails -MovieId $searchResult.Id -ApiKey $ApiKey

                if ($details) {
                    Write-Host "Found: $($details.Title) ($($details.Year))" -ForegroundColor Green
                    $nfoPath = Join-Path $videoFile.DirectoryName "$([System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)).nfo"
                    $null = New-MovieNFOFromTMDB -Metadata $details -NFOPath $nfoPath
                    $found++
                }
            } else {
                Write-Host "Not found: $($folder.Name)" -ForegroundColor Yellow
            }

            # Rate limiting - TMDB allows 40 requests per 10 seconds
            Start-Sleep -Milliseconds 300
        }

        Write-Progress -Activity "Fetching TMDB metadata" -Completed
        Write-Host "`nTMDB fetch complete: $found/$total movies matched" -ForegroundColor Cyan
        Write-Log "TMDB metadata fetch completed: $found/$total matched" "INFO"
    }
    catch {
        Write-Host "Error during TMDB fetch: $_" -ForegroundColor Red
        Write-Log "Error during TMDB fetch: $_" "ERROR"
    }
}

#============================================
# ARCHIVE FUNCTIONS
#============================================

<#
.SYNOPSIS
    Extracts all archives in the specified path
.PARAMETER Path
    The root path to search for archives
.PARAMETER DeleteAfterExtract
    Whether to delete archives after successful extraction
#>
function Expand-Archives {
    param(
        [string]$Path,
        [switch]$DeleteAfterExtract
    )

    Write-Host "Extracting archives..." -ForegroundColor Yellow
    Write-Log "Starting archive extraction in: $Path" "INFO"

    try {
        Set-Alias sz $script:Config.SevenZipPath -Scope Script

        $unzipQueue = @()
        foreach ($ext in $script:Config.ArchiveExtensions) {
            $archives = Get-ChildItem -Path $Path -Filter $ext -Recurse -ErrorAction SilentlyContinue
            if ($archives) {
                $unzipQueue += $archives
            }
        }

        $count = $unzipQueue.Count

        if ($count -gt 0) {
            Write-Host "Found $count archive(s) to extract" -ForegroundColor Cyan
            Write-Log "Found $count archive(s) to extract" "INFO"

            $current = 1
            foreach ($archive in $unzipQueue) {
                $percentComplete = ($current / $count) * 100
                Write-Progress -Activity "Extracting archives" -Status "[$current/$count] $($archive.Name)" -PercentComplete $percentComplete

                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] [$current/$count] Would extract: $($archive.Name)" -ForegroundColor Yellow
                    Write-Log "Would extract: $($archive.FullName)" "DRY-RUN"
                } else {
                    # Extract to archive's parent directory instead of root path
                    $extractPath = $archive.DirectoryName
                    $archiveSize = $archive.Length
                    $archiveSizeMB = [math]::Round($archiveSize / 1MB, 2)

                    Write-Host "[$current/$count] Extracting: $($archive.Name) ($archiveSizeMB MB) to $extractPath" -ForegroundColor Cyan
                    Write-Log "Extracting: $($archive.FullName) ($archiveSizeMB MB)" "INFO"

                    # Use Start-Process with timeout for better control
                    $timeoutSeconds = 1800  # 30 minute timeout
                    $process = Start-Process -FilePath $script:Config.SevenZipPath `
                        -ArgumentList "x", "-o`"$extractPath`"", "`"$($archive.FullName)`"", "-r", "-y" `
                        -NoNewWindow -PassThru -Wait:$false

                    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    $lastUpdate = 0

                    # Wait for process with progress updates
                    while (-not $process.HasExited) {
                        Start-Sleep -Milliseconds 500
                        $elapsedSeconds = [math]::Floor($stopwatch.Elapsed.TotalSeconds)

                        # Update every 5 seconds
                        if ($elapsedSeconds -ge ($lastUpdate + 5)) {
                            $lastUpdate = $elapsedSeconds
                            Write-Host "  Extracting... $elapsedSeconds seconds elapsed" -ForegroundColor Gray
                        }

                        # Check timeout
                        if ($stopwatch.Elapsed.TotalSeconds -gt $timeoutSeconds) {
                            Write-Host "  Timeout reached ($timeoutSeconds seconds), killing extraction process" -ForegroundColor Red
                            Write-Log "Extraction timeout for: $($archive.FullName)" "ERROR"
                            $process.Kill()
                            $script:Stats.ArchivesFailed++
                            $stopwatch.Stop()
                            continue
                        }
                    }
                    $stopwatch.Stop()

                    $exitCode = $process.ExitCode
                    $elapsedTime = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)

                    if ($exitCode -ne 0) {
                        Write-Host "Warning: Failed to extract $($archive.Name) (exit code: $exitCode)" -ForegroundColor Yellow
                        Write-Log "Failed to extract: $($archive.FullName) (exit code: $exitCode)" "WARNING"
                        $script:Stats.ArchivesFailed++
                    } else {
                        Write-Host "Successfully extracted $($archive.Name) in $elapsedTime seconds" -ForegroundColor Green
                        Write-Log "Successfully extracted: $($archive.FullName) in $elapsedTime seconds" "INFO"
                        $script:Stats.ArchivesExtracted++
                    }
                }
                $current++
            }
            Write-Progress -Activity "Extracting archives" -Completed

            # Delete archives after extraction
            if ($DeleteAfterExtract -and -not $script:Config.DryRun) {
                Write-Host "Deleting archive files..." -ForegroundColor Yellow
                foreach ($pattern in $script:Config.ArchiveCleanupPatterns) {
                    $archivesToDelete = Get-ChildItem -Path $Path -Include $pattern -Recurse -ErrorAction SilentlyContinue
                    foreach ($archiveFile in $archivesToDelete) {
                        $archiveSize = $archiveFile.Length
                        Remove-Item -Path $archiveFile.FullName -Force -ErrorAction SilentlyContinue
                        Write-Log "Deleted archive: $($archiveFile.FullName)" "INFO"
                        $script:Stats.FilesDeleted++
                        $script:Stats.BytesDeleted += $archiveSize
                    }
                }
            } elseif ($script:Config.DryRun) {
                Write-Host "[DRY-RUN] Would delete all archive files after extraction" -ForegroundColor Yellow
            }

            Write-Host "Archives processed" -ForegroundColor Green
        } else {
            Write-Host "No archives found to extract" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error processing archives: $_" -ForegroundColor Red
        Write-Log "Error processing archives: $_" "ERROR"
    }
}

#============================================
# FOLDER ORGANIZATION FUNCTIONS
#============================================

<#
.SYNOPSIS
    Detects and expands movie collection packs (trilogies, etc.) into individual movie folders
.DESCRIPTION
    Identifies folders containing multiple movie files (trilogies, collections, etc.) and
    extracts each movie into its own properly-named folder at the root level.
    Detects packs by: folder name patterns (trilogy, collection, etc.) or multiple large video files.
.PARAMETER Path
    The root path to scan for collection packs
#>
function Expand-MoviePacks {
    param(
        [string]$Path
    )

    Write-Host "`n--- Movie Pack Detection ---" -ForegroundColor Cyan

    # Patterns that indicate a collection pack
    $packPatterns = @(
        'trilogy',
        'duology',
        'quadrilogy',
        'pentalogy',
        'hexalogy',
        'collection',
        'saga',
        'complete',
        'boxset',
        'box set',
        '\d+-film',
        '\d+ film'
    )
    $packRegex = ($packPatterns | ForEach-Object { "($_)" }) -join '|'

    $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue

    $packsFound = 0
    $moviesExtracted = 0

    foreach ($folder in $folders) {
        $isPackByName = $folder.Name -match $packRegex

        # Get video files in this folder (not in subfolders)
        $videoFiles = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
            Where-Object {
                $script:Config.VideoExtensions -contains $_.Extension.ToLower() -and
                $_.Length -gt 500MB  # Only count substantial video files
            }

        $isPackByFileCount = $videoFiles.Count -gt 1

        if (-not ($isPackByName -or $isPackByFileCount)) {
            continue
        }

        if ($videoFiles.Count -lt 2) {
            continue  # Need at least 2 movies to be a pack
        }

        Write-Host "  Found pack: $($folder.Name)" -ForegroundColor Yellow
        Write-Host "    Contains $($videoFiles.Count) movies" -ForegroundColor Gray
        $packsFound++

        # Extract series/franchise name from pack folder for potential title inheritance
        # e.g., "Austin Powers Trilogy (1997-2002)" -> "Austin Powers"
        # e.g., "Rush Hour Collection" -> "Rush Hour"
        $packName = $folder.Name
        $seriesName = $null

        # Remove year ranges, quality tags, and pack indicators
        $seriesName = $packName -replace '\s*\([\d\-_]+[^)]*\)\s*$', ''  # Remove (1997-2002_1080p) etc
        $seriesName = $seriesName -replace '\s*(trilogy|duology|quadrilogy|pentalogy|hexalogy|collection|saga|complete|boxset|box set|\d+-?film)\s*$', ''
        $seriesName = $seriesName -replace '[\.\-_]', ' '
        $seriesName = $seriesName -replace '\s+', ' '
        $seriesName = $seriesName.Trim()

        if ($seriesName) {
            Write-Host "    Series name: $seriesName" -ForegroundColor DarkCyan
        }

        foreach ($videoFile in $videoFiles) {
            # Try to extract movie title and year from filename
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)

            # Try to find year in filename
            $year = $null
            if ($baseName -match '\b(19|20)\d{2}\b') {
                $year = $Matches[0]
            }

            # Clean up the title
            $title = $baseName
            # Remove quality tags and release info (order matters - do this before replacing dots)
            $title = $title -replace '\b(720p|1080p|2160p|4K|UHD|HDRip|DVDRip|DVD|BRRip|BDRip|BD[\s\.\-]?Rip|BluRay|Blu[\s\.\-]?Ray|WEB-DL|WEB[\s\.\-]?DL|WEBRip|WEB[\s\.\-]?Rip|HDTV|x264|x265|HEVC|AAC|AC3|DTS|10bit|H\.?264|H\.?265|REMUX|Atmos|TrueHD|DDP5\.1|DD5\.1|5\.1|7\.1)\b.*$', ''
            $title = $title -replace '[\.\-_]', ' '
            $title = $title -replace '\s+', ' '
            $title = $title.Trim()

            # Remove year from title if present (we'll add it back in proper format)
            if ($year) {
                $title = $title -replace "\s*\(?$year\)?", ''
                $title = $title.Trim()
            }

            # Try TMDB lookup to get the correct full title
            $tmdbTitle = $null
            if ($script:Config.TMDBApiKey -and $year) {
                # First try with just the extracted title
                $searchResult = Search-TMDBMovie -Title $title -Year $year -ApiKey $script:Config.TMDBApiKey

                # If that fails and we have a series name, try combining them
                if (-not $searchResult -and $seriesName -and $title -ne $seriesName) {
                    # Try: "Series Name: Subtitle" (e.g., "Austin Powers: Goldmember")
                    $combinedTitle = "$seriesName $title"
                    Write-Host "      TMDB lookup failed for '$title', trying '$combinedTitle'..." -ForegroundColor DarkGray
                    $searchResult = Search-TMDBMovie -Title $combinedTitle -Year $year -ApiKey $script:Config.TMDBApiKey

                    # Also try with common sequel patterns
                    if (-not $searchResult) {
                        # Try: "Series Name in Subtitle" (e.g., "Austin Powers in Goldmember")
                        $combinedTitle = "$seriesName in $title"
                        $searchResult = Search-TMDBMovie -Title $combinedTitle -Year $year -ApiKey $script:Config.TMDBApiKey
                    }
                }

                if ($searchResult -and $searchResult.Title) {
                    $tmdbTitle = $searchResult.Title
                    Write-Host "      TMDB match: $tmdbTitle" -ForegroundColor DarkGreen
                }
            }

            # Use TMDB title if found, otherwise fall back to extracted title
            $finalTitle = if ($tmdbTitle) { $tmdbTitle } else { $title }

            # Create proper folder name
            $newFolderName = if ($year) { "$finalTitle ($year)" } else { $finalTitle }

            # Sanitize folder name
            $newFolderName = $newFolderName -replace '[<>:"/\\|?*]', ''

            $newFolderPath = Join-Path $Path $newFolderName

            Write-Host "    Extracting: $($videoFile.Name)" -ForegroundColor Gray
            Write-Host "      -> $newFolderName" -ForegroundColor Green

            if ($script:Config.DryRun) {
                Write-Host "      [DRY RUN] Would create folder and move file" -ForegroundColor Yellow
                $moviesExtracted++
                continue
            }

            try {
                # Create the new folder
                if (-not (Test-Path $newFolderPath)) {
                    New-Item -Path $newFolderPath -ItemType Directory -Force | Out-Null
                }

                # Move the video file
                $newVideoPath = Join-Path $newFolderPath $videoFile.Name
                Move-Item -LiteralPath $videoFile.FullName -Destination $newVideoPath -Force

                # Look for matching subtitle files
                $videoBaseName = [System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)
                $subtitleFiles = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.Extension -match '\.(srt|sub|idx|ass|ssa|vtt)$' -and
                        $_.BaseName -like "$videoBaseName*"
                    }

                foreach ($subFile in $subtitleFiles) {
                    $newSubPath = Join-Path $newFolderPath $subFile.Name
                    Move-Item -LiteralPath $subFile.FullName -Destination $newSubPath -Force
                    Write-Host "      Moved subtitle: $($subFile.Name)" -ForegroundColor DarkGray
                }

                $moviesExtracted++
                Write-Log "Extracted from pack: $($videoFile.Name) -> $newFolderName" "INFO"
            }
            catch {
                Write-Host "      Error: $_" -ForegroundColor Red
                Write-Log "Error extracting from pack: $_" "ERROR"
            }
        }

        # Check if pack folder is now empty (or only has small files)
        $remainingFiles = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -gt 1MB }

        if ($remainingFiles.Count -eq 0) {
            if (-not $script:Config.DryRun) {
                Write-Host "    Removing empty pack folder" -ForegroundColor DarkGray
                Remove-Item -LiteralPath $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue
            } else {
                Write-Host "    [DRY RUN] Would remove empty pack folder" -ForegroundColor Yellow
            }
        }
    }

    if ($packsFound -eq 0) {
        Write-Host "  No movie packs detected" -ForegroundColor Gray
    } else {
        Write-Host "`n  Packs found: $packsFound" -ForegroundColor White
        Write-Host "  Movies extracted: $moviesExtracted" -ForegroundColor Green
    }
}

<#
.SYNOPSIS
    Creates individual folders for loose video files
.PARAMETER Path
    The root path containing loose files
#>
function New-FoldersForLooseFiles {
    param(
        [string]$Path
    )

    Write-Host "Creating folders for loose video files..." -ForegroundColor Yellow
    Write-Log "Starting folder creation for loose files in: $Path" "INFO"

    try {
        $looseFiles = Get-ChildItem -Path $Path -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($looseFiles) {
            foreach ($file in $looseFiles) {
                try {
                    $dir = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                    $fullDirPath = Join-Path -Path $Path -ChildPath $dir

                    if ($script:Config.DryRun) {
                        Write-Host "[DRY-RUN] Would create folder '$dir' and move: $($file.Name)" -ForegroundColor Yellow
                        Write-Log "Would create folder '$dir' and move: $($file.Name)" "DRY-RUN"
                    } else {
                        Write-Host "Processing: $($file.Name)" -ForegroundColor Cyan

                        if (-not (Test-Path $fullDirPath)) {
                            New-Item -Path $fullDirPath -ItemType Directory -ErrorAction Stop | Out-Null
                            $script:Stats.FoldersCreated++
                        }

                        Move-Item -Path $file.FullName -Destination $fullDirPath -Force -ErrorAction Stop
                        Write-Host "Moved to: $dir" -ForegroundColor Green
                        Write-Log "Moved '$($file.Name)' to folder '$dir'" "INFO"
                        $script:Stats.FilesMoved++
                    }
                }
                catch {
                    Write-Host "Warning: Could not process $($file.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error processing $($file.Name): $_" "ERROR"
                }
            }
        } else {
            Write-Host "No loose video files found" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error organizing loose files: $_" -ForegroundColor Red
        Write-Log "Error organizing loose files: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Cleans folder names by removing tags and formatting
.PARAMETER Path
    The root path containing folders to clean
#>
function Rename-CleanFolderNames {
    param(
        [string]$Path
    )

    Write-Host "Cleaning folder names..." -ForegroundColor Yellow
    Write-Log "Starting folder name cleaning in: $Path" "INFO"

    try {
        # Remove bracketed content from folder names (e.g., [Anime Time], [YTS.MX], [1080p])
        # This handles brackets at start, end, or anywhere in the name
        Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '\[' } |
            ForEach-Object {
                try {
                    $newName = $_.Name
                    # Remove all bracketed content
                    $newName = $newName -replace '\[[^\]]*\]', ''
                    # Remove any orphan brackets
                    $newName = $newName -replace '[\[\]]', ''
                    # Clean up whitespace
                    $newName = $newName -replace '\s+', ' '
                    $newName = $newName.Trim()

                    if ($newName -and $newName -ne $_.Name) {
                        if ($script:Config.DryRun) {
                            Write-Host "[DRY-RUN] Would rename '$($_.Name)' to '$newName'" -ForegroundColor Yellow
                            Write-Log "Would rename '$($_.Name)' to '$newName'" "DRY-RUN"
                        } else {
                            Rename-Item -LiteralPath $_.FullName -NewName $newName -ErrorAction Stop
                            Write-Log "Renamed '$($_.Name)' to '$newName'" "INFO"
                            $script:Stats.FoldersRenamed++
                        }
                    }
                }
                catch {
                    Write-Host "Warning: Could not rename $($_.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error renaming $($_.Name): $_" "ERROR"
                }
            }

        # Remove tags from folder names
        foreach ($tag in $script:Config.Tags) {
            Get-ChildItem -Path $Path -Directory -Filter "*$tag*" -ErrorAction SilentlyContinue |
                ForEach-Object {
                    try {
                        $newName = ($_.Name -split [regex]::Escape($tag))[0].TrimEnd(' ', '.', '-')
                        if ($newName -and $newName -ne $_.Name) {
                            if ($script:Config.DryRun) {
                                Write-Host "[DRY-RUN] Would rename '$($_.Name)' to '$newName'" -ForegroundColor Yellow
                                Write-Log "Would rename '$($_.Name)' to '$newName'" "DRY-RUN"
                            } else {
                                Rename-Item -Path $_.FullName -NewName $newName -ErrorAction Stop
                                Write-Log "Renamed '$($_.Name)' to '$newName'" "INFO"
                                $script:Stats.FoldersRenamed++
                            }
                        }
                    }
                    catch {
                        Write-Host "Warning: Could not rename $($_.Name): $_" -ForegroundColor Yellow
                        Write-Log "Error renaming $($_.Name): $_" "ERROR"
                    }
                }
        }

        # Replace dots with spaces in folder names
        Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '\.' } |
            ForEach-Object {
                try {
                    $newName = $_.Name -replace '\.', ' '
                    $newName = $newName -replace '\s+', ' '  # Remove multiple spaces
                    $newName = $newName.Trim()

                    if ($newName -ne $_.Name) {
                        if ($script:Config.DryRun) {
                            Write-Host "[DRY-RUN] Would rename '$($_.Name)' to '$newName'" -ForegroundColor Yellow
                            Write-Log "Would rename '$($_.Name)' to '$newName'" "DRY-RUN"
                        } else {
                            Rename-Item -Path $_.FullName -NewName $newName -Force -ErrorAction Stop
                            Write-Log "Renamed '$($_.Name)' to '$newName'" "INFO"
                            $script:Stats.FoldersRenamed++
                        }
                    }
                }
                catch {
                    Write-Host "Warning: Could not rename $($_.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error renaming $($_.Name): $_" "ERROR"
                }
            }

        # Format years with parentheses (both 19xx and 20xx)
        Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '\s(19|20)\d{2}$' } |
            ForEach-Object {
                try {
                    $newName = $_.Name -replace '\s((19|20)\d{2})$', ' ($1)'

                    if ($newName -ne $_.Name) {
                        if ($script:Config.DryRun) {
                            Write-Host "[DRY-RUN] Would rename '$($_.Name)' to '$newName'" -ForegroundColor Yellow
                            Write-Log "Would rename '$($_.Name)' to '$newName'" "DRY-RUN"
                        } else {
                            Rename-Item -Path $_.FullName -NewName $newName -ErrorAction Stop
                            Write-Log "Renamed '$($_.Name)' to '$newName'" "INFO"
                            $script:Stats.FoldersRenamed++
                        }
                    }
                }
                catch {
                    Write-Host "Warning: Could not rename $($_.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error renaming $($_.Name): $_" "ERROR"
                }
            }

        if ($script:Stats.FoldersRenamed -gt 0 -or $script:Config.DryRun) {
            Write-Host "Folder names cleaned" -ForegroundColor Green
        } else {
            Write-Host "No folders needed renaming" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error cleaning folder names: $_" -ForegroundColor Red
        Write-Log "Error cleaning folder names: $_" "ERROR"
    }
}

function Rename-CleanFileNames {
    param(
        [string]$Path
    )

    Write-Host "Cleaning file names (removing brackets)..." -ForegroundColor Yellow
    Write-Log "Starting file name cleaning in: $Path" "INFO"

    $filesRenamed = 0

    try {
        # Get all files with brackets in their names (recursively)
        Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '\[' } |
            ForEach-Object {
                try {
                    $originalName = $_.Name
                    $extension = $_.Extension
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($originalName)

                    # Remove all bracketed content
                    $newBaseName = $baseName -replace '\[[^\]]*\]', ''
                    # Remove any orphan brackets
                    $newBaseName = $newBaseName -replace '[\[\]]', ''
                    # Clean up whitespace
                    $newBaseName = $newBaseName -replace '\s+', ' '
                    $newBaseName = $newBaseName.Trim()

                    $newName = $newBaseName + $extension

                    if ($newName -and $newName -ne $originalName) {
                        $newPath = Join-Path $_.DirectoryName $newName

                        # Check if target already exists
                        if (Test-Path -LiteralPath $newPath) {
                            Write-Host "  Skipped: '$originalName' - target already exists" -ForegroundColor Yellow
                            Write-Log "Skipped renaming '$originalName' - target '$newName' already exists" "WARNING"
                        } else {
                            if ($script:Config.DryRun) {
                                Write-Host "  [DRY-RUN] Would rename '$originalName' to '$newName'" -ForegroundColor Yellow
                                Write-Log "Would rename file '$originalName' to '$newName'" "DRY-RUN"
                            } else {
                                Rename-Item -LiteralPath $_.FullName -NewName $newName -ErrorAction Stop
                                Write-Host "  Renamed: '$originalName' -> '$newName'" -ForegroundColor Gray
                                Write-Log "Renamed file '$originalName' to '$newName'" "INFO"
                                $filesRenamed++
                            }
                        }
                    }
                }
                catch {
                    Write-Host "  Warning: Could not rename $($_.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error renaming file $($_.Name): $_" "ERROR"
                }
            }

        if ($filesRenamed -gt 0 -or $script:Config.DryRun) {
            Write-Host "File names cleaned ($filesRenamed files renamed)" -ForegroundColor Green
        } else {
            Write-Host "No files needed bracket removal" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error cleaning file names: $_" -ForegroundColor Red
        Write-Log "Error cleaning file names: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Removes empty folders from the specified path
.PARAMETER Path
    The root path to search for empty folders
#>
function Remove-EmptyFolders {
    param(
        [string]$Path
    )

    Write-Host "Cleaning up empty folders..." -ForegroundColor Yellow
    Write-Log "Starting empty folder cleanup in: $Path" "INFO"

    try {
        # Get all empty directories (recursively, deepest first)
        $emptyFolders = Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue |
            Where-Object { (Get-ChildItem -Path $_.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0 } |
            Sort-Object { $_.FullName.Length } -Descending

        if ($emptyFolders) {
            Write-Host "Found $($emptyFolders.Count) empty folder(s)" -ForegroundColor Cyan

            foreach ($folder in $emptyFolders) {
                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would remove: $($folder.FullName)" -ForegroundColor Yellow
                    Write-Log "Would remove empty folder: $($folder.FullName)" "DRY-RUN"
                } else {
                    Remove-Item -Path $folder.FullName -Force -ErrorAction SilentlyContinue
                    Write-Log "Removed empty folder: $($folder.FullName)" "INFO"
                    $script:Stats.EmptyFoldersRemoved++
                }
            }
            Write-Host "Empty folders removed" -ForegroundColor Green
        } else {
            Write-Host "No empty folders found" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Warning: Could not remove all empty folders: $_" -ForegroundColor Yellow
        Write-Log "Error removing empty folders: $_" "ERROR"
    }
}

#============================================
# TV SHOW FUNCTIONS
#============================================

<#
.SYNOPSIS
    Parses a filename to extract season and episode information
.PARAMETER FileName
    The filename to parse
.OUTPUTS
    Hashtable with Season, Episode, Episodes (for multi-ep), Title, and EpisodeTitle
.DESCRIPTION
    Supports multiple naming formats:
    - S01E01, s01e01
    - S01E01E02, S01E01-E03 (multi-episode)
    - 1x01, 01x01
    - Season 1 Episode 1
    - Part 1, Pt.1
#>
function Get-EpisodeInfo {
    param(
        [string]$FileName
    )

    $info = @{
        Season = $null
        Episode = $null
        Episodes = @()      # For multi-episode files
        ShowTitle = $null
        EpisodeTitle = $null
        IsMultiEpisode = $false
    }

    # Remove extension
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

    # Pattern 1: S01E01 or S01E01E02 or S01E01-E03 (most common)
    if ($baseName -match '^(.+?)[.\s_-]+[Ss](\d{1,2})[Ee](\d{1,2})(?:[Ee-](\d{1,2}))?(?:[Ee-](\d{1,2}))?(.*)$') {
        $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
        $info.Season = [int]$Matches[2]
        $info.Episode = [int]$Matches[3]
        $info.Episodes += [int]$Matches[3]

        # Check for multi-episode
        if ($Matches[4]) {
            $info.Episodes += [int]$Matches[4]
            $info.IsMultiEpisode = $true
        }
        if ($Matches[5]) {
            $info.Episodes += [int]$Matches[5]
        }

        # Try to get episode title from remainder
        if ($Matches[6]) {
            $remainder = $Matches[6] -replace '^[.\s_-]+', '' -replace '\.', ' '
            # Remove quality tags from episode title
            $remainder = $remainder -replace '\s*(720p|1080p|2160p|4K|HDTV|WEB-DL|WEBRip|BluRay|x264|x265|HEVC|AAC|AC3).*$', ''
            if ($remainder.Trim()) {
                $info.EpisodeTitle = $remainder.Trim()
            }
        }
    }
    # Pattern 2: 1x01 format
    elseif ($baseName -match '^(.+?)[.\s_-]+(\d{1,2})x(\d{1,2})(.*)$') {
        $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
        $info.Season = [int]$Matches[2]
        $info.Episode = [int]$Matches[3]
        $info.Episodes += [int]$Matches[3]
    }
    # Pattern 3: Season 1 Episode 1
    elseif ($baseName -match '^(.+?)[.\s_-]+Season\s*(\d{1,2})[.\s_-]+Episode\s*(\d{1,2})(.*)$') {
        $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
        $info.Season = [int]$Matches[2]
        $info.Episode = [int]$Matches[3]
        $info.Episodes += [int]$Matches[3]
    }
    # Pattern 4: Show.Name.101 (season 1, episode 01)
    elseif ($baseName -match '^(.+?)[.\s_-]+(\d)(\d{2})[.\s_-](.*)$') {
        $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
        $info.Season = [int]$Matches[2]
        $info.Episode = [int]$Matches[3]
        $info.Episodes += [int]$Matches[3]
    }

    # Clean up show title
    if ($info.ShowTitle) {
        $info.ShowTitle = $info.ShowTitle.Trim()
    }

    return $info
}

<#
.SYNOPSIS
    Organizes TV show files into Season folders
.PARAMETER Path
    The root path of the TV show
.DESCRIPTION
    Scans video files, extracts season/episode info, and organizes into Season XX folders
#>
function Invoke-SeasonOrganization {
    param(
        [string]$Path
    )

    Write-Host "Organizing episodes into season folders..." -ForegroundColor Yellow
    Write-Log "Starting season organization in: $Path" "INFO"

    try {
        # Find all video files
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($videoFiles.Count -eq 0) {
            Write-Host "No video files found" -ForegroundColor Cyan
            return
        }

        Write-Host "Found $($videoFiles.Count) video file(s)" -ForegroundColor Cyan

        $organized = 0
        foreach ($file in $videoFiles) {
            $epInfo = Get-EpisodeInfo -FileName $file.Name

            if ($null -ne $epInfo.Season) {
                $seasonFolder = "Season {0:D2}" -f $epInfo.Season
                $seasonPath = Join-Path $Path $seasonFolder

                # Skip if already in correct season folder
                if ($file.Directory.Name -eq $seasonFolder) {
                    Write-Host "Already organized: $($file.Name)" -ForegroundColor Gray
                    continue
                }

                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would move to $seasonFolder : $($file.Name)" -ForegroundColor Yellow
                    Write-Log "Would move $($file.Name) to $seasonFolder" "DRY-RUN"
                } else {
                    # Create season folder if needed
                    if (-not (Test-Path $seasonPath)) {
                        New-Item -Path $seasonPath -ItemType Directory -Force | Out-Null
                        Write-Host "Created folder: $seasonFolder" -ForegroundColor Green
                        Write-Log "Created season folder: $seasonPath" "INFO"
                        $script:Stats.FoldersCreated++
                    }

                    # Move file
                    $destPath = Join-Path $seasonPath $file.Name
                    try {
                        Move-Item -Path $file.FullName -Destination $destPath -Force -ErrorAction Stop
                        Write-Host "Moved to $seasonFolder : $($file.Name)" -ForegroundColor Cyan
                        Write-Log "Moved $($file.Name) to $seasonFolder" "INFO"
                        $script:Stats.FilesMoved++
                        $organized++
                    }
                    catch {
                        Write-Host "Warning: Could not move $($file.Name): $_" -ForegroundColor Yellow
                        Write-Log "Error moving $($file.FullName): $_" "WARNING"
                    }
                }
            } else {
                Write-Host "Could not parse: $($file.Name)" -ForegroundColor Yellow
                Write-Log "Could not parse season/episode from: $($file.Name)" "WARNING"
            }
        }

        if ($organized -gt 0) {
            Write-Host "Organized $organized episode(s) into season folders" -ForegroundColor Green
        } else {
            Write-Host "No episodes needed organizing" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error organizing seasons: $_" -ForegroundColor Red
        Write-Log "Error organizing seasons: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Renames episode files to a consistent format
.PARAMETER Path
    The root path of the TV show
.PARAMETER Format
    The naming format to use (default: "{ShowTitle} - S{Season:D2}E{Episode:D2}")
.DESCRIPTION
    Standardizes episode filenames while preserving episode information
#>
function Rename-EpisodeFiles {
    param(
        [string]$Path,
        [string]$Format = "{0} - S{1:D2}E{2:D2}"
    )

    Write-Host "Standardizing episode filenames..." -ForegroundColor Yellow
    Write-Log "Starting episode rename in: $Path" "INFO"

    try {
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($videoFiles.Count -eq 0) {
            Write-Host "No video files found" -ForegroundColor Cyan
            return
        }

        $renamed = 0
        foreach ($file in $videoFiles) {
            $epInfo = Get-EpisodeInfo -FileName $file.Name

            if ($null -ne $epInfo.Season -and $epInfo.ShowTitle) {
                # Build new filename
                if ($epInfo.IsMultiEpisode) {
                    # Multi-episode: Show - S01E01-E03
                    $episodePart = "E" + ($epInfo.Episodes | ForEach-Object { "{0:D2}" -f $_ }) -join "-E"
                    $newName = "{0} - S{1:D2}{2}" -f $epInfo.ShowTitle, $epInfo.Season, $episodePart
                } else {
                    $newName = $Format -f $epInfo.ShowTitle, $epInfo.Season, $epInfo.Episode
                }

                # Add episode title if available
                if ($epInfo.EpisodeTitle) {
                    $newName += " - $($epInfo.EpisodeTitle)"
                }

                $newName += $file.Extension

                # Skip if already correctly named
                if ($file.Name -eq $newName) {
                    continue
                }

                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would rename:" -ForegroundColor Yellow
                    Write-Host "  From: $($file.Name)" -ForegroundColor Gray
                    Write-Host "  To:   $newName" -ForegroundColor Cyan
                    Write-Log "Would rename $($file.Name) to $newName" "DRY-RUN"
                } else {
                    try {
                        Rename-Item -Path $file.FullName -NewName $newName -ErrorAction Stop
                        Write-Host "Renamed: $($file.Name) -> $newName" -ForegroundColor Cyan
                        Write-Log "Renamed $($file.Name) to $newName" "INFO"
                        $script:Stats.FoldersRenamed++  # Reusing stat for file renames
                        $renamed++
                    }
                    catch {
                        Write-Host "Warning: Could not rename $($file.Name): $_" -ForegroundColor Yellow
                        Write-Log "Error renaming $($file.FullName): $_" "WARNING"
                    }
                }
            }
        }

        if ($renamed -gt 0) {
            Write-Host "Renamed $renamed episode file(s)" -ForegroundColor Green
        } else {
            Write-Host "No episodes needed renaming" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error renaming episodes: $_" -ForegroundColor Red
        Write-Log "Error renaming episodes: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Displays a summary of detected TV show episodes
.PARAMETER Path
    The root path of the TV show
#>
function Show-EpisodeSummary {
    param(
        [string]$Path
    )

    Write-Host "`nEpisode Summary:" -ForegroundColor Cyan
    Write-Log "Generating episode summary for: $Path" "INFO"

    try {
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($videoFiles.Count -eq 0) {
            Write-Host "No video files found" -ForegroundColor Cyan
            return
        }

        # Group by season
        $seasons = @{}
        foreach ($file in $videoFiles) {
            $epInfo = Get-EpisodeInfo -FileName $file.Name
            if ($null -ne $epInfo.Season) {
                $seasonKey = $epInfo.Season
                if (-not $seasons.ContainsKey($seasonKey)) {
                    $seasons[$seasonKey] = @()
                }
                $seasons[$seasonKey] += @{
                    Episode = $epInfo.Episode
                    Episodes = $epInfo.Episodes
                    FileName = $file.Name
                    IsMulti = $epInfo.IsMultiEpisode
                }
            }
        }

        if ($seasons.Count -eq 0) {
            Write-Host "Could not parse any episode information" -ForegroundColor Yellow
            return
        }

        # Display summary
        foreach ($season in ($seasons.Keys | Sort-Object)) {
            $episodes = $seasons[$season] | Sort-Object { $_.Episode }
            $episodeNums = ($episodes | ForEach-Object {
                if ($_.IsMulti) {
                    "E" + ($_.Episodes -join "-E")
                } else {
                    "E{0:D2}" -f $_.Episode
                }
            }) -join ", "

            Write-Host "  Season $season : $($episodes.Count) episode(s) - $episodeNums" -ForegroundColor White
        }

        # Check for gaps
        foreach ($season in ($seasons.Keys | Sort-Object)) {
            $episodes = $seasons[$season] | ForEach-Object { $_.Episode } | Sort-Object -Unique
            $expected = 1..($episodes | Measure-Object -Maximum).Maximum
            $missing = $expected | Where-Object { $_ -notin $episodes }
            if ($missing) {
                Write-Host "  Season $season missing: E$($missing -join ', E')" -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "Error generating summary: $_" -ForegroundColor Red
        Write-Log "Error generating episode summary: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Moves all video files to the root directory
.PARAMETER Path
    The root path to search for video files
#>
function Move-VideoFilesToRoot {
    param(
        [string]$Path
    )

    Write-Host "Moving video files to root..." -ForegroundColor Yellow
    Write-Log "Starting video file move to root in: $Path" "INFO"

    try {
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
            Where-Object { $_.DirectoryName -ne $Path }  # Exclude files already in root

        if ($videoFiles) {
            Write-Host "Found $($videoFiles.Count) video file(s)" -ForegroundColor Cyan
            Write-Log "Found $($videoFiles.Count) video file(s)" "INFO"

            $current = 1
            foreach ($file in $videoFiles) {
                $percentComplete = ($current / $videoFiles.Count) * 100
                Write-Progress -Activity "Moving video files" -Status "[$current/$($videoFiles.Count)] $($file.Name)" -PercentComplete $percentComplete

                try {
                    if ($script:Config.DryRun) {
                        Write-Host "[DRY-RUN] [$current/$($videoFiles.Count)] Would move: $($file.Name)" -ForegroundColor Yellow
                        Write-Log "Would move: $($file.FullName) to root" "DRY-RUN"
                    } else {
                        Move-Item -Path $file.FullName -Destination $Path -Force -ErrorAction Stop
                        Write-Host "[$current/$($videoFiles.Count)] Moved: $($file.Name)" -ForegroundColor Cyan
                        Write-Log "Moved: $($file.FullName) to root" "INFO"
                        $script:Stats.FilesMoved++
                    }
                }
                catch {
                    Write-Host "Warning: Could not move $($file.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error moving $($file.FullName): $_" "ERROR"
                }
                $current++
            }
            Write-Progress -Activity "Moving video files" -Completed
            Write-Host "Video files processed" -ForegroundColor Green
        } else {
            Write-Host "No video files found to move" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error moving video files: $_" -ForegroundColor Red
        Write-Log "Error moving video files: $_" "ERROR"
    }
}

#============================================
# MAIN PROCESSING FUNCTIONS
#============================================

<#
.SYNOPSIS
    Processes a movie library
.PARAMETER Path
    The root path of the movie library
#>
function Invoke-MovieProcessing {
    param(
        [string]$Path
    )

    Write-Host "`nMovie Routine" -ForegroundColor Magenta
    Write-Log "Movie routine started for path: $Path" "INFO"

    if ($script:Config.DryRun) {
        Write-Host "`n*** DRY-RUN MODE - Previewing changes only ***`n" -ForegroundColor Yellow
    }

    # Step 1: Verify 7-Zip
    if (-not (Test-SevenZipInstallation)) {
        Write-Host "Cannot proceed without 7-Zip" -ForegroundColor Red
        return
    }

    # Step 2: Process trailers (move to _Trailers folder before deletion)
    if ($script:Config.KeepTrailers) {
        Move-TrailersToFolder -Path $Path
    }

    # Step 3: Remove unnecessary files
    Remove-UnnecessaryFiles -Path $Path

    # Step 4: Process subtitles
    Invoke-SubtitleProcessing -Path $Path

    # Step 5: Extract archives
    Expand-Archives -Path $Path -DeleteAfterExtract

    # Step 6: Expand movie packs (trilogies, collections, etc.)
    Expand-MoviePacks -Path $Path

    # Step 7: Create folders for loose files
    New-FoldersForLooseFiles -Path $Path

    # Step 8: Clean folder names
    Rename-CleanFolderNames -Path $Path

    # Step 8b: Clean file names (remove brackets)
    Rename-CleanFileNames -Path $Path

    # Step 9: Generate NFO files (if enabled)
    Invoke-NFOGeneration -Path $Path

    # Step 10: Show existing NFO metadata summary
    Show-NFOMetadata -Path $Path

    # Step 11: Check for duplicates (if enabled)
    if ($script:Config.CheckDuplicates) {
        Show-DuplicateReport -Path $Path

        $removeDupes = Read-Host "`nRemove duplicate movies? (Y/N) [N]"
        if ($removeDupes -eq 'Y' -or $removeDupes -eq 'y') {
            Remove-DuplicateMovies -Path $Path
        }
    }

    # Step 12: Codec analysis
    Write-Host "`n--- Codec Analysis ---" -ForegroundColor Cyan
    $analysis = Invoke-CodecAnalysis -Path $Path

    if ($analysis -and $analysis.NeedTranscode.Count -gt 0) {
        Write-Host "`nFound $($analysis.NeedTranscode.Count) file(s) needing conversion:" -ForegroundColor Yellow

        $i = 1
        foreach ($file in $analysis.NeedTranscode | Select-Object -First 10) {
            $modeLabel = if ($file.TranscodeMode -eq "remux") { "[remux]" } else { "[transcode]" }
            Write-Host "  $i. $modeLabel $($file.FileName)" -ForegroundColor Gray
            $i++
        }
        if ($analysis.NeedTranscode.Count -gt 10) {
            Write-Host "  ... and $($analysis.NeedTranscode.Count - 10) more" -ForegroundColor Gray
        }

        Write-Host "`nOptions:" -ForegroundColor Cyan
        Write-Host "  1. Run now (transcoding may take a long time)"
        Write-Host "  2. Save script for later"
        Write-Host "  3. Skip"

        $transcodeChoice = Read-Host "`nSelect option [3]"

        switch ($transcodeChoice) {
            "1" {
                Write-Host "`nThis will transcode $($analysis.NeedTranscode.Count) file(s)." -ForegroundColor Yellow
                Write-Host "Original files will be replaced after successful conversion." -ForegroundColor Yellow
                $confirm = Read-Host "Are you sure you want to proceed? (Y/N) [N]"
                if ($confirm -match '^[Yy]') {
                    Invoke-Transcode -TranscodeQueue $analysis.NeedTranscode
                } else {
                    Write-Host "Transcode cancelled" -ForegroundColor Gray
                }
            }
            "2" {
                New-TranscodeScript -TranscodeQueue $analysis.NeedTranscode
            }
            default {
                Write-Host "Transcode skipped" -ForegroundColor Gray
            }
        }
    } else {
        Write-Host "All files use modern codecs" -ForegroundColor Green
    }

    # Step 13: Offer to transfer to main library (if this is an inbox and library path is configured)
    if ($script:Config.MoviesLibraryPath -and (Test-Path $script:Config.MoviesLibraryPath)) {
        # Check if we're processing the inbox
        $isInbox = $Path -eq $script:Config.MoviesInboxPath -or
                   $Path.TrimEnd('\', '/') -eq $script:Config.MoviesInboxPath.TrimEnd('\', '/')

        if ($isInbox) {
            # Count movies ready to transfer
            $movieCount = (Get-ChildItem -LiteralPath $Path -Directory -ErrorAction SilentlyContinue).Count

            if ($movieCount -gt 0) {
                Write-Host "`n--- Transfer to Library ---" -ForegroundColor Cyan
                Write-Host "Ready to move $movieCount movie(s) to: $($script:Config.MoviesLibraryPath)" -ForegroundColor White

                $transferChoice = Read-Host "`nTransfer movies to main library? (Y/N) [N]"

                if ($transferChoice -eq 'Y' -or $transferChoice -eq 'y') {
                    Move-MoviesToLibrary -InboxPath $Path -LibraryPath $script:Config.MoviesLibraryPath
                } else {
                    Write-Host "Transfer skipped" -ForegroundColor Gray
                }
            }
        }
    } elseif (-not $script:Config.MoviesLibraryPath) {
        # Prompt to configure library path if not set
        $isInbox = $script:Config.MoviesInboxPath -and
                   ($Path -eq $script:Config.MoviesInboxPath -or
                    $Path.TrimEnd('\', '/') -eq $script:Config.MoviesInboxPath.TrimEnd('\', '/'))

        if ($isInbox) {
            $movieCount = (Get-ChildItem -LiteralPath $Path -Directory -ErrorAction SilentlyContinue).Count
            if ($movieCount -gt 0) {
                Write-Host "`nTip: Set 'MoviesLibraryPath' in config to enable automatic transfer to your main collection." -ForegroundColor DarkGray
            }
        }
    }

    Write-Host "`nMovie routine completed!" -ForegroundColor Magenta
    Write-Log "Movie routine completed" "INFO"
}

<#
.SYNOPSIS
    Processes a TV show library
.PARAMETER Path
    The root path of the TV show library
#>
function Invoke-TVShowProcessing {
    param(
        [string]$Path
    )

    Write-Host "`nTV Show Routine" -ForegroundColor Magenta
    Write-Log "TV Show routine started for path: $Path" "INFO"

    if ($script:Config.DryRun) {
        Write-Host "`n*** DRY-RUN MODE - Previewing changes only ***`n" -ForegroundColor Yellow
    }

    # Step 1: Verify 7-Zip
    if (-not (Test-SevenZipInstallation)) {
        Write-Host "Cannot proceed without 7-Zip" -ForegroundColor Red
        return
    }

    # Step 2: Extract archives first (may contain episodes)
    Expand-Archives -Path $Path -DeleteAfterExtract

    # Step 3: Remove unnecessary files (samples, proofs, etc.)
    Remove-UnnecessaryFiles -Path $Path

    # Step 4: Process subtitles
    Invoke-SubtitleProcessing -Path $Path

    # Step 5: Organize into season folders (if enabled)
    if ($script:Config.OrganizeSeasons) {
        Invoke-SeasonOrganization -Path $Path
    }

    # Step 6: Rename episode files (if enabled)
    if ($script:Config.RenameEpisodes) {
        Rename-EpisodeFiles -Path $Path
    }

    # Step 7: Remove empty folders
    Remove-EmptyFolders -Path $Path

    # Step 8: Show episode summary
    Show-EpisodeSummary -Path $Path

    Write-Host "`nTV Show routine completed!" -ForegroundColor Magenta
    Write-Log "TV Show routine completed" "INFO"
}

#============================================
# MAIN SCRIPT EXECUTION
#============================================

# Initialize statistics
$script:Stats.StartTime = Get-Date

# Display header
Write-Host "`n=== LibraryLint v5.0 ===" -ForegroundColor Cyan
Write-Host "Automated media library organization tool`n" -ForegroundColor Gray

Write-Host "Log file: $($script:Config.LogFile)" -ForegroundColor Cyan
Write-Log "LibraryLint session started" "INFO"

# First-run setup check
$configExists = Test-Path $script:ConfigFilePath
if (-not $configExists) {
    Write-Host "`n" -NoNewline
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "  Welcome! First-time setup detected   " -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""

    # Check dependencies
    Write-Host "Checking dependencies..." -ForegroundColor Cyan

    $deps = @{
        "7-Zip" = Test-Path $script:Config.SevenZipPath
        "MediaInfo" = Test-MediaInfoInstallation
        "FFmpeg" = Test-FFmpegInstallation
        "yt-dlp" = Test-YtDlpInstallation
    }

    foreach ($dep in $deps.GetEnumerator()) {
        $status = if ($dep.Value) { "[OK]" } else { "[Missing]" }
        $color = if ($dep.Value) { "Green" } else { "Yellow" }
        Write-Host "  $($dep.Key): $status" -ForegroundColor $color
    }

    $missingDeps = $deps.GetEnumerator() | Where-Object { -not $_.Value }
    if ($missingDeps) {
        Write-Host ""
        Write-Host "Install missing dependencies with:" -ForegroundColor Gray
        if (-not $deps["MediaInfo"]) { Write-Host "  winget install MediaArea.MediaInfo" -ForegroundColor DarkGray }
        if (-not $deps["FFmpeg"]) { Write-Host "  winget install ffmpeg" -ForegroundColor DarkGray }
        if (-not $deps["yt-dlp"]) { Write-Host "  winget install yt-dlp" -ForegroundColor DarkGray }
    }

    Write-Host ""
    Write-Host "Let's configure your paths:" -ForegroundColor Cyan
    Write-Host ""

    # Movies inbox
    $defaultMoviesInbox = "G:\Inbox\_Movies"
    $moviesInbox = Read-Host "Movies inbox folder [$defaultMoviesInbox]"
    if (-not $moviesInbox) { $moviesInbox = $defaultMoviesInbox }
    $script:Config.MoviesInboxPath = $moviesInbox

    # Movies library
    $defaultMoviesLib = "G:\Movies"
    $moviesLib = Read-Host "Movies library folder [$defaultMoviesLib]"
    if (-not $moviesLib) { $moviesLib = $defaultMoviesLib }
    $script:Config.MoviesLibraryPath = $moviesLib

    # TV inbox
    $defaultTVInbox = "G:\Inbox\_Shows"
    $tvInbox = Read-Host "TV Shows inbox folder [$defaultTVInbox]"
    if (-not $tvInbox) { $tvInbox = $defaultTVInbox }
    $script:Config.TVShowsInboxPath = $tvInbox

    # TV library
    $defaultTVLib = "G:\Shows"
    $tvLib = Read-Host "TV Shows library folder [$defaultTVLib]"
    if (-not $tvLib) { $tvLib = $defaultTVLib }
    $script:Config.TVShowsLibraryPath = $tvLib

    Write-Host ""
    Write-Host "Optional: API Keys (press Enter to skip)" -ForegroundColor Cyan
    Write-Host "Get free keys from: themoviedb.org, fanart.tv" -ForegroundColor Gray
    Write-Host ""

    $tmdbKey = Read-Host "TMDB API Key (for metadata)"
    if ($tmdbKey) { $script:Config.TMDBApiKey = $tmdbKey }

    $fanartKey = Read-Host "Fanart.tv API Key (for artwork)"
    if ($fanartKey) { $script:Config.FanartTVApiKey = $fanartKey }

    # Save configuration
    Write-Host ""
    Export-Configuration
    Write-Host ""
    Write-Host "Setup complete! Configuration saved." -ForegroundColor Green
    Write-Host "You can edit settings anytime via Option 9 (Configuration)" -ForegroundColor Gray
    Write-Host ""
    Read-Host "Press Enter to continue to main menu"
}

# Media type selection
Write-Host "`n=== Main Menu ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "--- New Content ---" -ForegroundColor Yellow
Write-Host "1. Process New Movies"
Write-Host "2. Process New TV Shows"
Write-Host ""
Write-Host "--- Library Maintenance ---" -ForegroundColor Yellow
Write-Host "3. Health Check"
Write-Host "4. Fix Library Issues"
Write-Host "5. Codec Analysis"
Write-Host "6. Download Missing Subtitles"
Write-Host ""
Write-Host "--- Sync & Backup ---" -ForegroundColor Yellow
if ($script:SyncModuleLoaded) {
    Write-Host "S. SFTP Sync (download from seedbox)"
} else {
    Write-Host "S. SFTP Sync (module not loaded)" -ForegroundColor DarkGray
}
if ($script:MirrorModuleLoaded) {
    Write-Host "M. Mirror to Backup Drive"
} else {
    Write-Host "M. Mirror to Backup (module not loaded)" -ForegroundColor DarkGray
}
Write-Host ""
Write-Host "--- Utilities ---" -ForegroundColor Yellow
Write-Host "7. Export Library Report"
Write-Host "8. Undo Previous Session"
Write-Host "9. Configuration"

$type = Read-Host "`nSelect option"
Write-Log "User selected type: $type" "INFO"

# Process based on selection
switch ($type) {
    "1" {
        # Process New Movies
        Write-Host "`n=== Process New Movies ===" -ForegroundColor Cyan

        # Dry-run configuration
        $dryRunInput = Read-Host "Enable dry-run mode? (Y/N) [N]"
        if ($dryRunInput -match '^[Yy]') {
            $script:Config.DryRun = $true
            Write-Host "DRY-RUN MODE ENABLED - No changes will be made" -ForegroundColor Yellow
        } else {
            $script:Config.DryRun = $false
            Write-Host "Live mode - changes will be applied" -ForegroundColor Green
        }

        # Keep existing subtitles configuration
        $keepSubsInput = Read-Host "Keep existing subtitle files found in downloads? (Y/N) [Y]"
        if ($keepSubsInput -match '^[Nn]') {
            $script:Config.KeepSubtitles = $false
            Write-Host "Existing subtitles will be deleted" -ForegroundColor Yellow
        } else {
            $script:Config.KeepSubtitles = $true
            Write-Host "English subtitles will be kept" -ForegroundColor Green
        }

        # Keep existing trailers configuration
        $keepTrailersInput = Read-Host "Keep existing trailers (move to _Trailers folder)? (Y/N) [N]"
        if ($keepTrailersInput -match '^[Yy]') {
            $script:Config.KeepTrailers = $true
            Write-Host "Trailers will be moved to _Trailers folder" -ForegroundColor Green
        } else {
            $script:Config.KeepTrailers = $false
            Write-Host "Existing trailers will be deleted" -ForegroundColor Yellow
        }

        # Determine inbox path
        $path = $null
        if ($script:Config.MoviesInboxPath -and (Test-Path $script:Config.MoviesInboxPath)) {
            Write-Host "Movies inbox: $($script:Config.MoviesInboxPath)" -ForegroundColor Green
            $useConfigured = Read-Host "Use this folder? (Y/N) [Y]"
            if ($useConfigured -notmatch '^[Nn]') {
                $path = $script:Config.MoviesInboxPath
            }
        }

        if (-not $path) {
            $path = Select-FolderDialog -Description "Select your Movies inbox folder"
            if ($path) {
                # Offer to save as default
                $savePath = Read-Host "Save this as your default Movies inbox? (Y/N) [Y]"
                if ($savePath -notmatch '^[Nn]') {
                    $script:Config.MoviesInboxPath = $path
                    Export-Configuration
                    Write-Host "Movies inbox saved to configuration" -ForegroundColor Green
                }
            }
        }

        if ($path) {
            # Check TMDB API key for NFO generation
            if ($script:Config.GenerateNFO -or $true) {  # Always check since NFOs are important
                if (-not $script:Config.TMDBApiKey) {
                    Write-Host "`nTMDB API key not configured - NFO files will have limited metadata" -ForegroundColor Yellow
                    Write-Host "Get a free API key at: https://www.themoviedb.org/settings/api" -ForegroundColor Yellow
                    $apiKey = Read-Host "Enter your TMDB API key (or press Enter to skip)"
                    if ($apiKey) {
                        Write-Host "Validating API key..." -ForegroundColor Gray
                        if (Test-TMDBApiKey -ApiKey $apiKey) {
                            Write-Host "API key is valid" -ForegroundColor Green
                            $script:Config.TMDBApiKey = $apiKey

                            $saveKey = Read-Host "Save API key to config file? (Y/N) [Y]"
                            if ($saveKey -notmatch '^[Nn]') {
                                Export-Configuration
                            }
                        } else {
                            Write-Host "Invalid API key - continuing without TMDB" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "TMDB API key: configured" -ForegroundColor Green
                }

                # Check fanart.tv API key for additional artwork
                if (-not $script:Config.FanartTVApiKey) {
                    Write-Host "`nFanart.tv API key not configured - clearlogo/banner won't be downloaded" -ForegroundColor Yellow
                    Write-Host "Get a free API key at: https://fanart.tv/get-an-api-key/" -ForegroundColor Yellow
                    $fanartKey = Read-Host "Enter your fanart.tv API key (or press Enter to skip)"
                    if ($fanartKey) {
                        Write-Host "Validating API key..." -ForegroundColor Gray
                        if (Test-FanartTVApiKey -ApiKey $fanartKey) {
                            Write-Host "API key is valid" -ForegroundColor Green
                            $script:Config.FanartTVApiKey = $fanartKey

                            $saveKey = Read-Host "Save API key to config file? (Y/N) [Y]"
                            if ($saveKey -notmatch '^[Nn]') {
                                Export-Configuration
                            }
                        } else {
                            Write-Host "Invalid API key - continuing without fanart.tv" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "Fanart.tv API key: configured" -ForegroundColor Green
                }
            }

            $dupeInput = Read-Host "`nCheck for duplicate movies? (Y/N) [N]"
            if ($dupeInput -eq 'Y' -or $dupeInput -eq 'y') {
                $script:Config.CheckDuplicates = $true
                Write-Host "Duplicate detection enabled" -ForegroundColor Green
                Write-Log "Duplicate detection enabled" "INFO"
            } else {
                Write-Host "Duplicate detection disabled" -ForegroundColor Cyan
                Write-Log "Duplicate detection disabled" "INFO"
            }

            # NFO generation configuration
            $nfoDefault = if ($script:Config.GenerateNFO) { "Y" } else { "N" }
            $nfoInput = Read-Host "Generate Kodi NFO files for movies without them? (Y/N) [$nfoDefault]"
            if ($nfoInput -eq '') {
                if ($script:Config.GenerateNFO) {
                    Write-Host "NFO files will be generated (saved preference)" -ForegroundColor Green
                } else {
                    Write-Host "NFO generation skipped (saved preference)" -ForegroundColor Cyan
                }
            } elseif ($nfoInput -match '^[Yy]') {
                $script:Config.GenerateNFO = $true
                Write-Host "NFO files will be generated" -ForegroundColor Green
            } else {
                $script:Config.GenerateNFO = $false
                Write-Host "NFO generation skipped" -ForegroundColor Cyan
            }

            # Trailer download configuration (only if yt-dlp is available)
            if (Test-YtDlpInstallation) {
                $trailerDownloadDefault = if ($script:Config.DownloadTrailers) { "Y" } else { "N" }
                $trailerDownloadInput = Read-Host "Download movie trailers from YouTube? (Y/N) [$trailerDownloadDefault]"
                if ($trailerDownloadInput -eq '') {
                    if ($script:Config.DownloadTrailers) {
                        Write-Host "Trailers will be downloaded ($($script:Config.TrailerQuality)) (saved preference)" -ForegroundColor Green
                    } else {
                        Write-Host "Trailer downloads skipped (saved preference)" -ForegroundColor Cyan
                    }
                } elseif ($trailerDownloadInput -match '^[Yy]') {
                    $script:Config.DownloadTrailers = $true
                    Write-Host "Trailers will be downloaded ($($script:Config.TrailerQuality))" -ForegroundColor Green

                    # Prompt for browser if not already configured
                    if (-not $script:Config.YtDlpCookieBrowser -or $script:Config.YtDlpCookieBrowser -eq "chrome") {
                        Write-Host "`nYouTube may require authentication for age-restricted trailers." -ForegroundColor Yellow
                        Write-Host "Which browser should yt-dlp use for cookies?" -ForegroundColor Yellow
                        Write-Host "  1. Chrome"
                        Write-Host "  2. Firefox"
                        Write-Host "  3. Edge"
                        Write-Host "  4. Brave"
                        Write-Host "  5. None (skip age-restricted videos)"
                        $browserChoice = Read-Host "Select browser [1]"
                        $script:Config.YtDlpCookieBrowser = switch ($browserChoice) {
                            "2" { "firefox" }
                            "3" { "edge" }
                            "4" { "brave" }
                            "5" { $null }
                            default { "chrome" }
                        }
                        if ($script:Config.YtDlpCookieBrowser) {
                            Write-Host "Using $($script:Config.YtDlpCookieBrowser) cookies for YouTube authentication" -ForegroundColor Green
                        } else {
                            Write-Host "Age-restricted trailers will be skipped" -ForegroundColor Yellow
                        }
                    }
                } else {
                    $script:Config.DownloadTrailers = $false
                    Write-Host "Trailer downloads skipped" -ForegroundColor Cyan
                }
            }

            # Subtitle download configuration
            $subtitleDownloadDefault = if ($script:Config.DownloadSubtitles) { "Y" } else { "N" }
            $subtitleDownloadInput = Read-Host "Download subtitles from Subdl.com? (Y/N) [$subtitleDownloadDefault]"
            if ($subtitleDownloadInput -eq '') {
                if ($script:Config.DownloadSubtitles) {
                    Write-Host "Subtitles will be downloaded ($($script:Config.SubtitleLanguage)) (saved preference)" -ForegroundColor Green
                } else {
                    Write-Host "Subtitle downloads skipped (saved preference)" -ForegroundColor Cyan
                }
            } elseif ($subtitleDownloadInput -match '^[Yy]') {
                $script:Config.DownloadSubtitles = $true
                Write-Host "Subtitles will be downloaded ($($script:Config.SubtitleLanguage))" -ForegroundColor Green
            } else {
                $script:Config.DownloadSubtitles = $false
                Write-Host "Subtitle downloads skipped" -ForegroundColor Cyan
            }

            # Check for Subdl API key if subtitle download is enabled
            if ($script:Config.DownloadSubtitles -and -not $script:Config.SubdlApiKey) {
                Write-Host "`nSubdl.com API key not configured" -ForegroundColor Yellow
                Write-Host "Get a free API key at: https://subdl.com/panel/api" -ForegroundColor Yellow
                $apiKey = Read-Host "Enter your Subdl API key (or press Enter to skip subtitles)"
                if ($apiKey) {
                    $script:Config.SubdlApiKey = $apiKey
                    $saveKey = Read-Host "Save API key to config file? (Y/N) [Y]"
                    if ($saveKey -notmatch '^[Nn]') {
                        Export-Configuration
                        Write-Host "API key saved to configuration" -ForegroundColor Green
                    }
                } else {
                    $script:Config.DownloadSubtitles = $false
                    Write-Host "Subtitle downloads disabled - no API key" -ForegroundColor Yellow
                }
            }
        }

        if ($path) {
            # Track path for dry run report
            $script:ProcessedPath = $path
            $script:ProcessedMediaType = "Movies"

            Invoke-MovieProcessing -Path $path

            # Retry failed operations
            Invoke-RetryFailedOperations

            # Save undo manifest
            Save-UndoManifest

            # Generate dry run report if in dry run mode
            if ($script:Config.DryRun) {
                Export-DryRunReport -Path $path -MediaType "Movies"
            } else {
                # Offer to export report (only in live mode)
                $exportReport = Read-Host "`nExport library report? (Y/N) [N]"
                if ($exportReport -eq 'Y' -or $exportReport -eq 'y') {
                    Invoke-ExportMenu -Path $path -MediaType "Movies"
                }
            }

            # Offer to save config
            Invoke-ConfigurationSavePrompt
        }
    }
    "2" {
        # Process New TV Shows
        Write-Host "`n=== Process New TV Shows ===" -ForegroundColor Cyan

        # Dry-run configuration
        $dryRunInput = Read-Host "Enable dry-run mode? (Y/N) [N]"
        if ($dryRunInput -match '^[Yy]') {
            $script:Config.DryRun = $true
            Write-Host "DRY-RUN MODE ENABLED - No changes will be made" -ForegroundColor Yellow
        } else {
            $script:Config.DryRun = $false
            Write-Host "Live mode - changes will be applied" -ForegroundColor Green
        }

        # Keep existing subtitles configuration
        $keepSubsInput = Read-Host "Keep existing subtitle files found in downloads? (Y/N) [Y]"
        if ($keepSubsInput -match '^[Nn]') {
            $script:Config.KeepSubtitles = $false
            Write-Host "Existing subtitles will be deleted" -ForegroundColor Yellow
        } else {
            $script:Config.KeepSubtitles = $true
            Write-Host "English subtitles will be kept" -ForegroundColor Green
        }

        # Determine inbox path
        $path = $null
        if ($script:Config.TVShowsInboxPath -and (Test-Path $script:Config.TVShowsInboxPath)) {
            Write-Host "TV Shows inbox: $($script:Config.TVShowsInboxPath)" -ForegroundColor Green
            $useConfigured = Read-Host "Use this folder? (Y/N) [Y]"
            if ($useConfigured -notmatch '^[Nn]') {
                $path = $script:Config.TVShowsInboxPath
            }
        }

        if (-not $path) {
            $path = Select-FolderDialog -Description "Select your TV Shows inbox folder"
            if ($path) {
                # Offer to save as default
                $savePath = Read-Host "Save this as your default TV Shows inbox? (Y/N) [Y]"
                if ($savePath -notmatch '^[Nn]') {
                    $script:Config.TVShowsInboxPath = $path
                    Export-Configuration
                    Write-Host "TV Shows inbox saved to configuration" -ForegroundColor Green
                }
            }
        }

        if ($path) {
            $organizeInput = Read-Host "Organize episodes into Season folders? (Y/N) [Y]"
            if ($organizeInput -eq 'N' -or $organizeInput -eq 'n') {
                $script:Config.OrganizeSeasons = $false
                Write-Host "Season organization disabled" -ForegroundColor Cyan
                Write-Log "Season organization disabled" "INFO"
            } else {
                Write-Host "Episodes will be organized into Season folders" -ForegroundColor Green
                Write-Log "Season organization enabled" "INFO"
            }

            $renameInput = Read-Host "Rename episodes to standard format (ShowName - S01E01)? (Y/N) [N]"
            if ($renameInput -eq 'Y' -or $renameInput -eq 'y') {
                $script:Config.RenameEpisodes = $true
                Write-Host "Episodes will be renamed to standard format" -ForegroundColor Green
                Write-Log "Episode renaming enabled" "INFO"
            } else {
                Write-Host "Episode filenames will be preserved" -ForegroundColor Cyan
                Write-Log "Episode renaming disabled" "INFO"
            }
        }

        if ($path) {
            # Track path for dry run report
            $script:ProcessedPath = $path
            $script:ProcessedMediaType = "TVShows"

            Invoke-TVShowProcessing -Path $path

            # Retry failed operations
            Invoke-RetryFailedOperations

            # Save undo manifest
            Save-UndoManifest

            # Generate dry run report if in dry run mode
            if ($script:Config.DryRun) {
                Export-DryRunReport -Path $path -MediaType "TVShows"
            }

            # Offer to save config
            Invoke-ConfigurationSavePrompt
        }
    }
    "3" {
        # Health Check mode
        Write-Host "`n=== Health Check ===" -ForegroundColor Cyan
        $mediaTypeInput = Read-Host "Media type? (1=Movies, 2=TV Shows) [1]"
        $mediaType = if ($mediaTypeInput -eq '2') { "TVShows" } else { "Movies" }

        $path = Select-FolderDialog -Description "Select your media library folder"
        if ($path) {
            Invoke-LibraryHealthCheck -Path $path -MediaType $mediaType
        }
    }
    "4" {
        # Fix Library Issues - Submenu
        Write-Host "`n=== Fix Library Issues ===" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "1. Fix Missing Folder Years"
        Write-Host "2. Enhanced Duplicate Detection"
        Write-Host "3. Refresh/Repair Metadata"
        Write-Host "4. Download Missing Trailers"
        Write-Host "5. Fix Orphaned Subtitles"
        Write-Host "6. Remove Empty Folders"
        Write-Host "7. Run All Fixes"
        Write-Host "8. Back to Main Menu"

        $fixChoice = Read-Host "`nSelect option"

        # Get path first (used by all options except back)
        $path = $null
        if ($fixChoice -ne '8') {
            $mediaTypeInput = Read-Host "`nMedia type? (1=Movies, 2=TV Shows) [1]"
            $mediaType = if ($mediaTypeInput -eq '2') { "TVShows" } else { "Movies" }

            $path = Select-FolderDialog -Description "Select your media library folder"
        }

        if ($path) {
            # Check/prompt for TMDB API key (needed for most fixes)
            if ($fixChoice -in @('1', '3', '4', '7')) {
                if (-not $script:Config.TMDBApiKey) {
                    Write-Host "`nTMDB API key not configured" -ForegroundColor Yellow
                    Write-Host "Get a free API key at: https://www.themoviedb.org/settings/api" -ForegroundColor Yellow
                    $apiKey = Read-Host "Enter your TMDB API key (or press Enter to skip)"
                    if ($apiKey) {
                        Write-Host "Validating API key..." -ForegroundColor Gray
                        if (Test-TMDBApiKey -ApiKey $apiKey) {
                            Write-Host "API key is valid" -ForegroundColor Green
                            $script:Config.TMDBApiKey = $apiKey
                            $saveKey = Read-Host "Save API key to config file? (Y/N) [Y]"
                            if ($saveKey -notmatch '^[Nn]') {
                                Export-Configuration
                            }
                        } else {
                            Write-Host "Invalid API key - some fixes may be limited" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "TMDB API key: configured" -ForegroundColor Green
                }

                # Check fanart.tv key for metadata repair
                if ($fixChoice -in @('3', '7') -and -not $script:Config.FanartTVApiKey) {
                    Write-Host "Fanart.tv API key: not configured (clearlogo/banner won't be downloaded)" -ForegroundColor Yellow
                }

                # Check yt-dlp for trailer downloads
                if ($fixChoice -in @('4', '7')) {
                    if (-not (Test-YtDlpInstallation)) {
                        Write-Host "yt-dlp: not found (required for trailer downloads)" -ForegroundColor Yellow
                        Write-Host "Install with: winget install yt-dlp" -ForegroundColor Yellow
                    } else {
                        Write-Host "yt-dlp: available" -ForegroundColor Green
                    }
                }
            }

            switch ($fixChoice) {
                "1" {
                    # Fix Missing Folder Years
                    Write-Host "`n--- Fix Missing Folder Years ---" -ForegroundColor Yellow
                    $dryRunInput = Read-Host "Run in dry-run mode first? (Y/N) [Y]"
                    $dryRun = $dryRunInput -notmatch '^[Nn]'

                    if ($dryRun) {
                        Repair-MovieFolderYears -Path $path -WhatIf
                    } else {
                        Repair-MovieFolderYears -Path $path
                    }
                }
                "2" {
                    # Enhanced Duplicate Detection
                    Invoke-EnhancedDuplicateDetection -Path $path
                }
                "3" {
                    # Refresh/Repair Metadata
                    Write-Host "`n--- Refresh/Repair Metadata ---" -ForegroundColor Yellow
                    Write-Host "This will re-fetch NFO files and artwork for movies missing them" -ForegroundColor Gray

                    if (-not $script:Config.TMDBApiKey) {
                        Write-Host "TMDB API key required for metadata refresh" -ForegroundColor Red
                    } else {
                        $overwriteInput = Read-Host "Overwrite existing NFO files? (Y/N) [N]"
                        $overwrite = $overwriteInput -match '^[Yy]'

                        Invoke-MetadataRefresh -Path $path -MediaType $mediaType -Overwrite $overwrite
                    }
                }
                "4" {
                    # Download Missing Trailers
                    Write-Host "`nNote: For best results, run 'Refresh/Repair Metadata' first to ensure" -ForegroundColor Yellow
                    Write-Host "NFO files exist with trailer URLs. This reduces TMDB API calls." -ForegroundColor Yellow

                    $continueInput = Read-Host "`nContinue? (Y/N) [Y]"
                    if ($continueInput -match '^[Nn]') {
                        Write-Host "Cancelled." -ForegroundColor Gray
                    } else {
                        $qualityInput = Read-Host "Trailer quality? (1=1080p, 2=720p, 3=480p) [2]"
                        $quality = switch ($qualityInput) {
                            "1" { "1080p" }
                            "3" { "480p" }
                            default { "720p" }
                        }
                        Invoke-TrailerDownload -Path $path -Quality $quality
                    }
                }
                "5" {
                    # Fix Orphaned Subtitles
                    $dryRunInput = Read-Host "Run in dry-run mode first? (Y/N) [Y]"
                    $dryRun = $dryRunInput -notmatch '^[Nn]'

                    if ($dryRun) {
                        Repair-OrphanedSubtitles -Path $path -WhatIf
                    } else {
                        Repair-OrphanedSubtitles -Path $path
                    }
                }
                "6" {
                    # Remove Empty Folders
                    $dryRunInput = Read-Host "Run in dry-run mode first? (Y/N) [Y]"
                    $dryRun = $dryRunInput -notmatch '^[Nn]'

                    if ($dryRun) {
                        Remove-EmptyFolders -Path $path -WhatIf
                    } else {
                        Remove-EmptyFolders -Path $path
                    }
                }
                "7" {
                    # Run All Fixes
                    Write-Host "`n--- Running All Fixes ---" -ForegroundColor Yellow

                    Write-Host "`n[1/6] Fixing missing folder years..." -ForegroundColor Cyan
                    Repair-MovieFolderYears -Path $path

                    Write-Host "`n[2/6] Running duplicate detection..." -ForegroundColor Cyan
                    Invoke-EnhancedDuplicateDetection -Path $path

                    if ($script:Config.TMDBApiKey) {
                        Write-Host "`n[3/6] Refreshing metadata..." -ForegroundColor Cyan
                        Invoke-MetadataRefresh -Path $path -MediaType $mediaType -Overwrite $false

                        if (Test-YtDlpInstallation) {
                            Write-Host "`n[4/6] Downloading missing trailers..." -ForegroundColor Cyan
                            Invoke-TrailerDownload -Path $path
                        } else {
                            Write-Host "`n[4/6] Skipping trailer downloads (yt-dlp not installed)" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "`n[3/6] Skipping metadata refresh (no TMDB API key)" -ForegroundColor Yellow
                        Write-Host "`n[4/6] Skipping trailer downloads (no TMDB API key)" -ForegroundColor Yellow
                    }

                    Write-Host "`n[5/6] Fixing orphaned subtitles..." -ForegroundColor Cyan
                    Repair-OrphanedSubtitles -Path $path

                    Write-Host "`n[6/6] Removing empty folders..." -ForegroundColor Cyan
                    Remove-EmptyFolders -Path $path

                    Write-Host "`nAll fixes completed!" -ForegroundColor Green
                }
                "8" {
                    Write-Host "Returning to main menu..." -ForegroundColor Gray
                }
            }
        }
    }
    "5" {
        # Codec Analysis mode
        Write-Host "`n=== Codec Analysis ===" -ForegroundColor Cyan
        Write-Host "MediaInfo integration: $(if (Test-MediaInfoInstallation) { 'Available' } else { 'Not found (using filename parsing)' })" -ForegroundColor $(if (Test-MediaInfoInstallation) { 'Green' } else { 'Yellow' })
        Write-Host "FFmpeg: $(if (Test-FFmpegInstallation) { 'Available' } else { 'Not found (required for transcoding)' })" -ForegroundColor $(if (Test-FFmpegInstallation) { 'Green' } else { 'Yellow' })

        $path = Select-FolderDialog -Description "Select your media library folder"
        if ($path) {
            $analysis = Invoke-CodecAnalysis -Path $path

            if ($analysis -and $analysis.NeedTranscode.Count -gt 0) {
                Write-Host "`n--- Transcode Queue ---" -ForegroundColor Yellow
                Write-Host "Found $($analysis.NeedTranscode.Count) file(s) needing conversion:" -ForegroundColor White

                $i = 1
                foreach ($file in $analysis.NeedTranscode | Select-Object -First 10) {
                    $modeLabel = if ($file.TranscodeMode -eq "remux") { "[remux]" } else { "[transcode]" }
                    Write-Host "  $i. $modeLabel $($file.FileName)" -ForegroundColor Gray
                    $i++
                }
                if ($analysis.NeedTranscode.Count -gt 10) {
                    Write-Host "  ... and $($analysis.NeedTranscode.Count - 10) more" -ForegroundColor Gray
                }

                Write-Host "`nOptions:" -ForegroundColor Cyan
                Write-Host "  1. Run now (transcoding may take a long time)"
                Write-Host "  2. Save script for later"
                Write-Host "  3. Skip"

                $transcodeChoice = Read-Host "`nSelect option [3]"

                switch ($transcodeChoice) {
                    "1" {
                        Write-Host "`nThis will transcode $($analysis.NeedTranscode.Count) file(s)." -ForegroundColor Yellow
                        Write-Host "Original files will be replaced after successful conversion." -ForegroundColor Yellow
                        $confirm = Read-Host "Are you sure you want to proceed? (Y/N) [N]"
                        if ($confirm -match '^[Yy]') {
                            Invoke-Transcode -TranscodeQueue $analysis.NeedTranscode
                        } else {
                            Write-Host "Transcode cancelled" -ForegroundColor Gray
                        }
                    }
                    "2" {
                        New-TranscodeScript -TranscodeQueue $analysis.NeedTranscode
                    }
                    default {
                        Write-Host "Transcode skipped" -ForegroundColor Gray
                    }
                }
            }
        }
    }
    "6" {
        # Download Missing Subtitles
        Write-Host "`n=== Download Missing Subtitles ===" -ForegroundColor Cyan
        Write-Host "This will scan your movie library and download subtitles from Subdl.com" -ForegroundColor White
        Write-Host "for movies that don't already have them." -ForegroundColor White
        Write-Host "`nRequirements:" -ForegroundColor Yellow
        Write-Host "  - Subdl.com API key (free at https://subdl.com/panel/api)" -ForegroundColor Gray
        Write-Host "  - Movies must have NFO files (with IMDB ID for best results)" -ForegroundColor Gray
        Write-Host "`nNote: Subdl.com has a daily limit of ~1000 subtitles. For large libraries," -ForegroundColor Yellow
        Write-Host "you may need to run this over multiple days. Already downloaded subtitles" -ForegroundColor Yellow
        Write-Host "will be skipped on subsequent runs." -ForegroundColor Yellow

        # Check/prompt for Subdl API key
        if (-not $script:Config.SubdlApiKey) {
            Write-Host "`nSubdl.com API key not configured" -ForegroundColor Yellow
            Write-Host "Get a free API key at: https://subdl.com/panel/api" -ForegroundColor Yellow
            $apiKey = Read-Host "Enter your Subdl API key (or press Enter to cancel)"
            if ($apiKey) {
                $script:Config.SubdlApiKey = $apiKey
                $saveKey = Read-Host "Save API key to config file? (Y/N) [Y]"
                if ($saveKey -notmatch '^[Nn]') {
                    Export-Configuration
                    Write-Host "API key saved to configuration" -ForegroundColor Green
                }
            } else {
                Write-Host "Subtitle download cancelled - API key required" -ForegroundColor Yellow
                break
            }
        } else {
            Write-Host "`nSubdl.com API key: configured" -ForegroundColor Green
        }

        $path = Select-FolderDialog -Description "Select your movie library folder"
        if ($path) {
            # Language selection
            Write-Host "`nSubtitle language codes: en (English), es (Spanish), fr (French), de (German), etc." -ForegroundColor Gray
            $langInput = Read-Host "Enter language code [$($script:Config.SubtitleLanguage)]"
            $language = if ($langInput) { $langInput } else { $script:Config.SubtitleLanguage }

            # Confirm
            $confirm = Read-Host "`nStart downloading subtitles? (Y/N) [Y]"
            if ($confirm -notmatch '^[Nn]') {
                Invoke-SubtitleDownload -Path $path -Language $language
            } else {
                Write-Host "Subtitle download cancelled" -ForegroundColor Yellow
            }
        }
    }
    "7" {
        # Export Library Report
        Write-Host "`n=== Export Library Report ===" -ForegroundColor Cyan
        $mediaTypeInput = Read-Host "Media type? (1=Movies, 2=TV Shows) [1]"
        $mediaType = if ($mediaTypeInput -eq '2') { "TVShows" } else { "Movies" }

        $path = Select-FolderDialog -Description "Select your media library folder"
        if ($path) {
            Invoke-ExportMenu -Path $path -MediaType $mediaType
        }
    }
    "8" {
        # Undo Previous Session
        Write-Host "`n=== Undo Previous Session ===" -ForegroundColor Cyan

        # Find available undo manifests
        $manifests = Get-ChildItem -Path $script:UndoFolder -Filter "LibraryLint_Undo_*.json" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending

        if ($manifests.Count -eq 0) {
            Write-Host "No undo manifests found" -ForegroundColor Yellow
        } else {
            Write-Host "Available undo manifests:" -ForegroundColor Cyan
            $i = 1
            foreach ($manifest in $manifests | Select-Object -First 10) {
                Write-Host "  $i. $($manifest.Name) ($(Get-Date $manifest.LastWriteTime -Format 'yyyy-MM-dd HH:mm'))" -ForegroundColor White
                $i++
            }

            $selection = Read-Host "`nSelect manifest number to undo (or Enter to cancel)"
            if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $manifests.Count) {
                $selectedManifest = $manifests[[int]$selection - 1]
                Invoke-UndoOperations -ManifestPath $selectedManifest.FullName
            } else {
                Write-Host "Undo cancelled" -ForegroundColor Yellow
            }
        }
    }
    "9" {
        # Configuration Management
        Write-Host "`n=== Configuration ===" -ForegroundColor Cyan
        Write-Host "1. Save current configuration"
        Write-Host "2. Load configuration from file"
        Write-Host "3. Reset to defaults"
        Write-Host "4. View current configuration"

        $configChoice = Read-Host "`nSelect option"
        switch ($configChoice) {
            "1" { Export-Configuration }
            "2" {
                $customPath = Read-Host "Enter config file path (or Enter for default)"
                if ($customPath) {
                    Import-Configuration -Path $customPath
                } else {
                    Import-Configuration
                }
            }
            "3" {
                $script:Config = $script:DefaultConfig.Clone()
                $script:Config.LogFile = Join-Path $script:LogsFolder "LibraryLint_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
                Write-Host "Configuration reset to defaults" -ForegroundColor Green
            }
            "4" {
                Write-Host "`nCurrent Configuration:" -ForegroundColor Cyan
                $script:Config.GetEnumerator() | Where-Object { $_.Key -ne 'Tags' } | ForEach-Object {
                    $value = if ($_.Value -is [array]) { $_.Value -join ', ' } else { $_.Value }
                    Write-Host "  $($_.Key): $value" -ForegroundColor White
                }
            }
        }
    }
    { $_ -eq "S" -or $_ -eq "s" } {
        # SFTP Sync
        if (-not $script:SyncModuleLoaded) {
            Write-Host "`nSFTP Sync module not loaded!" -ForegroundColor Red
            Write-Host "Ensure modules/Sync.psm1 exists in the LibraryLint folder." -ForegroundColor Yellow
        } else {
            Write-Host "`n=== SFTP Sync ===" -ForegroundColor Cyan

            # Check if SFTP is configured
            if (-not $script:Config.SFTPHost) {
                Write-Host "SFTP not configured!" -ForegroundColor Red
                Write-Host "Please set the following in your config:" -ForegroundColor Yellow
                Write-Host "  SFTPHost, SFTPUsername, SFTPPassword (or SFTPPrivateKeyPath)" -ForegroundColor Gray
                Write-Host "  SFTPRemotePaths, MoviesInboxPath" -ForegroundColor Gray
            } else {
                $whatIfInput = Read-Host "Enable dry-run mode? (Y/N) [N]"
                $whatIf = $whatIfInput -match '^[Yy]'

                $listOnlyInput = Read-Host "List files only (no download)? (Y/N) [N]"
                $listOnly = $listOnlyInput -match '^[Yy]'

                # Determine local base path
                $localBase = if ($script:Config.MoviesInboxPath) {
                    Split-Path $script:Config.MoviesInboxPath -Parent
                } else {
                    "G:"
                }

                Write-Log "Starting SFTP Sync to $localBase" "INFO"

                $syncParams = @{
                    HostName = $script:Config.SFTPHost
                    Port = $script:Config.SFTPPort
                    Username = $script:Config.SFTPUsername
                    RemotePaths = $script:Config.SFTPRemotePaths
                    LocalBasePath = $localBase
                    WhatIf = $whatIf
                    ListOnly = $listOnly
                }

                if ($script:Config.SFTPPassword) {
                    $syncParams.Password = $script:Config.SFTPPassword
                }
                if ($script:Config.SFTPPrivateKeyPath) {
                    $syncParams.PrivateKeyPath = $script:Config.SFTPPrivateKeyPath
                }
                if ($script:Config.SFTPDeleteAfterDownload) {
                    $syncParams.DeleteAfterDownload = $true
                }

                $result = Invoke-SFTPSync @syncParams

                if ($result.Downloaded -gt 0 -and -not $whatIf -and -not $listOnly) {
                    Write-Host ""
                    $processNow = Read-Host "Process downloaded files now? (Y/N) [Y]"
                    if ($processNow -notmatch '^[Nn]') {
                        if ($script:Config.MoviesInboxPath -and (Test-Path $script:Config.MoviesInboxPath)) {
                            Invoke-MovieProcessing -Path $script:Config.MoviesInboxPath
                        }
                    }
                }

                Write-Log "SFTP Sync completed: $($result.Downloaded) downloaded, $($result.Failed) failed" "INFO"
            }
        }
    }
    { $_ -eq "M" -or $_ -eq "m" } {
        # Mirror to Backup
        if (-not $script:MirrorModuleLoaded) {
            Write-Host "`nMirror module not loaded!" -ForegroundColor Red
            Write-Host "Ensure modules/Mirror.psm1 exists in the LibraryLint folder." -ForegroundColor Yellow
        } else {
            Write-Host "`n=== Mirror to Backup ===" -ForegroundColor Cyan

            Write-Host "Source: $($script:Config.MirrorSourceDrive)" -ForegroundColor Gray
            Write-Host "Dest:   $($script:Config.MirrorDestDrive)" -ForegroundColor Gray
            Write-Host "Folders: $($script:Config.MirrorFolders -join ', ')" -ForegroundColor Gray
            Write-Host ""

            # Check if dest drive exists
            if (-not (Test-Path $script:Config.MirrorDestDrive)) {
                Write-Host "Destination drive not available: $($script:Config.MirrorDestDrive)" -ForegroundColor Red
                Write-Host "Please connect your backup drive and try again." -ForegroundColor Yellow
            } else {
                $whatIfInput = Read-Host "Enable dry-run mode? (Y/N) [N]"
                $whatIf = $whatIfInput -match '^[Yy]'

                Write-Log "Starting Mirror: $($script:Config.MirrorSourceDrive) -> $($script:Config.MirrorDestDrive)" "INFO"

                $result = Invoke-Mirror -SourceDrive $script:Config.MirrorSourceDrive `
                    -DestDrive $script:Config.MirrorDestDrive `
                    -Folders $script:Config.MirrorFolders `
                    -WhatIf:$whatIf

                Write-Log "Mirror completed: $($result.FilesCopied) copied, $($result.FilesDeleted) deleted" "INFO"
            }
        }
    }
    default {
        Write-Host "Invalid selection" -ForegroundColor Red
        Write-Log "Invalid type selection: $type" "ERROR"
    }
}

# Show statistics summary
Show-Statistics

Write-Host "`nLog file saved to: $($script:Config.LogFile)" -ForegroundColor Cyan
Write-Log "LibraryLint session ended" "INFO"
