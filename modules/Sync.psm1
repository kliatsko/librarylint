# Sync.psm1
# Downloads new files from SFTP server, tracking what's already been downloaded
# Requires WinSCP .NET assembly (install via: winget install WinSCP)
# Part of LibraryLint suite

#region Private Functions

function Format-SyncSize {
    param([long]$Bytes)

    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

function Get-SyncTrackingPath {
    param([string]$ConfigTrackingFile)

    if ($ConfigTrackingFile) {
        return $ConfigTrackingFile
    }
    return "$env:LOCALAPPDATA\LibraryLint\sftp_downloaded.json"
}

function Get-DownloadedFiles {
    param([string]$TrackingPath)

    $trackingDir = Split-Path $TrackingPath -Parent

    if (-not (Test-Path $trackingDir)) {
        New-Item -Path $trackingDir -ItemType Directory -Force | Out-Null
    }

    if (Test-Path $TrackingPath) {
        return Get-Content $TrackingPath -Raw | ConvertFrom-Json -AsHashtable
    }

    return @{}
}

function Save-DownloadedFiles {
    param(
        [hashtable]$Downloaded,
        [string]$TrackingPath
    )
    $Downloaded | ConvertTo-Json -Depth 3 | Set-Content $TrackingPath -Encoding UTF8
}

function Test-WinSCPCompatible {
    param([string]$DllPath)
    # Check if a WinSCPnet.dll is netstandard2.0 (compatible with PowerShell 7/.NET 5+)
    # .NET Framework builds reference mscorlib; netstandard builds reference netstandard
    try {
        $bytes = [System.IO.File]::ReadAllBytes($DllPath)
        $asm = [System.Reflection.Assembly]::Load($bytes)
        $refs = $asm.GetReferencedAssemblies()
        # If it references netstandard, it's compatible
        if ($refs | Where-Object { $_.Name -eq 'netstandard' }) { return $true }
        # If it references mscorlib but not netstandard, it's .NET Framework only
        if ($refs | Where-Object { $_.Name -eq 'mscorlib' }) { return $false }
        return $true  # Assume compatible if we can't tell
    } catch {
        return $true  # If we can't check, let it try and fail naturally
    }
}

function Install-WinSCPNetStandard {
    # Auto-download the netstandard2.0 WinSCPnet.dll from NuGet
    $targetDir = "$env:LOCALAPPDATA\LibraryLint"
    $targetDll = Join-Path $targetDir "WinSCPnet.dll"

    if (Test-Path $targetDll) {
        if (Test-WinSCPCompatible $targetDll) { return $targetDll }
        Remove-Item $targetDll -Force  # Remove incompatible version
    }

    Write-Host "  Downloading WinSCP .NET assembly (netstandard2.0)..." -ForegroundColor Cyan
    $tempFile = Join-Path $env:TEMP "winscp_nuget.zip"
    try {
        if (-not (Test-Path $targetDir)) { New-Item -Path $targetDir -ItemType Directory -Force | Out-Null }
        $nugetUrl = "https://www.nuget.org/api/v2/package/WinSCP"
        Invoke-WebRequest -Uri $nugetUrl -OutFile $tempFile -UseBasicParsing
        $tempExtract = Join-Path $env:TEMP "winscp_nuget_extract"
        if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
        Expand-Archive -Path $tempFile -DestinationPath $tempExtract -Force
        $netStdDll = Get-ChildItem -Path $tempExtract -Recurse -Filter "WinSCPnet.dll" |
            Where-Object { $_.FullName -match 'netstandard' } | Select-Object -First 1
        if ($netStdDll) {
            Copy-Item $netStdDll.FullName -Destination $targetDll -Force
            Write-Host "  Installed to: $targetDll" -ForegroundColor Green
            return $targetDll
        } else {
            Write-Host "  Could not find netstandard2.0 build in NuGet package" -ForegroundColor Red
            return $null
        }
    } catch {
        Write-Host "  Download failed: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    } finally {
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Test-WinSCPInstalled {
    param([string]$ModulePath)

    # Order matters: prioritize netstandard2.0 versions for .NET Core/5+ compatibility
    $winscpPaths = @(
        # LibraryLint folder (preferred - auto-downloaded netstandard2.0 version)
        "$env:LOCALAPPDATA\LibraryLint\WinSCPnet.dll",
        # NuGet netstandard2.0 (works with .NET Core/5+/PowerShell 7)
        "$env:USERPROFILE\.nuget\packages\winscp\*\lib\netstandard2.0\WinSCPnet.dll",
        # Downloads - netstandard2.0 subfolder
        "$env:USERPROFILE\Downloads\WinSCP-*-Automation\netstandard2.0\WinSCPnet.dll",
        "$env:USERPROFILE\Downloads\winscp_nuget\lib\netstandard2.0\WinSCPnet.dll",
        # Standard install locations (may be .NET Framework - checked for compatibility)
        "${env:ProgramFiles}\WinSCP\WinSCPnet.dll",
        "${env:ProgramFiles(x86)}\WinSCP\WinSCPnet.dll",
        # User profile locations
        "$env:LOCALAPPDATA\Programs\WinSCP\WinSCPnet.dll",
        "$env:USERPROFILE\WinSCP\WinSCPnet.dll"
    )

    if ($ModulePath) {
        $winscpPaths = @(Join-Path $ModulePath "WinSCPnet.dll") + $winscpPaths
    }

    # Expand wildcards and check paths - verify compatibility
    foreach ($pattern in $winscpPaths) {
        $resolved = Get-Item $pattern -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($resolved) {
            if (Test-WinSCPCompatible $resolved.FullName) {
                return $resolved.FullName
            }
            # Found a .NET Framework version - skip it, keep searching
        }
    }

    # No compatible version found - try auto-downloading
    $alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'WinSCPnet' }
    Write-Host "  No PowerShell 7-compatible WinSCP .NET assembly found." -ForegroundColor Yellow
    $downloaded = Install-WinSCPNetStandard
    if ($downloaded) {
        if ($alreadyLoaded) {
            Write-Host "  An incompatible version was already loaded in this session." -ForegroundColor Yellow
            Write-Host "  Please restart the script to use the new version." -ForegroundColor Yellow
            return $null
        }
        return $downloaded
    }

    return $null
}

function Connect-SFTPSession {
    param(
        [string]$DllPath,
        [string]$HostName,
        [int]$Port,
        [string]$Username,
        [string]$Password,
        [string]$PrivateKeyPath
    )

    # Skip Add-Type if WinSCPnet is already loaded in this session
    $loadedAsm = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'WinSCPnet' }
    if (-not $loadedAsm) {
        Add-Type -Path $DllPath
    }

    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $HostName
        PortNumber = $Port
        UserName = $Username
    }

    if ($PrivateKeyPath -and (Test-Path $PrivateKeyPath)) {
        $sessionOptions.SshPrivateKeyPath = $PrivateKeyPath
    } elseif ($Password) {
        $sessionOptions.Password = $Password
    } else {
        throw "No password or private key configured"
    }

    # Accept any host key (you may want to be more strict in production)
    $sessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true

    $session = New-Object WinSCP.Session

    # Suppress WinSCP.exe console progress output (prevents stale lines persisting in terminal)
    try { $session.DisableTransferProgress = $true } catch {}

    # Skip .NET version check — WinSCP.exe may not recognize newer .NET runtimes (e.g., .NET 10)
    try { $session.DisableVersionCheck = $true } catch {}

    # Set executable path - prefer the GUI install (version-matched), then fall back to DLL folder
    $dllFolder = Split-Path $DllPath -Parent
    $exePaths = @(
        "$env:LOCALAPPDATA\Programs\WinSCP\WinSCP.exe",
        "${env:ProgramFiles}\WinSCP\WinSCP.exe",
        "${env:ProgramFiles(x86)}\WinSCP\WinSCP.exe",
        (Join-Path $dllFolder "WinSCP.exe")
    )
    $exePath = $exePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($exePath) {
        $session.ExecutablePath = $exePath
    }

    $session.Open($sessionOptions)

    return $session
}

function Get-RemoteFilesRecursive {
    param(
        $Session,
        [string]$RemotePath,
        [string]$BasePath = $null,
        [int]$Depth = 0
    )

    # Track the base path for progress display
    if (-not $BasePath) {
        $BasePath = $RemotePath
    }

    $files = @()

    try {
        $directoryInfo = $Session.ListDirectory($RemotePath)

        foreach ($fileInfo in $directoryInfo.Files) {
            if ($fileInfo.IsDirectory) {
                if ($fileInfo.Name -ne "." -and $fileInfo.Name -ne "..") {
                    # Show folder name at depth 1 (immediate children of base path)
                    if ($Depth -eq 0) {
                        $label = "      Scanning: $($fileInfo.Name)"
                        $maxWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
                        if ($maxWidth -gt 0 -and $label.Length -gt $maxWidth) {
                            $label = $label.Substring(0, $maxWidth)
                        }
                        $padding = if ($maxWidth -gt $label.Length) { ' ' * ($maxWidth - $label.Length) } else { '' }
                        Write-Host "`r$label$padding" -ForegroundColor DarkGray -NoNewline
                    }
                    # Recurse into subdirectories
                    $subPath = "$RemotePath/$($fileInfo.Name)"
                    $files += Get-RemoteFilesRecursive -Session $Session -RemotePath $subPath -BasePath $BasePath -Depth ($Depth + 1)
                }
            } else {
                $files += [PSCustomObject]@{
                    FullPath = "$RemotePath/$($fileInfo.Name)"
                    Name = $fileInfo.Name
                    Size = $fileInfo.Length
                    LastModified = $fileInfo.LastWriteTime
                }
            }
        }
    } catch {
        Write-Host "  Error listing $RemotePath : $_" -ForegroundColor Red
    }

    return $files
}

function Get-SyncDestinationFolder {
    param(
        [string]$FileName,
        [string]$RemoteFullPath,
        [string[]]$RemotePaths,
        [long]$FileSize,
        [string]$LocalBasePath,
        [string[]]$MovieExtensions,
        [double]$MovieMinSizeGB,
        [hashtable]$FolderCategoryCache = @{}
    )

    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
    $sizeGB = $FileSize / 1GB

    # Extract relative path from remote (the folder structure after the base remote path)
    $relativePath = $RemoteFullPath
    foreach ($rp in $RemotePaths) {
        if ($relativePath.StartsWith($rp)) {
            $relativePath = $relativePath.Substring($rp.Length).TrimStart('/')
            break
        }
    }

    # Get the parent folder name (release/show folder)
    $pathParts = $relativePath -split '/'

    # Strip common remote category directories (e.g., "movies/Movie Name/file.mkv" → "Movie Name/file.mkv")
    # These add an unnecessary nesting level since we already classify into _Movies/_Shows/_Downloads
    # Also handles loose files like "movies/file.srt" — the category dir is used as a classification hint
    $categoryDirNames = @('movies', 'films', 'film', 'tv', 'shows', 'tv shows', 'television', 'anime', 'documentaries', 'docs', 'media', 'radarr', 'sonarr')
    $movieCategoryNames = @('movies', 'films', 'film', 'radarr')
    $tvCategoryNames = @('tv', 'shows', 'tv shows', 'television', 'sonarr')
    $categoryHint = $null
    if ($pathParts.Count -ge 2 -and $pathParts[0].ToLower() -in $categoryDirNames) {
        $categoryHint = $pathParts[0].ToLower()
        $pathParts = $pathParts[1..($pathParts.Count - 1)]
    }

    $parentFolder = if ($pathParts.Count -gt 1) { $pathParts[0] } else { $null }

    # Companion files - these follow the video file's category, not their own
    # NFO, subtitles, images, etc. should stay with their video
    $companionExtensions = @(".nfo", ".srt", ".sub", ".idx", ".ass", ".ssa", ".jpg", ".jpeg", ".png", ".txt", ".tmp")

    # If this is a companion file and we've already categorized this folder, use that
    if ($companionExtensions -contains $ext -and $parentFolder -and $FolderCategoryCache.ContainsKey($parentFolder)) {
        $baseFolder = $FolderCategoryCache[$parentFolder]
        return Join-Path $baseFolder $parentFolder
    }

    # Extension-based routing for non-video content (only for standalone files)
    $musicExtensions = @(".mp3", ".flac", ".m4a", ".aac", ".ogg", ".wav", ".wma", ".alac", ".aiff")
    $bookExtensions = @(".epub", ".mobi", ".azw", ".azw3", ".pdf", ".djvu", ".cbr", ".cbz")

    if ($musicExtensions -contains $ext) {
        $baseFolder = Join-Path $LocalBasePath "_Music"
        if ($parentFolder) {
            $FolderCategoryCache[$parentFolder] = $baseFolder
            return Join-Path $baseFolder $parentFolder
        }
        return $baseFolder
    }

    if ($bookExtensions -contains $ext) {
        $baseFolder = Join-Path $LocalBasePath "_Books"
        if ($parentFolder) {
            $FolderCategoryCache[$parentFolder] = $baseFolder
            return Join-Path $baseFolder $parentFolder
        }
        return $baseFolder
    }

    # Determine category based on file AND parent folder characteristics
    $isMovie = $false
    $isTVShow = $false

    # Large video files are likely movies
    if ($MovieExtensions -contains $ext -and $sizeGB -ge $MovieMinSizeGB) {
        $isMovie = $true
    }

    # TV show detection - check filename AND parent folder name
    # Patterns: S01E01, S01.E01, 1x01, E01, Season 1, "Complete Series", etc.
    $tvPatterns = 'S\d{1,2}E\d{1,2}|S\d{1,2}\.E\d{1,2}|\d{1,2}x\d{2}|Season\s*\d+|Complete\s*Series|Episode\s*\d+'
    if ($FileName -match $tvPatterns) {
        $isTVShow = $true
    } elseif ($parentFolder -and $parentFolder -match $tvPatterns) {
        $isTVShow = $true
    }
    # Also check for standalone episode numbers like "E01 - Title" at the start
    if ($FileName -match '^E\d{1,3}\s*[-.]') {
        $isTVShow = $true
    }

    # TV shows take priority over movie detection (episodes can be large)
    if ($isTVShow) {
        $baseFolder = Join-Path $LocalBasePath "_Shows"
    } elseif ($isMovie) {
        $baseFolder = Join-Path $LocalBasePath "_Movies"
    } elseif ($categoryHint -and $categoryHint -in $movieCategoryNames) {
        # File came from a remote movies/ directory — route to _Movies
        $baseFolder = Join-Path $LocalBasePath "_Movies"
    } elseif ($categoryHint -and $categoryHint -in $tvCategoryNames) {
        # File came from a remote tv/ directory — route to _Shows
        $baseFolder = Join-Path $LocalBasePath "_Shows"
    } else {
        # For unknown companion files without a cached folder, default to _Downloads
        $baseFolder = Join-Path $LocalBasePath "_Downloads"
    }

    # Cache the folder category for companion files
    if ($parentFolder) {
        $FolderCategoryCache[$parentFolder] = $baseFolder
        return Join-Path $baseFolder $parentFolder
    }

    return $baseFolder
}

<#
.SYNOPSIS
    Builds a flat lookup of every file under the given roots, keyed by
    "Name|Size", value is the first-seen FullName.
.DESCRIPTION
    One-time build is far cheaper than per-file recursive scans during sync
    (261 files * recursive walk of E:\Movies = tens of thousands of FS hits;
    one walk + O(1) lookups = thousands).

    Missing or unreadable roots are silently skipped so the caller can pass a
    mix of inbox and library paths without per-path defensiveness.

    Collisions on Name|Size keep the first occurrence — sufficient for the
    "do I already have this file?" question.
#>
function Build-LocalFileIndex {
    param(
        [string[]]$RootPaths
    )

    $index = @{}
    foreach ($root in $RootPaths) {
        if (-not $root) { continue }
        if (-not (Test-Path -LiteralPath $root)) { continue }
        Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
            $key = "$($_.Name)|$($_.Length)"
            if (-not $index.ContainsKey($key)) {
                $index[$key] = $_.FullName
            }
        }
    }
    return $index
}

<#
.SYNOPSIS
    Returns a case-insensitive set of immediate-child folder names under the
    given roots — typically the canonical movie/show folders in the library.
.DESCRIPTION
    Folder-name match is the primary "I already have this" signal for the
    library case: a release file inside a seedbox folder named
    "Pride & Prejudice (2005)" should be considered already-have if the
    library has a sibling folder with the same name, even when the local
    video file has been renamed away from its release name and so won't
    match by Name|Size.
#>
function Build-LocalFolderSet {
    param(
        [string[]]$RootPaths
    )

    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($root in $RootPaths) {
        if (-not $root) { continue }
        if (-not (Test-Path -LiteralPath $root)) { continue }
        Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            [void]$set.Add($_.Name)
        }
    }
    return $set
}

<#
.SYNOPSIS
    Returns whether a remote file is already present locally and how it was
    matched.
.DESCRIPTION
    Two-tier check, strongest signal first:
      1. Name|Size match in the file index — covers manually-transferred
         files in the inbox and any library file that kept its release name.
      2. Folder-name match in the folder set — covers the common library case
         where the user's local file was renamed during processing but the
         containing release folder still uses the canonical name that the
         seedbox also uses.
    Returns @{ Found = $bool; LocalPath = $string?; MatchType = 'NameSize'|'Folder'? }.
#>
function Test-RemoteFileAlreadyHave {
    param(
        [string]$FileName,
        [long]$FileSize,
        [string]$RemoteParentName,
        $FileIndex,
        $FolderSet
    )

    if ($FileIndex) {
        $key = "$FileName|$FileSize"
        if ($FileIndex.ContainsKey($key)) {
            return @{ Found = $true; LocalPath = $FileIndex[$key]; MatchType = 'NameSize' }
        }
    }

    if ($FolderSet -and $RemoteParentName -and $FolderSet.Contains($RemoteParentName)) {
        return @{ Found = $true; LocalPath = $null; MatchType = 'Folder' }
    }

    return @{ Found = $false }
}

function Invoke-FileDownload {
    param(
        $Session,
        [string]$RemotePath,
        [string]$LocalPath,
        [int]$SpeedLimitKBps = 0
    )

    $localDir = Split-Path $LocalPath -Parent
    if (-not (Test-Path $localDir)) {
        New-Item -Path $localDir -ItemType Directory -Force | Out-Null
    }

    # Delete any partial file from interrupted transfer
    if (Test-Path $LocalPath) {
        Remove-Item $LocalPath -Force -ErrorAction SilentlyContinue
    }

    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
    # Enable resume support for interrupted transfers
    $transferOptions.ResumeSupport.State = [WinSCP.TransferResumeSupportState]::On
    # Apply speed limit if configured (WinSCP uses KB/s)
    if ($SpeedLimitKBps -gt 0) {
        $transferOptions.SpeedLimit = $SpeedLimitKBps
    }

    $escapedRemote = [WinSCP.RemotePath]::EscapeFileMask($RemotePath)
    $result = $Session.GetFiles($escapedRemote, $LocalPath, $false, $transferOptions)
    $result.Check()

    return $result.IsSuccess
}

#endregion

#region Public Functions

<#
.SYNOPSIS
    Syncs files from an SFTP server to local storage
.DESCRIPTION
    Connects to an SFTP server and downloads new files, tracking what's
    already been downloaded to avoid duplicates.
.PARAMETER HostName
    SFTP server hostname
.PARAMETER Port
    SFTP server port (default: 22)
.PARAMETER Username
    SFTP username
.PARAMETER Password
    SFTP password (use this OR PrivateKeyPath)
.PARAMETER PrivateKeyPath
    Path to SSH private key (use this OR Password)
.PARAMETER RemotePaths
    Array of remote paths to scan for files
.PARAMETER LocalBasePath
    Local base path for downloads (files organized into subfolders)
.PARAMETER TrackingFile
    Path to tracking file (default: AppData\LibraryLint\sftp_downloaded.json)
.PARAMETER MovieExtensions
    File extensions to consider as movies
.PARAMETER MovieMinSizeGB
    Minimum size in GB to consider a file a movie
.PARAMETER DeleteAfterDownload
    Delete files from server after successful download
.PARAMETER SpeedLimitKBps
    Maximum download speed in KB/s (0 = unlimited)
.PARAMETER ExcludePatterns
    Array of wildcard patterns to exclude from download (e.g., "*.nfo", "*.txt", "Sample*")
.PARAMETER WhatIf
    Preview without downloading (shows what would be downloaded and where)
.PARAMETER Force
    Re-download files even if already tracked
.EXAMPLE
    Invoke-SFTPSync -HostName "server.com" -Username "user" -Password "pass" -RemotePaths @("/downloads") -LocalBasePath "G:"
.EXAMPLE
    Invoke-SFTPSync -HostName "server.com" -Username "user" -Password "pass" -RemotePaths @("/downloads") -LocalBasePath "G:" -SpeedLimitKBps 5120 -ExcludePatterns @("*.nfo", "*.txt", "Sample*")
#>
<#
.SYNOPSIS
    Removes a remote directory only if it contains no entries beyond '.'/'..'.
.DESCRIPTION
    Best-effort empty-directory cleanup for the SFTP delete flows. Lists the
    directory, filters synthetic '.' and '..' entries that WinSCP includes,
    and only calls RemoveFiles when nothing else is present — so a folder
    holding untracked files (samples the user keeps, partial uploads, anything
    the script doesn't own) stays put.

    There's a tiny TOCTOU window between the list and the remove. Acceptable
    on a user-owned seedbox where the only writer is the script itself; not
    suitable for actively-written paths.
.OUTPUTS
    Hashtable: Removed (bool), and one of Reason (non-empty) or Error (failure).
#>
function Remove-RemoteDirIfEmpty {
    param(
        $Session,
        [string]$RemotePath
    )

    try {
        $listing = $Session.ListDirectory($RemotePath)
        $remaining = @($listing.Files | Where-Object { $_.Name -ne '.' -and $_.Name -ne '..' })
        if ($remaining.Count -eq 0) {
            # Trailing slash + EscapeFileMask so WinSCP treats the path as a
            # single directory entry (not a wildcard mask) even when the
            # release name contains brackets or other special chars.
            $pathForRemoval = $RemotePath.TrimEnd('/') + '/'
            $escaped = [WinSCP.RemotePath]::EscapeFileMask($pathForRemoval)
            $Session.RemoveFiles($escaped).Check()
            return @{ Removed = $true }
        }
        return @{ Removed = $false; Reason = "$($remaining.Count) item(s) remain" }
    } catch {
        return @{ Removed = $false; Error = $_.ToString() }
    }
}

<#
.SYNOPSIS
    Returns true if the given remote path equals or is an ancestor of any of
    the supplied root paths — i.e. removing it would damage a configured root.
.DESCRIPTION
    Used by the sync delete-after-download cleanup to make sure we never
    bubble up the empty-folder removal past the seedbox roots the user
    configured. Forward-slash, case-sensitive comparison (typical Linux
    seedbox semantics).
#>
function Test-RemotePathAtOrAboveRoot {
    param(
        [string]$Path,
        [string[]]$Roots
    )

    $p = $Path.TrimEnd('/')
    foreach ($root in $Roots) {
        $r = $root.TrimEnd('/')
        if ($p -eq $r) { return $true }
        if ($r.StartsWith($p + '/')) { return $true }
    }
    return $false
}

function Invoke-SFTPSync {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HostName,

        [int]$Port = 22,

        [Parameter(Mandatory=$true)]
        [string]$Username,

        [string]$Password,
        [string]$PrivateKeyPath,

        [Parameter(Mandatory=$true)]
        [string[]]$RemotePaths,

        [Parameter(Mandatory=$true)]
        [string]$LocalBasePath,

        [string[]]$LibraryPaths = @(),

        [string]$TrackingFile,

        [string[]]$MovieExtensions = @(".mkv", ".mp4", ".avi", ".m4v"),
        [double]$MovieMinSizeGB = 1,

        [switch]$DeleteAfterDownload,
        [int]$SpeedLimitKBps = 0,
        [string[]]$ExcludePatterns = @(),
        [switch]$WhatIf,
        [switch]$Force
    )

    # Header
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "                    SFTP SYNC                          " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] No files will be downloaded" -ForegroundColor Yellow
        Write-Host ""
    }

    # Check WinSCP .NET assembly
    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        Write-Host ""
        Write-Host "  The WinSCP GUI app doesn't include the .NET assembly by default." -ForegroundColor Yellow
        Write-Host "  To enable SFTP sync:" -ForegroundColor Yellow
        Write-Host "    1. Download from: https://winscp.net/eng/downloads.php" -ForegroundColor Gray
        Write-Host "    2. Get the 'Automation' package or '.NET assembly / COM library'" -ForegroundColor Gray
        Write-Host "    3. Extract WinSCPnet.dll to one of:" -ForegroundColor Gray
        Write-Host "       - $env:LOCALAPPDATA\LibraryLint\" -ForegroundColor Gray
        Write-Host "       - $env:ProgramFiles\WinSCP\" -ForegroundColor Gray
        Write-Host ""
        return @{ Downloaded = 0; Failed = 0; Error = "WinSCP .NET assembly not installed" }
    }

    # Tracking file path
    $trackingPath = Get-SyncTrackingPath -ConfigTrackingFile $TrackingFile

    Write-Host "  Host: ${HostName}:${Port}" -ForegroundColor Gray
    Write-Host "  User: $Username" -ForegroundColor Gray
    Write-Host "  Dest: $LocalBasePath" -ForegroundColor Gray
    if ($SpeedLimitKBps -gt 0) {
        $limitDisplay = if ($SpeedLimitKBps -ge 1024) { "{0:N1} MB/s" -f ($SpeedLimitKBps / 1024) } else { "$SpeedLimitKBps KB/s" }
        Write-Host "  Speed: $limitDisplay" -ForegroundColor Gray
    }
    if ($ExcludePatterns.Count -gt 0) {
        Write-Host "  Exclude: $($ExcludePatterns -join ', ')" -ForegroundColor Gray
    }
    Write-Host ""

    # Connect
    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ Downloaded = 0; Failed = 0; Error = $_.ToString() }
    }

    $result = @{
        Downloaded = 0
        Failed = 0
        SkippedDuplicates = 0
        BytesDownloaded = 0
        Duration = $null
        FoldersRemoved = 0
    }

    # Parents we deleted from this run. Walked in a single post-loop sweep so
    # nested empties cascade up (release/sample/ removed first, then release/).
    $touchedParents = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)

    try {
        # Scan remote paths
        Write-Host ""
        Write-Host "  Scanning remote paths..." -ForegroundColor Gray

        $allFiles = @()
        foreach ($remotePath in $RemotePaths) {
            $files = Get-RemoteFilesRecursive -Session $session -RemotePath $remotePath
            # Clear the scanning progress line before printing the summary
            $clearWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
            if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
            Write-Host "    $remotePath ($($files.Count) files)" -ForegroundColor Gray
            $allFiles += $files
        }

        # Get tracking data
        $downloaded = Get-DownloadedFiles -TrackingPath $trackingPath

        # Apply exclusion patterns
        $excludedCount = 0
        if ($ExcludePatterns.Count -gt 0) {
            $beforeCount = $allFiles.Count
            $allFiles = $allFiles | Where-Object {
                $name = $_.Name
                $excluded = $false
                foreach ($pattern in $ExcludePatterns) {
                    if ($name -like $pattern) { $excluded = $true; break }
                }
                -not $excluded
            }
            $excludedCount = $beforeCount - $allFiles.Count
        }

        # Filter out already downloaded
        $newFiles = $allFiles | Where-Object {
            $Force -or (-not $downloaded.ContainsKey($_.FullPath))
        }
        $skippedCount = $allFiles.Count - $newFiles.Count

        # Sort files so the largest video files are processed first - this ensures the folder
        # category cache is populated by the main movie/episode (not trailers or samples) before
        # companion files (NFO, subtitles, artwork) are processed
        $newFiles = $newFiles | Sort-Object {
            $ext = [System.IO.Path]::GetExtension($_.Name).ToLower()
            if ($MovieExtensions -contains $ext) { 0 } else { 1 }
        }, { -$_.Size }, Name

        # Extract unique folder names from new files
        $newFolders = $newFiles | ForEach-Object {
            # Get the first folder after the remote path (the release/movie folder)
            $relativePath = $_.FullPath
            foreach ($rp in $RemotePaths) {
                if ($relativePath.StartsWith($rp)) {
                    $relativePath = $relativePath.Substring($rp.Length).TrimStart('/')
                    break
                }
            }
            # Get the top-level folder name
            $parts = $relativePath -split '/'
            if ($parts.Count -gt 1) { $parts[0] } else { $null }
        } | Where-Object { $_ } | Select-Object -Unique | Sort-Object

        Write-Host ""
        Write-Host "  Total files:    $($allFiles.Count + $excludedCount)" -ForegroundColor White
        if ($excludedCount -gt 0) {
            Write-Host "  Excluded:       $excludedCount" -ForegroundColor Gray
        }
        Write-Host "  Already synced: $skippedCount" -ForegroundColor Gray
        Write-Host "  New files:      $($newFiles.Count)" -ForegroundColor $(if ($newFiles.Count -gt 0) { 'Green' } else { 'Gray' })

        if ($newFiles.Count -eq 0) {
            Write-Host ""
            Write-Host "  Nothing new to download!" -ForegroundColor Green
            return $result
        }

        # Show folder names
        if ($newFolders.Count -gt 0) {
            Write-Host ""
            Write-Host "  New folders to download:" -ForegroundColor Cyan
            foreach ($folder in $newFolders) {
                # Count files and size for this folder
                $folderFiles = $newFiles | Where-Object { $_.FullPath -like "*/$folder/*" -or $_.FullPath -like "*/$folder" }
                $folderSize = ($folderFiles | Measure-Object -Property Size -Sum).Sum
                Write-Host "    - $folder " -NoNewline -ForegroundColor White
                Write-Host "($(Format-SyncSize $folderSize))" -ForegroundColor Gray
            }
        }

        $totalSize = ($newFiles | Measure-Object -Property Size -Sum).Sum
        Write-Host ""
        Write-Host "  Total size:     $(Format-SyncSize $totalSize)" -ForegroundColor Cyan

        # Check local drive capacity
        $driveRoot = [System.IO.Path]::GetPathRoot($LocalBasePath)
        $diskInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($driveRoot.TrimEnd('\'))'" -ErrorAction SilentlyContinue
        if ($diskInfo) {
            $freeSpace = $diskInfo.FreeSpace
            $totalSpace = $diskInfo.Size
            $freePercent = [math]::Round(($freeSpace / $totalSpace) * 100, 1)

            Write-Host "  Drive space:    $(Format-SyncSize $freeSpace) free ($freePercent%)" -ForegroundColor $(if ($freePercent -lt 10) { 'Red' } elseif ($freePercent -lt 20) { 'Yellow' } else { 'Gray' })

            # Warn if low disk space
            if ($freePercent -lt 10) {
                Write-Host ""
                Write-Host "  WARNING: Disk space is critically low ($freePercent% free)!" -ForegroundColor Red
            }

            # Check if enough space for download
            if ($totalSize -gt $freeSpace) {
                Write-Host ""
                Write-Host "  ERROR: Not enough disk space!" -ForegroundColor Red
                Write-Host "  Need $(Format-SyncSize $totalSize), only $(Format-SyncSize $freeSpace) available." -ForegroundColor Red
                Write-Host "  Short by $(Format-SyncSize ($totalSize - $freeSpace))." -ForegroundColor Red
                Write-Host ""
                $continue = Read-Host "  Continue anyway? (Y/N) [N]"
                if ($continue -notmatch '^[Yy]') {
                    Write-Host "  Download cancelled." -ForegroundColor Yellow
                    return $result
                }
            }
        }

        Write-Host ""

        # Build the local "do I already have this?" lookup. The file index
        # covers inbox roots + library roots (catches manually-transferred
        # files in either, including library files that kept their release
        # name). The folder set covers library-only and catches the common
        # case where the local video file was renamed during processing but
        # its release-folder name still matches the seedbox folder.
        $inboxRoots = @(
            Join-Path $LocalBasePath "_Movies"
            Join-Path $LocalBasePath "_Shows"
            Join-Path $LocalBasePath "_Music"
            Join-Path $LocalBasePath "_Books"
            Join-Path $LocalBasePath "_Downloads"
        )
        $indexRoots = @($inboxRoots) + @($LibraryPaths)
        Write-Host "  Indexing local files..." -ForegroundColor Gray -NoNewline
        $localIndex = Build-LocalFileIndex -RootPaths $indexRoots
        $localFolders = Build-LocalFolderSet -RootPaths $LibraryPaths
        Write-Host " $($localIndex.Count) files, $($localFolders.Count) library folders" -ForegroundColor Gray
        Write-Host ""

        # Download files
        Write-Host "  Downloading..." -ForegroundColor Yellow
        Write-Host ""

        $downloadCount = 0
        $failCount = 0
        $skippedDupes = 0
        $downloadedBytes = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        # Cache folder categories so companion files (NFO, SRT, images) go to same place as video
        $folderCategoryCache = @{}

        foreach ($file in $newFiles) {
            $destFolder = Get-SyncDestinationFolder -FileName $file.Name -RemoteFullPath $file.FullPath `
                -RemotePaths $RemotePaths -FileSize $file.Size -LocalBasePath $LocalBasePath `
                -MovieExtensions $MovieExtensions -MovieMinSizeGB $MovieMinSizeGB -FolderCategoryCache $folderCategoryCache
            $localPath = Join-Path $destFolder $file.Name

            $truncatedName = if ($file.Name.Length -gt 45) { $file.Name.Substring(0, 42) + "..." } else { $file.Name }
            $progress = "[$($downloadCount + $failCount + $skippedDupes + 1)/$($newFiles.Count)]"

            Write-Host "  $progress $truncatedName" -ForegroundColor White -NoNewline
            Write-Host " ($(Format-SyncSize $file.Size))" -ForegroundColor Gray -NoNewline

            # Check if file already exists locally (manual transfer or
            # already-processed library copy)
            $remoteParent = Split-Path $file.FullPath -Parent
            $remoteParentName = if ($remoteParent) { Split-Path $remoteParent -Leaf } else { $null }
            $haveCheck = Test-RemoteFileAlreadyHave -FileName $file.Name -FileSize $file.Size `
                -RemoteParentName $remoteParentName -FileIndex $localIndex -FolderSet $localFolders
            if ($haveCheck.Found) {
                $skipReason = if ($haveCheck.MatchType -eq 'Folder') { 'already in library' } else { 'already exists' }
                Write-Host " SKIP ($skipReason)" -ForegroundColor Cyan
                # Track it so we don't check again next time
                $downloaded[$file.FullPath] = @{
                    LocalPath = $haveCheck.LocalPath
                    Size = $file.Size
                    DownloadedAt = (Get-Date).ToString("o")
                    ManualTransfer = $true
                    MatchType = $haveCheck.MatchType
                }
                Save-DownloadedFiles -Downloaded $downloaded -TrackingPath $trackingPath
                $skippedDupes++
                continue
            }

            if ($WhatIf) {
                Write-Host " [would download to $destFolder]" -ForegroundColor Yellow
                $downloadCount++
                continue
            }

            # Retry logic - try up to 3 times
            $maxRetries = 3
            $attempt = 0
            $success = $false

            while ($attempt -lt $maxRetries -and -not $success) {
                $attempt++
                try {
                    if ($attempt -gt 1) {
                        Write-Host " retry $attempt..." -ForegroundColor Yellow -NoNewline
                        Start-Sleep -Seconds 2  # Brief pause before retry
                    }

                    $success = Invoke-FileDownload -Session $session -RemotePath $file.FullPath -LocalPath $localPath -SpeedLimitKBps $SpeedLimitKBps

                    if ($success) {
                        Write-Host " OK" -ForegroundColor Green

                        # Track the download
                        $downloaded[$file.FullPath] = @{
                            LocalPath = $localPath
                            Size = $file.Size
                            DownloadedAt = (Get-Date).ToString("o")
                        }
                        Save-DownloadedFiles -Downloaded $downloaded -TrackingPath $trackingPath

                        $downloadCount++
                        $downloadedBytes += $file.Size

                        # Delete from server if configured
                        if ($DeleteAfterDownload) {
                            $session.RemoveFiles([WinSCP.RemotePath]::EscapeFileMask($file.FullPath)).Check()
                            $parent = Split-Path $file.FullPath -Parent
                            if ($parent) { [void]$touchedParents.Add($parent) }
                        }
                    } elseif ($attempt -eq $maxRetries) {
                        Write-Host " FAILED" -ForegroundColor Red
                        $failCount++
                    }
                } catch {
                    if ($attempt -eq $maxRetries) {
                        Write-Host " ERROR: $_" -ForegroundColor Red
                        $failCount++
                    }
                }
            }
        }

        # Empty-folder cleanup. Sort deepest-first so when we remove a sample/
        # subfolder, its now-empty parent release/ folder also gets a turn in
        # this same sweep. Bounded by $RemotePaths so we never walk up to (or
        # past) a configured seedbox root.
        $foldersRemoved = 0
        if ($DeleteAfterDownload -and $touchedParents.Count -gt 0) {
            $sortedParents = @($touchedParents) | Sort-Object -Property Length -Descending
            $cleanupQueue = New-Object 'System.Collections.Generic.Queue[string]'
            foreach ($p in $sortedParents) { $cleanupQueue.Enqueue($p) }

            while ($cleanupQueue.Count -gt 0) {
                $candidate = $cleanupQueue.Dequeue()
                if (Test-RemotePathAtOrAboveRoot -Path $candidate -Roots $RemotePaths) { continue }

                $outcome = Remove-RemoteDirIfEmpty -Session $session -RemotePath $candidate
                if ($outcome.Removed) {
                    $foldersRemoved++
                    # Cascade: the parent of what we just removed might now be
                    # empty too. Push it onto the queue.
                    $grandparent = Split-Path $candidate -Parent
                    if ($grandparent -and -not (Test-RemotePathAtOrAboveRoot -Path $grandparent -Roots $RemotePaths)) {
                        $cleanupQueue.Enqueue($grandparent)
                    }
                }
            }
        }

        $stopwatch.Stop()

        # Summary
        Write-Host ""
        Write-Host "======================================================" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Downloaded: $downloadCount files ($(Format-SyncSize $downloadedBytes))" -ForegroundColor Green
        if ($skippedDupes -gt 0) {
            Write-Host "  Skipped:    $skippedDupes files (already in inbox)" -ForegroundColor Cyan
        }
        if ($failCount -gt 0) {
            Write-Host "  Failed:     $failCount files" -ForegroundColor Red
        }
        if ($foldersRemoved -gt 0) {
            Write-Host "  Cleaned:    $foldersRemoved empty folder(s) removed from server" -ForegroundColor Green
        }
        Write-Host "  Time:       $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
        Write-Host ""

        $result.Downloaded = $downloadCount
        $result.Failed = $failCount
        $result.SkippedDuplicates = $skippedDupes
        $result.BytesDownloaded = $downloadedBytes
        $result.Duration = $stopwatch.Elapsed
        $result.FoldersRemoved = $foldersRemoved

    } finally {
        $session.Dispose()
    }

    return $result
}

<#
.SYNOPSIS
    Prunes old files from SFTP server that were downloaded more than X days ago
.DESCRIPTION
    Checks the tracking file for files that were downloaded, and deletes from
    the remote server any files older than the specified number of days.
.PARAMETER HostName
    SFTP server hostname
.PARAMETER Port
    SFTP server port (default: 22)
.PARAMETER Username
    SFTP username
.PARAMETER Password
    SFTP password (use this OR PrivateKeyPath)
.PARAMETER PrivateKeyPath
    Path to SSH private key (use this OR Password)
.PARAMETER DaysOld
    Delete files downloaded more than this many days ago
.PARAMETER TrackingFile
    Path to tracking file (default: AppData\LibraryLint\sftp_downloaded.json)
.PARAMETER WhatIf
    Preview without deleting
.EXAMPLE
    Invoke-SFTPPrune -HostName "server.com" -Username "user" -Password "pass" -DaysOld 7
#>
function Invoke-SFTPPrune {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HostName,

        [int]$Port = 22,

        [Parameter(Mandatory=$true)]
        [string]$Username,

        [string]$Password,
        [string]$PrivateKeyPath,

        [Parameter(Mandatory=$true)]
        [int]$DaysOld,

        [string]$TrackingFile,

        [string[]]$RemotePaths = @(),

        [hashtable]$RadarrImportedPaths = $null,  # TMDB ID → movie info from Radarr

        [switch]$WhatIf
    )

    # Header
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "                  SFTP PRUNE                           " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] No files will be deleted" -ForegroundColor Yellow
        Write-Host ""
    }

    $hasRadarrVerify = $null -ne $RadarrImportedPaths
    if ($hasRadarrVerify) {
        Write-Host "  Radarr verification: enabled" -ForegroundColor Green
        Write-Host "  Files will only be pruned if Radarr has imported them." -ForegroundColor Gray
        Write-Host ""
    }

    Write-Host "  Pruning files downloaded more than $DaysOld days ago" -ForegroundColor Yellow
    Write-Host ""

    # Check WinSCP .NET assembly
    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        return @{ Deleted = 0; Failed = 0; Error = "WinSCP .NET assembly not installed" }
    }

    # Load tracking data
    $trackingPath = Get-SyncTrackingPath -ConfigTrackingFile $TrackingFile
    if (-not (Test-Path $trackingPath)) {
        Write-Host "  No tracking file found - nothing to prune" -ForegroundColor Yellow
        return @{ Deleted = 0; Failed = 0 }
    }

    $downloaded = Get-DownloadedFiles -TrackingPath $trackingPath
    if ($downloaded.Count -eq 0) {
        Write-Host "  No tracked downloads - nothing to prune" -ForegroundColor Yellow
        return @{ Deleted = 0; Failed = 0 }
    }

    # Find files older than threshold
    $cutoffDate = (Get-Date).AddDays(-$DaysOld)
    $filesToPrune = @()

    foreach ($entry in $downloaded.GetEnumerator()) {
        $downloadedAt = [DateTime]::Parse($entry.Value.DownloadedAt)
        if ($downloadedAt -lt $cutoffDate) {
            # Resolve the actual remote path — the tracked path may have changed if remote paths config was updated
            $trackedPath = $entry.Key
            $resolvedPath = $trackedPath

            # Check if the tracked path starts with any current remote path
            $matchesCurrent = $false
            foreach ($rp in $RemotePaths) {
                if ($trackedPath.StartsWith($rp)) { $matchesCurrent = $true; break }
            }

            # If it doesn't match, try to remap by extracting the relative portion
            # e.g., tracked: /home/user/downloads/radarr/radarr/Movie/file → relative: Movie/file
            # current remote path: /home/user/downloads/radarr/ → resolved: /home/user/downloads/radarr/Movie/file
            if (-not $matchesCurrent) {
                # Find the deepest matching portion of the path against any remote path
                $pathParts = $trackedPath.Split('/')
                # Try progressively shorter prefixes to find the relative path
                for ($i = 1; $i -lt $pathParts.Count - 1; $i++) {
                    $candidateRelative = ($pathParts[$i..($pathParts.Count - 1)] -join '/')
                    foreach ($rp in $RemotePaths) {
                        $candidateFull = "$($rp.TrimEnd('/'))/$candidateRelative"
                        if ($candidateFull -ne $trackedPath) {
                            $resolvedPath = $candidateFull
                            $matchesCurrent = $true
                            break
                        }
                    }
                    if ($matchesCurrent) { break }
                }
            }

            $filesToPrune += @{
                RemotePath = $resolvedPath
                TrackingKey = $trackedPath
                LocalPath = $entry.Value.LocalPath
                Size = $entry.Value.Size
                DownloadedAt = $downloadedAt
                Age = [math]::Floor(((Get-Date) - $downloadedAt).TotalDays)
            }
        }
    }

    if ($filesToPrune.Count -eq 0) {
        Write-Host "  No files older than $DaysOld days" -ForegroundColor Green
        return @{ Deleted = 0; Failed = 0; Skipped = 0 }
    }

    # Radarr verification: separate safe-to-prune from not-yet-imported
    $notImported = @()
    if ($hasRadarrVerify) {
        # Build a lookup of Radarr-imported movie titles (normalized) for matching
        $radarrTitleLookup = @{}
        foreach ($tmdbId in $RadarrImportedPaths.Keys) {
            $info = $RadarrImportedPaths[$tmdbId]
            if ($info.HasFile) {
                $normTitle = ($info.Title -replace '[^\w\s]', ' ' -replace '\s+', ' ').Trim().ToLower()
                $key = "$normTitle|$($info.Year)"
                $radarrTitleLookup[$key] = $true
                # Also index by path leaf for direct path matching
                if ($info.Path) {
                    $pathLeaf = Split-Path $info.Path -Leaf
                    $radarrTitleLookup["path:$($pathLeaf.ToLower())"] = $true
                }
            }
        }

        $safeToPrune = @()
        foreach ($file in $filesToPrune) {
            # Get the folder name for this file (the movie release folder)
            $folderName = Split-Path (Split-Path $file.RemotePath -Parent) -Leaf
            if (-not $folderName -or $folderName -eq '/') {
                $folderName = Split-Path $file.RemotePath -Parent
            }

            # Try to match against Radarr's imported movies
            $isImported = $false

            # Method 1: Direct path leaf match
            if ($radarrTitleLookup.ContainsKey("path:$($folderName.ToLower())")) {
                $isImported = $true
            }

            # Method 2: Normalize folder name and match title+year
            if (-not $isImported) {
                # Extract title and year from release folder name
                $normFolder = ($folderName -replace '[\.\-_]', ' ' -replace '\s+', ' ').Trim().ToLower()
                # Try to extract year
                $folderYear = $null
                if ($normFolder -match '((?:19|20)\d{2})') { $folderYear = $Matches[1] }
                # Strip everything from the year onward for title matching
                $folderTitle = ($normFolder -replace '\s*(?:19|20)\d{2}.*$', '').Trim()
                $folderTitle = $folderTitle -replace '\s*(1080p|720p|2160p|4k|uhd|bluray|web|remux|x264|x265|hevc|aac|dts|ac3|atmos|hdr|10bit).*$', ''
                $folderTitle = $folderTitle.Trim()

                if ($folderTitle -and $folderYear) {
                    $lookupKey = "$folderTitle|$folderYear"
                    if ($radarrTitleLookup.ContainsKey($lookupKey)) {
                        $isImported = $true
                    }
                }
            }

            if ($isImported) {
                $safeToPrune += $file
            } else {
                $notImported += $file
            }
        }

        if ($notImported.Count -gt 0) {
            Write-Host "  Skipping $($notImported.Count) file(s) not yet imported by Radarr:" -ForegroundColor Yellow
            $notImportedFolders = $notImported | Group-Object { Split-Path $_.RemotePath -Parent } | Select-Object -First 10
            foreach ($group in $notImportedFolders) {
                $folderName = Split-Path $group.Name -Leaf
                if (-not $folderName) { $folderName = $group.Name }
                $truncated = if ($folderName.Length -gt 55) { $folderName.Substring(0, 52) + "..." } else { $folderName }
                Write-Host "    - $truncated ($($group.Count) files)" -ForegroundColor DarkYellow
            }
            if ($notImportedFolders.Count -lt ($notImported | Group-Object { Split-Path $_.RemotePath -Parent }).Count) {
                Write-Host "    ... and more" -ForegroundColor DarkGray
            }
            Write-Host ""
            Write-Host "  Note: These files may still be seeding or awaiting import." -ForegroundColor DarkGray
            Write-Host "  They will be eligible for pruning once Radarr imports them." -ForegroundColor DarkGray
            Write-Host ""
        }

        $filesToPrune = $safeToPrune
        if ($filesToPrune.Count -eq 0) {
            Write-Host "  No verified-imported files older than $DaysOld days to prune." -ForegroundColor Green
            return @{ Deleted = 0; Failed = 0; Skipped = $notImported.Count }
        }
    }

    # Group files by parent folder for display
    $totalSize = ($filesToPrune | Measure-Object -Property Size -Sum).Sum
    $folderGroups = $filesToPrune | Group-Object { Split-Path $_.RemotePath -Parent }

    foreach ($group in $folderGroups | Sort-Object { ($_.Group | Measure-Object -Property Age -Maximum).Maximum } -Descending) {
        $folderName = Split-Path $group.Name -Leaf
        if (-not $folderName) { $folderName = $group.Name }
        $truncatedName = if ($folderName.Length -gt 50) { $folderName.Substring(0, 47) + "..." } else { $folderName }
        $folderSize = ($group.Group | Measure-Object -Property Size -Sum).Sum
        $maxAge = ($group.Group | Measure-Object -Property Age -Maximum).Maximum
        Write-Host "    $($maxAge) days old: " -NoNewline -ForegroundColor Gray
        Write-Host "$truncatedName " -NoNewline -ForegroundColor White
        Write-Host "($($group.Count) files, $(Format-SyncSize $folderSize))" -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host "  Found $($filesToPrune.Count) files in $($folderGroups.Count) folders to prune ($(Format-SyncSize $totalSize))" -ForegroundColor Cyan
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] Would delete $($filesToPrune.Count) files" -ForegroundColor Yellow
        return @{ Deleted = $filesToPrune.Count; Failed = 0; Skipped = $notImported.Count; WhatIf = $true }
    }

    # Connect and delete
    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ Deleted = 0; Failed = 0; Error = $_.ToString() }
    }

    $deletedCount = 0
    $failedCount = 0
    $deletedBytes = 0
    $foldersRemoved = 0

    try {
        Write-Host ""
        Write-Host "  Deleting files..." -ForegroundColor Yellow
        Write-Host ""

        $deleteFolderGroups = $filesToPrune | Group-Object { Split-Path $_.RemotePath -Parent }

        foreach ($group in $deleteFolderGroups) {
            $folderName = Split-Path $group.Name -Leaf
            if (-not $folderName) { $folderName = $group.Name }
            $truncatedName = if ($folderName.Length -gt 45) { $folderName.Substring(0, 42) + "..." } else { $folderName }
            $folderFileCount = $group.Count
            $folderDeleted = 0
            $folderFailed = 0
            $folderNotFound = 0

            foreach ($file in $group.Group) {
                try {
                    $escapedPath = [WinSCP.RemotePath]::EscapeFileMask($file.RemotePath)
                    $trackKey = if ($file.TrackingKey) { $file.TrackingKey } else { $file.RemotePath }
                    if ($session.FileExists($escapedPath)) {
                        $session.RemoveFiles($escapedPath).Check()
                        $folderDeleted++
                        $deletedCount++
                        $deletedBytes += $file.Size
                        $downloaded.Remove($trackKey)
                    } else {
                        $folderNotFound++
                        $downloaded.Remove($trackKey)
                    }
                } catch {
                    $folderFailed++
                    $failedCount++
                }
            }

            # Try to remove the now-empty release folder. Only when no per-file
            # failures, so a failed delete doesn't leave the parent looking
            # falsely empty by virtue of our own incomplete cleanup.
            $folderRemovalNote = ''
            if ($folderFailed -eq 0 -and $group.Name) {
                $outcome = Remove-RemoteDirIfEmpty -Session $session -RemotePath $group.Name
                if ($outcome.Removed) {
                    $foldersRemoved++
                    $folderRemovalNote = ', folder removed'
                } elseif ($outcome.Reason) {
                    $folderRemovalNote = ", folder kept ($($outcome.Reason))"
                }
            }

            Write-Host "    $truncatedName " -NoNewline -ForegroundColor White
            if ($folderFailed -gt 0) {
                Write-Host "($folderDeleted/$folderFileCount deleted, $folderFailed failed)" -ForegroundColor Red
            } elseif ($folderNotFound -gt 0) {
                Write-Host "($folderDeleted deleted, $folderNotFound already gone$folderRemovalNote)" -ForegroundColor Gray
            } else {
                Write-Host "($folderDeleted files deleted$folderRemovalNote)" -ForegroundColor Green
            }
        }

        # Save updated tracking file
        Save-DownloadedFiles -Downloaded $downloaded -TrackingPath $trackingPath

    } finally {
        $session.Dispose()
    }

    # Summary
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Deleted: $deletedCount files ($(Format-SyncSize $deletedBytes))" -ForegroundColor Green
    if ($failedCount -gt 0) {
        Write-Host "  Failed:  $failedCount files" -ForegroundColor Red
    }
    if ($foldersRemoved -gt 0) {
        Write-Host "  Cleaned: $foldersRemoved empty folder(s) removed" -ForegroundColor Green
    }
    Write-Host ""

    return @{
        Deleted = $deletedCount
        Failed = $failedCount
        Skipped = $notImported.Count
        BytesDeleted = $deletedBytes
        FoldersRemoved = $foldersRemoved
    }
}

<#
.SYNOPSIS
    Initializes SFTP tracking by marking all existing remote files as already downloaded
.DESCRIPTION
    Scans the remote server and adds all files to the tracking file without downloading.
    Use this on first run to prevent downloading everything that's already on the server.
.PARAMETER HostName
    SFTP server hostname
.PARAMETER Port
    SFTP server port (default: 22)
.PARAMETER Username
    SFTP username
.PARAMETER Password
    SFTP password (use this OR PrivateKeyPath)
.PARAMETER PrivateKeyPath
    Path to SSH private key (use this OR Password)
.PARAMETER RemotePaths
    Array of remote paths to scan
.PARAMETER TrackingFile
    Path to tracking file (default: AppData\LibraryLint\sftp_downloaded.json)
.EXAMPLE
    Initialize-SFTPTracking -HostName "server.com" -Username "user" -Password "pass" -RemotePaths @("/downloads")
#>
function Initialize-SFTPTracking {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HostName,

        [int]$Port = 22,

        [Parameter(Mandatory=$true)]
        [string]$Username,

        [string]$Password,
        [string]$PrivateKeyPath,

        [Parameter(Mandatory=$true)]
        [string[]]$RemotePaths,

        [string]$TrackingFile
    )

    # Header
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "            SFTP TRACKING INITIALIZATION               " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  This will mark all existing remote files as 'already downloaded'" -ForegroundColor Yellow
    Write-Host "  so they won't be downloaded on subsequent syncs." -ForegroundColor Yellow
    Write-Host ""

    # Check WinSCP .NET assembly
    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        return @{ Initialized = 0; Error = "WinSCP .NET assembly not installed" }
    }

    # Tracking file path
    $trackingPath = Get-SyncTrackingPath -ConfigTrackingFile $TrackingFile

    # Check if tracking file already exists
    if (Test-Path $trackingPath) {
        $existing = Get-DownloadedFiles -TrackingPath $trackingPath
        if ($existing.Count -gt 0) {
            Write-Host "  Tracking file already exists with $($existing.Count) entries." -ForegroundColor Yellow
            Write-Host "  New files will be added to existing tracking data." -ForegroundColor Gray
            Write-Host ""
        }
    }

    Write-Host "  Host: ${HostName}:${Port}" -ForegroundColor Gray
    Write-Host "  User: $Username" -ForegroundColor Gray
    Write-Host ""

    # Connect
    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ Initialized = 0; Error = $_.ToString() }
    }

    try {
        # Scan remote paths
        Write-Host ""
        Write-Host "  Scanning remote paths..." -ForegroundColor Gray

        $allFiles = @()
        foreach ($remotePath in $RemotePaths) {
            $files = Get-RemoteFilesRecursive -Session $session -RemotePath $remotePath
            # Clear the scanning progress line before printing the summary
            $clearWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
            if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
            Write-Host "    $remotePath ($($files.Count) files)" -ForegroundColor Gray
            $allFiles += $files
        }

        # Extract unique folder names
        $folders = $allFiles | ForEach-Object {
            $relativePath = $_.FullPath
            foreach ($rp in $RemotePaths) {
                if ($relativePath.StartsWith($rp)) {
                    $relativePath = $relativePath.Substring($rp.Length).TrimStart('/')
                    break
                }
            }
            $parts = $relativePath -split '/'
            if ($parts.Count -gt 1) { $parts[0] } else { $null }
        } | Where-Object { $_ } | Select-Object -Unique | Sort-Object

        # Show folders found
        if ($folders.Count -gt 0) {
            Write-Host ""
            Write-Host "  Folders found:" -ForegroundColor Cyan
            foreach ($folder in $folders) {
                $folderFiles = $allFiles | Where-Object { $_.FullPath -like "*/$folder/*" -or $_.FullPath -like "*/$folder" }
                $folderSize = ($folderFiles | Measure-Object -Property Size -Sum).Sum
                Write-Host "    - $folder " -NoNewline -ForegroundColor White
                Write-Host "($(Format-SyncSize $folderSize))" -ForegroundColor Gray
            }
        }

        $totalSize = ($allFiles | Measure-Object -Property Size -Sum).Sum
        Write-Host ""
        Write-Host "  Total: $($allFiles.Count) files ($(Format-SyncSize $totalSize))" -ForegroundColor White
        Write-Host ""

        # Get existing tracking data
        $downloaded = Get-DownloadedFiles -TrackingPath $trackingPath
        $newCount = 0

        # Add all files to tracking using the file's actual modification date
        Write-Host "  Adding to tracking file..." -ForegroundColor Gray
        foreach ($file in $allFiles) {
            if (-not $downloaded.ContainsKey($file.FullPath)) {
                # Use the file's LastModified date so prune can work on old files
                $fileDate = if ($file.LastModified) { $file.LastModified.ToString("o") } else { (Get-Date).ToString("o") }
                $downloaded[$file.FullPath] = @{
                    LocalPath = "[initialized - not downloaded]"
                    Size = $file.Size
                    DownloadedAt = $fileDate
                    Initialized = $true
                }
                $newCount++
            }
        }

        # Save tracking file
        Save-DownloadedFiles -Downloaded $downloaded -TrackingPath $trackingPath

        Write-Host ""
        Write-Host "======================================================" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Initialized: $newCount new files added to tracking" -ForegroundColor Green
        Write-Host "  Total tracked: $($downloaded.Count) files" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Future syncs will only download NEW files added after this point." -ForegroundColor Yellow
        Write-Host ""

        return @{
            Initialized = $newCount
            TotalTracked = $downloaded.Count
        }

    } finally {
        $session.Dispose()
    }
}

<#
.SYNOPSIS
    Reconciles the SFTP tracking file with what's already in the local
    library / inbox by matching remote files against local files and
    folders.
.DESCRIPTION
    The complement to Initialize-SFTPTracking. Initialize marks every file
    currently visible on the seedbox as "already downloaded", which works
    for the cold-start case but not when the seedbox already has new
    arrivals you don't want to re-fetch (because you have them locally
    from earlier).

    This function walks the seedbox AND the local roots, then for each
    remote file checks whether the user already has it locally — by
    Name|Size match (file index across inbox + library), or by release
    folder name (folder set across library only). Matches get added to
    the tracking file as ManualTransfer with a MatchType field, so future
    syncs treat them as already-have without downloading.

    Doesn't touch local files, doesn't touch remote files. Tracking-only.
.PARAMETER HostName
    SFTP server hostname
.PARAMETER Port
    SFTP server port (default: 22)
.PARAMETER Username
    SFTP username
.PARAMETER Password
    SFTP password (use this OR PrivateKeyPath)
.PARAMETER PrivateKeyPath
    Path to SSH private key (use this OR Password)
.PARAMETER RemotePaths
    Array of remote paths to scan
.PARAMETER LocalBasePath
    Inbox base (the parent of _Movies, _Shows, etc.)
.PARAMETER LibraryPaths
    Long-term library roots (MoviesLibraryPath, TVShowsLibraryPath, etc.)
.PARAMETER TrackingFile
    Path to tracking file (default: AppData\LibraryLint\sftp_downloaded.json)
#>
function Update-SFTPTrackingFromLocal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HostName,

        [int]$Port = 22,

        [Parameter(Mandatory=$true)]
        [string]$Username,

        [string]$Password,
        [string]$PrivateKeyPath,

        [Parameter(Mandatory=$true)]
        [string[]]$RemotePaths,

        [string]$LocalBasePath,
        [string[]]$LibraryPaths = @(),

        [string]$TrackingFile
    )

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "       RECONCILE TRACKING WITH LOCAL LIBRARY           " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        return @{ MatchedFiles = 0; Error = "WinSCP .NET assembly not installed" }
    }

    $trackingPath = Get-SyncTrackingPath -ConfigTrackingFile $TrackingFile

    Write-Host "  Host: ${HostName}:${Port}" -ForegroundColor Gray
    Write-Host "  User: $Username" -ForegroundColor Gray
    if ($LibraryPaths.Count -gt 0) {
        Write-Host "  Library roots:" -ForegroundColor Gray
        foreach ($lp in $LibraryPaths) { Write-Host "    - $lp" -ForegroundColor DarkGray }
    }
    Write-Host ""

    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ MatchedFiles = 0; Error = $_.ToString() }
    }

    try {
        Write-Host ""
        Write-Host "  Scanning remote paths..." -ForegroundColor Gray
        $allFiles = @()
        foreach ($remotePath in $RemotePaths) {
            $files = Get-RemoteFilesRecursive -Session $session -RemotePath $remotePath
            $clearWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
            if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
            Write-Host "    $remotePath ($($files.Count) files)" -ForegroundColor Gray
            $allFiles += $files
        }

        $inboxRoots = if ($LocalBasePath) {
            @(
                Join-Path $LocalBasePath "_Movies"
                Join-Path $LocalBasePath "_Shows"
                Join-Path $LocalBasePath "_Music"
                Join-Path $LocalBasePath "_Books"
                Join-Path $LocalBasePath "_Downloads"
            )
        } else { @() }
        $indexRoots = @($inboxRoots) + @($LibraryPaths)

        Write-Host ""
        Write-Host "  Indexing local files..." -ForegroundColor Gray -NoNewline
        $localIndex = Build-LocalFileIndex -RootPaths $indexRoots
        $localFolders = Build-LocalFolderSet -RootPaths $LibraryPaths
        Write-Host " $($localIndex.Count) files, $($localFolders.Count) library folders" -ForegroundColor Gray
        Write-Host ""

        $downloaded = Get-DownloadedFiles -TrackingPath $trackingPath
        $matchedNameSize = 0
        $matchedFolder = 0
        $alreadyTracked = 0
        $noMatch = 0
        # Per-folder counts for reporting
        $byFolder = @{}

        foreach ($file in $allFiles) {
            if ($downloaded.ContainsKey($file.FullPath)) {
                $alreadyTracked++
                continue
            }

            $remoteParent = Split-Path $file.FullPath -Parent
            $remoteParentName = if ($remoteParent) { Split-Path $remoteParent -Leaf } else { $null }
            $haveCheck = Test-RemoteFileAlreadyHave -FileName $file.Name -FileSize $file.Size `
                -RemoteParentName $remoteParentName -FileIndex $localIndex -FolderSet $localFolders

            if ($haveCheck.Found) {
                $downloaded[$file.FullPath] = @{
                    LocalPath = $haveCheck.LocalPath
                    Size = $file.Size
                    DownloadedAt = (Get-Date).ToString("o")
                    ManualTransfer = $true
                    MatchType = $haveCheck.MatchType
                    Reconciled = $true
                }
                if ($haveCheck.MatchType -eq 'NameSize') { $matchedNameSize++ } else { $matchedFolder++ }

                if ($remoteParentName) {
                    if (-not $byFolder.ContainsKey($remoteParentName)) {
                        $byFolder[$remoteParentName] = @{ Files = 0; MatchType = $haveCheck.MatchType }
                    }
                    $byFolder[$remoteParentName].Files++
                }
            } else {
                $noMatch++
            }
        }

        Save-DownloadedFiles -Downloaded $downloaded -TrackingPath $trackingPath

        Write-Host ""
        Write-Host "======================================================" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Already tracked:        $alreadyTracked" -ForegroundColor Gray
        Write-Host "  Matched (name+size):    $matchedNameSize" -ForegroundColor Green
        Write-Host "  Matched (folder name):  $matchedFolder" -ForegroundColor Green
        Write-Host "  Still considered new:   $noMatch" -ForegroundColor $(if ($noMatch -gt 0) { 'Yellow' } else { 'Gray' })
        Write-Host ""

        if ($byFolder.Count -gt 0 -and ($matchedNameSize + $matchedFolder) -gt 0) {
            Write-Host "  Folders matched (top 30):" -ForegroundColor Cyan
            $byFolder.GetEnumerator() | Sort-Object Name | Select-Object -First 30 | ForEach-Object {
                $tag = if ($_.Value.MatchType -eq 'Folder') { 'folder' } else { 'name+size' }
                Write-Host "    - $($_.Key) ($($_.Value.Files) files, matched by $tag)" -ForegroundColor DarkGray
            }
            if ($byFolder.Count -gt 30) {
                Write-Host "    ... and $($byFolder.Count - 30) more" -ForegroundColor DarkGray
            }
            Write-Host ""
        }

        return @{
            MatchedFiles      = $matchedNameSize + $matchedFolder
            MatchedNameSize   = $matchedNameSize
            MatchedFolder     = $matchedFolder
            AlreadyTracked    = $alreadyTracked
            StillNew          = $noMatch
            FoldersMatched    = $byFolder.Count
        }

    } finally {
        $session.Dispose()
    }
}

<#
.SYNOPSIS
    Checks an SFTP server for new files without downloading
.DESCRIPTION
    Connects to an SFTP server, scans remote paths, and reports what files are
    new (not yet tracked). Useful for a quick "anything new?" check before
    committing to a full sync.
.PARAMETER HostName
    SFTP server hostname
.PARAMETER Port
    SFTP server port (default: 22)
.PARAMETER Username
    SFTP username
.PARAMETER Password
    SFTP password (use this OR PrivateKeyPath)
.PARAMETER PrivateKeyPath
    Path to SSH private key (use this OR Password)
.PARAMETER RemotePaths
    Array of remote paths to scan for files
.PARAMETER TrackingFile
    Path to tracking file (default: AppData\LibraryLint\sftp_downloaded.json)
.EXAMPLE
    Get-SFTPNewFiles -HostName "server.com" -Username "user" -Password "pass" -RemotePaths @("/downloads")
#>
function Get-SFTPNewFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HostName,

        [int]$Port = 22,

        [Parameter(Mandatory=$true)]
        [string]$Username,

        [string]$Password,
        [string]$PrivateKeyPath,

        [Parameter(Mandatory=$true)]
        [string[]]$RemotePaths,

        [string]$TrackingFile
    )

    # Header
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "              CHECK FOR NEW FILES                      " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    # Check WinSCP .NET assembly
    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        return @{ NewFiles = 0; NewFolders = @(); TotalSize = 0; Error = "WinSCP .NET assembly not installed" }
    }

    # Tracking file path
    $trackingPath = Get-SyncTrackingPath -ConfigTrackingFile $TrackingFile

    Write-Host "  Host: ${HostName}:${Port}" -ForegroundColor Gray
    Write-Host "  User: $Username" -ForegroundColor Gray
    Write-Host ""

    # Connect
    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ NewFiles = 0; NewFolders = @(); TotalSize = 0; Error = $_.ToString() }
    }

    try {
        # Scan remote paths
        Write-Host ""
        Write-Host "  Scanning remote paths..." -ForegroundColor Gray

        $allFiles = @()
        foreach ($remotePath in $RemotePaths) {
            $files = Get-RemoteFilesRecursive -Session $session -RemotePath $remotePath
            # Clear the scanning progress line before printing the summary
            $clearWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
            if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
            Write-Host "    $remotePath ($($files.Count) files)" -ForegroundColor Gray
            $allFiles += $files
        }

        # Get tracking data
        $downloaded = Get-DownloadedFiles -TrackingPath $trackingPath
        $isFirstSync = ($downloaded.Count -eq 0)

        # Filter to new files only
        $newFiles = $allFiles | Where-Object { -not $downloaded.ContainsKey($_.FullPath) }
        $trackedCount = $allFiles.Count - $newFiles.Count

        if ($isFirstSync) {
            Write-Host ""
            Write-Host "  First sync - no download history found." -ForegroundColor Yellow
            Write-Host "  Showing all remote content. Files will be tracked after download." -ForegroundColor Yellow
        }

        Write-Host ""
        Write-Host "  Total on server: $($allFiles.Count) files" -ForegroundColor White
        if (-not $isFirstSync) {
            Write-Host "  Already synced:  $trackedCount" -ForegroundColor Gray

            # Breakdown by download date
            if ($trackedCount -gt 0) {
                $syncedFiles = $allFiles | Where-Object { $downloaded.ContainsKey($_.FullPath) }
                $byDate = @{}
                foreach ($file in $syncedFiles) {
                    $entry = $downloaded[$file.FullPath]
                    try {
                        $date = [DateTime]::Parse($entry.DownloadedAt).ToString("yyyy-MM-dd")
                    } catch {
                        $date = "Unknown"
                    }
                    if (-not $byDate.ContainsKey($date)) {
                        $byDate[$date] = @{ Files = 0; Size = [long]0 }
                    }
                    $byDate[$date].Files++
                    $byDate[$date].Size += $file.Size
                }
                foreach ($day in $byDate.GetEnumerator() | Sort-Object Name -Descending) {
                    Write-Host "    $($day.Key): $($day.Value.Files) files, $(Format-SyncSize $day.Value.Size)" -ForegroundColor DarkGray
                }
            }
        }
        Write-Host "  New files:       $($newFiles.Count)" -ForegroundColor $(if ($newFiles.Count -gt 0) { 'Green' } else { 'Gray' })

        if ($newFiles.Count -eq 0) {
            Write-Host ""
            Write-Host "  Nothing new on the server." -ForegroundColor Green
            Write-Host ""
            return @{ NewFiles = 0; NewFolders = @(); TotalSize = 0 }
        }

        $totalSize = ($newFiles | Measure-Object -Property Size -Sum).Sum

        # Group files by movie folder
        $byFolder = @{}
        foreach ($file in $newFiles) {
            $relativePath = $file.FullPath
            foreach ($rp in $RemotePaths) {
                if ($relativePath.StartsWith($rp)) {
                    $relativePath = $relativePath.Substring($rp.Length).TrimStart('/')
                    break
                }
            }
            $parts = $relativePath -split '/'
            $folderName = if ($parts.Count -gt 1) { $parts[0] } else { "(root)" }
            if (-not $byFolder.ContainsKey($folderName)) {
                $byFolder[$folderName] = @{ Files = 0; Size = [long]0 }
            }
            $byFolder[$folderName].Files++
            $byFolder[$folderName].Size += $file.Size
        }

        Write-Host ""
        $contentLabel = if ($isFirstSync) { "Remote content:" } else { "New content:" }
        Write-Host "  $contentLabel" -ForegroundColor Cyan
        $sortedFolders = @($byFolder.GetEnumerator() | Sort-Object { $_.Value.Size } -Descending)
        foreach ($folder in $sortedFolders) {
            $truncated = if ($folder.Key.Length -gt 55) { $folder.Key.Substring(0, 52) + "..." } else { $folder.Key }
            Write-Host "    $truncated " -NoNewline -ForegroundColor White
            Write-Host "($($folder.Value.Files) files, $(Format-SyncSize $folder.Value.Size))" -ForegroundColor Gray
        }

        Write-Host ""
        $totalLabel = if ($isFirstSync) { "Total on server" } else { "Total new" }
        Write-Host "  ${totalLabel}: $($sortedFolders.Count) movies, $(Format-SyncSize $totalSize)" -ForegroundColor Cyan
        Write-Host ""

        return @{
            NewFiles = $newFiles.Count
            NewFolders = @()
            TotalSize = $totalSize
        }

    } finally {
        $session.Dispose()
    }
}

<#
.SYNOPSIS
    Scans remote server for incomplete downloads
.DESCRIPTION
    Connects to the SFTP server and looks for signs of incomplete downloads:
    - Files with .part, .!qB, .!ut, .partial, .downloading extensions
    - Zero-byte video files
    - Suspiciously small video files (e.g. a 1080p movie under 50MB)
    Reports findings so the user can decide whether to wait, re-download, or clean up.
#>
function Find-SFTPIncompleteFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HostName,

        [int]$Port = 22,

        [Parameter(Mandatory=$true)]
        [string]$Username,

        [string]$Password,
        [string]$PrivateKeyPath,

        [Parameter(Mandatory=$true)]
        [string[]]$RemotePaths,

        [string[]]$VideoExtensions = @(".mkv", ".mp4", ".avi", ".m4v", ".wmv", ".ts", ".mpg", ".mpeg"),
        [long]$SuspiciousMaxBytes = 50MB
    )

    # Header
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "           FIND INCOMPLETE DOWNLOADS                   " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    # Check WinSCP .NET assembly
    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        return @{ Incomplete = @(); Error = "WinSCP .NET assembly not installed" }
    }

    Write-Host "  Host: ${HostName}:${Port}" -ForegroundColor Gray
    Write-Host "  User: $Username" -ForegroundColor Gray
    Write-Host ""

    # Connect
    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ Incomplete = @(); Error = $_.ToString() }
    }

    try {
        # Scan remote paths
        Write-Host ""
        Write-Host "  Scanning remote paths..." -ForegroundColor Gray

        $allFiles = @()
        foreach ($remotePath in $RemotePaths) {
            $files = Get-RemoteFilesRecursive -Session $session -RemotePath $remotePath
            # Clear the scanning progress line before printing the summary
            $clearWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
            if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
            Write-Host "    $remotePath ($($files.Count) files)" -ForegroundColor Gray
            $allFiles += $files
        }

        # Partial download extensions
        $partialExtensions = @(".part", ".!qb", ".!ut", ".partial", ".downloading", ".crdownload", ".tmp")

        $incomplete = @()

        foreach ($file in $allFiles) {
            $ext = [System.IO.Path]::GetExtension($file.Name).ToLower()
            $reason = $null

            # Check 1: Known partial-download extensions
            if ($partialExtensions -contains $ext) {
                $reason = "Partial download ($ext)"
            }
            # Check 2: File has a video extension tacked on before a partial ext
            # e.g. "movie.mkv.part" - the base name itself has a video ext
            elseif ($ext -and $partialExtensions -contains $ext) {
                $reason = "Partial download ($ext)"
            }

            # Skip non-media files that are expected to be small or empty
            $ignoredExtensions = @(".nfo", ".txt", ".srt", ".sub", ".idx", ".ass", ".ssa", ".jpg", ".jpeg", ".png")
            $isIgnored = $ignoredExtensions -contains $ext

            # Skip trailers — filenames commonly contain [TRAILER], TRAILER-, or "Sample"
            $isTrailer = $file.Name -match '(?i)\btrailer\b|(?i)\bsample\b'

            # Check 3: Zero-byte files (skip non-media metadata files)
            if (-not $reason -and $file.Size -eq 0 -and -not $isIgnored) {
                $reason = "Zero-byte file"
            }

            # Check 4: Suspiciously small video files (skip trailers and samples)
            if (-not $reason -and -not $isTrailer) {
                $baseExt = $ext
                # For partial files, check the underlying extension
                if ($partialExtensions -contains $ext) {
                    $baseExt = [System.IO.Path]::GetExtension(
                        [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                    ).ToLower()
                }
                if ($VideoExtensions -contains $baseExt -and $file.Size -gt 0 -and $file.Size -lt $SuspiciousMaxBytes) {
                    $reason = "Suspiciously small video ($(Format-SyncSize $file.Size))"
                }
            }

            if ($reason) {
                $incomplete += [PSCustomObject]@{
                    FullPath = $file.FullPath
                    Name     = $file.Name
                    Size     = $file.Size
                    Reason   = $reason
                }
            }
        }

        # Display results
        Write-Host ""
        Write-Host "  Total files scanned: $($allFiles.Count)" -ForegroundColor White

        if ($incomplete.Count -eq 0) {
            Write-Host ""
            Write-Host "  No incomplete downloads found." -ForegroundColor Green
            Write-Host ""
        } else {
            Write-Host "  Potentially incomplete: $($incomplete.Count)" -ForegroundColor Yellow
            Write-Host ""

            # Group by reason for cleaner output
            $grouped = $incomplete | Group-Object Reason

            foreach ($group in $grouped) {
                Write-Host "  $($group.Name) ($($group.Count)):" -ForegroundColor Yellow
                foreach ($item in $group.Group) {
                    $sizeStr = if ($item.Size -gt 0) { Format-SyncSize $item.Size } else { "0 bytes" }
                    Write-Host "    - $($item.Name) " -NoNewline -ForegroundColor White
                    Write-Host "($sizeStr)" -ForegroundColor Gray

                    # Show path relative to remote root for context
                    $relPath = $item.FullPath
                    foreach ($rp in $RemotePaths) {
                        if ($relPath.StartsWith($rp)) {
                            $relPath = $relPath.Substring($rp.Length).TrimStart('/')
                            break
                        }
                    }
                    $parentFolder = ($relPath -split '/')[0]
                    if ($parentFolder -ne $item.Name) {
                        Write-Host "      in: $parentFolder" -ForegroundColor DarkGray
                    }
                }
                Write-Host ""
            }
        }

        return @{
            Incomplete = $incomplete
            TotalScanned = $allFiles.Count
        }

    } finally {
        $session.Dispose()
    }
}

#endregion

# Export public functions
Export-ModuleMember -Function Invoke-SFTPSync, Invoke-SFTPPrune, Initialize-SFTPTracking, Update-SFTPTrackingFromLocal, Get-SFTPNewFiles, Find-SFTPIncompleteFiles, Test-WinSCPInstalled, Connect-SFTPSession, Get-RemoteFilesRecursive, Invoke-FileDownload, Get-DownloadedFiles, Save-DownloadedFiles, Get-SyncTrackingPath, Format-SyncSize
