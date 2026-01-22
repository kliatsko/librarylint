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

function Test-WinSCPInstalled {
    param([string]$ModulePath)

    # Order matters: prioritize netstandard2.0 versions for .NET Core/5+ compatibility
    $winscpPaths = @(
        # LibraryLint folder (preferred - user should place netstandard2.0 version here)
        "$env:LOCALAPPDATA\LibraryLint\WinSCPnet.dll",
        # NuGet netstandard2.0 (works with .NET Core/5+/PowerShell 7)
        "$env:USERPROFILE\.nuget\packages\winscp\*\lib\netstandard2.0\WinSCPnet.dll",
        # Downloads - netstandard2.0 subfolder
        "$env:USERPROFILE\Downloads\WinSCP-*-Automation\netstandard2.0\WinSCPnet.dll",
        "$env:USERPROFILE\Downloads\winscp_nuget\lib\netstandard2.0\WinSCPnet.dll",
        # Standard install locations
        "${env:ProgramFiles}\WinSCP\WinSCPnet.dll",
        "${env:ProgramFiles(x86)}\WinSCP\WinSCPnet.dll",
        # User profile locations
        "$env:LOCALAPPDATA\Programs\WinSCP\WinSCPnet.dll",
        "$env:USERPROFILE\WinSCP\WinSCPnet.dll"
        # Note: Removed net40 paths as they don't work with PowerShell 7/.NET 5+
    )

    if ($ModulePath) {
        $winscpPaths = @(Join-Path $ModulePath "WinSCPnet.dll") + $winscpPaths
    }

    # Expand wildcards and check paths
    foreach ($pattern in $winscpPaths) {
        $resolved = Get-Item $pattern -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($resolved) {
            return $resolved.FullName
        }
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

    Add-Type -Path $DllPath

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

    # Set executable path - look for WinSCP.exe in same folder as DLL or common locations
    $dllFolder = Split-Path $DllPath -Parent
    $exePaths = @(
        (Join-Path $dllFolder "WinSCP.exe"),
        "${env:ProgramFiles}\WinSCP\WinSCP.exe",
        "${env:ProgramFiles(x86)}\WinSCP\WinSCP.exe",
        "$env:LOCALAPPDATA\Programs\WinSCP\WinSCP.exe"
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
                        Write-Host "      Scanning: $($fileInfo.Name)" -ForegroundColor DarkGray
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
    $parentFolder = if ($pathParts.Count -gt 1) { $pathParts[0] } else { $null }

    # Companion files - these follow the video file's category, not their own
    # NFO, subtitles, images, etc. should stay with their video
    $companionExtensions = @(".nfo", ".srt", ".sub", ".idx", ".ass", ".ssa", ".jpg", ".jpeg", ".png", ".txt")

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

function Test-FileExistsInInbox {
    <#
    .SYNOPSIS
        Checks if a file with the same name and size already exists in the inbox folders
    .DESCRIPTION
        Searches _Movies, _Shows, and _Downloads folders recursively for a matching file.
        Used to detect manual transfers and avoid re-downloading.
    #>
    param(
        [string]$FileName,
        [long]$FileSize,
        [string]$LocalBasePath
    )

    $inboxFolders = @(
        Join-Path $LocalBasePath "_Movies"
        Join-Path $LocalBasePath "_Shows"
        Join-Path $LocalBasePath "_Music"
        Join-Path $LocalBasePath "_Books"
        Join-Path $LocalBasePath "_Downloads"
    )

    foreach ($folder in $inboxFolders) {
        if (Test-Path $folder) {
            $existing = Get-ChildItem -Path $folder -Recurse -File -Filter $FileName -ErrorAction SilentlyContinue |
                Where-Object { $_.Length -eq $FileSize } |
                Select-Object -First 1

            if ($existing) {
                return $existing.FullName
            }
        }
    }

    return $null
}

function Invoke-FileDownload {
    param(
        $Session,
        [string]$RemotePath,
        [string]$LocalPath
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

    $result = $Session.GetFiles($RemotePath, $LocalPath, $false, $transferOptions)
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
.PARAMETER WhatIf
    Preview without downloading (shows what would be downloaded and where)
.PARAMETER Force
    Re-download files even if already tracked
.EXAMPLE
    Invoke-SFTPSync -HostName "server.com" -Username "user" -Password "pass" -RemotePaths @("/downloads") -LocalBasePath "G:"
#>
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

        [string]$TrackingFile,

        [string[]]$MovieExtensions = @(".mkv", ".mp4", ".avi", ".m4v"),
        [double]$MovieMinSizeGB = 1,

        [switch]$DeleteAfterDownload,
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
    }

    try {
        # Scan remote paths
        Write-Host ""
        Write-Host "  Scanning remote paths..." -ForegroundColor Gray

        $allFiles = @()
        foreach ($remotePath in $RemotePaths) {
            Write-Host "    $remotePath" -ForegroundColor Gray -NoNewline
            $files = Get-RemoteFilesRecursive -Session $session -RemotePath $remotePath
            Write-Host " ($($files.Count) files)" -ForegroundColor Gray
            $allFiles += $files
        }

        # Get tracking data
        $downloaded = Get-DownloadedFiles -TrackingPath $trackingPath

        # Filter out already downloaded
        $newFiles = $allFiles | Where-Object {
            $Force -or (-not $downloaded.ContainsKey($_.FullPath))
        }
        $skippedCount = $allFiles.Count - $newFiles.Count

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
        Write-Host "  Total files:    $($allFiles.Count)" -ForegroundColor White
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

            # Check if file already exists locally (manual transfer detection)
            $existingFile = Test-FileExistsInInbox -FileName $file.Name -FileSize $file.Size -LocalBasePath $LocalBasePath
            if ($existingFile) {
                Write-Host " SKIP (already exists)" -ForegroundColor Cyan
                # Track it so we don't check again next time
                $downloaded[$file.FullPath] = @{
                    LocalPath = $existingFile
                    Size = $file.Size
                    DownloadedAt = (Get-Date).ToString("o")
                    ManualTransfer = $true
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

                    $success = Invoke-FileDownload -Session $session -RemotePath $file.FullPath -LocalPath $localPath

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
                            $session.RemoveFiles($file.FullPath).Check()
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
        Write-Host "  Time:       $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
        Write-Host ""

        $result.Downloaded = $downloadCount
        $result.Failed = $failCount
        $result.SkippedDuplicates = $skippedDupes
        $result.BytesDownloaded = $downloadedBytes
        $result.Duration = $stopwatch.Elapsed

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
            $filesToPrune += @{
                RemotePath = $entry.Key
                LocalPath = $entry.Value.LocalPath
                Size = $entry.Value.Size
                DownloadedAt = $downloadedAt
                Age = [math]::Floor(((Get-Date) - $downloadedAt).TotalDays)
            }
        }
    }

    if ($filesToPrune.Count -eq 0) {
        Write-Host "  No files older than $DaysOld days" -ForegroundColor Green
        return @{ Deleted = 0; Failed = 0 }
    }

    # Show files to prune
    $totalSize = ($filesToPrune | Measure-Object -Property Size -Sum).Sum
    Write-Host "  Found $($filesToPrune.Count) files to prune ($(Format-SyncSize $totalSize))" -ForegroundColor Cyan
    Write-Host ""

    foreach ($file in $filesToPrune | Sort-Object Age -Descending) {
        $fileName = Split-Path $file.RemotePath -Leaf
        $truncatedName = if ($fileName.Length -gt 50) { $fileName.Substring(0, 47) + "..." } else { $fileName }
        Write-Host "    $($file.Age) days old: " -NoNewline -ForegroundColor Gray
        Write-Host $truncatedName -ForegroundColor White
    }
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] Would delete $($filesToPrune.Count) files" -ForegroundColor Yellow
        return @{ Deleted = $filesToPrune.Count; Failed = 0; WhatIf = $true }
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

    try {
        Write-Host ""
        Write-Host "  Deleting files..." -ForegroundColor Yellow
        Write-Host ""

        foreach ($file in $filesToPrune) {
            $fileName = Split-Path $file.RemotePath -Leaf
            $truncatedName = if ($fileName.Length -gt 45) { $fileName.Substring(0, 42) + "..." } else { $fileName }

            Write-Host "    $truncatedName" -NoNewline -ForegroundColor White

            try {
                # Check if file still exists on remote
                if ($session.FileExists($file.RemotePath)) {
                    $session.RemoveFiles($file.RemotePath).Check()
                    Write-Host " deleted" -ForegroundColor Green
                    $deletedCount++
                    $deletedBytes += $file.Size

                    # Remove from tracking
                    $downloaded.Remove($file.RemotePath)
                } else {
                    Write-Host " not found (already deleted?)" -ForegroundColor Gray
                    # Still remove from tracking since it's gone
                    $downloaded.Remove($file.RemotePath)
                }
            } catch {
                Write-Host " FAILED: $_" -ForegroundColor Red
                $failedCount++
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
    Write-Host ""

    return @{
        Deleted = $deletedCount
        Failed = $failedCount
        BytesDeleted = $deletedBytes
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
            Write-Host "    $remotePath" -ForegroundColor Gray -NoNewline
            $files = Get-RemoteFilesRecursive -Session $session -RemotePath $remotePath
            Write-Host " ($($files.Count) files)" -ForegroundColor Gray
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

#endregion

# Export public functions
Export-ModuleMember -Function Invoke-SFTPSync, Invoke-SFTPPrune, Initialize-SFTPTracking
