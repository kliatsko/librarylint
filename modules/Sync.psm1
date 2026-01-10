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

    $winscpPaths = @(
        "${env:ProgramFiles}\WinSCP\WinSCPnet.dll",
        "${env:ProgramFiles(x86)}\WinSCP\WinSCPnet.dll"
    )

    if ($ModulePath) {
        $winscpPaths = @(Join-Path $ModulePath "WinSCPnet.dll") + $winscpPaths
    }

    $winscpPath = $winscpPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($winscpPath) {
        return $winscpPath
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
    $session.Open($sessionOptions)

    return $session
}

function Get-RemoteFilesRecursive {
    param(
        $Session,
        [string]$RemotePath
    )

    $files = @()

    try {
        $directoryInfo = $Session.ListDirectory($RemotePath)

        foreach ($fileInfo in $directoryInfo.Files) {
            if ($fileInfo.IsDirectory) {
                if ($fileInfo.Name -ne "." -and $fileInfo.Name -ne "..") {
                    # Recurse into subdirectories
                    $subPath = "$RemotePath/$($fileInfo.Name)"
                    $files += Get-RemoteFilesRecursive -Session $Session -RemotePath $subPath
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
        [long]$FileSize,
        [string]$LocalBasePath,
        [string[]]$MovieExtensions,
        [double]$MovieMinSizeGB
    )

    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
    $sizeGB = $FileSize / 1GB

    # Large video files are likely movies
    if ($MovieExtensions -contains $ext -and $sizeGB -ge $MovieMinSizeGB) {
        return Join-Path $LocalBasePath "Inbox\_Movies"
    }

    # TV show detection based on naming patterns (S01E01, etc.)
    if ($FileName -match 'S\d{2}E\d{2}|Season\s*\d+') {
        return Join-Path $LocalBasePath "Inbox\_Shows"
    }

    # Default to a Downloads folder
    return Join-Path $LocalBasePath "Inbox\_Downloads"
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

    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

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
    Preview without downloading
.PARAMETER Force
    Re-download files even if already tracked
.PARAMETER ListOnly
    Just list new files without downloading
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
        [switch]$Force,
        [switch]$ListOnly
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

    if ($ListOnly) {
        Write-Host "  [LIST MODE] Just showing new files" -ForegroundColor Yellow
        Write-Host ""
    }

    # Check WinSCP
    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP not found!" -ForegroundColor Red
        Write-Host "  Install with: winget install WinSCP" -ForegroundColor Yellow
        return @{ Downloaded = 0; Failed = 0; Error = "WinSCP not installed" }
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

        Write-Host ""
        Write-Host "  Total files:    $($allFiles.Count)" -ForegroundColor White
        Write-Host "  Already synced: $skippedCount" -ForegroundColor Gray
        Write-Host "  New files:      $($newFiles.Count)" -ForegroundColor $(if ($newFiles.Count -gt 0) { 'Green' } else { 'Gray' })

        if ($newFiles.Count -eq 0) {
            Write-Host ""
            Write-Host "  Nothing new to download!" -ForegroundColor Green
            return $result
        }

        $totalSize = ($newFiles | Measure-Object -Property Size -Sum).Sum
        Write-Host "  Total size:     $(Format-SyncSize $totalSize)" -ForegroundColor Cyan
        Write-Host ""

        # List mode - just show files
        if ($ListOnly) {
            Write-Host "  New files:" -ForegroundColor Yellow
            foreach ($file in $newFiles | Sort-Object Size -Descending) {
                Write-Host "    $(Format-SyncSize $file.Size)".PadRight(12) -ForegroundColor Gray -NoNewline
                Write-Host $file.Name -ForegroundColor White
            }
            return $result
        }

        # Download files
        Write-Host "  Downloading..." -ForegroundColor Yellow
        Write-Host ""

        $downloadCount = 0
        $failCount = 0
        $downloadedBytes = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        foreach ($file in $newFiles) {
            $destFolder = Get-SyncDestinationFolder -FileName $file.Name -FileSize $file.Size `
                -LocalBasePath $LocalBasePath -MovieExtensions $MovieExtensions -MovieMinSizeGB $MovieMinSizeGB
            $localPath = Join-Path $destFolder $file.Name

            $truncatedName = if ($file.Name.Length -gt 45) { $file.Name.Substring(0, 42) + "..." } else { $file.Name }
            $progress = "[$($downloadCount + $failCount + 1)/$($newFiles.Count)]"

            Write-Host "  $progress $truncatedName" -ForegroundColor White -NoNewline
            Write-Host " ($(Format-SyncSize $file.Size))" -ForegroundColor Gray -NoNewline

            if ($WhatIf) {
                Write-Host " [would download to $destFolder]" -ForegroundColor Yellow
                $downloadCount++
                continue
            }

            try {
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
                } else {
                    Write-Host " FAILED" -ForegroundColor Red
                    $failCount++
                }
            } catch {
                Write-Host " ERROR: $_" -ForegroundColor Red
                $failCount++
            }
        }

        $stopwatch.Stop()

        # Summary
        Write-Host ""
        Write-Host "======================================================" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Downloaded: $downloadCount files ($(Format-SyncSize $downloadedBytes))" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "  Failed:     $failCount files" -ForegroundColor Red
        }
        Write-Host "  Time:       $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
        Write-Host ""

        $result.Downloaded = $downloadCount
        $result.Failed = $failCount
        $result.BytesDownloaded = $downloadedBytes
        $result.Duration = $stopwatch.Elapsed

    } finally {
        $session.Dispose()
    }

    return $result
}

#endregion

# Export public functions
Export-ModuleMember -Function Invoke-SFTPSync
