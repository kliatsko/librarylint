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
            Copy-Item -LiteralPath $netStdDll.FullName -Destination $targetDll -Force
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

    # Default Session.Timeout is 60s, which is too short for recursive deletes
    # of folders that contain hundreds of files (the timeout is the IPC channel
    # to WinSCP.exe, not the SFTP-level operation). Bump to 10 minutes so a slow
    # remote-side delete doesn't cascade-kill the session for everything that follows.
    try { $session.Timeout = [TimeSpan]::FromMinutes(10) } catch {}

    $session.Open($sessionOptions)

    return $session
}

function Get-RemoteFilesRecursive {
    param(
        $Session,
        [string]$RemotePath,
        [string]$BasePath = $null,
        [int]$Depth = 0,
        # Subtrees to skip — any subdirectory whose normalized path equals or
        # falls inside one of these is not descended into. Saves network round
        # trips when scanning a working dir that contains a known prune subtree
        # (we'd just filter it out post-scan anyway).
        [string[]]$ExcludePaths = @()
    )

    # Track the base path for progress display
    if (-not $BasePath) {
        $BasePath = $RemotePath
    }

    $files = @()

    # Avoid double slashes when callers pass a path with a trailing /
    $cleanRemote = $RemotePath.TrimEnd('/')

    # Pre-normalize exclude paths so the per-subdir check below is a cheap
    # string compare instead of repeating the normalization each iteration.
    $excludeNormalized = @()
    if ($ExcludePaths -and $ExcludePaths.Count -gt 0) {
        $excludeNormalized = @($ExcludePaths | ForEach-Object {
            (($_ -replace '\\', '/') -replace '/+', '/').TrimEnd('/')
        })
    }

    try {
        $directoryInfo = $Session.ListDirectory($RemotePath)

        foreach ($fileInfo in $directoryInfo.Files) {
            if ($fileInfo.IsDirectory) {
                if ($fileInfo.Name -ne "." -and $fileInfo.Name -ne "..") {
                    $subPath = "$cleanRemote/$($fileInfo.Name)"

                    # Pre-empt descent into excluded subtrees
                    $skipThis = $false
                    if ($excludeNormalized.Count -gt 0) {
                        $subNorm = (($subPath -replace '\\', '/') -replace '/+', '/').TrimEnd('/')
                        foreach ($excl in $excludeNormalized) {
                            if ($subNorm -eq $excl -or $subNorm.StartsWith($excl + '/')) {
                                $skipThis = $true
                                break
                            }
                        }
                    }
                    if ($skipThis) { continue }

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
                    # Recurse into subdirectories (passing exclude list down)
                    $files += Get-RemoteFilesRecursive -Session $Session -RemotePath $subPath -BasePath $BasePath -Depth ($Depth + 1) -ExcludePaths $ExcludePaths
                }
            } else {
                $files += [PSCustomObject]@{
                    FullPath = "$cleanRemote/$($fileInfo.Name)"
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

# Subfolder names that are part of a release (not the release itself).
# When a file's immediate parent matches one of these, walk up one level
# so files in Movie/Subs/x.srt group under Movie, not under Subs.
# Module-scope so multiple functions can share without redefining.
$script:ReleaseComponentSubfolders = @('subs','subtitles','sample','samples','featurettes',
    'extras','proof','cd1','cd2','cd3','cd4','disc1','disc2')

function Get-ReleaseFolderPath {
    param([string]$FilePath, [string[]]$Components)
    # Normalize separators — remote paths use forward slashes but Split-Path
    # on Windows can introduce backslashes that confuse downstream string ops.
    $normalized = ($FilePath -replace '\\', '/') -replace '/+', '/'
    $parent = Split-Path $normalized -Parent
    if (-not $parent) { return $null }
    $parent = ($parent -replace '\\', '/') -replace '/+', '/'
    $parentLeaf = Split-Path $parent -Leaf
    if ($parentLeaf -and $Components -contains $parentLeaf.ToLower()) {
        $parent = Split-Path $parent -Parent
        if ($parent) { $parent = ($parent -replace '\\', '/') -replace '/+', '/' }
    }
    return $parent
}

function Get-NormalizedReleaseKey {
    param([string]$FolderLeaf)
    if (-not $FolderLeaf) { return $null }
    $norm = ($FolderLeaf -replace '[\.\-_]', ' ' -replace '\s+', ' ').Trim().ToLower()
    $year = $null
    if ($norm -match '((?:19|20)\d{2})') { $year = $Matches[1] }
    $title = ($norm -replace '\s*(?:19|20)\d{2}.*$', '').Trim()
    $title = $title -replace '\s*(1080p|720p|2160p|4k|uhd|bluray|webrip|web-dl|web|brrip|bdrip|remux|x264|x265|hevc|aac|dts|ac3|atmos|hdr|10bit).*$', ''
    $title = $title.Trim()
    $title = $title -replace '^(the|a|an)\s+', ''
    if (-not $title) { return $null }
    if ($year) { return "$title|$year" }
    return $title
}

function Group-IntoReleaseFolders {
    param($Files, [string[]]$Components, [string[]]$VideoExts)
    $folders = @{}
    foreach ($f in $Files) {
        $releaseFolder = Get-ReleaseFolderPath -FilePath $f.FullPath -Components $Components
        if (-not $releaseFolder) { continue }

        if (-not $folders.ContainsKey($releaseFolder)) {
            $folders[$releaseFolder] = [PSCustomObject]@{
                Path = $releaseFolder
                Leaf = Split-Path $releaseFolder -Leaf
                PrimarySize = 0L
                PrimaryName = $null
                TotalVideoBytes = 0L
                VideoFileCount = 0
                # Tracks the newest mtime across ALL files in the release (not
                # just videos) — any recent activity is enough to flag the
                # folder as "in use" for age-based prune skips.
                NewestMtime = [DateTime]::MinValue
            }
        }
        $entry = $folders[$releaseFolder]
        if ($f.LastModified -and $f.LastModified -gt $entry.NewestMtime) {
            $entry.NewestMtime = $f.LastModified
        }
        $ext = [System.IO.Path]::GetExtension($f.Name).ToLower()
        if ($VideoExts -contains $ext) {
            $entry.VideoFileCount++
            $entry.TotalVideoBytes += $f.Size
            if ($f.Size -gt $entry.PrimarySize) {
                $entry.PrimarySize = $f.Size
                $entry.PrimaryName = $f.Name
            }
        }
    }
    return $folders
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

        # Filter out already downloaded. Tracking-file matches the EXACT
        # remote path of a previous download — never re-download these even
        # under -Force, since they're literally the same file we already
        # have. -Force only loosens the per-file local-have check below
        # (specifically the folder-name match, which can false-positive on
        # quality upgrades).
        $newFiles = $allFiles | Where-Object {
            -not $downloaded.ContainsKey($_.FullPath)
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
            # already-processed library copy). Two tiers:
            #   - Name+size match in inbox/library — guaranteed identical
            #     bytes; always skip even under -Force (no point re-fetching
            #     a file we already have).
            #   - Folder-name match in the library — a folder with this
            #     release name exists locally but the file inside might be a
            #     different cut. -Force suppresses ONLY this tier, so a
            #     Radarr re-acquisition / quality upgrade run can pull the
            #     new version that lives in a folder we already have.
            $remoteParent = Split-Path $file.FullPath -Parent
            $remoteParentName = if ($remoteParent) { Split-Path $remoteParent -Leaf } else { $null }
            $effectiveFolderSet = if ($Force) { $null } else { $localFolders }
            $haveCheck = Test-RemoteFileAlreadyHave -FileName $file.Name -FileSize $file.Size `
                -RemoteParentName $remoteParentName -FileIndex $localIndex -FolderSet $effectiveFolderSet
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
.PARAMETER RemotePaths
    Prune-target roots (e.g. rTorrent's complete folder).
.PARAMETER LibraryPaths
    Sync-source roots (e.g. Radarr's renamed library mirror). Tracked files
    that live under these paths are pruned in place. Without this, library
    files get silently skipped because their tracked paths don't match any
    PrunePath and the remap fallback rewrites them onto a non-existent twin.
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

        [string[]]$LibraryPaths = @(),

        # Local library paths used for auto-discovery: untracked files on the
        # seedbox under LibraryPaths are matched to local folders by name +
        # main-video size, then added to tracking so the local-verified path
        # picks them up. Without this, the prune only sees files LibraryLint
        # synced itself.
        [string[]]$LocalLibraryPaths = @(),

        [hashtable]$RadarrImportedPaths = $null,  # TMDB ID → movie info from Radarr

        # After deleting a prune-folder torrent's data, also erase its now-dead
        # entry from rTorrent's session. Only affects torrents whose data lives
        # under a prune-path root — library-mirror prunes never touch torrents
        # (those are hardlinks; the torrent keeps seeding from its own copy).
        [switch]$EraseTorrents,

        [switch]$WhatIf
    )

    # Union of every root the prune is allowed to touch. Tracked paths under
    # any of these roots are pruned in place; only paths outside them get
    # remapped onto a current root.
    $allRoots = @()
    if ($RemotePaths)  { $allRoots += $RemotePaths }
    if ($LibraryPaths) { $allRoots += $LibraryPaths }
    $allRoots = @($allRoots | Where-Object { $_ } | Select-Object -Unique)

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
    # Tracking file may not exist yet — auto-discovery (below) can still find
    # untracked library files. Build an empty index if absent.
    $downloaded = if (Test-Path $trackingPath) { Get-DownloadedFiles -TrackingPath $trackingPath } else { @{} }

    # Auto-discovery of untracked library files. Walks the seedbox LibraryPaths
    # and matches folder-by-folder against the local libraries by main-video
    # size. Files locally renamed by Radarr (whose names no longer match the
    # seedbox release naming) are still found because the folder name + primary
    # file size together identify the movie. Matches get added to the tracking
    # file as `AutoDiscovered=true` entries so the eligibility loop below sees
    # them and the local-verified fast path can prune them.
    $autoDiscovered = 0
    if ($LibraryPaths.Count -gt 0 -and $LocalLibraryPaths.Count -gt 0) {
        Write-Host "  Auto-discovering untracked library files..." -ForegroundColor Gray

        # Local index: top-level folder name (case-insensitive) →
        # @{ Path; FileSizes (HashSet<long>) }. The FileSizes set contains
        # every video file size found recursively under that top-level
        # folder. Movies (one video per folder) and TV shows (many episodes
        # per show folder, nested in seasons) both index naturally — the
        # set is size 1 for a movie folder, size ~N for an N-episode show.
        $videoExts = @('.mkv', '.mp4', '.avi', '.m4v')
        $junkRegex = '(?i)(-trailer|\.trailer|-sample|\.sample|-featurette|behindthescenes|extras?)$'
        $localTopFolders = @{}
        foreach ($root in $LocalLibraryPaths) {
            if (-not $root -or -not (Test-Path -LiteralPath $root)) { continue }
            foreach ($folder in (Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
                $sizes = New-Object 'System.Collections.Generic.HashSet[long]'
                Get-ChildItem -LiteralPath $folder.FullName -Recurse -File -ErrorAction SilentlyContinue |
                    Where-Object {
                        $videoExts -contains $_.Extension.ToLower() -and
                        $_.BaseName -notmatch $junkRegex
                    } | ForEach-Object {
                        [void]$sizes.Add([long]$_.Length)
                    }
                if ($sizes.Count -gt 0) {
                    $localTopFolders[$folder.Name.ToLower()] = @{
                        Path = $folder.FullName
                        FileSizes = $sizes
                    }
                }
            }
        }

        if ($localTopFolders.Count -eq 0) {
            Write-Host "    No local library folders found at $($LocalLibraryPaths -join ', ')" -ForegroundColor DarkYellow
        } else {
            # Open a scan-only session — the main delete flow opens its own
            # session later. Disposed at the end of discovery.
            $scanSession = $null
            try {
                $scanSession = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
                    -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
            } catch {
                Write-Host "    Scan-session connect failed: $_" -ForegroundColor Yellow
                Write-Host "    Skipping auto-discovery; proceeding with tracked-only prune." -ForegroundColor Yellow
            }

            if ($scanSession) {
                try {
                    # Pull all files under the seedbox LibraryPaths
                    $remoteFiles = @()
                    foreach ($lp in $LibraryPaths) {
                        $remoteFiles += Get-RemoteFilesRecursive -Session $scanSession -RemotePath $lp
                        $clearWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
                        if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
                    }

                    # Pass 1: match each remote VIDEO file by walking up its
                    # parent chain looking for an ancestor whose name is a
                    # known local top-level folder. If the file's size is in
                    # that folder's size set, it's a match. Innermost ancestor
                    # wins so the most-specific local folder is preferred.
                    #
                    # Movies:   /…/media/Movies/Almost Famous (2000)/movie.mkv
                    #           ancestors: [Almost Famous (2000), Movies, …]
                    #           → match at "Almost Famous (2000)"
                    #
                    # TV shows: /…/media/Shows/Samurai Champloo (2004)/Season 01/ep.mkv
                    #           ancestors: [Season 01, Samurai Champloo (2004), Shows, …]
                    #           → "Season 01" not in index, walk further → match at "Samurai Champloo (2004)"
                    $matchedByParent = @{}  # remote parent → list of matched video file objects, used by Pass 2
                    foreach ($f in $remoteFiles) {
                        if ($downloaded.ContainsKey($f.FullPath)) { continue }
                        $ext = [System.IO.Path]::GetExtension($f.Name).ToLower()
                        if ($videoExts -notcontains $ext) { continue }
                        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                        if ($baseName -match $junkRegex) { continue }

                        $normalized = ($f.FullPath -replace '\\', '/') -replace '/+', '/'
                        $parts = $normalized.Split('/')
                        $matchedFolder = $null
                        # Walk ancestors innermost-first. Skip ancestors whose
                        # leaf isn't a known local top-level folder.
                        for ($i = $parts.Count - 2; $i -ge 1; $i--) {
                            $leaf = $parts[$i].ToLower()
                            if ($localTopFolders.ContainsKey($leaf)) {
                                $entry = $localTopFolders[$leaf]
                                if ($entry.FileSizes.Contains([long]$f.Size)) {
                                    $matchedFolder = $entry
                                    break
                                }
                                # Folder-name hit but size miss — could be a
                                # different cut/encode. Stop here rather than
                                # continuing up; we found the right local
                                # folder, we just don't have THIS file.
                                break
                            }
                        }

                        if ($matchedFolder) {
                            $downloaded[$f.FullPath] = @{
                                LocalPath      = $matchedFolder.Path
                                Size           = $f.Size
                                DownloadedAt   = (Get-Date).ToString('o')
                                AutoDiscovered = $true
                                DiscoveryMatch = 'AncestorPlusSize'
                            }
                            $autoDiscovered++

                            # Remember which folder we matched in, scoped to
                            # the immediate parent so Pass 2 (sibling sweep)
                            # can find companion .nfo / .srt / .jpg files.
                            $idx = $normalized.LastIndexOf('/')
                            if ($idx -gt 0) {
                                $remoteParent = $normalized.Substring(0, $idx)
                                if (-not $matchedByParent.ContainsKey($remoteParent)) {
                                    $matchedByParent[$remoteParent] = @{
                                        LocalFolder = $matchedFolder.Path
                                        Videos      = @()
                                    }
                                }
                                $matchedByParent[$remoteParent].Videos += $f
                            }
                        }
                    }

                    # Pass 2: sibling sweep. For each remote parent folder
                    # that had at least one video matched in Pass 1, pull in
                    # the non-video companions (.nfo, .srt, artwork, etc.)
                    # so the prune can delete the whole folder cleanly.
                    #
                    # Movie-shaped folder (exactly 1 video matched): sweep
                    # ALL non-video siblings — single movie, no episode
                    # disambiguation needed.
                    #
                    # TV-shaped folder (2+ videos matched): only sweep non-
                    # video siblings whose basename stems match one of the
                    # matched videos. This protects metadata for unsynced
                    # episodes (which we DON'T want to prune) while still
                    # cleaning up the matched episodes' own companions.
                    foreach ($f in $remoteFiles) {
                        if ($downloaded.ContainsKey($f.FullPath)) { continue }
                        $ext = [System.IO.Path]::GetExtension($f.Name).ToLower()
                        if ($videoExts -contains $ext) { continue }

                        $normalized = ($f.FullPath -replace '\\', '/') -replace '/+', '/'
                        $idx = $normalized.LastIndexOf('/')
                        if ($idx -le 0) { continue }
                        $parent = $normalized.Substring(0, $idx)
                        if (-not $matchedByParent.ContainsKey($parent)) { continue }

                        $entry = $matchedByParent[$parent]
                        $isMultiVideo = ($entry.Videos.Count -gt 1)
                        $shouldSweep = $false

                        if (-not $isMultiVideo) {
                            # Single-video folder — movie pattern, sweep all
                            $shouldSweep = $true
                        } else {
                            # Multi-video folder — TV pattern, sweep only by
                            # name-prefix match against matched video stems
                            $stem = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                            foreach ($v in $entry.Videos) {
                                $vStem = [System.IO.Path]::GetFileNameWithoutExtension($v.Name)
                                if ($stem -eq $vStem -or
                                    $stem.StartsWith($vStem + '.') -or
                                    $stem.StartsWith($vStem + ' ') -or
                                    $stem.StartsWith($vStem + '-')) {
                                    $shouldSweep = $true
                                    break
                                }
                            }
                        }

                        if ($shouldSweep) {
                            $downloaded[$f.FullPath] = @{
                                LocalPath      = $entry.LocalFolder
                                Size           = $f.Size
                                DownloadedAt   = (Get-Date).ToString('o')
                                AutoDiscovered = $true
                                DiscoveryMatch = if ($isMultiVideo) { 'TVSibling' } else { 'MovieSibling' }
                            }
                            $autoDiscovered++
                        }
                    }

                    if ($autoDiscovered -gt 0) {
                        Save-DownloadedFiles -Downloaded $downloaded -TrackingPath $trackingPath
                    }
                    Write-Host "    Local top-level folders indexed: $($localTopFolders.Count) | Remote files scanned: $($remoteFiles.Count) | Newly tracked: $autoDiscovered" -ForegroundColor Gray
                } finally {
                    $scanSession.Dispose()
                }
            }
        }
        Write-Host ""
    }

    if ($downloaded.Count -eq 0) {
        Write-Host "  No tracked downloads - nothing to prune" -ForegroundColor Yellow
        return @{ Deleted = 0; Failed = 0 }
    }

    # Eligibility:
    #   - Library files (path under LibraryPaths): eligible AS SOON AS the local
    #     copy is confirmed (file exists with matching size). The library mirror
    #     is usually a hardlink farm — removing the entry doesn't touch the
    #     bytes the torrent still references. No age requirement. Falls back
    #     to age-based eligibility if the local file is missing/mismatched
    #     (rare: deleted-locally, drive moved, etc.).
    #   - Prune-path files (rTorrent complete/working): age-only, hit-and-run
    #     window typically applies. Existing $DaysOld threshold governs.
    $cutoffDate = (Get-Date).AddDays(-$DaysOld)
    $filesToPrune = @()
    $libraryRoots = if ($LibraryPaths) { @($LibraryPaths) } else { @() }

    # Helper: does the tracked path live under any of the given root paths?
    # Forward-slash semantics; tracked paths are POSIX seedbox paths.
    $isUnderRoots = {
        param([string]$Path, [string[]]$Roots)
        foreach ($r in $Roots) { if ($Path.StartsWith($r)) { return $true } }
        return $false
    }

    foreach ($entry in $downloaded.GetEnumerator()) {
        $trackedPath = $entry.Key
        $downloadedAt = [DateTime]::Parse($entry.Value.DownloadedAt)
        $isLibrary = (& $isUnderRoots $trackedPath $libraryRoots)

        $eligible = $false
        $eligibilityReason = $null

        if ($isLibrary) {
            # Local-presence check: local file must exist AND match the tracked
            # size. Catches partial downloads, deletions, and bit-rot at the
            # size level (full hash check would be too slow for a 50TB library;
            # size mismatch is the practical floor).
            #
            # Auto-discovered entries (from the discovery pass above) already
            # had their match verified at the folder level — the LocalPath
            # points at the local FOLDER, not a file with matching size.
            # Re-verify that the folder still exists and skip the per-file
            # size check.
            $localOk = $false
            if ($entry.Value.AutoDiscovered) {
                if ($entry.Value.LocalPath -and (Test-Path -LiteralPath $entry.Value.LocalPath)) {
                    $localOk = $true
                }
            } elseif ($entry.Value.LocalPath -and (Test-Path -LiteralPath $entry.Value.LocalPath)) {
                try {
                    $localSize = (Get-Item -LiteralPath $entry.Value.LocalPath -ErrorAction Stop).Length
                    if ($localSize -eq [long]$entry.Value.Size) {
                        $localOk = $true
                    }
                } catch {}
            }

            if ($localOk) {
                $eligible = $true
                $eligibilityReason = 'LocalVerified'
            } elseif ($downloadedAt -lt $cutoffDate) {
                # No local match, but age-based safety net still applies.
                $eligible = $true
                $eligibilityReason = 'LibraryAgeFallback'
            }
        } else {
            # Non-library tracked file: age-based only (existing behavior).
            if ($downloadedAt -lt $cutoffDate) {
                $eligible = $true
                $eligibilityReason = 'Age'
            }
        }

        if (-not $eligible) { continue }

        # Resolve the actual remote path — the tracked path may have changed
        # if remote paths config was updated.
        $resolvedPath = $trackedPath
        $matchesCurrent = $false
        foreach ($rp in $allRoots) {
            if ($trackedPath.StartsWith($rp)) { $matchesCurrent = $true; break }
        }
        if (-not $matchesCurrent) {
            # Find the deepest matching portion of the path against any configured root
            $pathParts = $trackedPath.Split('/')
            for ($i = 1; $i -lt $pathParts.Count - 1; $i++) {
                $candidateRelative = ($pathParts[$i..($pathParts.Count - 1)] -join '/')
                foreach ($rp in $allRoots) {
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
            RemotePath        = $resolvedPath
            TrackingKey       = $trackedPath
            LocalPath         = $entry.Value.LocalPath
            Size              = $entry.Value.Size
            DownloadedAt      = $downloadedAt
            Age               = [math]::Floor(((Get-Date) - $downloadedAt).TotalDays)
            EligibilityReason = $eligibilityReason
        }
    }

    if ($filesToPrune.Count -eq 0) {
        Write-Host "  No files eligible for prune (none older than $DaysOld days, no library files confirmed locally)" -ForegroundColor Green
        return @{ Deleted = 0; Failed = 0; Skipped = 0 }
    }

    # Surface how many files are coming from each eligibility path. Helps the
    # user see at a glance whether the local-verified fast path is doing
    # work or whether everything is falling back to age-based.
    $byReason = $filesToPrune | Group-Object EligibilityReason | ForEach-Object { @{ Reason = $_.Name; Count = $_.Count } }
    Write-Host "  Eligibility breakdown:" -ForegroundColor DarkGray
    foreach ($g in $byReason) {
        $label = switch ($g.Reason) {
            'LocalVerified'      { "library (local copy verified)" }
            'LibraryAgeFallback' { "library (age fallback, local missing/mismatch)" }
            'Age'                { "prune folder (age >= $DaysOld days)" }
            default              { $g.Reason }
        }
        Write-Host "    $($g.Count) — $label" -ForegroundColor DarkGray
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
            # LocalVerified library files have ALREADY passed the strongest
            # possible check (we have the bytes on disk with matching size).
            # Requiring Radarr to also confirm import here is redundant — and
            # would falsely block prunes when Radarr is briefly out of sync
            # with what's actually in its library.
            if ($file.EligibilityReason -eq 'LocalVerified') {
                $safeToPrune += $file
                continue
            }

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
        # Tag with eligibility: if any file in the group is LocalVerified, the
        # folder is being pruned because the local copy is confirmed — that's
        # more meaningful than "0 days old" for a same-day sync.
        $groupReasons = @($group.Group | ForEach-Object { $_.EligibilityReason } | Select-Object -Unique)
        $tag = if ($groupReasons -contains 'LocalVerified') { '[local OK]' }
               elseif ($groupReasons -contains 'LibraryAgeFallback') { '[lib age]' }
               else { "$($maxAge)d old" }
        Write-Host "    $tag : " -NoNewline -ForegroundColor Gray
        Write-Host "$truncatedName " -NoNewline -ForegroundColor White
        Write-Host "($($group.Count) files, $(Format-SyncSize $folderSize))" -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host "  Found $($filesToPrune.Count) files in $($folderGroups.Count) folders to prune ($(Format-SyncSize $totalSize))" -ForegroundColor Cyan
    Write-Host ""

    # Age-fallback confirmation. LocalVerified files have proof on disk and
    # PrunePaths files passed the explicit age threshold the user chose. The
    # remaining bucket — library files where the local copy is missing — is
    # genuinely risky: a torrent that finished on the seedbox but was never
    # synced locally (user away for weeks, sync skipped, etc.) looks
    # identical to a file the user deliberately moved. Require an explicit
    # opt-in here, defaulting to N so a stray Enter doesn't drop anything.
    $libraryAgeFiles = @($filesToPrune | Where-Object { $_.EligibilityReason -eq 'LibraryAgeFallback' })
    if ($libraryAgeFiles.Count -gt 0 -and -not $WhatIf) {
        Write-Host "  WARNING — age-fallback prune candidates" -ForegroundColor Yellow
        Write-Host "  $($libraryAgeFiles.Count) library file(s) have NO matching local copy and would be" -ForegroundColor Yellow
        Write-Host "  pruned only because they're older than the age threshold." -ForegroundColor Yellow
        Write-Host "  If you haven't synced these yet, they will be permanently removed" -ForegroundColor Yellow
        Write-Host "  and need to be re-acquired." -ForegroundColor Yellow
        Write-Host ""

        $ageFolderGroups = $libraryAgeFiles | Group-Object { Split-Path $_.RemotePath -Parent }
        foreach ($g in $ageFolderGroups | Sort-Object { ($_.Group | Measure-Object -Property Age -Maximum).Maximum } -Descending | Select-Object -First 20) {
            $folderName = Split-Path $g.Name -Leaf
            if (-not $folderName) { $folderName = $g.Name }
            $truncatedName = if ($folderName.Length -gt 60) { $folderName.Substring(0, 57) + "..." } else { $folderName }
            $maxAge = ($g.Group | Measure-Object -Property Age -Maximum).Maximum
            $folderSize = ($g.Group | Measure-Object -Property Size -Sum).Sum
            Write-Host "    $($maxAge)d old: " -NoNewline -ForegroundColor DarkYellow
            Write-Host "$truncatedName " -NoNewline -ForegroundColor White
            Write-Host "($($g.Count) files, $(Format-SyncSize $folderSize))" -ForegroundColor DarkGray
            # Also surface the tracked LocalPath that didn't resolve, so the
            # user can tell at a glance whether they recognize where it was
            # supposed to be.
            $sampleLocal = ($g.Group | Where-Object { $_.LocalPath } | Select-Object -First 1).LocalPath
            if ($sampleLocal) {
                Write-Host "      expected local: $sampleLocal" -ForegroundColor DarkGray
            }
        }
        if ($ageFolderGroups.Count -gt 20) {
            Write-Host "    ... and $($ageFolderGroups.Count - 20) more folder(s)" -ForegroundColor DarkGray
        }
        Write-Host ""

        $ageConfirm = Read-Host "Proceed with deleting these $($libraryAgeFiles.Count) age-fallback file(s)? (Y/N) [N]"
        if ($ageConfirm -notmatch '^[Yy]') {
            # Drop only the age-fallback bucket; keep LocalVerified + Age pruning.
            $filesToPrune = @($filesToPrune | Where-Object { $_.EligibilityReason -ne 'LibraryAgeFallback' })
            $deferredAgeCount = $libraryAgeFiles.Count
            Write-Host "  Skipping $deferredAgeCount age-fallback file(s). Other eligible files will still be pruned." -ForegroundColor DarkGray
            Write-Host ""

            if ($filesToPrune.Count -eq 0) {
                Write-Host "  Nothing else eligible — exiting." -ForegroundColor Green
                return @{ Deleted = 0; Failed = 0; Skipped = $notImported.Count; SkippedAgeFallback = $deferredAgeCount }
            }
        }
    }

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
    $torrentsErased = 0

    try {
        # rTorrent snapshot for the torrent-erase integration. Capture every
        # torrent whose data currently lives under a prune-path root, BEFORE
        # any deletes. After the deletes, any of these whose data has gone
        # missing was killed by this prune — its dead entry gets erased.
        # Snapshotting first (rather than diffing the whole session) keeps
        # the scope tight: working-dir torrents and anything not under a
        # prune root are never considered.
        $torrentSnapshot = @{}
        if ($EraseTorrents) {
            Write-Host "  Snapshotting rTorrent session..." -ForegroundColor Gray -NoNewline
            $snap = Get-SeedboxTorrents -Session $session
            if ($snap.Success) {
                foreach ($t in $snap.Torrents) {
                    if ($t.path_exists -and (Test-RTorrentPathUnderRoots -Path $t.base_path -Roots $RemotePaths)) {
                        $torrentSnapshot[$t.hash] = $t.name
                    }
                }
                Write-Host " $($torrentSnapshot.Count) torrent(s) under prune path(s)" -ForegroundColor Gray
            } else {
                Write-Host " failed" -ForegroundColor Yellow
                Write-Host "  rTorrent snapshot failed ($($snap.Error)) — torrent erase skipped this run." -ForegroundColor Yellow
            }
        }

        Write-Host ""
        Write-Host "  Deleting files..." -ForegroundColor Yellow
        Write-Host ""

        $deleteFolderGroups = $filesToPrune | Group-Object { Split-Path $_.RemotePath -Parent }
        $touchedParents = New-Object 'System.Collections.Generic.HashSet[string]'

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

            if ($group.Name) { [void]$touchedParents.Add($group.Name) }

            Write-Host "    $truncatedName " -NoNewline -ForegroundColor White
            if ($folderFailed -gt 0) {
                Write-Host "($folderDeleted/$folderFileCount deleted, $folderFailed failed)" -ForegroundColor Red
            } elseif ($folderNotFound -gt 0) {
                Write-Host "($folderDeleted deleted, $folderNotFound already gone)" -ForegroundColor Gray
            } else {
                Write-Host "($folderDeleted files deleted)" -ForegroundColor Green
            }
        }

        # Empty-folder cleanup. Sort deepest-first so a sub/ that goes empty
        # gets removed before we evaluate its parent release/. Cascade up via
        # queue so a release/ becoming empty after its sub/ is removed gets
        # its own turn. Bounded by $allRoots — never walk past a configured
        # root. Remove-RemoteDirIfEmpty's own emptiness check handles the
        # failed-delete case (folder won't be empty if a file delete failed).
        if ($touchedParents.Count -gt 0) {
            $sortedParents = @($touchedParents) | Sort-Object -Property Length -Descending
            $cleanupQueue = New-Object 'System.Collections.Generic.Queue[string]'
            foreach ($p in $sortedParents) { $cleanupQueue.Enqueue($p) }

            while ($cleanupQueue.Count -gt 0) {
                $candidate = $cleanupQueue.Dequeue()
                if ($allRoots.Count -gt 0 -and (Test-RemotePathAtOrAboveRoot -Path $candidate -Roots $allRoots)) { continue }

                $outcome = Remove-RemoteDirIfEmpty -Session $session -RemotePath $candidate
                if ($outcome.Removed) {
                    $foldersRemoved++
                    $grandparent = Split-Path $candidate -Parent
                    if ($grandparent -and ($allRoots.Count -eq 0 -or -not (Test-RemotePathAtOrAboveRoot -Path $grandparent -Roots $allRoots))) {
                        $cleanupQueue.Enqueue($grandparent)
                    }
                }
            }
        }

        # Save updated tracking file
        Save-DownloadedFiles -Downloaded $downloaded -TrackingPath $trackingPath

        # Torrent-erase: re-query rTorrent and erase any snapshotted torrent
        # whose data has now gone missing — those are exactly the torrents
        # this prune killed. Past the age threshold the user chose, so the
        # hit-and-run window has elapsed; erasing the dead entry is safe.
        if ($EraseTorrents -and $torrentSnapshot.Count -gt 0) {
            Write-Host ""
            Write-Host "  Erasing rTorrent entries for fully-pruned torrents..." -ForegroundColor Yellow
            $postSnap = Get-SeedboxTorrents -Session $session
            if ($postSnap.Success) {
                foreach ($t in $postSnap.Torrents) {
                    if ($torrentSnapshot.ContainsKey($t.hash) -and -not $t.path_exists) {
                        $tname = $torrentSnapshot[$t.hash]
                        $shortName = if ($tname.Length -gt 55) { $tname.Substring(0, 52) + '...' } else { $tname }
                        $rm = Remove-SeedboxTorrent -Session $session -Hash $t.hash
                        if ($rm.Success) {
                            Write-Host "    erased: $shortName" -ForegroundColor Green
                            $torrentsErased++
                        } else {
                            Write-Host "    erase failed: $shortName — $($rm.Error)" -ForegroundColor Yellow
                        }
                    }
                }
                if ($torrentsErased -eq 0) {
                    Write-Host "    (no torrent had ALL its data pruned this run — entries kept)" -ForegroundColor DarkGray
                }
            } else {
                Write-Host "    Could not re-query rTorrent: $($postSnap.Error)" -ForegroundColor Yellow
            }
        }

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
    if ($EraseTorrents) {
        Write-Host "  Torrents erased: $torrentsErased rTorrent entry(ies)" -ForegroundColor Green
    }
    Write-Host ""

    return @{
        Deleted = $deletedCount
        Failed = $failedCount
        Skipped = $notImported.Count
        BytesDeleted = $deletedBytes
        FoldersRemoved = $foldersRemoved
        TorrentsErased = $torrentsErased
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
    Detects incomplete downloads by comparing video file sizes between sync and prune folders.
.DESCRIPTION
    With a typical Radarr+rTorrent seedbox setup, the same video file lives in
    both the sync folder (Radarr's renamed library) and the prune folder
    (rTorrent's completed downloads). When Radarr hardlinks on import, both
    copies share inodes and are byte-identical. A size mismatch therefore
    indicates a real problem: a broken hardlink, a re-download that only
    replaced one side, or a corrupt copy.

    Algorithm:
      1. Recursively list both trees, group files into release folders.
      2. For each release folder, identify the largest video file (the "primary").
      3. Match prune folders to sync folders in two passes:
         - Pass 1: exact primary-file size match (catches all hardlinked pairs).
         - Pass 2: title+year normalization on remaining folders (catches the
           bug case where sizes differ but it's the same release).
      4. Flag matched pairs whose primary sizes differ as incomplete.
      5. Report unmatched prune folders separately (likely not yet imported).
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
        [string[]]$SyncPaths,

        [string[]]$PrunePaths = @(),

        [string[]]$WorkingPaths = @(),

        [string[]]$VideoExtensions = @(".mkv", ".mp4", ".avi", ".m4v", ".wmv", ".ts", ".mpg", ".mpeg")
    )

    # Header
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "           FIND INCOMPLETE DOWNLOADS                   " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Two checks:" -ForegroundColor Gray
    Write-Host "    1. Working dir (rTorrent in-progress) — anything here is" -ForegroundColor Gray
    Write-Host "       not yet finished by rTorrent or is pending the move." -ForegroundColor Gray
    Write-Host "    2. Sync vs prune — primary video file sizes should match" -ForegroundColor Gray
    Write-Host "       byte-for-byte under hardlinks; mismatches mean corruption." -ForegroundColor Gray
    Write-Host ""

    if ((-not $PrunePaths -or $PrunePaths.Count -eq 0) -and
        (-not $WorkingPaths -or $WorkingPaths.Count -eq 0)) {
        Write-Host "  Neither SFTPPrunePaths nor SFTPWorkingPaths is configured." -ForegroundColor Red
        Write-Host "  Configure at least one and retry." -ForegroundColor Gray
        Write-Host ""
        return @{ Mismatches = @(); InProgress = @(); Error = "No prune or working paths configured"; TotalScanned = 0 }
    }

    # Check WinSCP .NET assembly
    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        return @{ Mismatches = @(); Error = "WinSCP .NET assembly not installed"; TotalScanned = 0 }
    }

    Write-Host "  Host: ${HostName}:${Port}" -ForegroundColor Gray
    Write-Host "  User: $Username" -ForegroundColor Gray
    Write-Host ""

    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ Mismatches = @(); Error = $_.ToString(); TotalScanned = 0 }
    }

    # Component-subfolder list lives at module scope (see top of file). These
    # helpers (Get-ReleaseFolderPath, Group-IntoReleaseFolders) are also
    # module-scope so they can be reused by Invoke-SFTPPruneWorkingDir.
    $componentSubfolders = $script:ReleaseComponentSubfolders

    try {
        $scanTree = {
            param($paths, $label, $excludePaths)
            Write-Host ""
            Write-Host "  Scanning $label..." -ForegroundColor Gray
            $collected = @()
            foreach ($remotePath in $paths) {
                $rrParams = @{ Session = $session; RemotePath = $remotePath }
                if ($excludePaths -and $excludePaths.Count -gt 0) { $rrParams.ExcludePaths = $excludePaths }
                $files = Get-RemoteFilesRecursive @rrParams
                $clearWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
                if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
                Write-Host "    $remotePath ($($files.Count) files)" -ForegroundColor Gray
                $collected += $files
            }
            return ,$collected
        }

        # Scan working (in-progress) side — skip prune subtree to avoid
        # re-enumerating the same files we'll scan separately as prune.
        $workingFiles = @()
        if ($WorkingPaths -and $WorkingPaths.Count -gt 0) {
            $workingFiles = & $scanTree $WorkingPaths "working (rTorrent in-progress)" $PrunePaths
        }

        # Scan sync side
        $syncFiles = & $scanTree $SyncPaths "sync (library)" @()

        # Scan prune side
        $pruneFiles = @()
        if ($PrunePaths -and $PrunePaths.Count -gt 0) {
            $pruneFiles = & $scanTree $PrunePaths "prune (rTorrent complete)" @()
        }

        # Group all three into release folders
        $workingFolders = Group-IntoReleaseFolders -Files $workingFiles -Components $componentSubfolders -VideoExts $VideoExtensions
        $syncFolders = Group-IntoReleaseFolders -Files $syncFiles -Components $componentSubfolders -VideoExts $VideoExtensions
        $pruneFolders = Group-IntoReleaseFolders -Files $pruneFiles -Components $componentSubfolders -VideoExts $VideoExtensions

        $workingWithVideo = @{}
        foreach ($k in $workingFolders.Keys) {
            if ($workingFolders[$k].VideoFileCount -gt 0) { $workingWithVideo[$k] = $workingFolders[$k] }
        }

        # Many seedboxes nest the prune (rTorrent move-on-completion target)
        # subfolder INSIDE the working dir — e.g. working = /downloads/rtorrent/
        # and prune = /downloads/rtorrent/complete/radarr/. A naive recursive
        # working scan then surfaces completed releases as "in-progress." Drop
        # any working folder that lives inside a configured prune path.
        $excludedFromWorking = 0
        if ($PrunePaths -and $PrunePaths.Count -gt 0 -and $workingWithVideo.Count -gt 0) {
            $pruneRoots = @($PrunePaths | ForEach-Object {
                (($_ -replace '\\', '/') -replace '/+', '/').TrimEnd('/')
            })
            $filtered = @{}
            foreach ($k in $workingWithVideo.Keys) {
                $folderPath = (($workingWithVideo[$k].Path -replace '\\', '/') -replace '/+', '/').TrimEnd('/')
                $insidePrune = $false
                foreach ($pr in $pruneRoots) {
                    if ($folderPath -eq $pr -or $folderPath.StartsWith($pr + '/')) {
                        $insidePrune = $true
                        break
                    }
                }
                if ($insidePrune) {
                    $excludedFromWorking++
                } else {
                    $filtered[$k] = $workingWithVideo[$k]
                }
            }
            $workingWithVideo = $filtered
        }
        $syncWithVideo = @{}
        foreach ($k in $syncFolders.Keys) {
            if ($syncFolders[$k].VideoFileCount -gt 0) { $syncWithVideo[$k] = $syncFolders[$k] }
        }
        $pruneWithVideo = @{}
        foreach ($k in $pruneFolders.Keys) {
            if ($pruneFolders[$k].VideoFileCount -gt 0) { $pruneWithVideo[$k] = $pruneFolders[$k] }
        }

        Write-Host ""

        # Pass 1: index sync by primary size; match prune folders whose primary
        # size hits an unclaimed sync folder of the same size
        $syncBySize = @{}
        foreach ($k in $syncWithVideo.Keys) {
            $sz = $syncWithVideo[$k].PrimarySize
            if (-not $syncBySize.ContainsKey($sz)) {
                $syncBySize[$sz] = New-Object System.Collections.Generic.List[object]
            }
            $syncBySize[$sz].Add($syncWithVideo[$k])
        }

        $pairs = @()
        $matchedSync = @{}
        $deferred = @()

        foreach ($pk in $pruneWithVideo.Keys) {
            $prune = $pruneWithVideo[$pk]
            $sz = $prune.PrimarySize
            $matched = $false
            if ($syncBySize.ContainsKey($sz)) {
                $candidates = @($syncBySize[$sz] | Where-Object { -not $matchedSync.ContainsKey($_.Path) })
                if ($candidates.Count -ge 1) {
                    $sync = if ($candidates.Count -eq 1) {
                        $candidates[0]
                    } else {
                        # Multiple sync folders with the same primary size — disambiguate by name
                        $pruneKey = Get-NormalizedReleaseKey -FolderLeaf $prune.Leaf
                        $byKey = $candidates | Where-Object {
                            (Get-NormalizedReleaseKey -FolderLeaf $_.Leaf) -eq $pruneKey
                        } | Select-Object -First 1
                        if ($byKey) { $byKey } else { $candidates[0] }
                    }
                    $matchedSync[$sync.Path] = $true
                    $pairs += [PSCustomObject]@{ Prune = $prune; Sync = $sync }
                    $matched = $true
                }
            }
            if (-not $matched) { $deferred += $prune }
        }

        # Pass 2: remaining prune folders matched to remaining sync folders by
        # normalized title+year. Any pair found here has a size mismatch by
        # definition (otherwise Pass 1 would have caught it).
        $unmatchedPrune = @()
        if ($deferred.Count -gt 0) {
            $unmatchedSync = @($syncWithVideo.Values | Where-Object { -not $matchedSync.ContainsKey($_.Path) })
            $syncByKey = @{}
            foreach ($s in $unmatchedSync) {
                $key = Get-NormalizedReleaseKey -FolderLeaf $s.Leaf
                if ($key -and -not $syncByKey.ContainsKey($key)) { $syncByKey[$key] = $s }
            }

            foreach ($prune in $deferred) {
                $key = Get-NormalizedReleaseKey -FolderLeaf $prune.Leaf
                if ($key -and $syncByKey.ContainsKey($key)) {
                    $sync = $syncByKey[$key]
                    $matchedSync[$sync.Path] = $true
                    $syncByKey.Remove($key)
                    $pairs += [PSCustomObject]@{ Prune = $prune; Sync = $sync }
                } else {
                    $unmatchedPrune += $prune
                }
            }
        }

        $mismatches = @($pairs | Where-Object { $_.Prune.PrimarySize -ne $_.Sync.PrimarySize })
        $healthy = @($pairs | Where-Object { $_.Prune.PrimarySize -eq $_.Sync.PrimarySize })
        $unmatchedSyncCount = @($syncWithVideo.Values | Where-Object { -not $matchedSync.ContainsKey($_.Path) }).Count

        # Display results
        Write-Host ""
        Write-Host "  RESULTS" -ForegroundColor Cyan
        Write-Host "  -------" -ForegroundColor Cyan

        # Section 1: anything in the working dir is by definition incomplete or
        # pending move. rTorrent pre-allocates the full file size, so the size
        # number does not indicate completion — it just means the file slot exists.
        if ($WorkingPaths -and $WorkingPaths.Count -gt 0) {
            Write-Host ""
            if ($workingWithVideo.Count -eq 0) {
                Write-Host "  In-progress (working dir): none" -ForegroundColor Green
            } else {
                Write-Host "  IN-PROGRESS / PENDING MOVE ($($workingWithVideo.Count)):" -ForegroundColor Yellow
                Write-Host "    Anything here is not yet finished by rTorrent (or waiting" -ForegroundColor DarkGray
                Write-Host "    to be moved). File sizes may be pre-allocated and not reflect" -ForegroundColor DarkGray
                Write-Host "    actual progress — verify via the ruTorrent UI." -ForegroundColor DarkGray
                foreach ($wf in ($workingWithVideo.Values | Sort-Object Leaf)) {
                    Write-Host ""
                    Write-Host "    $($wf.Leaf)" -ForegroundColor White
                    Write-Host "      primary : $($wf.PrimaryName) ($(Format-SyncSize $wf.PrimarySize))" -ForegroundColor Gray
                    Write-Host "      path    : $($wf.Path)" -ForegroundColor DarkGray
                }
            }
        }

        # Section 2: sync vs prune size comparison (only if prune is configured)
        if ($PrunePaths -and $PrunePaths.Count -gt 0) {
            Write-Host ""
            if ($mismatches.Count -eq 0) {
                Write-Host "  Sync vs prune sizes: no mismatches found" -ForegroundColor Green
            } else {
                Write-Host "  SIZE MISMATCH SYNC VS PRUNE ($($mismatches.Count)):" -ForegroundColor Yellow
                foreach ($m in $mismatches) {
                    $prune = $m.Prune
                    $sync = $m.Sync
                    $delta = $prune.PrimarySize - $sync.PrimarySize
                    $sign = if ($delta -gt 0) { '+' } else { '-' }
                    $absDelta = [math]::Abs($delta)
                    $pct = if ($sync.PrimarySize -gt 0) {
                        [math]::Round(($prune.PrimarySize / $sync.PrimarySize) * 100, 1)
                    } else { 0 }
                    Write-Host ""
                    Write-Host "    $($sync.Leaf)" -ForegroundColor White
                    Write-Host "      sync  : $(Format-SyncSize $sync.PrimarySize)" -ForegroundColor Gray
                    Write-Host "      prune : $(Format-SyncSize $prune.PrimarySize) " -NoNewline -ForegroundColor Yellow
                    Write-Host "($sign$(Format-SyncSize $absDelta), $pct% of sync)" -ForegroundColor DarkYellow
                    Write-Host "      sync  path : $($sync.Path)" -ForegroundColor DarkGray
                    Write-Host "      prune path : $($prune.Path)" -ForegroundColor DarkGray
                }
            }

            Write-Host ""
            Write-Host "  Healthy (matched, sizes equal): $($healthy.Count)" -ForegroundColor Green
            if ($unmatchedPrune.Count -gt 0) {
                Write-Host "  Prune folders with no library counterpart: $($unmatchedPrune.Count)" -ForegroundColor DarkYellow
                Write-Host "    (likely not yet imported by Radarr — not flagged as incomplete)" -ForegroundColor DarkGray
            }
            if ($unmatchedSyncCount -gt 0) {
                Write-Host "  Library folders with no prune counterpart: $unmatchedSyncCount" -ForegroundColor DarkGray
                Write-Host "    (rTorrent download already pruned, manual import, etc.)" -ForegroundColor DarkGray
            }
        }

        # Final diagnostics — always at the bottom, where the terminal scroll
        # buffer leaves them visible even when the in-progress list is long.
        Write-Host ""
        Write-Host "  Scan diagnostics" -ForegroundColor Cyan
        Write-Host "  ----------------" -ForegroundColor Cyan
        if ($WorkingPaths -and $WorkingPaths.Count -gt 0) {
            Write-Host "  Working path(s):  $($WorkingPaths -join ', ')" -ForegroundColor DarkGray
            $workingMsg = "  Working release folders: $($workingWithVideo.Count)"
            if ($excludedFromWorking -gt 0) {
                $workingMsg += " ($excludedFromWorking inside prune path excluded)"
            } elseif ($PrunePaths -and $PrunePaths.Count -gt 0) {
                $workingMsg += " (0 excluded — no working folders fell inside the prune path)"
            }
            Write-Host $workingMsg -ForegroundColor Gray
        }
        Write-Host "  Sync path(s):     $($SyncPaths -join ', ')" -ForegroundColor DarkGray
        Write-Host "  Sync release folders: $($syncWithVideo.Count)" -ForegroundColor Gray
        if ($PrunePaths -and $PrunePaths.Count -gt 0) {
            Write-Host "  Prune path(s):    $($PrunePaths -join ', ')" -ForegroundColor DarkGray
            Write-Host "  Prune release folders: $($pruneWithVideo.Count)" -ForegroundColor Gray
        }
        Write-Host ""

        return @{
            Mismatches = $mismatches
            Healthy = $healthy.Count
            UnmatchedPrune = $unmatchedPrune.Count
            UnmatchedSync = $unmatchedSyncCount
            InProgress = @($workingWithVideo.Values)
            TotalScanned = ($syncFiles.Count + $pruneFiles.Count + $workingFiles.Count)
        }

    } finally {
        $session.Dispose()
    }
}

<#
.SYNOPSIS
    Prunes release folders from the rTorrent working dir whose primary video
    file matches a folder in the sync (library) tree by exact byte size.
.DESCRIPTION
    rTorrent's "Move on completion" plugin sometimes fails to relocate a
    finished torrent into the complete folder — leaving a release sitting in
    the working dir indefinitely. The regular prune (which scans the complete
    folder) never sees those, and they accumulate.

    Safe deletion criterion: a working-dir folder whose primary video file
    has the EXACT same byte size as a video file in the sync (library) tree.
    Under hardlinks (the standard Radarr setup) the bytes are shared with
    the library copy, so removing the working-dir reference doesn't touch
    the data — the library hardlink keeps the file alive. Under independent
    copies, the library copy survives independently. Either way, safe.

    Folders with NO library size match are left alone — they may still be
    actively downloading, paused, or stuck for reasons that need human
    attention via the ruTorrent UI.

    Folders living inside any configured PrunePaths subtree are skipped —
    those are the regular-prune's territory.

    Folders whose newest file mtime is within DaysOld days of now are also
    skipped — a fresh download is presumed to still be seeding, and rTorrent
    treats the working-dir copy as authoritative. Deleting it mid-seed
    breaks the torrent. Default 14 days; set 0 to disable the age guard.
#>
function Invoke-SFTPPruneWorkingDir {
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
        [string[]]$WorkingPaths,

        [Parameter(Mandatory=$true)]
        [string[]]$SyncPaths,

        [string[]]$PrunePaths = @(),

        # Local library roots used as an additional "we have this" signal.
        # If aggressive pruning has emptied the seedbox library mirror,
        # working-dir releases have nothing remote to match against. With
        # LocalLibraryPaths set, the working-dir prune also indexes the
        # local libraries by primary-video size, so files we have locally
        # (and so safely backed up off the seedbox) become eligible.
        [string[]]$LocalLibraryPaths = @(),

        [string[]]$VideoExtensions = @(".mkv", ".mp4", ".avi", ".m4v", ".wmv", ".ts", ".mpg", ".mpeg"),

        [int]$DaysOld = 14,

        [switch]$WhatIf
    )

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "             SFTP PRUNE — WORKING DIR                  " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Looks for finished-but-unmoved releases in the rTorrent" -ForegroundColor Gray
    Write-Host "  working dir, matched against the library by exact video" -ForegroundColor Gray
    Write-Host "  file size. Only matched folders are eligible for deletion." -ForegroundColor Gray
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] No folders will be deleted" -ForegroundColor Yellow
        Write-Host ""
    }

    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        return @{ Deleted = 0; Failed = 0; Skipped = 0; Error = "WinSCP .NET assembly not installed" }
    }

    Write-Host "  Host: ${HostName}:${Port}" -ForegroundColor Gray
    Write-Host "  User: $Username" -ForegroundColor Gray
    Write-Host ""

    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ Deleted = 0; Failed = 0; Skipped = 0; Error = $_.ToString() }
    }

    try {
        # Scan working tree, pre-empting descent into prune subtrees so we
        # don't enumerate thousands of completed-folder files just to throw
        # them away later.
        Write-Host ""
        Write-Host "  Scanning working dir..." -ForegroundColor Gray
        $workingFiles = @()
        foreach ($remotePath in $WorkingPaths) {
            $files = Get-RemoteFilesRecursive -Session $session -RemotePath $remotePath -ExcludePaths $PrunePaths
            $clearWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
            if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
            Write-Host "    $remotePath ($($files.Count) files)" -ForegroundColor Gray
            $workingFiles += $files
        }

        # Scan sync tree
        Write-Host ""
        Write-Host "  Scanning sync (library)..." -ForegroundColor Gray
        $syncFiles = @()
        foreach ($remotePath in $SyncPaths) {
            $files = Get-RemoteFilesRecursive -Session $session -RemotePath $remotePath
            $clearWidth = [Math]::Max(0, [Console]::WindowWidth - 1)
            if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
            Write-Host "    $remotePath ($($files.Count) files)" -ForegroundColor Gray
            $syncFiles += $files
        }

        $components = $script:ReleaseComponentSubfolders
        $workingFolders = Group-IntoReleaseFolders -Files $workingFiles -Components $components -VideoExts $VideoExtensions
        $syncFolders    = Group-IntoReleaseFolders -Files $syncFiles    -Components $components -VideoExts $VideoExtensions

        # Belt-and-suspenders: even with the descent-skip above, also filter
        # at the folder level in case any prune-subtree leak slipped through
        # (path normalization edge cases, symlinks, etc.).
        $pruneRoots = @()
        if ($PrunePaths -and $PrunePaths.Count -gt 0) {
            $pruneRoots = @($PrunePaths | ForEach-Object {
                (($_ -replace '\\', '/') -replace '/+', '/').TrimEnd('/')
            })
        }

        # Hard blacklist — these names are aggregator/category folders, never
        # individual releases. Even if a stray video file ends up directly
        # inside one (sample, broken hardlink, etc.) we will not delete the
        # whole container. Worst case the user has to clean it up manually.
        $categoryBlacklist = @('radarr','sonarr','movies','tv','shows','tvshows',
            'music','books','complete','incoming','downloads','rtorrent',
            'seedbox','prowlarr','bazarr','lidarr','readarr','staging','watch')

        $workingCandidates = @{}
        $excludedInPrune = 0
        $excludedAsCategory = 0
        $loosePromoted = 0

        # Index of files by their immediate parent. Used below to recover
        # per-file info for loose videos sitting at the root of category
        # folders — the folder grouping aggregates them into one entry,
        # losing the individual size/mtime needed to match each as its own
        # release.
        $filesByParent = @{}
        foreach ($f in $workingFiles) {
            $normalized = ($f.FullPath -replace '\\', '/') -replace '/+', '/'
            $parent = $normalized.Substring(0, $normalized.LastIndexOf('/'))
            if (-not $filesByParent.ContainsKey($parent)) {
                $filesByParent[$parent] = @()
            }
            $filesByParent[$parent] += $f
        }

        foreach ($k in $workingFolders.Keys) {
            $wf = $workingFolders[$k]
            if ($wf.VideoFileCount -eq 0) { continue }

            # Category-container guard. The folder itself never goes into the
            # delete list (we'd nuke /rtorrent/ otherwise), but any loose
            # video file living directly inside one is still individually
            # prunable — match each on its own size and delete just that
            # file. Common case: old completed downloads that bypassed
            # rTorrent's move-on-completion and ended up at the working-tree
            # root.
            if ($wf.Leaf -and $categoryBlacklist -contains $wf.Leaf.ToLower()) {
                $excludedAsCategory++
                $folderPathNorm = (($wf.Path -replace '\\', '/') -replace '/+', '/').TrimEnd('/')
                if ($filesByParent.ContainsKey($folderPathNorm)) {
                    foreach ($file in $filesByParent[$folderPathNorm]) {
                        $ext = [System.IO.Path]::GetExtension($file.Name).ToLower()
                        if ($VideoExtensions -notcontains $ext) { continue }
                        # Use the file path itself as the candidate key —
                        # deletion later will see IsLooseFile=$true and
                        # remove just this file, leaving the parent folder
                        # alone.
                        $workingCandidates[$file.FullPath] = [PSCustomObject]@{
                            Path            = $file.FullPath
                            Leaf            = $file.Name
                            PrimarySize     = $file.Size
                            PrimaryName     = $file.Name
                            TotalVideoBytes = $file.Size
                            VideoFileCount  = 1
                            NewestMtime     = $file.LastModified
                            IsLooseFile     = $true
                        }
                        $loosePromoted++
                    }
                }
                continue
            }

            $folderPath = (($wf.Path -replace '\\', '/') -replace '/+', '/').TrimEnd('/')
            $insidePrune = $false
            foreach ($pr in $pruneRoots) {
                if ($folderPath -eq $pr -or $folderPath.StartsWith($pr + '/')) {
                    $insidePrune = $true
                    break
                }
            }
            if ($insidePrune) { $excludedInPrune++; continue }
            $workingCandidates[$k] = $wf
        }

        # Build sync size index. Multiple sync folders can share a primary size
        # (rare but possible) — we just need to know if ANY match exists.
        $syncBySize = @{}
        foreach ($k in $syncFolders.Keys) {
            $sf = $syncFolders[$k]
            if ($sf.VideoFileCount -eq 0) { continue }
            if (-not $syncBySize.ContainsKey($sf.PrimarySize)) {
                $syncBySize[$sf.PrimarySize] = New-Object System.Collections.Generic.List[object]
            }
            $syncBySize[$sf.PrimarySize].Add($sf)
        }

        # Build a local-library size index. Recursive walk — every (non-junk)
        # video file under any LocalLibraryPath contributes its size. For
        # movies (one video per folder) this matches the previous behavior.
        # For TV shows (multiple episodes nested in season subfolders) it
        # also indexes each episode individually, so working-dir release
        # folders containing a single episode can match against any local
        # episode by exact byte count.
        $localBySize = @{}
        if ($LocalLibraryPaths -and $LocalLibraryPaths.Count -gt 0) {
            $junkRegex = '(?i)(-trailer|\.trailer|-sample|\.sample|-featurette|behindthescenes|extras?)$'
            foreach ($root in $LocalLibraryPaths) {
                if (-not $root -or -not (Test-Path -LiteralPath $root)) { continue }
                Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
                    Where-Object {
                        $VideoExtensions -contains $_.Extension.ToLower() -and
                        $_.BaseName -notmatch $junkRegex
                    } | ForEach-Object {
                        $file = $_
                        if (-not $localBySize.ContainsKey([long]$file.Length)) {
                            $localBySize[[long]$file.Length] = New-Object System.Collections.Generic.List[object]
                        }
                        $localBySize[[long]$file.Length].Add([PSCustomObject]@{
                            Path = $file.Directory.FullName
                            Leaf = $file.Directory.Name
                            PrimarySize = $file.Length
                            PrimaryName = $file.Name
                        })
                    }
            }
        }

        # Match. Eligible if either side has a primary-video size hit.
        $deletable = @()
        $unmatched = @()
        $tooRecent = @()
        $matchedRemote = 0
        $matchedLocal = 0
        $ageCutoff = if ($DaysOld -gt 0) { (Get-Date).AddDays(-$DaysOld) } else { $null }
        foreach ($k in $workingCandidates.Keys) {
            $wf = $workingCandidates[$k]
            $remoteHit = $syncBySize.ContainsKey($wf.PrimarySize)
            $localHit  = $localBySize.ContainsKey($wf.PrimarySize)
            if ($remoteHit -or $localHit) {
                # Age guard: a fresh release is presumed to still be seeding.
                # Skip even though the library has it.
                if ($ageCutoff -and $wf.NewestMtime -gt $ageCutoff) {
                    $tooRecent += $wf
                    continue
                }
                $matchSource = if ($remoteHit) { $syncBySize[$wf.PrimarySize][0] } else { $localBySize[$wf.PrimarySize][0] }
                $matchOrigin = if ($remoteHit) { 'remote' } else { 'local' }
                if ($remoteHit) { $matchedRemote++ } else { $matchedLocal++ }
                $deletable += [PSCustomObject]@{
                    Working     = $wf
                    SyncMatch   = $matchSource
                    MatchOrigin = $matchOrigin
                }
            } else {
                $unmatched += $wf
            }
        }

        Write-Host ""
        Write-Host "  Working folders considered:    $($workingCandidates.Count)" -ForegroundColor Gray
        if ($excludedInPrune -gt 0) {
            Write-Host "    ($excludedInPrune skipped — inside prune subtree, regular prune handles)" -ForegroundColor DarkGray
        }
        if ($excludedAsCategory -gt 0) {
            $catNote = "$excludedAsCategory category-container folder(s) skipped"
            if ($loosePromoted -gt 0) {
                $catNote += " — $loosePromoted loose video(s) inside them promoted to per-file candidates"
            }
            Write-Host "    ($catNote)" -ForegroundColor DarkGray
        }
        Write-Host "  With library size match:       $($deletable.Count)" -ForegroundColor Green
        if (($matchedRemote + $matchedLocal + $tooRecent.Count) -gt 0) {
            $localCount = $localBySize.Count
            $remoteCount = $syncBySize.Count
            Write-Host "    -- seedbox library index: $remoteCount folders | local library index: $localCount unique file sizes (incl. TV episodes)" -ForegroundColor DarkGray
            if ($matchedLocal -gt 0) {
                Write-Host "    -- matched against seedbox: $matchedRemote | matched against local: $matchedLocal" -ForegroundColor DarkGray
            }
        }
        if ($tooRecent.Count -gt 0) {
            Write-Host "  Too recent (< $DaysOld days, kept): $($tooRecent.Count)" -ForegroundColor Cyan
            # Oldest-first so the user sees what's closest to becoming
            # eligible for prune at the top of the list.
            foreach ($wf in ($tooRecent | Sort-Object { $_.NewestMtime } | Select-Object -First 10)) {
                $ageDays = [math]::Floor(((Get-Date) - $wf.NewestMtime).TotalDays)
                $truncated = if ($wf.Leaf.Length -gt 60) { $wf.Leaf.Substring(0, 57) + "..." } else { $wf.Leaf }
                Write-Host "    - $truncated ($ageDays day(s) old)" -ForegroundColor DarkCyan
            }
            if ($tooRecent.Count -gt 10) {
                Write-Host "    ... and $($tooRecent.Count - 10) more" -ForegroundColor DarkGray
            }
        }
        Write-Host "  Without match (left alone):    $($unmatched.Count)" -ForegroundColor DarkYellow
        Write-Host ""

        if ($deletable.Count -eq 0) {
            Write-Host "  Nothing to prune from the working dir." -ForegroundColor Green
            Write-Host ""
            return @{ Deleted = 0; Failed = 0; Skipped = 0; TooRecent = $tooRecent.Count }
        }

        # Show the candidates
        $totalBytes = 0L
        Write-Host "  Eligible for deletion:" -ForegroundColor Yellow
        foreach ($d in ($deletable | Sort-Object { $_.Working.Leaf })) {
            $totalBytes += $d.Working.PrimarySize
            $matchLabel = if ($d.MatchOrigin -eq 'local') { 'local library' } else { 'seedbox library' }
            $pathLabel  = if ($d.MatchOrigin -eq 'local') { 'local  ' } else { 'library' }
            Write-Host "    $($d.Working.Leaf)" -ForegroundColor White
            Write-Host "      $(Format-SyncSize $d.Working.PrimarySize) — matched in $matchLabel" -ForegroundColor DarkGray
            Write-Host "      working: $($d.Working.Path)" -ForegroundColor DarkGray
            Write-Host "      $pathLabel : $($d.SyncMatch.Path)" -ForegroundColor DarkGray
        }
        Write-Host ""
        Write-Host "  Total to delete: $($deletable.Count) folder(s), $(Format-SyncSize $totalBytes)" -ForegroundColor Cyan
        Write-Host ""

        if ($WhatIf) {
            Write-Host "  [DRY RUN] No folders deleted." -ForegroundColor Yellow
            return @{ Deleted = 0; Failed = 0; Skipped = $deletable.Count; TooRecent = $tooRecent.Count; WhatIf = $true; Eligible = $deletable.Count }
        }

        # Delete each matched folder. Folder candidates get the trailing-slash
        # + EscapeFileMask form so WinSCP treats them as directory entries
        # (not a wildcard mask), even when the release name contains brackets.
        # Loose-file candidates (promoted from category folders above) get the
        # plain file path so only that specific file is removed — the parent
        # folder (e.g. /rtorrent/) stays put.
        $deleted = 0
        $failed = 0
        $sessionLost = $false
        foreach ($d in $deletable) {
            # If a previous iteration killed the session, every subsequent
            # RemoveFiles call would fail with "Session is not opened". Stop
            # immediately rather than emitting a wall of failure noise.
            if ($sessionLost -or -not $session.Opened) {
                $sessionLost = $true
                break
            }
            if ($d.Working.IsLooseFile) {
                $targetPath = $d.Working.Path
                $kindLabel  = "File"
            } else {
                $targetPath = $d.Working.Path.TrimEnd('/') + '/'
                $kindLabel  = "Deleted"
            }
            $escaped = [WinSCP.RemotePath]::EscapeFileMask($targetPath)
            try {
                $session.RemoveFiles($escaped).Check()
                Write-Host "  $($kindLabel): $($d.Working.Leaf)" -ForegroundColor Green
                $deleted++
            } catch {
                $errMsg = $_.Exception.Message
                Write-Host "  FAILED: $($d.Working.Leaf) — $errMsg" -ForegroundColor Red
                $failed++
                # Session-lost / WinSCP-IPC-timeout failure modes are terminal —
                # the session will not recover and every subsequent delete will
                # cascade-fail. Flag for the next iteration to break.
                if ($errMsg -match 'Session is not opened|Timeout waiting for WinSCP') {
                    $sessionLost = $true
                }
            }
        }

        $notAttempted = $deletable.Count - $deleted - $failed
        Write-Host ""
        if ($sessionLost) {
            Write-Host "  Session lost — aborted. $notAttempted folder(s) not attempted." -ForegroundColor Red
            Write-Host "  Re-run the prune to continue from where it stopped." -ForegroundColor DarkGray
        }
        $recentMsg = if ($tooRecent.Count -gt 0) { "   Too recent: $($tooRecent.Count)" } else { "" }
        Write-Host "  Deleted: $deleted   Failed: $failed   Not attempted: $notAttempted   Left alone: $($unmatched.Count)$recentMsg" -ForegroundColor Cyan
        Write-Host ""

        return @{
            Deleted = $deleted
            Failed = $failed
            NotAttempted = $notAttempted
            SessionLost = $sessionLost
            Skipped = $unmatched.Count
            TooRecent = $tooRecent.Count
            Eligible = $deletable.Count
        }
    } finally {
        $session.Dispose()
    }
}

#endregion

#region Remote unrar extraction
# Test scaffolding for handling multi-part RAR releases on the seedbox.
# Stays out of Invoke-SFTPSync for now — invoke Invoke-SFTPExtractRarReleases
# manually (via the SFTP menu) to verify before integrating into auto-sync.

<#
.SYNOPSIS
    Inspects top-level files of a remote folder to determine whether it's a
    multi-part RAR release that needs unrar to produce a playable video.
.DESCRIPTION
    A folder qualifies as a RAR release when:
      - One or more top-level files match the RAR-chain pattern
        (`.rar`, `.r##`, or `.part##.rar`)
      - No playable top-level video file of meaningful size exists
        (i.e., not just a `Sample/` in a subfolder)

    First archive in the chain is selected by precedence:
      1. `<name>.rar` (chain entry-point when present)
      2. `<name>.part01.rar`
      3. `<name>.r00` (lowest-numbered .r##)
.OUTPUTS
    Hashtable: IsRarRelease (bool), and when true, FirstArchive (full path),
    FirstArchiveName (leaf), PartCount (int), ChainPattern (rar/partN/rNN),
    Basename (stem common to the chain).
#>
function Get-RarReleaseInfo {
    param(
        [array]$TopLevelFiles,
        [string[]]$VideoExtensions = @(".mkv", ".mp4", ".avi", ".m4v"),
        [long]$MinVideoBytes = 50MB
    )

    # A playable top-level video disqualifies the folder — extraction isn't
    # needed if the release already has its content sitting there ready.
    foreach ($f in $TopLevelFiles) {
        $ext = [System.IO.Path]::GetExtension($f.Name).ToLower()
        if ($VideoExtensions -contains $ext -and $f.Size -ge $MinVideoBytes) {
            return @{ IsRarRelease = $false }
        }
    }

    $rarRegex = '^(?<base>.+?)\.(?<ext>rar|r\d{2,}|part\d+\.rar)$'
    $rarFiles = @()
    foreach ($f in $TopLevelFiles) {
        if ($f.Name -match $rarRegex) {
            $rarFiles += [PSCustomObject]@{
                File = $f
                Base = $Matches.base
                Ext  = $Matches.ext.ToLower()
            }
        }
    }
    if ($rarFiles.Count -eq 0) {
        return @{ IsRarRelease = $false }
    }

    # Pick chain entry-point. Prefer `<name>.rar`, then `partN.rar`, then `.r##`.
    $bareRar = $rarFiles | Where-Object { $_.Ext -eq 'rar' } | Select-Object -First 1
    if ($bareRar) {
        $first = $bareRar
        $pattern = 'rar'
    } else {
        $partFiles = @($rarFiles | Where-Object { $_.Ext -match '^part\d+\.rar$' } |
            Sort-Object { [int]([regex]::Match($_.Ext, 'part(\d+)\.rar').Groups[1].Value) })
        if ($partFiles.Count -gt 0) {
            $first = $partFiles[0]
            $pattern = 'partN'
        } else {
            $rNNFiles = @($rarFiles | Where-Object { $_.Ext -match '^r\d+$' } |
                Sort-Object { [int]([regex]::Match($_.Ext, 'r(\d+)').Groups[1].Value) })
            if ($rNNFiles.Count -eq 0) {
                return @{ IsRarRelease = $false }
            }
            $first = $rNNFiles[0]
            $pattern = 'rNN'
        }
    }

    # Completeness sanity-check on the chain BEFORE extraction. Two cheap
    # checks that don't require reading file contents:
    #
    #   1. Zero-byte parts — sparse-allocated paused/incomplete downloads
    #      leave individual parts at 0 bytes. Preallocate-style would still
    #      pass this check (size reflects allocation, not actual data),
    #      but the user's paused-rtorrent case typically shows 0-byte parts.
    #
    #   2. Numbering gaps in `.r##` and `.partNN.rar` chains — if part 02
    #      is missing between 01 and 03, that file simply doesn't exist on
    #      disk yet.
    #
    # If either fires, we report Complete=$false with a reason so the
    # orchestrator can skip extraction (which would have failed fast with
    # an unhelpful unrar error anyway).
    $complete = $true
    $incompleteReason = $null

    # Check 1: zero-byte parts
    $zeroParts = @($rarFiles | Where-Object { $_.File.Size -eq 0 })
    if ($zeroParts.Count -gt 0) {
        $complete = $false
        $names = ($zeroParts | ForEach-Object { $_.File.Name } | Select-Object -First 3) -join ', '
        $more = if ($zeroParts.Count -gt 3) { " +$($zeroParts.Count - 3) more" } else { '' }
        $incompleteReason = "$($zeroParts.Count) zero-byte part(s): $names$more"
    }

    # Check 2: sequential numbering gaps. Only meaningful for chains where
    # we can index parts by integer (r## and partNN.rar). The .rar entry-
    # point doesn't carry its own number — it implicitly precedes r00.
    if ($complete) {
        if ($pattern -eq 'rNN') {
            $indices = @($rarFiles | Where-Object { $_.Ext -match '^r\d+$' } |
                ForEach-Object { [int]([regex]::Match($_.Ext, 'r(\d+)').Groups[1].Value) } |
                Sort-Object)
            if ($indices.Count -gt 0) {
                $expected = $indices[0]
                foreach ($idx in $indices) {
                    if ($idx -ne $expected) {
                        $complete = $false
                        $incompleteReason = "missing part(s) in r## chain — jump from r$('{0:00}' -f ($expected - 1)) to r$('{0:00}' -f $idx)"
                        break
                    }
                    $expected++
                }
            }
        } elseif ($pattern -eq 'partN') {
            $indices = @($rarFiles | Where-Object { $_.Ext -match '^part\d+\.rar$' } |
                ForEach-Object { [int]([regex]::Match($_.Ext, 'part(\d+)\.rar').Groups[1].Value) } |
                Sort-Object)
            if ($indices.Count -gt 0) {
                $expected = $indices[0]
                foreach ($idx in $indices) {
                    if ($idx -ne $expected) {
                        $complete = $false
                        $incompleteReason = "missing part(s) in partN chain — jump from part$($expected - 1) to part$idx"
                        break
                    }
                    $expected++
                }
            }
        }
    }

    return @{
        IsRarRelease     = $true
        FirstArchive     = $first.File.FullPath
        FirstArchiveName = $first.File.Name
        PartCount        = $rarFiles.Count
        ChainPattern     = $pattern
        Basename         = $first.Base
        Complete         = $complete
        IncompleteReason = $incompleteReason
    }
}

<#
.SYNOPSIS
    Scans remote paths and returns the folders that look like multi-part RAR
    releases — candidates for remote extraction.
.DESCRIPTION
    Uses Get-RemoteFilesRecursive to walk the tree, groups files by their
    immediate parent directory, and runs Get-RarReleaseInfo against each
    folder's top-level files. Sample/ subfolders are naturally segregated
    because Get-RemoteFilesRecursive returns them under their own parent.
#>
function Find-SFTPRarReleases {
    param(
        $Session,
        [string[]]$RemotePaths,
        [string[]]$VideoExtensions = @(".mkv", ".mp4", ".avi", ".m4v"),
        [long]$MinVideoBytes = 50MB
    )

    $allFiles = @()
    foreach ($p in $RemotePaths) {
        $allFiles += Get-RemoteFilesRecursive -Session $Session -RemotePath $p
    }

    # Group by parent folder (forward-slash semantics — these are remote POSIX paths).
    $byParent = @{}
    foreach ($f in $allFiles) {
        $normalized = ($f.FullPath -replace '\\', '/')
        $idx = $normalized.LastIndexOf('/')
        if ($idx -le 0) { continue }
        $parent = $normalized.Substring(0, $idx)
        if (-not $byParent.ContainsKey($parent)) { $byParent[$parent] = @() }
        $byParent[$parent] += $f
    }

    $releases = @()
    foreach ($folder in $byParent.Keys) {
        $info = Get-RarReleaseInfo -TopLevelFiles $byParent[$folder] `
            -VideoExtensions $VideoExtensions -MinVideoBytes $MinVideoBytes
        if ($info.IsRarRelease) {
            $obj = [PSCustomObject]@{
                Folder           = $folder
                Leaf             = Split-Path $folder -Leaf
                FirstArchive     = $info.FirstArchive
                FirstArchiveName = $info.FirstArchiveName
                PartCount        = $info.PartCount
                ChainPattern     = $info.ChainPattern
                Complete         = $info.Complete
                IncompleteReason = $info.IncompleteReason
                Basename         = $info.Basename
            }
            $releases += $obj
        }
    }
    return $releases
}

<#
.SYNOPSIS
    Single-quotes a string for safe inclusion in a POSIX shell command.
.DESCRIPTION
    Wraps in single quotes and escapes any embedded single quote as `'\''`.
    Use for every variable interpolated into a shell string — release names
    with spaces, brackets, hyphens, etc. all pass through cleanly.
#>
function ConvertTo-PosixShellArg {
    param([string]$Value)
    return "'$($Value -replace "'", "'\''")'"
}

<#
.SYNOPSIS
    Executes unrar on the seedbox to extract a multi-part RAR chain into a
    destination folder, leaving the original torrent files (still seeding)
    untouched.
.DESCRIPTION
    Uses WinSCP's Session.ExecuteCommand for SSH-level shell execution.
    The remote command creates the destination folder if needed, cds into
    it, and runs `unrar x -inul -o+ <first archive>`. Switches:
      -inul  : suppress unrar's interactive output
      -o+    : overwrite-on-existing (idempotent re-run is safe)
.OUTPUTS
    Hashtable: Success (bool), ExitStatus (int), Output (string),
    ErrorOutput (string), Command (the literal remote command we ran).
#>
function Invoke-SFTPRemoteUnrar {
    param(
        $Session,
        [string]$FirstArchive,
        [string]$DestinationDir,
        [string]$UnrarBinary = 'unrar'
    )

    $qDest = ConvertTo-PosixShellArg $DestinationDir
    $qArchive = ConvertTo-PosixShellArg $FirstArchive
    # unrar binary itself is not quoted — usually a bare command name. If the
    # user configures it with a path containing spaces, they need to quote it
    # themselves in config.
    #
    # Switches:
    #   x      - extract with full paths
    #   -idq   - quiet mode but keep error messages (vs `-inul` which silences
    #            errors too and would have us reporting "FAILED" with no clue)
    #   -o+    - overwrite existing files (idempotent re-runs)
    #   -y     - assume yes on all prompts (don't wait for input on a missing
    #            volume — fail fast instead)
    $cmd = "mkdir -p $qDest && cd $qDest && $UnrarBinary x -idq -o+ -y $qArchive"

    try {
        $result = $Session.ExecuteCommand($cmd)
        return @{
            Success     = ($result.ExitStatus -eq 0)
            ExitStatus  = $result.ExitStatus
            Output      = $result.Output
            ErrorOutput = $result.ErrorOutput
            Command     = $cmd
        }
    } catch {
        return @{
            Success     = $false
            ExitStatus  = -1
            Output      = ''
            ErrorOutput = $_.ToString()
            Command     = $cmd
        }
    }
}

<#
.SYNOPSIS
    Tells Radarr to scan a folder for downloaded movies and import any that
    match monitored entries. Used after a successful remote unrar so the
    fresh .mkv lands in the library and the Radarr queue item self-clears.
.DESCRIPTION
    POSTs Radarr's `DownloadedMoviesScan` command with the target path.
    Radarr scans the path, parses video files against monitored movies,
    moves/copies/hardlinks matches into its library, and clears the queue
    entry that originated the download.

    PATH CAVEAT: the path is sent verbatim. If Radarr is on the same machine
    as the file, it works as-is. If Radarr is remote (your case: Radarr local
    + seedbox at /home/nullpointr/...), you need Radarr's Remote Path
    Mappings (Settings > Download Clients) to translate. Any unmapped path
    surfaces as Radarr's own "path is not a valid directory" error here.
.OUTPUTS
    Hashtable: Success (bool), CommandId (int) on success, or Error string.
#>
function Invoke-RadarrDownloadedScan {
    param(
        [Parameter(Mandatory=$true)][string]$RadarrUrl,
        [Parameter(Mandatory=$true)][string]$ApiKey,
        [Parameter(Mandatory=$true)][string]$Path,
        [ValidateSet('Move', 'Copy', 'Hardlink', 'Auto')]
        [string]$ImportMode = 'Auto'
    )

    $headers = @{ 'X-Api-Key' = $ApiKey }
    $body = @{
        name       = 'DownloadedMoviesScan'
        path       = $Path
        importMode = $ImportMode
    } | ConvertTo-Json -Depth 3

    try {
        $response = Invoke-RestMethod -Uri "$($RadarrUrl.TrimEnd('/'))/api/v3/command" `
            -Method Post -Headers $headers -Body $body -ContentType 'application/json' -ErrorAction Stop
        return @{
            Success   = $true
            CommandId = $response.id
            Status    = $response.status
        }
    } catch {
        # Try to peel out the Radarr-side error body — its JSON has a clear
        # "errorMessage" field that's far more useful than the .NET wrapper's
        # generic exception text.
        $errText = $_.ToString()
        try {
            if ($_.Exception.Response) {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $errBody = $reader.ReadToEnd()
                if ($errBody) {
                    $parsed = $errBody | ConvertFrom-Json -ErrorAction SilentlyContinue
                    if ($parsed.errorMessage) { $errText = $parsed.errorMessage }
                    elseif ($parsed.message)  { $errText = $parsed.message }
                }
            }
        } catch {}
        return @{ Success = $false; Error = $errText }
    }
}

<#
.SYNOPSIS
    Test entry-point: connect to the seedbox, find multi-part RAR releases,
    extract each remotely via unrar. Standalone for now — not invoked by
    Invoke-SFTPSync.
.PARAMETER UnrarBinary
    Remote command/path to unrar. Default 'unrar' assumes it's on PATH.
.PARAMETER ExtractedSuffix
    Appended to the release folder path to form the extraction destination.
    Default '.extracted' yields a sibling folder that's easy to spot and
    leaves the seeded torrent folder untouched.
.PARAMETER WhatIf
    Scan and report only — no remote command is issued.
.PARAMETER NotifyRadarr
    After each successful extraction, POST a DownloadedMoviesScan command
    to Radarr against the .extracted/ folder so Radarr imports the .mkv
    and clears the queue item. Requires RadarrUrl + RadarrApiKey.
.PARAMETER RadarrUrl
    Base URL for Radarr (e.g. http://localhost:7878). Required when
    NotifyRadarr is set.
.PARAMETER RadarrApiKey
    Radarr API key. Required when NotifyRadarr is set.
#>
function Invoke-SFTPExtractRarReleases {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$HostName,
        [int]$Port = 22,
        [Parameter(Mandatory=$true)][string]$Username,
        [string]$Password,
        [string]$PrivateKeyPath,
        [Parameter(Mandatory=$true)][string[]]$RemotePaths,
        [string]$UnrarBinary = 'unrar',
        [string]$ExtractedSuffix = '.extracted',
        [switch]$NotifyRadarr,
        [string]$RadarrUrl,
        [string]$RadarrApiKey,
        [switch]$WhatIf
    )

    if ($NotifyRadarr -and (-not $RadarrUrl -or -not $RadarrApiKey)) {
        Write-Host "  NotifyRadarr requested but RadarrUrl/RadarrApiKey not set — disabling." -ForegroundColor Yellow
        $NotifyRadarr = $false
    }

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "       SFTP REMOTE UNRAR (TEST)                       " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] No extractions will be performed" -ForegroundColor Yellow
        Write-Host ""
    }

    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        return @{ Extracted = 0; Failed = 0; Error = "WinSCP .NET assembly not installed" }
    }

    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ Extracted = 0; Failed = 0; Error = $_.ToString() }
    }

    try {
        # Probe that the server allows shell exec and that unrar is available.
        # Cheap up-front check beats a wall of identical extraction failures.
        #
        # We use `command -v` (POSIX shell builtin) rather than running unrar
        # itself — `unrar -?` writes help to stderr in most versions and the
        # exit code varies, so checking output is unreliable. `command -v`
        # prints the resolved path on stdout for installed commands and exits
        # with non-zero otherwise.
        #
        # If the configured name isn't on the non-interactive PATH (common on
        # seedboxes where unrar lives in ~/bin/), we fall back to common
        # absolute paths. Whichever one resolves becomes the binary we use
        # for the rest of the run.
        Write-Host "  Probing remote unrar..." -ForegroundColor Gray -NoNewline

        $resolvedUnrar = $null

        # Wrap in try/catch because ExecuteCommand can throw if the server
        # rejects the exec channel entirely (SFTP-only chroots, etc.).
        try {
            $probe = $session.ExecuteCommand("command -v $UnrarBinary 2>/dev/null")
            $resolvedUnrar = ($probe.Output).Trim()
        } catch {
            Write-Host " failed" -ForegroundColor Red
            Write-Host "  Session.ExecuteCommand error: $_" -ForegroundColor Red
            Write-Host "  Your SFTP server may not allow SSH command execution." -ForegroundColor Yellow
            return @{ Extracted = 0; Failed = 0; Error = "Session.ExecuteCommand failed: $_" }
        }

        if (-not $resolvedUnrar) {
            # Fallback: scan common absolute paths. Most seedboxes have unrar
            # at one of these. ~/bin is per-user (Ultra.cc style); $HOME is
            # expanded by the remote shell.
            $candidates = @('/usr/bin/unrar', '/usr/local/bin/unrar', '/bin/unrar', '$HOME/bin/unrar', '$HOME/.local/bin/unrar')
            foreach ($candidate in $candidates) {
                try {
                    $test = $session.ExecuteCommand("test -x $candidate && echo $candidate")
                    $hit = ($test.Output).Trim()
                    if ($hit) { $resolvedUnrar = $hit; break }
                } catch {}
            }
        }

        if (-not $resolvedUnrar) {
            Write-Host " not found" -ForegroundColor Red
            Write-Host "  unrar isn't on PATH or at common locations." -ForegroundColor Yellow
            Write-Host "  Find it on the seedbox (e.g. 'which unrar' or 'ls ~/bin') and set" -ForegroundColor Yellow
            Write-Host "  SFTPUnrarCommand to the full path." -ForegroundColor Yellow
            return @{ Extracted = 0; Failed = 0; Error = "unrar not available on seedbox" }
        }

        # Get a version line for human confirmation. unrar writes its banner
        # to stderr OR stdout depending on version, so capture both.
        try {
            $ver = $session.ExecuteCommand("$resolvedUnrar -V 2>&1 | head -n 1")
            $verLine = ($ver.Output).Trim()
            if ($verLine) {
                Write-Host " $verLine [at $resolvedUnrar]" -ForegroundColor Green
            } else {
                Write-Host " found at $resolvedUnrar" -ForegroundColor Green
            }
        } catch {
            Write-Host " found at $resolvedUnrar" -ForegroundColor Green
        }

        # If the resolved path differs from the configured command, use it for
        # the rest of this run. The user can update SFTPUnrarCommand to make
        # it stick.
        if ($resolvedUnrar -ne $UnrarBinary) {
            $UnrarBinary = $resolvedUnrar
        }

        Write-Host "  Scanning: $($RemotePaths -join ', ')" -ForegroundColor Gray
        $releases = Find-SFTPRarReleases -Session $session -RemotePaths $RemotePaths

        # Dedup releases that surface twice because two configured scan paths
        # point at the same physical directory (e.g. Ultra.cc's /home and
        # /home16 mounts, or overlapping prune/working paths). Same release
        # name + same archive count = same release; keep the shortest folder
        # path (typically the canonical one) and drop the rest.
        if ($releases.Count -gt 1) {
            $bySignature = @{}
            foreach ($r in $releases) {
                $sig = "$($r.Basename)|$($r.PartCount)|$($r.FirstArchiveName)"
                if (-not $bySignature.ContainsKey($sig)) {
                    $bySignature[$sig] = $r
                } elseif ($r.Folder.Length -lt $bySignature[$sig].Folder.Length) {
                    $bySignature[$sig] = $r
                }
            }
            $deduped = @($bySignature.Values)
            if ($deduped.Count -lt $releases.Count) {
                Write-Host "  Deduplicated $($releases.Count - $deduped.Count) release(s) found under multiple scan paths" -ForegroundColor DarkGray
            }
            $releases = $deduped
        }

        if ($releases.Count -eq 0) {
            Write-Host "  No multi-part RAR releases found." -ForegroundColor Green
            Write-Host ""
            return @{ Extracted = 0; Failed = 0; Skipped = 0 }
        }

        # Split into extractable vs. incomplete so the user sees what's being
        # skipped (and why) without it counting as a noisy failure. Paused
        # downloads with 0-byte parts or chain-numbering gaps fail fast in
        # unrar anyway with unhelpful errors; skip them up-front.
        $incomplete = @($releases | Where-Object { -not $_.Complete })
        $extractable = @($releases | Where-Object { $_.Complete })

        Write-Host ""
        if ($incomplete.Count -gt 0) {
            Write-Host "  $($incomplete.Count) release(s) skipped — incomplete RAR set:" -ForegroundColor Yellow
            foreach ($r in $incomplete) {
                Write-Host "    $($r.Leaf)" -ForegroundColor DarkYellow
                Write-Host "      $($r.IncompleteReason)" -ForegroundColor DarkGray
                Write-Host "      folder: $($r.Folder)" -ForegroundColor DarkGray
            }
            Write-Host ""
        }

        if ($extractable.Count -eq 0) {
            Write-Host "  No complete RAR releases to extract." -ForegroundColor Green
            Write-Host ""
            return @{ Extracted = 0; Failed = 0; Skipped = 0; Incomplete = $incomplete.Count }
        }

        Write-Host "  Found $($extractable.Count) complete RAR release(s):" -ForegroundColor Cyan
        foreach ($r in $extractable) {
            Write-Host "    $($r.Leaf)" -ForegroundColor White
            Write-Host "      $($r.PartCount) part(s), chain: $($r.ChainPattern), entry: $($r.FirstArchiveName)" -ForegroundColor DarkGray
            Write-Host "      folder: $($r.Folder)" -ForegroundColor DarkGray
        }
        Write-Host ""

        if ($WhatIf) {
            Write-Host "  [DRY RUN] No extractions performed." -ForegroundColor Yellow
            return @{ Extracted = 0; Failed = 0; Skipped = 0; Incomplete = $incomplete.Count; WhatIf = $true; Eligible = $extractable.Count }
        }

        # The extraction loop uses $releases for the per-release iteration —
        # narrow to extractable only.
        $releases = $extractable

        $extracted = 0
        $failed = 0
        $skipped = 0

        # Helper closure: returns the first video file in $destDir, or $null.
        # Used both for idempotency (already extracted on a prior run) and for
        # confirming an extraction succeeded — we trust the presence of an
        # output video over unrar's exit code, which WinSCP's ExecuteCommand
        # doesn't always receive cleanly after long-running shell exec.
        $getExtractedVideo = {
            param($Session, $Path)
            try {
                $listing = $Session.ListDirectory($Path)
                $videos = @($listing.Files | Where-Object {
                    -not $_.IsDirectory -and ($_.Name -match '\.(mkv|mp4|avi|m4v)$')
                } | Sort-Object Length -Descending)
                if ($videos.Count -gt 0) { return $videos[0] }
            } catch {
                # Path missing or not listable — caller treats as "not present".
            }
            return $null
        }

        foreach ($r in $releases) {
            $destDir = ($r.Folder -replace '\\', '/').TrimEnd('/') + $ExtractedSuffix

            Write-Host "  $($r.Leaf)" -ForegroundColor White

            # Idempotency: skip if a prior run already extracted here.
            $existing = & $getExtractedVideo $session $destDir
            if ($existing) {
                Write-Host "    Already extracted: $($existing.Name) — skipped" -ForegroundColor DarkGray
                $skipped++
                continue
            }

            Write-Host "    Extracting to $destDir ..." -ForegroundColor Yellow
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

            # ExecuteCommand can throw "Element session@0 already read to the
            # end" after a long-running command leaves the IPC channel in a
            # bad state. Catch it here so the loop can recover for the next
            # release instead of crashing.
            $result = $null
            $execThrew = $null
            try {
                $result = Invoke-SFTPRemoteUnrar -Session $session `
                    -FirstArchive $r.FirstArchive -DestinationDir $destDir `
                    -UnrarBinary $UnrarBinary
            } catch {
                $execThrew = $_
            }
            $stopwatch.Stop()
            $elapsed = $stopwatch.Elapsed.ToString('hh\:mm\:ss')

            # Long-running ExecuteCommand calls corrupt the WinSCP IPC channel
            # — the session technically stays "open" but subsequent commands
            # throw. Reconnect proactively so the NEXT iteration starts fresh
            # AND we can list the destination to verify what just happened.
            try { $session.Dispose() } catch {}
            try {
                $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
                    -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
            } catch {
                Write-Host "    Reconnect failed: $_" -ForegroundColor Red
                Write-Host "    Aborting remaining releases." -ForegroundColor Red
                $failed++
                break
            }

            # Verify by side effect: did we get a video file? This is more
            # trustworthy than $result.ExitStatus, which can be $null even on
            # a successful run if WinSCP didn't receive the SSH exit-status
            # packet.
            $produced = & $getExtractedVideo $session $destDir
            if ($produced) {
                Write-Host "    OK ($elapsed) -> $($produced.Name)" -ForegroundColor Green
                $extracted++

                if ($NotifyRadarr) {
                    # Ask Radarr to scan the .extracted/ folder so it imports
                    # the freshly-extracted .mkv and clears its queue entry.
                    $radarrResult = Invoke-RadarrDownloadedScan -RadarrUrl $RadarrUrl `
                        -ApiKey $RadarrApiKey -Path $destDir
                    if ($radarrResult.Success) {
                        Write-Host "      Radarr scan queued (cmd id=$($radarrResult.CommandId))" -ForegroundColor Cyan
                    } else {
                        Write-Host "      Radarr scan request failed: $($radarrResult.Error)" -ForegroundColor Yellow
                        Write-Host "      (Check Radarr's Remote Path Mappings if it can't see $destDir)" -ForegroundColor DarkYellow
                    }
                }
            } else {
                Write-Host "    FAILED ($elapsed)" -ForegroundColor Red
                if ($execThrew) {
                    Write-Host "      $execThrew" -ForegroundColor DarkRed
                } elseif ($result) {
                    if ($null -ne $result.ExitStatus -and $result.ExitStatus -ne '') {
                        Write-Host "      exit status: $($result.ExitStatus)" -ForegroundColor DarkRed
                    }
                    if ($result.ErrorOutput) {
                        $errLines = ($result.ErrorOutput -split "`n" | Where-Object { $_.Trim() } | Select-Object -First 3)
                        foreach ($line in $errLines) {
                            Write-Host "      $line" -ForegroundColor DarkRed
                        }
                    }
                }
                $failed++
            }
        }

        $incompleteMsg = if ($incomplete.Count -gt 0) { "   Incomplete: $($incomplete.Count)" } else { "" }
        Write-Host ""
        Write-Host "======================================================" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Extracted: $extracted   Failed: $failed   Skipped (already done): $skipped$incompleteMsg" -ForegroundColor Cyan
        Write-Host ""

        return @{
            Extracted  = $extracted
            Failed     = $failed
            Skipped    = $skipped
            Incomplete = $incomplete.Count
            Eligible   = $releases.Count
        }
    } finally {
        $session.Dispose()
    }
}

#endregion

#region rTorrent control
# Talks to rTorrent's XML-RPC interface so LibraryLint can clean up dead
# torrent entries (downloads whose data was already pruned). rTorrent here
# exposes XML-RPC only on a local Unix socket (network.scgi.open_local =
# ~/.config/rtorrent/socket) — no TCP port, and the host has no rtxmlrpc/
# rtcontrol CLI. So we ship a small Python3 SCGI client and run it over the
# same SSH channel the unrar feature uses.

# Python3 SCGI/XML-RPC client. Sent to the seedbox base64-encoded (avoids all
# shell-quoting problems) and run as `python3 - <action> [hash]`. Emits a
# single JSON line on stdout: {ok, torrents|erased|error}.
$script:RTorrentScgiPython = @'
import socket, os, sys, json, xmlrpc.client

SOCKET_PATH = os.path.expanduser("~/.config/rtorrent/socket")

def scgi_call(xml_body):
    body = xml_body.encode("utf-8")
    # SCGI: netstring(headers) + body. CONTENT_LENGTH must be first;
    # an SCGI=1 header is required.
    headers = b"CONTENT_LENGTH\x00" + str(len(body)).encode() + b"\x00SCGI\x001\x00"
    request = str(len(headers)).encode() + b":" + headers + b"," + body
    s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
    s.settimeout(30)
    try:
        s.connect(SOCKET_PATH)
        s.sendall(request)
        resp = b""
        while True:
            chunk = s.recv(65536)
            if not chunk:
                break
            resp += chunk
    finally:
        s.close()
    # rTorrent replies with CGI-style headers, blank line, then the XML body.
    for sep in (b"\r\n\r\n", b"\n\n"):
        if sep in resp:
            return resp.split(sep, 1)[1]
    return resp

def call(method, params):
    xml = xmlrpc.client.dumps(tuple(params), method)
    result, _ = xmlrpc.client.loads(scgi_call(xml))
    return result[0]

def main():
    action = sys.argv[1] if len(sys.argv) > 1 else "list"
    try:
        if action == "list":
            rows = call("d.multicall2", ["", "main",
                "d.hash=", "d.name=", "d.base_path=", "d.is_open=",
                "d.complete=", "d.bytes_done=", "d.size_bytes=",
                "d.message=", "d.state="])
            out = []
            for r in rows:
                bp = r[2]
                out.append({
                    "hash": r[0], "name": r[1], "base_path": bp,
                    "is_open": r[3], "complete": r[4],
                    "bytes_done": r[5], "size_bytes": r[6],
                    "message": r[7], "state": r[8],
                    # Stat the data path here — saves a second round trip.
                    # A torrent whose base_path is gone has had its data
                    # pruned and is the safe-to-erase "dead" case.
                    "path_exists": bool(bp) and os.path.exists(bp),
                })
            print(json.dumps({"ok": True, "torrents": out}))
        elif action == "erase":
            if len(sys.argv) < 3:
                print(json.dumps({"ok": False, "error": "erase requires a hash"}))
                return
            h = sys.argv[2]
            call("d.erase", [h])
            print(json.dumps({"ok": True, "erased": h}))
        else:
            print(json.dumps({"ok": False, "error": "unknown action: " + action}))
    except Exception as e:
        print(json.dumps({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}))

main()
'@

<#
.SYNOPSIS
    Returns whether a path falls at or under any of the given root paths,
    tolerant of the seedbox's two namespace projections.
.DESCRIPTION
    rTorrent (inside the chroot) reports data paths as `/home/nullpointr/...`
    while LibraryLint's configured prune roots are SFTP-side
    `/home16/nullpointr/...`. They're the same disk. This collapses any
    `/home<digits>` prefix to a canonical `/home` on both sides before the
    prefix comparison so a torrent's base path can be matched against a
    prune root.
#>
function Test-RTorrentPathUnderRoots {
    param([string]$Path, [string[]]$Roots)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $normPath = ($Path -replace '^/home\d*/', '/home/')
    foreach ($r in $Roots) {
        if (-not $r) { continue }
        $normRoot = (($r -replace '^/home\d*/', '/home/')).TrimEnd('/')
        if ($normPath -eq $normRoot -or $normPath.StartsWith($normRoot + '/')) {
            return $true
        }
    }
    return $false
}

<#
.SYNOPSIS
    Runs an rTorrent XML-RPC action over SSH via the embedded Python SCGI
    client. Low-level helper behind Get-SeedboxTorrents / Remove-SeedboxTorrent.
.PARAMETER Session
    An open WinSCP session (Connect-SFTPSession).
.PARAMETER Action
    'list' — enumerate all torrents; 'erase' — remove one by hash.
.PARAMETER Hash
    Info-hash for the 'erase' action.
.OUTPUTS
    PSCustomObject parsed from the Python's JSON: .ok plus .torrents / .erased
    / .error depending on outcome.
#>
function Invoke-RTorrentCommand {
    param(
        $Session,
        [ValidateSet('list', 'erase')]
        [string]$Action,
        [string]$Hash
    )

    $pyB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($script:RTorrentScgiPython))
    $argPart = if ($Action -eq 'erase') { "erase $Hash" } else { 'list' }
    # base64 → decode → python3 reading script from stdin, args after the `-`
    $cmd = "echo $pyB64 | base64 -d | python3 - $argPart"

    try {
        $res = $Session.ExecuteCommand($cmd)
        $stdout = if ($res.Output) { $res.Output.Trim() } else { '' }
        if (-not $stdout) {
            $stderr = if ($res.ErrorOutput) { $res.ErrorOutput.Trim() } else { '<none>' }
            return [PSCustomObject]@{ ok = $false; error = "no stdout (stderr: $stderr)" }
        }
        # The Python prints exactly one JSON line; guard against any stray
        # leading shell noise by taking the last non-empty line.
        $jsonLine = ($stdout -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 1)
        return ($jsonLine | ConvertFrom-Json)
    } catch {
        return [PSCustomObject]@{ ok = $false; error = "ExecuteCommand/parse failed: $_" }
    }
}

<#
.SYNOPSIS
    Lists every torrent in the seedbox's rTorrent session.
.DESCRIPTION
    Each returned object carries hash, name, base_path, completion/state
    fields, and a path_exists flag (computed on the seedbox) — path_exists
    = $false marks a "dead" torrent whose data has already been pruned.
.OUTPUTS
    Hashtable: Success (bool), Torrents (array) on success, or Error (string).
#>
function Get-SeedboxTorrents {
    param($Session)

    $result = Invoke-RTorrentCommand -Session $Session -Action 'list'
    if (-not $result.ok) {
        return @{ Success = $false; Error = $result.error }
    }
    return @{ Success = $true; Torrents = @($result.torrents) }
}

<#
.SYNOPSIS
    Removes a single torrent from the rTorrent session by info-hash.
.DESCRIPTION
    Calls rTorrent's d.erase. This removes the torrent ENTRY from the
    session — it does not delete data files (the prune already does that).
    Erasing a torrent stops it seeding, so callers must only erase
    torrents that are past their seed obligation or whose data is gone.
.OUTPUTS
    Hashtable: Success (bool), Hash (string) on success, or Error (string).
#>
function Remove-SeedboxTorrent {
    param(
        $Session,
        [Parameter(Mandatory=$true)]
        [string]$Hash
    )

    $result = Invoke-RTorrentCommand -Session $Session -Action 'erase' -Hash $Hash
    if (-not $result.ok) {
        return @{ Success = $false; Hash = $Hash; Error = $result.error }
    }
    return @{ Success = $true; Hash = $Hash }
}

<#
.SYNOPSIS
    Removes "dead" torrent entries from rTorrent — torrents whose data files
    no longer exist on the seedbox (already pruned by LibraryLint or deleted
    otherwise).
.DESCRIPTION
    A dead torrent is one where rTorrent reports no usable data path
    (path_exists = false): the data is gone, so the torrent is no longer
    seeding anything and the entry is pure clutter in the client. Erasing it
    is inherently hit-and-run-safe — you cannot seed data you do not have.

    Torrents with data still present are never touched. This is the standalone
    cleanup pass; it does not coordinate with the prune (a future prune-
    integrated path will erase torrents as their data is pruned).
.PARAMETER WhatIf
    List the dead torrents without erasing anything.
.OUTPUTS
    Hashtable: Erased, Failed, DeadFound, TotalTorrents counts (or Error).
#>
function Invoke-SeedboxDeadTorrentCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$HostName,
        [int]$Port = 22,
        [Parameter(Mandatory=$true)][string]$Username,
        [string]$Password,
        [string]$PrivateKeyPath,
        [switch]$WhatIf
    )

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "       SEEDBOX — CLEAN DEAD TORRENTS                  " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Removes rTorrent entries whose data files are gone" -ForegroundColor Gray
    Write-Host "  (already pruned). Torrents with data present are" -ForegroundColor Gray
    Write-Host "  left alone — only dead entries are erased." -ForegroundColor Gray
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] No torrents will be erased" -ForegroundColor Yellow
        Write-Host ""
    }

    $modulePath = Split-Path $PSScriptRoot -Parent
    $winscpPath = Test-WinSCPInstalled -ModulePath $modulePath
    if (-not $winscpPath) {
        Write-Host "  WinSCP .NET assembly not found!" -ForegroundColor Red
        return @{ Erased = 0; Failed = 0; Error = "WinSCP .NET assembly not installed" }
    }

    Write-Host "  Connecting..." -ForegroundColor Gray -NoNewline
    try {
        $session = Connect-SFTPSession -DllPath $winscpPath -HostName $HostName -Port $Port `
            -Username $Username -Password $Password -PrivateKeyPath $PrivateKeyPath
        Write-Host " connected" -ForegroundColor Green
    } catch {
        Write-Host " failed" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        return @{ Erased = 0; Failed = 0; Error = $_.ToString() }
    }

    try {
        Write-Host "  Querying rTorrent session..." -ForegroundColor Gray
        $listResult = Get-SeedboxTorrents -Session $session
        if (-not $listResult.Success) {
            Write-Host "  Could not list torrents: $($listResult.Error)" -ForegroundColor Red
            return @{ Erased = 0; Failed = 0; Error = $listResult.Error }
        }

        $allTorrents = @($listResult.Torrents)
        $dead = @($allTorrents | Where-Object { -not $_.path_exists })

        Write-Host "  Torrents in session: $($allTorrents.Count)  |  data missing (dead): $($dead.Count)" -ForegroundColor Gray
        Write-Host ""

        if ($dead.Count -eq 0) {
            Write-Host "  No dead torrents — nothing to clean up." -ForegroundColor Green
            Write-Host ""
            return @{ Erased = 0; Failed = 0; DeadFound = 0; TotalTorrents = $allTorrents.Count }
        }

        Write-Host "  Dead torrents (data already gone):" -ForegroundColor Yellow
        foreach ($t in ($dead | Sort-Object name)) {
            $name = if ($t.name.Length -gt 64) { $t.name.Substring(0, 61) + "..." } else { $t.name }
            Write-Host "    $name" -ForegroundColor White
            Write-Host "      hash: $($t.hash)" -ForegroundColor DarkGray
        }
        Write-Host ""

        if ($WhatIf) {
            Write-Host "  [DRY RUN] Would erase $($dead.Count) dead torrent entry(ies)." -ForegroundColor Yellow
            Write-Host ""
            return @{ Erased = 0; Failed = 0; DeadFound = $dead.Count; TotalTorrents = $allTorrents.Count; WhatIf = $true }
        }

        Write-Host "  Erasing $($dead.Count) dead torrent entry(ies)..." -ForegroundColor Yellow
        Write-Host ""
        $erased = 0
        $failed = 0
        foreach ($t in $dead) {
            $name = if ($t.name.Length -gt 55) { $t.name.Substring(0, 52) + "..." } else { $t.name }
            $rm = Remove-SeedboxTorrent -Session $session -Hash $t.hash
            if ($rm.Success) {
                Write-Host "    erased: $name" -ForegroundColor Green
                $erased++
            } else {
                Write-Host "    FAILED: $name — $($rm.Error)" -ForegroundColor Red
                $failed++
            }
        }

        Write-Host ""
        Write-Host "======================================================" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Erased: $erased   Failed: $failed   (of $($dead.Count) dead, $($allTorrents.Count) total)" -ForegroundColor Cyan
        Write-Host ""

        return @{
            Erased        = $erased
            Failed        = $failed
            DeadFound     = $dead.Count
            TotalTorrents = $allTorrents.Count
        }
    } finally {
        $session.Dispose()
    }
}

#endregion

# Export public functions
Export-ModuleMember -Function Invoke-SFTPSync, Invoke-SFTPPrune, Invoke-SFTPPruneWorkingDir, Initialize-SFTPTracking, Update-SFTPTrackingFromLocal, Get-SFTPNewFiles, Find-SFTPIncompleteFiles, Test-WinSCPInstalled, Connect-SFTPSession, Get-RemoteFilesRecursive, Invoke-FileDownload, Get-DownloadedFiles, Save-DownloadedFiles, Get-SyncTrackingPath, Format-SyncSize, Get-RarReleaseInfo, Find-SFTPRarReleases, Invoke-SFTPRemoteUnrar, Invoke-SFTPExtractRarReleases, Invoke-RadarrDownloadedScan, Get-SeedboxTorrents, Remove-SeedboxTorrent, Invoke-SeedboxDeadTorrentCleanup
