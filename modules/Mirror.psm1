# Mirror.psm1
# Mirrors media folders to backup destination using robocopy
# Part of LibraryLint suite

#region Private Functions

function Format-MirrorSize {
    param([long]$Bytes)

    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

function Format-TimeSpan {
    param([int]$Seconds)

    if ($Seconds -lt 60) {
        return "${Seconds}s"
    } elseif ($Seconds -lt 3600) {
        $min = [math]::Floor($Seconds / 60)
        $sec = $Seconds % 60
        return "${min}m ${sec}s"
    } else {
        $hr = [math]::Floor($Seconds / 3600)
        $min = [math]::Floor(($Seconds % 3600) / 60)
        return "${hr}h ${min}m"
    }
}

function Get-DriveSpace {
    param([string]$Path)

    try {
        # Handle both local drives and UNC paths
        if ($Path -match '^\\\\') {
            # UNC path - try to get space using .NET
            $drive = [System.IO.DriveInfo]::GetDrives() | Where-Object {
                $Path.StartsWith($_.Name, [StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1

            if (-not $drive) {
                # For network paths, try using a temporary file approach or just return null
                return $null
            }
            return @{
                FreeSpace = $drive.AvailableFreeSpace
                TotalSpace = $drive.TotalSize
            }
        } else {
            # Local drive
            $driveLetter = $Path.Substring(0, 2)
            $diskInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -ErrorAction SilentlyContinue
            if ($diskInfo) {
                return @{
                    FreeSpace = $diskInfo.FreeSpace
                    TotalSpace = $diskInfo.Size
                }
            }
        }
    } catch {
        # Silently fail - space check is nice to have but not required
    }
    return $null
}

function Connect-MirrorShare {
    # Authenticates the SMB share that backs a UNC mirror destination
    # using stored credentials. Mirrors the inline net-use block from
    # the menu handler so the Status pre-flight (Get-MirrorPendingChanges)
    # can reach a credential-gated share. Returns @{Authenticated, Skipped,
    # ShareRoot, Error}. Skipped=$true when the dest isn't UNC or no creds
    # were provided — caller proceeds with whatever ambient auth Windows
    # offers (current Windows session credentials).
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$DestDrive,
        [string]$NetworkUser,
        [string]$NetworkPass
    )

    $result = @{ Authenticated = $false; Skipped = $false; ShareRoot = $null; Error = $null }

    $isUnc = $DestDrive -match '^\\\\'
    if (-not $isUnc -or -not $NetworkUser) {
        $result.Skipped = $true
        return $result
    }

    $shareRoot = ($DestDrive -replace '^(\\\\[^\\]+\\[^\\]+).*', '$1')
    $result.ShareRoot = $shareRoot

    # Drop ALL existing connections to this server first. Windows SMB
    # rejects multiple sessions to the same server with different creds
    # (System error 1219), so a stale Explorer mapping or any prior run
    # blocks an explicit-creds attempt with a misleading "unreachable".
    if ($shareRoot -match '^\\\\([^\\]+)') {
        $serverName = $Matches[1]
        $serverPattern = '\\\\' + [regex]::Escape($serverName) + '\\\S+'
        $netUseListing = (& net use 2>&1) | Out-String
        $existingShares = [regex]::Matches($netUseListing, $serverPattern) |
            ForEach-Object { $_.Value } |
            Select-Object -Unique
        foreach ($existingShare in $existingShares) {
            & net use $existingShare /delete /y *>&1 | Out-Null
        }
    }

    $netOutput = & net use $shareRoot /user:$NetworkUser $NetworkPass 2>&1
    if ($LASTEXITCODE -eq 0) {
        $result.Authenticated = $true
    } else {
        $result.Error = ($netOutput | Out-String).Trim()
    }
    return $result
}

function Test-MirrorDestAlive {
    # Watchdog probe for the mirror destination. Deliberately avoids
    # Test-Path for UNC roots — Test-Path on a dead SMB session can itself
    # hang for minutes, which would defeat the watchdog. For UNC we TCP-
    # probe the server's SMB port with a hard 5s cap; for local paths a
    # plain Test-Path is safe.
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$DestRoot)

    if ($DestRoot -match '^\\\\([^\\]+)') {
        $server = $Matches[1]
        $tcp = $null
        try {
            $tcp = [System.Net.Sockets.TcpClient]::new()
            $async = $tcp.BeginConnect($server, 445, $null, $null)
            if ($async.AsyncWaitHandle.WaitOne(5000)) {
                $tcp.EndConnect($async)
                return $true
            }
            return $false
        } catch {
            return $false
        } finally {
            if ($tcp) { $tcp.Dispose() }
        }
    }
    return (Test-Path -LiteralPath $DestRoot)
}

function Get-MirrorNetworkBytesSent {
    # Sums BytesSent across all non-loopback "Up" adapters. We use this as the
    # speed signal instead of parsing robocopy's stdout because robocopy with
    # /MT:16 emits per-file announcements in bursts — the parser systematically
    # lags real disk writes (each new-file line treats the previous as done,
    # but with 16 parallel threads many files are in flight simultaneously).
    # The adapter counter is what Task Manager reads, so the displayed speed
    # matches what the user sees in Task Manager. Bytes from unrelated traffic
    # (browser, updates) are included, but during a mirror run the SMB copy is
    # the dominant flow and the noise is negligible.
    try {
        $total = [long]0
        foreach ($nic in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
            if ($nic.OperationalStatus -ne 'Up') { continue }
            if ($nic.NetworkInterfaceType -eq 'Loopback') { continue }
            $total += $nic.GetIPv4Statistics().BytesSent
        }
        return $total
    } catch {
        return [long]0
    }
}

function ConvertFrom-RobocopyOutput {
    param([string[]]$Output)

    $stats = @{
        FilesCopied = 0
        FilesSkipped = 0
        FilesDeleted = 0
        FilesFailed = 0
        BytesCopied = 0
    }

    foreach ($line in $Output) {
        # Parse the summary lines
        if ($line -match "^\s*Files\s*:\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)") {
            $stats.FilesCopied = [int]$Matches[2]
            $stats.FilesSkipped = [int]$Matches[3]
            $stats.FilesFailed = [int]$Matches[5]
        }
        if ($line -match "^\s*Bytes\s*:\s*([\d.]+)\s*([tgmk]?)\s+([\d.]+)\s*([tgmk]?)") {
            $value = [double]$Matches[3]
            $unit = $Matches[4]
            $stats.BytesCopied = switch ($unit) {
                "t" { $value * 1TB }
                "g" { $value * 1GB }
                "m" { $value * 1MB }
                "k" { $value * 1KB }
                default { $value }
            }
        }
        if ($line -match "^\s*Extras\s*:\s*(\d+)") {
            $stats.FilesDeleted = [int]$Matches[1]
        }
    }

    return $stats
}

#endregion

#region Public Functions

<#
.SYNOPSIS
    Mirrors media folders to a backup destination
.DESCRIPTION
    Uses robocopy with /MIR to create exact mirrors of source folders,
    deleting files at destination that don't exist in source.
.PARAMETER SourceDrive
    The source drive letter (e.g., "G:") or UNC path (e.g., "\\NAS\Media")
.PARAMETER DestDrive
    The destination drive letter (e.g., "F:") or UNC path (e.g., "\\NAS\Backup")
.PARAMETER Folders
    Array of folder names to mirror (e.g., @("Movies", "Shows"))
.PARAMETER WhatIf
    Preview changes without making them
.EXAMPLE
    Invoke-Mirror -SourceDrive "G:" -DestDrive "F:" -Folders @("Movies", "Shows")
.EXAMPLE
    Invoke-Mirror -SourceDrive "G:" -DestDrive "\\NAS\Backup" -Folders @("Movies") -WhatIf
#>
function Invoke-Mirror {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourceDrive,

        [Parameter(Mandatory=$true)]
        [string]$DestDrive,

        [Parameter(Mandatory=$true)]
        [string[]]$Folders,

        [switch]$WhatIf
    )

    # Header
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "                    MIRROR BACKUP                      " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Source:      $SourceDrive" -ForegroundColor White
    Write-Host "  Destination: $DestDrive" -ForegroundColor White
    Write-Host "  Folders:     $($Folders -join ', ')" -ForegroundColor White
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN MODE] No changes will be made" -ForegroundColor Yellow
        Write-Host ""
    }

    Write-Host "  Press Ctrl+C to cancel at any time" -ForegroundColor DarkGray
    Write-Host ""

    # Validate paths before doing anything
    Write-Host "--- Path Validation ---" -ForegroundColor Yellow
    Write-Host ""
    $pathsOK = $true

    # Test source drive
    if (Test-Path $SourceDrive) {
        Write-Host "  Source drive:       " -NoNewline; Write-Host "OK ($SourceDrive)" -ForegroundColor Green
    } else {
        Write-Host "  Source drive:       " -NoNewline; Write-Host "NOT FOUND ($SourceDrive)" -ForegroundColor Red
        $pathsOK = $false
    }

    # Test destination drive
    if (Test-Path $DestDrive) {
        Write-Host "  Destination drive:  " -NoNewline; Write-Host "OK ($DestDrive)" -ForegroundColor Green
    } else {
        Write-Host "  Destination drive:  " -NoNewline; Write-Host "NOT FOUND ($DestDrive)" -ForegroundColor Red
        $pathsOK = $false
    }

    # Test each source folder
    foreach ($folder in $Folders) {
        $sourcePath = Join-Path $SourceDrive $folder
        $destPath = Join-Path $DestDrive $folder
        if (Test-Path $sourcePath) {
            Write-Host "  Source\$folder`:     " -NoNewline; Write-Host "OK" -ForegroundColor Green
        } else {
            Write-Host "  Source\$folder`:     " -NoNewline; Write-Host "NOT FOUND ($sourcePath)" -ForegroundColor Red
            $pathsOK = $false
        }
        if (Test-Path $destPath) {
            Write-Host "  Dest\$folder`:       " -NoNewline; Write-Host "OK" -ForegroundColor Green
        } else {
            Write-Host "  Dest\$folder`:       " -NoNewline; Write-Host "WILL BE CREATED" -ForegroundColor Yellow
        }
    }

    Write-Host ""

    if (-not $pathsOK) {
        Write-Host "  Cannot proceed - fix the paths above and try again." -ForegroundColor Red
        Write-Host "  Check that drives are mounted and network shares are accessible." -ForegroundColor Yellow
        Write-Host ""
        return @{
            FilesCopied = 0
            FilesDeleted = 0
            FilesFailed = 0
            BytesCopied = 0
            Cancelled = $true
        }
    }

    # Set up cancellation handling
    $cancelled = $false
    $currentProcess = $null

    $cancelHandler = {
        $script:cancelled = $true
        if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
            $script:currentProcess.Kill()
        }
    }

    [Console]::TreatControlCAsInput = $false
    $null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action $cancelHandler

    # Pre-scan phase — use robocopy /L (list-only) for fast diffing instead of slow PowerShell enumeration
    Write-Host "--- Scanning ---" -ForegroundColor Yellow
    Write-Host ""

    # Robocopy flags optimized for large media libraries over LAN/Samba
    # /MIR    = Mirror mode (sync + delete extras)
    # /R:2    = Retry 2 times (LAN — if it fails twice, retrying won't help)
    # /W:5    = Wait 5 seconds between retries
    # /MT:16  = 16 threads for parallel copying (saturate gigabit LAN with large files)
    # /XJD    = Exclude junction points for directories
    # /NDL    = Don't log directory names
    # /NC     = Don't log file classes
    # /BYTES  = Show sizes in bytes for parsing
    # /COPY:DT = Copy Data and Timestamps only (no Security — needed for network/Samba shares)
    # /DCOPY:T = Copy directory Timestamps only
    # /FFT    = FAT file time granularity (2-second tolerance — prevents false re-copies on Samba/Linux)
    # /NP = No progress (suppresses robocopy's own % and "Removed X of Y" console output;
    #        we calculate progress ourselves from file size lines)
    $robocopyBaseArgs = @("/MIR", "/R:2", "/W:5", "/MT:16", "/XJD", "/NP", "/NDL", "/NC", "/BYTES", "/COPY:DT", "/DCOPY:T", "/FFT")

    $totalFilesToCopy = 0
    $totalBytesToCopy = [long]0
    $totalExtras = 0
    $folderStats = @{}

    foreach ($folder in $Folders) {
        $sourcePath = Join-Path $SourceDrive $folder
        $destPath = Join-Path $DestDrive $folder

        if (-not (Test-Path $sourcePath)) {
            Write-Host "  $folder - Source not found!" -ForegroundColor Red
            continue
        }

        Write-Host ""

        # Use robocopy /L to quickly diff source vs dest without copying
        $scanArgs = "`"$sourcePath`" `"$destPath`" " + (($robocopyBaseArgs + @("/L", "/NP")) -join ' ')
        $scanProcess = New-Object System.Diagnostics.Process
        $scanProcess.StartInfo.FileName = "robocopy"
        $scanProcess.StartInfo.Arguments = $scanArgs
        $scanProcess.StartInfo.UseShellExecute = $false
        $scanProcess.StartInfo.RedirectStandardOutput = $true
        $scanProcess.StartInfo.RedirectStandardError = $true
        $scanProcess.StartInfo.CreateNoWindow = $true
        $scanProcess.Start() | Out-Null

        # Stream output line-by-line with live progress
        $scanLines = @()
        $filesToCopy = 0
        $bytesToCopy = [long]0
        $scanFilesProcessed = 0
        $scanLastUpdate = [DateTime]::MinValue
        $scanStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        while (-not $scanProcess.StandardOutput.EndOfStream) {
            $scanLine = $scanProcess.StandardOutput.ReadLine()
            if ($null -eq $scanLine) { break }
            $scanLines += $scanLine

            # Count files that would be copied
            if ($scanLine -match '^\s+(\d+)\s+.+$') {
                $filesToCopy++
                $bytesToCopy += [long]$Matches[1]
            }

            $scanFilesProcessed++

            # Update progress display every 300ms
            $now = [DateTime]::Now
            if (($now - $scanLastUpdate).TotalMilliseconds -ge 300) {
                $scanLastUpdate = $now
                $elapsed = $scanStopwatch.Elapsed.TotalSeconds
                $rate = if ($elapsed -gt 0) { [math]::Round($scanFilesProcessed / $elapsed) } else { 0 }
                Write-Host "`r  $folder - scanning... $scanFilesProcessed files compared ($rate/s) | $filesToCopy to copy     " -NoNewline -ForegroundColor DarkGray
            }
        }

        if (-not $scanProcess.HasExited) { $scanProcess.WaitForExit() }
        $scanStopwatch.Stop()

        $scanStats = ConvertFrom-RobocopyOutput $scanLines

        $folderStats[$folder] = @{
            FilesToCopy = $filesToCopy
            BytesToCopy = $bytesToCopy
            Extras = $scanStats.FilesDeleted
            Skipped = $scanStats.FilesSkipped
        }

        $totalFilesToCopy += $filesToCopy
        $totalBytesToCopy += $bytesToCopy
        $totalExtras += $scanStats.FilesDeleted

        # Clear the live scan line completely before writing final result.
        # Pad the visible portion only — the leading `\r` is a control char
        # that counts toward .PadRight() width and would eat the separator
        # for any folder name >=17 chars ("Movie Set Artwork" was rendering
        # as "Movie Set Artwork5 to copy" with no space before the count).
        Write-Host "`r$(' ' * 120)" -NoNewline
        Write-Host ("`r  " + "$folder ".PadRight(22)) -NoNewline
        if ($filesToCopy -eq 0 -and $scanStats.FilesDeleted -eq 0) {
            Write-Host "up to date ($($scanStats.FilesSkipped) files)" -ForegroundColor Green
        } else {
            $parts = @()
            if ($filesToCopy -gt 0) { $parts += "$filesToCopy to copy ($(Format-MirrorSize $bytesToCopy))" }
            if ($scanStats.FilesDeleted -gt 0) { $parts += "$($scanStats.FilesDeleted) to delete" }
            Write-Host ($parts -join " | ") -ForegroundColor Yellow
        }
    }

    Write-Host ""
    if ($totalFilesToCopy -eq 0 -and $totalExtras -eq 0) {
        Write-Host "  Everything is in sync!" -ForegroundColor Green
    } else {
        if ($totalFilesToCopy -gt 0) {
            Write-Host "  To copy:   $totalFilesToCopy files ($(Format-MirrorSize $totalBytesToCopy))" -ForegroundColor Cyan
        }
        if ($totalExtras -gt 0) {
            Write-Host "  To delete: $totalExtras files (no longer in source)" -ForegroundColor Magenta
        }
    }

    # Check destination drive capacity
    $destSpace = Get-DriveSpace $DestDrive
    if ($destSpace) {
        $freePercent = [math]::Round(($destSpace.FreeSpace / $destSpace.TotalSpace) * 100, 1)
        Write-Host "  Dest space: $(Format-MirrorSize $destSpace.FreeSpace) free ($freePercent%)" -ForegroundColor $(if ($freePercent -lt 10) { 'Red' } elseif ($freePercent -lt 20) { 'Yellow' } else { 'Gray' })

        if ($totalBytesToCopy -gt $destSpace.FreeSpace) {
            Write-Host ""
            Write-Host "  WARNING: Not enough space!" -ForegroundColor Red
            Write-Host "  Need: $(Format-MirrorSize $totalBytesToCopy) | Free: $(Format-MirrorSize $destSpace.FreeSpace)" -ForegroundColor Red
            Write-Host ""
            $continue = Read-Host "  Continue anyway? (Y/N) [N]"
            if ($continue -notmatch '^[Yy]') {
                Write-Host "  Cancelled." -ForegroundColor Yellow
                return @{
                    FilesCopied = 0
                    FilesDeleted = 0
                    FilesFailed = 0
                    BytesCopied = 0
                    Cancelled = $true
                }
            }
        }
    }

    # Skip mirror phase if nothing to do
    if ($totalFilesToCopy -eq 0 -and $totalExtras -eq 0) {
        Write-Host ""
        $overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $overallStopwatch.Stop()
        $grandTotalCopied = 0; $grandTotalSkipped = 0; $grandTotalDeleted = 0; $grandTotalFailed = 0; $grandTotalBytesCopied = 0
    } else {

    Write-Host ""

    # Mirror phase
    Write-Host "--- Mirroring ---" -ForegroundColor Yellow
    Write-Host ""

    if ($WhatIf) {
        $robocopyBaseArgs += "/L"
    }

    $grandTotalCopied = 0
    $grandTotalSkipped = 0
    $grandTotalDeleted = 0
    $grandTotalFailed = 0
    $grandTotalBytesCopied = 0
    $overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $folderIndex = 0

    foreach ($folder in $Folders) {
        $folderIndex++
        $source = Join-Path $SourceDrive $folder
        $dest = Join-Path $DestDrive $folder

        if (-not (Test-Path $source)) {
            continue
        }

        # Read this folder's planned work BEFORE the header so we don't show
        # the previous iteration's leftover counts (Shows was rendering
        # "(2,494 files, 0 bytes)" — that was Movies' count from iteration 1).
        $folderSize = if ($folderStats[$folder]) { $folderStats[$folder].BytesToCopy } else { 0 }
        $folderTotalFiles = if ($folderStats[$folder]) { $folderStats[$folder].FilesToCopy } else { 0 }

        Write-Host "  [$folderIndex/$($Folders.Count)] $folder -> $dest" -NoNewline -ForegroundColor Yellow
        if ($folderTotalFiles -gt 0) {
            Write-Host " ($($folderTotalFiles.ToString('N0')) files, $(Format-MirrorSize $folderSize))" -ForegroundColor DarkGray
        } else {
            Write-Host " (syncing)" -ForegroundColor DarkGray
        }

        $folderStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $folderBytesCopied = 0
        $folderFilesCopied = 0
        $deleteErrors = 0
        $copyErrors = 0
        $pendingErrorFile = $null
        $lastProgressUpdate = [DateTime]::MinValue
        $currentFileName = ""
        $currentFileSize = 0
        $currentFilePct = 0

        # Rolling-window throughput sampled from the NIC bytes-sent counter
        # (same source Task Manager reads). We previously sampled the robocopy-
        # announcement-driven $effectiveBytes, but with /MT:16 those announce-
        # ments are bursty and systematically lag real disk writes — the
        # cumulative-avg fallback then trended down through the whole copy
        # while Task Manager held steady. NIC bytes are wall-clock authoritative.
        $speedSamples       = New-Object 'System.Collections.Generic.Queue[object]'
        $speedWindowSec     = 5.0
        $smoothedSpeed      = 0.0
        $folderInitialBytes = Get-MirrorNetworkBytesSent
        $folderInitialTime  = [DateTime]::Now

        # Dead-destination watchdog: if throughput stays below the stall
        # threshold for the full window, probe the destination host. A dest
        # that died mid-copy (HTPC idle-shutdown, ejected drive) leaves
        # robocopy hung on a dead SMB session for up to an hour — the
        # watchdog converts that into a 90-second detection + clean abort.
        $stallThresholdBps  = 100KB
        $stallWindowSec     = 90
        $stallSince         = $null
        $destLost           = $false
        $aliveStallStrikes  = 0
        $loopError          = $null

        # Every file robocopy announces this run, with its dest mapping.
        # Used by the cancel path to repair timestamp drift: a killed
        # robocopy never stamps source mtimes onto completed copies, so
        # the server stamps close-time instead and every subsequent scan
        # re-flags those files ("Older") forever.
        $announcedFiles     = [System.Collections.Generic.List[object]]::new()

        # Run robocopy
        $outputLines = @()
        $preCopyRow = [Console]::CursorTop  # Save position to clean up robocopy's direct console writes

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo.FileName = "robocopy"
        $process.StartInfo.Arguments = "`"$source`" `"$dest`" " + (($robocopyBaseArgs | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' ')
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true
        $process.StartInfo.CreateNoWindow = $true

        # Async output collection. Robocopy's /NP flag suppresses per-file %
        # lines (we set it on purpose — it stops robocopy's own \r-rewriting
        # from fighting our progress display), so without a side-channel,
        # the bar would only tick when files complete. With async we can
        # poll the destination file's size BETWEEN robocopy lines and use
        # actual bytes-on-disk for the in-flight file's progress.
        $outputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $outputSub = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -MessageData $outputQueue -Action {
            if ($null -ne $EventArgs.Data) {
                $Event.MessageData.Enqueue($EventArgs.Data)
            }
        }

        $process.Start() | Out-Null
        $process.BeginOutputReadLine()
        $currentProcess = $process

        # Track the dest path of the currently-displayed in-flight file so
        # we can poll its size for smooth progress during big copies.
        $currentDestFilePath = $null

        try {
            while (-not $process.HasExited -or -not $outputQueue.IsEmpty) {
                # Check for key press (Ctrl+C or Q to quit). KeyAvailable can
                # throw in some host states (VSCode terminal resize/focus
                # events, detached console) — treat that as "no key", never
                # as an error: before this guard, such a throw escaped to the
                # loop's catch and was misreported as a user cancel.
                $keyAvailable = try { [Console]::KeyAvailable } catch { $false }
                if ($keyAvailable) {
                    $key = [Console]::ReadKey($true)
                    if ($key.Key -eq 'Q' -or ($key.Modifiers -eq 'Control' -and $key.Key -eq 'C')) {
                        $cancelled = $true
                        Write-Host ""
                        Write-Host "  Cancelling..." -ForegroundColor Yellow
                        if (-not $process.HasExited) {
                            $process.Kill()
                        }
                        break
                    }
                }

                # Drain any queued robocopy output. Same per-line parsing as
                # before, just sourced from the queue instead of ReadLine().
                $processedLine = $false
                $line = $null
                while ($outputQueue.TryDequeue([ref]$line)) {
                    $processedLine = $true
                    $outputLines += $line

                    # Show errors, retries, and warnings
                    if ($line -match 'ERROR.*Deleting Extra File|ERROR.*Access is denied') {
                        # Silently count delete permission errors (common on Samba/Kodi shares)
                        $deleteErrors++
                    }
                    elseif ($line -match 'ERROR\s+\d+\s*\(0x[0-9A-Fa-f]+\)\s+Copying File\s+(.+)') {
                        # File copy error — capture filename, wait for reason on next line(s)
                        $pendingErrorFile = $Matches[1].Trim()
                    }
                    elseif ($line -match 'ERROR\s+\d+\s*\(0x[0-9A-Fa-f]+\)\s+(.+)') {
                        # Non-file error (generic)
                        Write-Host "`r$(' ' * 120)" -NoNewline
                        Write-Host "`r       ERROR: $($Matches[1])" -ForegroundColor Red
                        $pendingErrorFile = $null
                    }
                    elseif ($pendingErrorFile -and $line -match '(The process cannot access|The network path|The specified network|network name is no longer available|Access is denied|being used by another process)') {
                        # Reason line following a file copy error — combine into one message
                        $shortFile = Split-Path $pendingErrorFile -Leaf
                        $reason = $line.Trim()
                        Write-Host "`r$(' ' * 120)" -NoNewline
                        Write-Host "`r       ERROR: $shortFile - $reason" -ForegroundColor Red
                        $copyErrors++
                        $pendingErrorFile = $null
                    }
                    elseif ($line -match '(The process cannot access|The network path|The specified network|network name is no longer available)') {
                        Write-Host "`r$(' ' * 120)" -NoNewline
                        Write-Host "`r       ERROR: $($line.Trim())" -ForegroundColor Red
                    }
                    elseif ($line -match 'Waiting\s+(\d+)\s+seconds') {
                        Write-Host "`r$(' ' * 120)" -NoNewline
                        Write-Host "`r       Waiting $($Matches[1])s before retry..." -ForegroundColor DarkYellow
                    }
                    elseif ($line -match 'Retrying\.\.\.') {
                        Write-Host "`r$(' ' * 120)" -NoNewline
                        Write-Host "`r       Retrying..." -ForegroundColor DarkYellow
                    }
                    # Parse per-file progress lines (format: "  <percentage>%" — only
                    # emitted if /NP is off; left in for forward compat).
                    elseif ($line -match '^\s+([\d\.]+)%') {
                        $currentFilePct = [math]::Min(100, [double]$Matches[1])
                    }
                    # Parse new file lines (format: "   <size> <path>")
                    elseif ($line -match '^\s+(\d+)\s+(.+)$') {
                        # Previous file finished — count it
                        if ($currentFileSize -gt 0) {
                            $folderBytesCopied += $currentFileSize
                            $folderFilesCopied++
                        }

                        $currentFileSize = [long]$Matches[1]
                        $currentFileName = $Matches[2].Trim()
                        $currentFilePct = 0

                        # Compute the dest path so we can poll its size for
                        # smooth in-file progress. Robocopy emits the SOURCE
                        # full path; substitute the source root with dest.
                        if ($currentFileName.StartsWith($source, [System.StringComparison]::OrdinalIgnoreCase)) {
                            $relativePath = $currentFileName.Substring($source.Length).TrimStart('\', '/')
                            $currentDestFilePath = Join-Path $dest $relativePath
                            $announcedFiles.Add(@{ Source = $currentFileName; Dest = $currentDestFilePath })
                        } else {
                            $currentDestFilePath = $null
                        }
                    }
                }

                # Side-channel: poll the dest file size to update the in-flight
                # file's percentage. With /MT this only tracks the most-recent
                # file announced (other parallel files are invisible until they
                # finish), but it's a huge improvement over the previous
                # behavior where the bar was stuck for the entire duration of
                # a multi-GB copy.
                if ($currentDestFilePath -and $currentFileSize -gt 0) {
                    try {
                        $destItem = Get-Item -LiteralPath $currentDestFilePath -ErrorAction SilentlyContinue
                        if ($destItem -and $destItem.Length -gt 0) {
                            $polledPct = [math]::Min(100, [double]($destItem.Length * 100 / $currentFileSize))
                            # Only let polled value INCREASE the percentage —
                            # never let a stale Get-Item.Length read regress
                            # progress (Windows can report sizes that briefly
                            # lag actual writes during heavy I/O).
                            if ($polledPct -gt $currentFilePct) {
                                $currentFilePct = $polledPct
                            }
                        }
                    } catch { }
                }

                # Update progress display (throttled to reduce flicker)
                $now = [DateTime]::Now
                if (($now - $lastProgressUpdate).TotalMilliseconds -ge 200) {
                    $lastProgressUpdate = $now

                    # Total bytes = completed files + partial progress on current file
                    $effectiveBytes = $folderBytesCopied + [long]($currentFileSize * $currentFilePct / 100)
                    $pctBytes = if ($folderSize -gt 0) { [math]::Min(100, [math]::Round(($effectiveBytes / $folderSize) * 100, 0)) } else { 0 }

                    # Progress bar (25 chars wide)
                    $barFilled = [math]::Floor($pctBytes / 4)
                    $barEmpty = 25 - $barFilled
                    $progressBar = "[" + ("#" * $barFilled) + ("-" * $barEmpty) + "]"

                    # Rolling-window speed sampled from the NIC bytes-sent
                    # counter — same authoritative source Task Manager uses,
                    # so the displayed speed matches what's actually on the
                    # wire. We keep a queue of (wall-clock-time, bytes-sent)
                    # samples and trim entries older than $speedWindowSec.
                    $elapsed = $folderStopwatch.Elapsed.TotalSeconds
                    $eta = ""
                    $speedStr = ""
                    if ($elapsed -gt 1) {
                        $nicBytes    = Get-MirrorNetworkBytesSent
                        $nicElapsed  = ($now - $folderInitialTime).TotalSeconds
                        $bytesOnWire = $nicBytes - $folderInitialBytes
                        if ($bytesOnWire -lt 0) { $bytesOnWire = 0 }  # adapter swap / counter wrap

                        $speedSamples.Enqueue([PSCustomObject]@{ T = $nicElapsed; B = [long]$bytesOnWire })

                        # Trim samples older than the window. Always keep at
                        # least one so we have something to subtract against.
                        while ($speedSamples.Count -gt 1 -and ($nicElapsed - $speedSamples.Peek().T) -gt $speedWindowSec) {
                            $speedSamples.Dequeue() | Out-Null
                        }

                        if ($speedSamples.Count -ge 2) {
                            $oldest = $speedSamples.Peek()
                            $deltaB = $bytesOnWire - $oldest.B
                            if ($deltaB -lt 0) { $deltaB = 0 }
                            $deltaT = $nicElapsed - $oldest.T
                            if ($deltaT -gt 0) {
                                $smoothedSpeed = $deltaB / $deltaT
                            }
                        }

                        if ($smoothedSpeed -gt 0) {
                            # No [math]::Max here: PowerShell binds the Int32
                            # overload from the literal 0 and then overflows
                            # converting a multi-GB byte count — this exact
                            # line crashed (and, under the old swallow-all
                            # catch, masqueraded as "cancelled by user") on
                            # any folder with >2.1 GB left to copy.
                            $remainingBytes = [long]$folderSize - [long]$effectiveBytes
                            if ($remainingBytes -lt 0) { $remainingBytes = 0 }
                            $remaining = [math]::Round($remainingBytes / $smoothedSpeed)
                            if ($remaining -gt 0) {
                                $eta = " ETA $(Format-TimeSpan $remaining)"
                            }
                            $speedStr = "$(Format-MirrorSize ([long]$smoothedSpeed))/s"
                        } else {
                            $speedStr = "warming up"
                        }

                        # Watchdog: throughput below the stall threshold starts
                        # the clock; sustained for the full window -> probe the
                        # destination host. Recovered traffic resets the clock,
                        # so a slow-but-alive transfer never trips it.
                        if ($smoothedSpeed -lt $stallThresholdBps) {
                            if ($null -eq $stallSince) {
                                $stallSince = $now
                            } elseif (($now - $stallSince).TotalSeconds -ge $stallWindowSec) {
                                if (-not (Test-MirrorDestAlive -DestRoot $dest)) {
                                    Write-Host ""
                                    Write-Host "  Destination went offline mid-run ($dest) — aborting mirror." -ForegroundColor Red
                                    Write-Host "  (No traffic for $stallWindowSec s and the host stopped answering.)" -ForegroundColor DarkYellow
                                    $destLost = $true
                                    if (-not $process.HasExited) { $process.Kill() }
                                    break
                                }
                                # Host answers but nothing is moving. A wedged
                                # SMB session (stale handle from an earlier
                                # network drop) presents exactly like this —
                                # port 445 alive, zero bytes forever. Give it
                                # three windows, then abort with a diagnosis
                                # instead of "warming up" until doomsday.
                                $aliveStallStrikes++
                                if ($aliveStallStrikes -ge 3) {
                                    Write-Host ""
                                    Write-Host "  Transfer stalled for $([int](3 * $stallWindowSec / 60)) min although $dest answers — aborting mirror." -ForegroundColor Red
                                    Write-Host "  Likely a stale SMB session. Try: net use \\$(if ($dest -match '^\\\\([^\\]+)') { $Matches[1] } else { 'server' }) /delete, then re-run." -ForegroundColor DarkYellow
                                    $destLost = $true
                                    if (-not $process.HasExited) { $process.Kill() }
                                    break
                                }
                                $stallSince = $now
                            }
                        } else {
                            $stallSince = $null
                            $aliveStallStrikes = 0
                        }
                    }

                    # Current file (truncated)
                    $displayName = if ($currentFileName) { Split-Path $currentFileName -Leaf } else { "scanning..." }
                    $truncName = if ($displayName.Length -gt 40) { $displayName.Substring(0, 37) + "..." } else { $displayName }

                    # File counter (e.g., "1,247/3,485")
                    $fileCounter = if ($folderTotalFiles -gt 0) {
                        "$($folderFilesCopied.ToString('N0'))/$($folderTotalFiles.ToString('N0'))"
                    } else {
                        "$($folderFilesCopied.ToString('N0')) files"
                    }

                    # Build progress line: [####-----] 42% 1.2 GB (1,247/3,485) | 45 MB/s ETA 2m 30s | filename.mkv
                    $progressLine = "       $progressBar $pctBytes% $(Format-MirrorSize $effectiveBytes) ($fileCounter)"
                    if ($speedStr) { $progressLine += " | $speedStr" }
                    if ($eta) { $progressLine += $eta }
                    if ($folderFilesCopied -gt 0 -or $currentFileName) {
                        $progressLine += " | $truncName"
                    }

                    # Read the live console width each tick so a mid-run
                    # terminal resize doesn't leave orphan progress lines
                    # scattered across the screen. The fixed 120-char pad
                    # was wider than narrow terminals, causing the line
                    # to wrap to multiple visual rows — `\r` then only
                    # returned cursor to the LAST wrapped row, and every
                    # subsequent draw left the previous attempt visible.
                    $consoleWidth = try { [Console]::BufferWidth - 1 } catch { 120 }
                    if ($consoleWidth -lt 40) { $consoleWidth = 40 }  # sanity floor
                    if ($progressLine.Length -gt $consoleWidth) {
                        # Truncate rather than wrap — losing the trailing
                        # filename is preferable to corrupting the display.
                        $progressLine = $progressLine.Substring(0, $consoleWidth)
                    } else {
                        $progressLine = $progressLine.PadRight($consoleWidth)
                    }
                    Write-Host "`r$progressLine" -NoNewline -ForegroundColor Cyan
                }

                # Yield CPU briefly when nothing came through this iteration.
                # Without this the loop spins on TryDequeue at full CPU when
                # robocopy is mid-file (no stdout) but not yet exited.
                if (-not $processedLine) {
                    Start-Sleep -Milliseconds 50
                }
            }
        }
        catch {
            # A real exception escaped the copy loop. This used to be
            # silently converted into $cancelled = $true, which printed
            # "Mirror cancelled by user" for failures the user never
            # triggered — masking the actual error entirely. Record it;
            # the post-loop handling below reports it as what it is.
            $loopError = $_
        }
        finally {
            # Tear down the event subscription. Leaking these accumulates
            # over multiple mirror runs in the same PowerShell session.
            if ($outputSub) {
                Unregister-Event -SourceIdentifier $outputSub.Name -ErrorAction SilentlyContinue
                Remove-Job -Id $outputSub.Id -Force -ErrorAction SilentlyContinue
            }
        }

        # Count final file
        if ($currentFileSize -gt 0) {
            $folderBytesCopied += $currentFileSize
            $folderFilesCopied++
        }

        if (-not $process.HasExited) {
            $process.WaitForExit()
        }

        # Drain any output that came in after the dequeue loop's last
        # IsEmpty check. The async OutputDataReceived events deliver lines
        # via a thread-pool callback that can fire after the process has
        # exited; without this trailing drain, robocopy's "Files:" /
        # "Bytes:" summary lines were being missed by the parser and
        # everything got reported as Copied: 0 / Bytes: 0. The parameterless
        # WaitForExit() above is the documented signal that all async
        # handlers have completed, so by here the queue is final.
        $drainLine = $null
        while ($outputQueue.TryDequeue([ref]$drainLine)) {
            $outputLines += $drainLine
        }

        $exitCode = $process.ExitCode
        $folderStopwatch.Stop()
        $currentProcess = $null

        if ($loopError) {
            Write-Host ""
            Write-Host "  Mirror loop failed: $($loopError.Exception.Message)" -ForegroundColor Red
            Write-Host "  (Reported as an error — this was NOT a user cancel.)" -ForegroundColor DarkYellow
            Write-Host "  At: $($loopError.InvocationInfo.PositionMessage)" -ForegroundColor DarkGray
            if (-not $process.HasExited) { try { $process.Kill() } catch {} }
            # Route through the cancel path below so completed files get the
            # same timestamp repair — they shouldn't re-copy next run just
            # because the loop crashed.
            $cancelled = $true
        }

        if ($cancelled -or $destLost) {
            Write-Host ""
            if ($destLost) {
                Write-Host "  Mirror aborted — destination went offline" -ForegroundColor Red
            } elseif ($loopError) {
                Write-Host "  Mirror stopped after loop error (see above)" -ForegroundColor Red
            } else {
                Write-Host "  Mirror cancelled by user" -ForegroundColor Yellow
            }

            # Timestamp repair: a killed robocopy never stamps source mtimes
            # onto files whose data finished copying — the server stamps
            # close-time instead, so every later scan re-flags them ("Older")
            # and re-copies gigabytes that already transferred. For every
            # file announced this run whose dest size matches the source,
            # copy the source mtime across. Skipped when the dest is gone
            # (nothing reachable to repair).
            if (-not $destLost -and $announcedFiles.Count -gt 0) {
                $repaired = 0
                foreach ($af in $announcedFiles) {
                    try {
                        $srcItem = Get-Item -LiteralPath $af.Source -ErrorAction SilentlyContinue
                        $dstItem = Get-Item -LiteralPath $af.Dest -ErrorAction SilentlyContinue
                        if ($srcItem -and $dstItem -and
                            $srcItem.Length -eq $dstItem.Length -and
                            $srcItem.LastWriteTime -ne $dstItem.LastWriteTime) {
                            $dstItem.LastWriteTime = $srcItem.LastWriteTime
                            $repaired++
                        }
                    } catch {
                        # Repair is best-effort — an unreachable file just
                        # stays flagged for the next full run.
                    }
                }
                if ($repaired -gt 0) {
                    Write-Host "  Repaired timestamps on $repaired completed file(s) so they won't re-copy next run." -ForegroundColor Cyan
                }
            }
            Write-Host ""
            break
        }

        # Clear progress line and any robocopy direct-console output (e.g., "Removed X of Y" from /MIR + /MT)
        $postCopyRow = [Console]::CursorTop
        for ($row = $preCopyRow; $row -le $postCopyRow; $row++) {
            [Console]::SetCursorPosition(0, $row)
            Write-Host (' ' * [Math]::Min(120, [Console]::BufferWidth - 1)) -NoNewline
        }
        [Console]::SetCursorPosition(0, $preCopyRow)

        # Parse results
        $stats = ConvertFrom-RobocopyOutput $outputLines

        $grandTotalCopied += $stats.FilesCopied
        $grandTotalSkipped += $stats.FilesSkipped
        $grandTotalDeleted += $stats.FilesDeleted
        $grandTotalFailed += $stats.FilesFailed
        $grandTotalBytesCopied += $stats.BytesCopied

        # Display folder results
        $statusIcon = if ($exitCode -ge 8) { "X" } elseif ($stats.FilesCopied -gt 0) { "+" } else { "=" }
        $statusColor = if ($exitCode -ge 8) { "Red" } elseif ($stats.FilesCopied -gt 0) { "Green" } else { "Gray" }

        # Show error details when robocopy fails
        if ($exitCode -ge 8) {
            $errText = $process.StandardError.ReadToEnd()
            if ($errText) {
                Write-Host "       ERROR: $($errText.Trim())" -ForegroundColor Red
            }
            $exitReasons = @{
                8  = "Some files could not be copied (retries exceeded)"
                16 = "Fatal error - no files were copied (check path and permissions)"
            }
            $reason = if ($exitCode -ge 16) { $exitReasons[16] } else { $exitReasons[8] }
            Write-Host "       Robocopy exit code $exitCode`: $reason" -ForegroundColor Red
            if (-not (Test-Path $dest)) {
                Write-Host "       Destination path not reachable: $dest" -ForegroundColor Red
            }
        }

        $duration = [int]$folderStopwatch.Elapsed.TotalSeconds
        $avgSpeed = if ($duration -gt 0 -and $stats.BytesCopied -gt 0) { " @ $(Format-MirrorSize ([long]($stats.BytesCopied / $duration)))/s" } else { "" }
        Write-Host "       $statusIcon Copied: $($stats.FilesCopied)" -ForegroundColor $statusColor -NoNewline
        if ($stats.BytesCopied -gt 0) {
            Write-Host " ($(Format-MirrorSize $stats.BytesCopied)$avgSpeed)" -ForegroundColor $statusColor -NoNewline
        }
        Write-Host " | Skipped: $($stats.FilesSkipped)" -ForegroundColor Gray -NoNewline
        Write-Host " | Deleted: $($stats.FilesDeleted)" -ForegroundColor $(if ($stats.FilesDeleted -gt 0) { "Magenta" } else { "Gray" }) -NoNewline
        if ($stats.FilesFailed -gt 0) {
            Write-Host " | Failed: $($stats.FilesFailed)" -ForegroundColor Red -NoNewline
        }
        Write-Host " | $(Format-TimeSpan $duration)" -ForegroundColor DarkGray
        if ($copyErrors -gt 0) {
            Write-Host "       $copyErrors file(s) skipped (in use by another process)" -ForegroundColor DarkYellow
        }
        if ($deleteErrors -gt 0) {
            Write-Host "       $deleteErrors file(s) could not be deleted (permission denied on dest)" -ForegroundColor DarkYellow
        }
        Write-Host ""
    }

    } # End of else (has work to do)

    $overallStopwatch.Stop()

    # Clean up event handler
    Unregister-Event -SourceIdentifier PowerShell.Exiting -ErrorAction SilentlyContinue

    # Summary
    if ($destLost) {
        Write-Host "--- Aborted (destination offline or stalled) ---" -ForegroundColor Red
    } elseif ($loopError) {
        Write-Host "--- Failed (loop error) ---" -ForegroundColor Red
    } elseif ($cancelled) {
        Write-Host "--- Cancelled ---" -ForegroundColor Yellow
    } else {
        Write-Host "--- Complete ---" -ForegroundColor Yellow
    }
    Write-Host ""

    $totalDuration = [int]$overallStopwatch.Elapsed.TotalSeconds
    $overallSpeed = if ($totalDuration -gt 0 -and $grandTotalBytesCopied -gt 0) { "$(Format-MirrorSize ([long]($grandTotalBytesCopied / $totalDuration)))/s avg" } else { $null }

    Write-Host "  Files copied:  $grandTotalCopied" -NoNewline -ForegroundColor White
    if ($grandTotalBytesCopied -gt 0) {
        Write-Host " ($(Format-MirrorSize $grandTotalBytesCopied))" -ForegroundColor Cyan
    } else {
        Write-Host ""
    }
    Write-Host "  Files skipped: $grandTotalSkipped (already synced)" -ForegroundColor Gray
    Write-Host "  Files deleted: $grandTotalDeleted (removed from backup)" -ForegroundColor $(if ($grandTotalDeleted -gt 0) { "Magenta" } else { "Gray" })

    if ($grandTotalFailed -gt 0) {
        Write-Host "  Files failed:  $grandTotalFailed" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "  Total time:    $(Format-TimeSpan $totalDuration)" -NoNewline -ForegroundColor Cyan
    if ($overallSpeed) {
        Write-Host " ($overallSpeed)" -ForegroundColor DarkCyan
    } else {
        Write-Host ""
    }
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] No changes were made" -ForegroundColor Yellow
        Write-Host ""
    }

    # Return stats for caller
    return @{
        FilesCopied = $grandTotalCopied
        FilesSkipped = $grandTotalSkipped
        FilesDeleted = $grandTotalDeleted
        FilesFailed = $grandTotalFailed
        BytesCopied = $grandTotalBytesCopied
        Duration = $overallStopwatch.Elapsed
        Cancelled = $cancelled
    }
}

<#
.SYNOPSIS
    Counts pending mirror changes via robocopy /L without copying anything.
.DESCRIPTION
    The Status dashboard needs a quick "is the mirror in sync, or how far
    behind is it?" reading. Invoke-Mirror has a pre-scan that does exactly
    this internally before the copy phase, but it's wrapped in the full
    mirror UI (drive checks, free-space warnings, user prompts, copy
    execution). This helper extracts just the per-folder /L diff and
    returns counts so the dashboard can render a one-line summary.
.PARAMETER SourceDrive
    Local library root (e.g. "G:\").
.PARAMETER DestDrive
    Mirror destination root — local drive or UNC share. The dest must be
    reachable for the scan to produce real numbers; if Test-Path fails
    the result returns Reachable=$false and zero counts (the caller can
    surface this as "destination offline" rather than "in sync").
.PARAMETER Folders
    Subfolders under SourceDrive to scan (e.g. @("Movies","Shows")).
.OUTPUTS
    @{
        Reachable           = bool
        Error               = string  # populated on auth/reachability failure
        Folders             = [PSCustomObject[]] (per-folder counts)
        TotalFilesToCopy    = int
        TotalBytesToCopy    = long
        TotalToDelete       = int
        ReleaseFoldersToCopy = [PSCustomObject[]] (top-level release folders with pending copies,
                              sorted by total bytes desc; @{Name, Files, Bytes, Tier})
    }
#>
function Get-MirrorPendingChanges {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$SourceDrive,
        [Parameter(Mandatory)] [string]$DestDrive,
        [Parameter(Mandatory)] [string[]]$Folders,
        [string]$NetworkUser,
        [string]$NetworkPass
    )

    $result = @{
        Reachable            = $false
        Error                = $null
        Folders              = @()
        TotalFilesToCopy     = 0
        TotalBytesToCopy     = [long]0
        TotalToDelete        = 0
        ReleaseFoldersToCopy = @()
    }

    # Authenticate the SMB share if creds are configured. Without this,
    # Test-Path on \\server\share returns $false when the share requires
    # explicit credentials (and the ambient Windows session doesn't have
    # them) — which previously showed up as a misleading "dest unreachable"
    # even though the box was up and serving Kodi.
    $conn = Connect-MirrorShare -DestDrive $DestDrive -NetworkUser $NetworkUser -NetworkPass $NetworkPass
    if ($conn.Error) {
        $result.Error = "SMB auth failed for $($conn.ShareRoot): $($conn.Error)"
        return $result
    }

    if (-not (Test-Path -LiteralPath $DestDrive)) {
        # Auth was either skipped (non-UNC / no creds) or succeeded but the
        # specific share path still isn't readable. Could be: share name
        # wrong, share offline, current Windows session lacks access, or
        # the backing drive on the remote host is offline/ejected.
        $result.Error = "Test-Path failed for $DestDrive (share offline, wrong name, or session lacks access)"
        return $result
    }
    $result.Reachable = $true

    # Same flags as Invoke-Mirror's pre-scan path. /L makes it list-only —
    # no files are actually copied. /NJH+/NJS would also hide the summary,
    # but Invoke-Mirror parses that summary for FilesDeleted, so we keep it
    # and let ConvertFrom-RobocopyOutput pick out the counts.
    $robocopyBaseArgs = @("/MIR", "/R:0", "/W:0", "/MT:8", "/XJD", "/NP", "/NDL", "/NC", "/BYTES", "/COPY:DT", "/DCOPY:T", "/FFT", "/L")

    # Cross-tier accumulator: each pending file gets bucketed by the first
    # path segment under its source-tier root (e.g. "Movies/Some.Movie/file.mkv"
    # → release "Some.Movie", tier "Movies"). Status dashboard renders these
    # as a flat top-N list of human-recognizable release names.
    $releaseAcc = @{}

    foreach ($folder in $Folders) {
        $sourcePath = Join-Path $SourceDrive $folder
        $destPath   = Join-Path $DestDrive $folder
        if (-not (Test-Path -LiteralPath $sourcePath)) { continue }

        # Normalize the prefix once per folder so the per-line strip is a
        # cheap StartsWith + Substring instead of repeated Join-Path math.
        $sourcePrefix = $sourcePath.TrimEnd('\','/') + '\'

        $scanArgs = "`"$sourcePath`" `"$destPath`" " + ($robocopyBaseArgs -join ' ')
        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo.FileName               = "robocopy"
        $proc.StartInfo.Arguments              = $scanArgs
        $proc.StartInfo.UseShellExecute        = $false
        $proc.StartInfo.RedirectStandardOutput = $true
        $proc.StartInfo.RedirectStandardError  = $true
        $proc.StartInfo.CreateNoWindow         = $true
        $proc.Start() | Out-Null

        $lines       = @()
        $filesToCopy = 0
        $bytesToCopy = [long]0
        while (-not $proc.StandardOutput.EndOfStream) {
            $line = $proc.StandardOutput.ReadLine()
            if ($null -eq $line) { break }
            $lines += $line
            # Same pattern Invoke-Mirror uses to count file-list entries:
            # leading whitespace + size + path. Capture the path too so we
            # can group by release folder for the dashboard listing.
            if ($line -match '^\s+(\d+)\s+(.+)$') {
                $size = [long]$Matches[1]
                $path = $Matches[2].Trim()
                $filesToCopy++
                $bytesToCopy += $size

                # Strip the source-tier prefix to get a release-relative
                # path. First segment is the release folder (movie name
                # or show name). Files directly at the tier root (loose)
                # are bucketed under "(root)".
                if ($path.StartsWith($sourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $relative = $path.Substring($sourcePrefix.Length)
                    $firstSep = $relative.IndexOfAny([char[]]@('\','/'))
                    $releaseName = if ($firstSep -gt 0) { $relative.Substring(0, $firstSep) } else { "(root)" }
                    $key = "$folder|$releaseName"
                    if (-not $releaseAcc.ContainsKey($key)) {
                        $releaseAcc[$key] = @{ Name = $releaseName; Tier = $folder; Files = 0; Bytes = [long]0 }
                    }
                    $releaseAcc[$key].Files++
                    $releaseAcc[$key].Bytes += $size
                }
            }
        }
        if (-not $proc.HasExited) { $proc.WaitForExit() }

        $summary = ConvertFrom-RobocopyOutput $lines

        $result.Folders += [PSCustomObject]@{
            Name        = $folder
            FilesToCopy = $filesToCopy
            BytesToCopy = $bytesToCopy
            ToDelete    = $summary.FilesDeleted
        }
        $result.TotalFilesToCopy += $filesToCopy
        $result.TotalBytesToCopy += $bytesToCopy
        $result.TotalToDelete    += $summary.FilesDeleted
    }

    $result.ReleaseFoldersToCopy = @(
        $releaseAcc.Values | Sort-Object { $_.Bytes } -Descending | ForEach-Object {
            [PSCustomObject]@{
                Name  = $_.Name
                Tier  = $_.Tier
                Files = $_.Files
                Bytes = $_.Bytes
            }
        }
    )

    return $result
}

#endregion

# Export public functions
Export-ModuleMember -Function Invoke-Mirror, Get-MirrorPendingChanges
