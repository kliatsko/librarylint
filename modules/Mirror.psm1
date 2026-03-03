# Mirror.psm1
# Mirrors media folders to backup destination using robocopy
# Part of LibraryLint suite

#region Private Functions

function Get-MirrorFolderStats {
    param(
        [string]$Path,
        [switch]$ShowProgress
    )

    if (-not (Test-Path $Path)) {
        return @{ Files = 0; Size = 0 }
    }

    $fileCount = 0
    $totalSize = [long]0
    $lastUpdate = [DateTime]::MinValue

    Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        $fileCount++
        $totalSize += $_.Length

        if ($ShowProgress) {
            $now = [DateTime]::Now
            if (($now - $lastUpdate).TotalMilliseconds -ge 300) {
                $lastUpdate = $now
                Write-Host "`r  Scanning... $fileCount files ($(Format-MirrorSize $totalSize))     " -NoNewline -ForegroundColor DarkGray
            }
        }
    }

    return @{
        Files = $fileCount
        Size = $totalSize
    }
}

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
        if ($line -match "^\s*Bytes\s*:\s*[\d.]+\s*\w*\s+([\d.]+)\s*(\w*)") {
            $value = [double]$Matches[1]
            $unit = $Matches[2]
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
    $robocopyBaseArgs = @("/MIR", "/R:2", "/W:5", "/MT:16", "/XJD", "/NDL", "/NC", "/BYTES", "/COPY:DT", "/DCOPY:T", "/FFT")

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

        Write-Host "`r  $folder".PadRight(25) -NoNewline
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

        $folderSize = if ($folderStats[$folder]) { $folderStats[$folder].BytesToCopy } else { 0 }

        Write-Host "  [$folderIndex/$($Folders.Count)] $folder -> $dest" -ForegroundColor Yellow

        $folderStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $folderBytesCopied = 0
        $folderFilesCopied = 0
        $deleteErrors = 0
        $lastProgressUpdate = [DateTime]::MinValue
        $currentFileName = ""
        $currentFileSize = 0
        $currentFilePct = 0

        # Run robocopy
        $outputLines = @()

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo.FileName = "robocopy"
        $process.StartInfo.Arguments = "`"$source`" `"$dest`" " + (($robocopyBaseArgs | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' ')
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true
        $process.StartInfo.CreateNoWindow = $true

        $process.Start() | Out-Null
        $currentProcess = $process

        try {
            while (-not $process.StandardOutput.EndOfStream) {
                # Check for key press (Ctrl+C or Q to quit)
                if ([Console]::KeyAvailable) {
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

                $line = $process.StandardOutput.ReadLine()
                if ($null -eq $line) { break }
                $outputLines += $line

                # Show errors, retries, and warnings
                if ($line -match 'ERROR.*Deleting Extra File|ERROR.*Access is denied') {
                    # Silently count delete permission errors (common on Samba/Kodi shares)
                    $deleteErrors++
                }
                elseif ($line -match 'ERROR\s+\d+\s*\(0x[0-9A-Fa-f]+\)\s+(.+)') {
                    Write-Host "`r".PadRight(100) -NoNewline
                    Write-Host "`r       ERROR: $($Matches[1])" -ForegroundColor Red
                }
                elseif ($line -match 'Waiting\s+(\d+)\s+seconds') {
                    Write-Host "`r".PadRight(100) -NoNewline
                    Write-Host "`r       Waiting $($Matches[1])s before retry..." -ForegroundColor Yellow
                }
                elseif ($line -match 'Retrying\.\.\.') {
                    Write-Host "       Retrying..." -ForegroundColor Yellow
                }
                elseif ($line -match '(The process cannot access|The network path|The specified network|network name is no longer available)') {
                    Write-Host "`r".PadRight(100) -NoNewline
                    Write-Host "`r       $line" -ForegroundColor Red
                }
                # Parse per-file progress lines (format: "  <percentage>%" — emitted during large file copies)
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

                    # ETA calculation
                    $elapsed = $folderStopwatch.Elapsed.TotalSeconds
                    $eta = ""
                    if ($elapsed -gt 3 -and $effectiveBytes -gt 0) {
                        $speed = $effectiveBytes / $elapsed
                        $remaining = if ($speed -gt 0) { [math]::Max(0, [math]::Round(($folderSize - $effectiveBytes) / $speed)) } else { 0 }
                        if ($remaining -gt 0) {
                            $eta = " ETA $(Format-TimeSpan $remaining)"
                        }
                        $speedStr = "$(Format-MirrorSize ([long]$speed))/s"
                    } else {
                        $speedStr = ""
                    }

                    # Current file (truncated)
                    $displayName = if ($currentFileName) { Split-Path $currentFileName -Leaf } else { "scanning..." }
                    $truncName = if ($displayName.Length -gt 30) { $displayName.Substring(0, 27) + "..." } else { $displayName }

                    # Build progress line: [####-----] 42% 1.2 GB | 45 MB/s ETA 2m 30s | filename.mkv
                    $progressLine = "       $progressBar $pctBytes% $(Format-MirrorSize $effectiveBytes)"
                    if ($speedStr) { $progressLine += " | $speedStr" }
                    if ($eta) { $progressLine += $eta }
                    if ($folderFilesCopied -gt 0 -or $currentFileName) {
                        $progressLine += " | $truncName"
                    }

                    Write-Host "`r$($progressLine.PadRight(100))" -NoNewline -ForegroundColor Cyan
                }
            }
        }
        catch {
            # Handle interruption
            $cancelled = $true
        }

        # Count final file
        if ($currentFileSize -gt 0) {
            $folderBytesCopied += $currentFileSize
            $folderFilesCopied++
        }

        if (-not $process.HasExited) {
            $process.WaitForExit()
        }
        $exitCode = $process.ExitCode
        $folderStopwatch.Stop()
        $currentProcess = $null

        if ($cancelled) {
            Write-Host ""
            Write-Host "  Mirror cancelled by user" -ForegroundColor Yellow
            Write-Host ""
            break
        }

        # Clear progress line
        Write-Host "`r".PadRight(95) -NoNewline
        Write-Host "`r" -NoNewline

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
    if ($cancelled) {
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

#endregion

# Export public functions
Export-ModuleMember -Function Invoke-Mirror
