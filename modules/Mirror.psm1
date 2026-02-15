# Mirror.psm1
# Mirrors media folders to backup destination using robocopy
# Part of LibraryLint suite

#region Private Functions

function Get-MirrorFolderStats {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return @{ Files = 0; Size = 0 }
    }

    $items = Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue
    $totalSize = ($items | Measure-Object -Property Length -Sum).Sum
    return @{
        Files = $items.Count
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

    # Pre-scan phase
    Write-Host "--- Scanning ---" -ForegroundColor Yellow
    Write-Host ""

    $totalSourceFiles = 0
    $totalSourceSize = 0
    $totalDestFiles = 0
    $totalDestSize = 0
    $folderStats = @{}

    foreach ($folder in $Folders) {
        $sourcePath = Join-Path $SourceDrive $folder
        $destPath = Join-Path $DestDrive $folder

        Write-Host "  $folder" -NoNewline

        if (-not (Test-Path $sourcePath)) {
            Write-Host " - Source not found!" -ForegroundColor Red
            continue
        }

        Write-Host " - scanning..." -ForegroundColor Gray -NoNewline

        $sourceStats = Get-MirrorFolderStats $sourcePath
        $destStats = Get-MirrorFolderStats $destPath

        $totalSourceFiles += $sourceStats.Files
        $totalSourceSize += $sourceStats.Size
        $totalDestFiles += $destStats.Files
        $totalDestSize += $destStats.Size

        $folderStats[$folder] = @{
            SourceFiles = $sourceStats.Files
            SourceSize = $sourceStats.Size
            DestFiles = $destStats.Files
            DestSize = $destStats.Size
        }

        Write-Host "`r  $folder".PadRight(25) -NoNewline
        Write-Host "$($sourceStats.Files) files " -ForegroundColor Cyan -NoNewline
        Write-Host "($(Format-MirrorSize $sourceStats.Size))" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "  Total: $totalSourceFiles files ($(Format-MirrorSize $totalSourceSize))" -ForegroundColor Cyan

    # Check destination drive capacity
    $destSpace = Get-DriveSpace $DestDrive
    if ($destSpace) {
        $freePercent = [math]::Round(($destSpace.FreeSpace / $destSpace.TotalSpace) * 100, 1)
        Write-Host "  Dest space: $(Format-MirrorSize $destSpace.FreeSpace) free ($freePercent%)" -ForegroundColor $(if ($freePercent -lt 10) { 'Red' } elseif ($freePercent -lt 20) { 'Yellow' } else { 'Gray' })

        # Estimate space needed
        $estimatedNeeded = $totalSourceSize - $totalDestSize
        if ($estimatedNeeded -gt 0 -and $estimatedNeeded -gt $destSpace.FreeSpace) {
            Write-Host ""
            Write-Host "  WARNING: May not have enough space!" -ForegroundColor Red
            Write-Host "  Estimated need: $(Format-MirrorSize $estimatedNeeded)" -ForegroundColor Red
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

    Write-Host ""

    # Mirror phase
    Write-Host "--- Mirroring ---" -ForegroundColor Yellow
    Write-Host ""

    # /MIR = Mirror mode (sync + delete extras)
    # /R:5 = Retry 5 times on failure
    # /W:10 = Wait 10 seconds between retries
    # /MT:8 = 8 threads for parallel copying
    # /XJD = Exclude junction points for directories
    # /NP = No progress percentage (we handle our own)
    # /NDL = Don't log directory names
    # /NC = Don't log file classes
    # /BYTES = Show sizes in bytes for parsing
    # /Z = Restartable mode (resume interrupted copies)
    # /IPG:100 = Inter-packet gap (ms) to reduce network congestion
    $robocopyBaseArgs = @("/MIR", "/R:5", "/W:10", "/MT:8", "/XJD", "/NP", "/NDL", "/NC", "/BYTES", "/Z", "/IPG:100")

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

        $folderSize = if ($folderStats[$folder]) { $folderStats[$folder].SourceSize } else { 0 }

        Write-Host "  [$folderIndex/$($Folders.Count)] $folder -> $dest" -ForegroundColor Yellow

        $folderStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $folderBytesProcessed = 0
        $folderFilesProcessed = 0
        $lastProgressUpdate = [DateTime]::MinValue

        # Run robocopy
        $robocopyArgs = @($source, $dest) + $robocopyBaseArgs
        $outputLines = @()

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo.FileName = "robocopy"
        $process.StartInfo.Arguments = ($robocopyArgs | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' '
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
                if ($line -match 'ERROR\s+\d+\s*\(0x[0-9A-Fa-f]+\)\s+(.+)') {
                    Write-Host ""
                    Write-Host "       ERROR: $($Matches[1])" -ForegroundColor Red
                }
                elseif ($line -match 'Waiting\s+(\d+)\s+seconds') {
                    Write-Host ""
                    Write-Host "       Waiting $($Matches[1])s before retry..." -ForegroundColor Yellow
                }
                elseif ($line -match 'Retrying\.\.\.') {
                    # Show what file is being retried (previous line usually has the path)
                    Write-Host "       Retrying..." -ForegroundColor Yellow
                }
                elseif ($line -match '(The process cannot access|Access is denied|The network path|The specified network|network name is no longer available)') {
                    Write-Host ""
                    Write-Host "       $line" -ForegroundColor Red
                }
                # Parse file copy lines (format: "   <size> <path>")
                elseif ($line -match '^\s+(\d+)\s+(.+)$') {
                    $fileSize = [long]$Matches[1]
                    $filePath = $Matches[2].Trim()
                    $folderBytesProcessed += $fileSize
                    $folderFilesProcessed++

                    # Update progress every 250ms to reduce flicker
                    $now = [DateTime]::Now
                    if (($now - $lastProgressUpdate).TotalMilliseconds -ge 250) {
                        $lastProgressUpdate = $now

                        # Calculate progress
                        $pctBytes = if ($folderSize -gt 0) { [math]::Min(100, [math]::Round(($folderBytesProcessed / $folderSize) * 100, 0)) } else { 0 }

                        # Progress bar (20 chars wide)
                        $barFilled = [math]::Floor($pctBytes / 5)
                        $barEmpty = 20 - $barFilled
                        $progressBar = "[" + ("=" * $barFilled) + (" " * $barEmpty) + "]"

                        # ETA calculation
                        $elapsed = $folderStopwatch.Elapsed.TotalSeconds
                        $eta = ""
                        if ($elapsed -gt 2 -and $pctBytes -gt 0 -and $pctBytes -lt 100) {
                            $estimatedTotal = $elapsed / ($pctBytes / 100)
                            $remaining = [math]::Max(0, [math]::Round($estimatedTotal - $elapsed))
                            $eta = " ETA: $(Format-TimeSpan $remaining)"
                        }

                        # Current file (truncated)
                        $fileName = Split-Path $filePath -Leaf
                        $truncName = if ($fileName.Length -gt 35) { $fileName.Substring(0, 32) + "..." } else { $fileName }

                        Write-Host "`r       $progressBar $pctBytes% $(Format-MirrorSize $folderBytesProcessed)$eta - $truncName".PadRight(90) -NoNewline -ForegroundColor Cyan
                    }
                }
            }
        }
        catch {
            # Handle interruption
            $cancelled = $true
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

        Write-Host "       $statusIcon Copied: $($stats.FilesCopied)" -ForegroundColor $statusColor -NoNewline
        if ($stats.BytesCopied -gt 0) {
            Write-Host " ($(Format-MirrorSize $stats.BytesCopied))" -ForegroundColor $statusColor -NoNewline
        }
        Write-Host " | Skipped: $($stats.FilesSkipped)" -ForegroundColor Gray -NoNewline
        Write-Host " | Deleted: $($stats.FilesDeleted)" -ForegroundColor $(if ($stats.FilesDeleted -gt 0) { "Magenta" } else { "Gray" }) -NoNewline
        if ($stats.FilesFailed -gt 0) {
            Write-Host " | Failed: $($stats.FilesFailed)" -ForegroundColor Red -NoNewline
        }
        Write-Host " | $(Format-TimeSpan ([int]$folderStopwatch.Elapsed.TotalSeconds))" -ForegroundColor DarkGray
        Write-Host ""
    }

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
    Write-Host "  Total time: $(Format-TimeSpan ([int]$overallStopwatch.Elapsed.TotalSeconds))" -ForegroundColor Cyan
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
