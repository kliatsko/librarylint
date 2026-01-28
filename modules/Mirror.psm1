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
    The source drive letter (e.g., "G:")
.PARAMETER DestDrive
    The destination drive letter (e.g., "F:")
.PARAMETER Folders
    Array of folder names to mirror (e.g., @("Movies", "Shows"))
.PARAMETER WhatIf
    Preview changes without making them
.EXAMPLE
    Invoke-Mirror -SourceDrive "G:" -DestDrive "F:" -Folders @("Movies", "Shows")
.EXAMPLE
    Invoke-Mirror -SourceDrive "G:" -DestDrive "F:" -Folders @("Movies", "Shows") -WhatIf
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
    Write-Host "                    MIRROR MEDIA                       " -ForegroundColor Cyan
    Write-Host "        $SourceDrive (Source) --> $DestDrive (Backup)  " -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] No changes will be made" -ForegroundColor Yellow
        Write-Host ""
    }

    # Pre-scan phase
    Write-Host "--- Scanning Folders ---" -ForegroundColor DarkGray
    Write-Host ""

    $totalSourceFiles = 0
    $totalSourceSize = 0
    $totalDestFiles = 0
    $totalDestSize = 0

    foreach ($folder in $Folders) {
        $sourcePath = Join-Path $SourceDrive $folder
        $destPath = Join-Path $DestDrive $folder

        Write-Host "  Scanning $folder..." -ForegroundColor Gray -NoNewline

        $sourceStats = Get-MirrorFolderStats $sourcePath
        $destStats = Get-MirrorFolderStats $destPath

        $totalSourceFiles += $sourceStats.Files
        $totalSourceSize += $sourceStats.Size
        $totalDestFiles += $destStats.Files
        $totalDestSize += $destStats.Size

        Write-Host " done" -ForegroundColor Green
        Write-Host "    Source: $($sourceStats.Files) files ($(Format-MirrorSize $sourceStats.Size))" -ForegroundColor White
        Write-Host "    Dest:   $($destStats.Files) files ($(Format-MirrorSize $destStats.Size))" -ForegroundColor White
        Write-Host ""
    }

    Write-Host "  ─────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  TOTAL SOURCE: $totalSourceFiles files ($(Format-MirrorSize $totalSourceSize))" -ForegroundColor Cyan
    Write-Host "  TOTAL DEST:   $totalDestFiles files ($(Format-MirrorSize $totalDestSize))" -ForegroundColor Cyan
    Write-Host ""

    # Mirror phase
    Write-Host "--- Mirroring ---" -ForegroundColor DarkGray
    Write-Host ""

    # Use single-threaded for better progress display (MT causes buffered/batched output)
    # /NP removes the percentage spam, we'll show our own progress
    $robocopyBaseArgs = @("/MIR", "/R:3", "/W:5", "/XJD", "/BYTES", "/NC", "/NS", "/NDL", "/NP")

    if ($WhatIf) {
        $robocopyBaseArgs += "/L"
    }

    $grandTotalCopied = 0
    $grandTotalSkipped = 0
    $grandTotalDeleted = 0
    $grandTotalFailed = 0
    $grandTotalBytesCopied = 0
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($folder in $Folders) {
        $source = Join-Path $SourceDrive $folder
        $dest = Join-Path $DestDrive $folder

        if (-not (Test-Path $source)) {
            Write-Host "  X Source not found: $source" -ForegroundColor Red
            continue
        }

        Write-Host "  [$folder]" -ForegroundColor Yellow
        Write-Host "  $source --> $dest" -ForegroundColor Gray
        Write-Host ""

        $folderStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        # First, do a quick scan to count files to copy
        Write-Host "  Analyzing..." -ForegroundColor Gray -NoNewline
        $scanArgs = @($source, $dest) + $robocopyBaseArgs + @("/L")
        $scanOutput = & robocopy @scanArgs 2>&1 | Out-String -Stream
        $filesToProcess = ($scanOutput | Where-Object { $_ -match '^\s+\d+\s+' -and $_ -notmatch '^\s*(Files|Bytes|Times|Ended|Started|Speed|Total)' }).Count
        Write-Host " $filesToProcess files to process" -ForegroundColor Gray

        # Run robocopy and show progress
        $robocopyArgs = @($source, $dest) + $robocopyBaseArgs
        $outputLines = @()
        $filesProcessed = 0
        $bytesCopied = 0
        $copyStartTime = [DateTime]::Now

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo.FileName = "robocopy"
        $process.StartInfo.Arguments = ($robocopyArgs | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' '
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true
        $process.StartInfo.CreateNoWindow = $true

        $process.Start() | Out-Null

        while (-not $process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            $outputLines += $line

            # Check if this is a file being copied (has size prefix and file path)
            if ($line -match '^\s+(\d+)\s+(.+)$') {
                $fileSize = [long]$Matches[1]
                $filePath = $Matches[2].Trim()
                $fileName = Split-Path $filePath -Leaf
                $filesProcessed++
                $bytesCopied += $fileSize

                # Update progress display on every file for responsive feedback
                $now = [DateTime]::Now
                $elapsed = $now - $copyStartTime
                $progressPct = if ($filesToProcess -gt 0) { [math]::Round(($filesProcessed / $filesToProcess) * 100, 1) } else { 0 }

                # Calculate ETA
                $eta = ""
                if ($filesProcessed -gt 0 -and $filesToProcess -gt $filesProcessed) {
                    $filesRemaining = $filesToProcess - $filesProcessed
                    $avgTimePerFile = $elapsed.TotalSeconds / $filesProcessed
                    $secondsRemaining = [math]::Round($avgTimePerFile * $filesRemaining)
                    if ($secondsRemaining -lt 60) {
                        $eta = " ETA: ${secondsRemaining}s"
                    } elseif ($secondsRemaining -lt 3600) {
                        $eta = " ETA: $([math]::Floor($secondsRemaining / 60))m"
                    } else {
                        $eta = " ETA: $([math]::Floor($secondsRemaining / 3600))h $([math]::Floor(($secondsRemaining % 3600) / 60))m"
                    }
                }

                $truncatedName = if ($fileName.Length -gt 35) { $fileName.Substring(0, 32) + "..." } else { $fileName }
                $sizeStr = Format-MirrorSize $bytesCopied

                # Clear line and show detailed progress
                Write-Host "`r  $filesProcessed/$filesToProcess ($progressPct%) $sizeStr$eta - $truncatedName".PadRight(90) -ForegroundColor Cyan -NoNewline
            }
        }

        $process.WaitForExit()
        $exitCode = $process.ExitCode

        # Clear progress line
        Write-Host "`r".PadRight(95) -NoNewline
        Write-Host "`r" -NoNewline

        $folderStopwatch.Stop()

        # Parse results
        $stats = ConvertFrom-RobocopyOutput $outputLines

        $grandTotalCopied += $stats.FilesCopied
        $grandTotalSkipped += $stats.FilesSkipped
        $grandTotalDeleted += $stats.FilesDeleted
        $grandTotalFailed += $stats.FilesFailed
        $grandTotalBytesCopied += $stats.BytesCopied

        # Display folder results
        $statusColor = if ($exitCode -ge 8) { "Red" } elseif ($stats.FilesCopied -gt 0 -or $stats.FilesDeleted -gt 0) { "Green" } else { "Gray" }
        $statusIcon = if ($exitCode -ge 8) { "X" } else { "OK" }

        Write-Host "  [$statusIcon] Copied:  $($stats.FilesCopied) files ($(Format-MirrorSize $stats.BytesCopied))" -ForegroundColor $statusColor
        Write-Host "       Skipped: $($stats.FilesSkipped) files (already up to date)" -ForegroundColor Gray
        Write-Host "       Deleted: $($stats.FilesDeleted) files (not in source)" -ForegroundColor $(if ($stats.FilesDeleted -gt 0) { "Magenta" } else { "Gray" })

        if ($stats.FilesFailed -gt 0) {
            Write-Host "       Failed:  $($stats.FilesFailed) files" -ForegroundColor Red
        }

        Write-Host "       Time:    $($folderStopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor DarkGray
        Write-Host ""
    }

    $stopwatch.Stop()

    # Summary
    Write-Host "--- Summary ---" -ForegroundColor DarkGray
    Write-Host ""

    $summaryColor = if ($grandTotalFailed -gt 0) { "Yellow" } else { "Green" }

    Write-Host "  Mirror operation complete" -ForegroundColor $summaryColor
    Write-Host ""
    Write-Host "    Files copied:  $grandTotalCopied ($(Format-MirrorSize $grandTotalBytesCopied))" -ForegroundColor White
    Write-Host "    Files skipped: $grandTotalSkipped" -ForegroundColor Gray
    Write-Host "    Files deleted: $grandTotalDeleted" -ForegroundColor $(if ($grandTotalDeleted -gt 0) { "Magenta" } else { "Gray" })

    if ($grandTotalFailed -gt 0) {
        Write-Host "    Files failed:  $grandTotalFailed" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "    Total time:    $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
    Write-Host ""

    if ($WhatIf) {
        Write-Host "  [DRY RUN] No changes were made. Run without -WhatIf to apply." -ForegroundColor Yellow
        Write-Host ""
    }

    # Return stats for caller
    return @{
        FilesCopied = $grandTotalCopied
        FilesSkipped = $grandTotalSkipped
        FilesDeleted = $grandTotalDeleted
        FilesFailed = $grandTotalFailed
        BytesCopied = $grandTotalBytesCopied
        Duration = $stopwatch.Elapsed
    }
}

#endregion

# Export public functions
Export-ModuleMember -Function Invoke-Mirror
