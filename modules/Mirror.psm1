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

function Get-MirrorProcessWriteBytes {
    # Bytes the robocopy process has handed to WriteFile so far, from its
    # own I/O counters. This is the one honest progress source. The three
    # things tried before it all lie: robocopy pre-allocates every
    # destination file to full size the moment a thread starts it, so the
    # dest file's size reads 100% immediately; robocopy's redirected stdout
    # is block-buffered, so file announcements arrive in flushes that can be
    # an hour apart on big files; and the NIC bytes-sent counter is reported
    # once per NDIS filter driver stacked on the adapter (WFP, QoS,
    # VirtualBox — five copies on the machine this was found on) plus
    # whatever else is on the wire. Returns $null once the process is gone.
    [CmdletBinding()]
    param([Parameter(Mandatory)] [int]$ProcessId)

    try {
        $proc = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $ProcessId" -ErrorAction Stop
        if ($proc) { return [long]$proc.WriteTransferCount }
    } catch {
        # Fall through: a failed query is indistinguishable from "gone" for
        # the caller, which keeps its last good value either way.
    }
    return $null
}

function Update-MirrorSpeedWindow {
    # Adds one (elapsed, bytes) sample to a rolling window and returns the
    # bytes-per-second over that window. Kept pure so the arithmetic — the
    # part that used to trend the ETA wrong — can be tested without a copy.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [System.Collections.Generic.Queue[object]]$Samples,
        [Parameter(Mandatory)] [double]$ElapsedSec,
        [Parameter(Mandatory)] [long]$Bytes,
        [double]$WindowSec = 5.0
    )

    $Samples.Enqueue([PSCustomObject]@{ T = $ElapsedSec; B = $Bytes })
    # Trim samples older than the window, always keeping one to subtract
    # against.
    while ($Samples.Count -gt 1 -and ($ElapsedSec - $Samples.Peek().T) -gt $WindowSec) {
        [void]$Samples.Dequeue()
    }
    if ($Samples.Count -lt 2) { return [double]0 }
    $oldest = $Samples.Peek()
    $deltaBytes = $Bytes - $oldest.B
    if ($deltaBytes -lt 0) { $deltaBytes = [long]0 }
    $deltaSec = $ElapsedSec - $oldest.T
    if ($deltaSec -le 0) { return [double]0 }
    return [double]($deltaBytes / $deltaSec)
}

function New-MirrorLogTail {
    # Cursor over a robocopy /UNILOG file that is still being written.
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$Path)

    return @{ Path = $Path; Position = [long]0 }
}

function Read-MirrorLogTail {
    # Returns the complete lines written to the log since the last call and
    # advances the cursor past them. Robocopy writes the log line by line as
    # each thread starts a file, which is what makes it usable for live
    # progress where stdout is not. /UNILOG is UTF-16LE with a BOM; only
    # whole code units are consumed and a trailing partial line is left in
    # the file for the next read, so a write caught mid-line never yields a
    # torn line or a torn character.
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable]$Tail)

    $lines = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path -LiteralPath $Tail.Path)) { return , $lines }

    $stream = $null
    try {
        $stream = [IO.File]::Open($Tail.Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        if ($Tail.Position -eq 0 -and $stream.Length -ge 2) {
            $bom = New-Object byte[] 2
            [void]$stream.Read($bom, 0, 2)
            if ($bom[0] -eq 0xFF -and $bom[1] -eq 0xFE) { $Tail.Position = 2 }
        }
        $available = $stream.Length - $Tail.Position
        if ($available -lt 2) { return , $lines }
        $count = [int]($available - ($available % 2))
        $buffer = New-Object byte[] $count
        $stream.Position = $Tail.Position
        $read = 0
        while ($read -lt $count) {
            $chunk = $stream.Read($buffer, $read, $count - $read)
            if ($chunk -le 0) { break }
            $read += $chunk
        }
        if ($read % 2 -eq 1) { $read-- }
        $text = [Text.Encoding]::Unicode.GetString($buffer, 0, $read)
        $lastNewline = $text.LastIndexOf("`n")
        if ($lastNewline -lt 0) { return , $lines }
        $complete = $text.Substring(0, $lastNewline + 1)
        foreach ($line in $complete.Split("`n")) {
            $lines.Add($line.TrimEnd("`r"))
        }
        # Split leaves one empty entry after the final newline.
        $lines.RemoveAt($lines.Count - 1)
        $Tail.Position += ($lastNewline + 1) * 2
    } catch {
        # A read that races a write simply retries next tick.
    } finally {
        if ($stream) { $stream.Dispose() }
    }
    return , $lines
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
        # Parse the summary lines. Columns: Total Copied Skipped Mismatch
        # FAILED Extras — the extras (files purged from the destination)
        # live in the sixth column of the Files row; there is no separate
        # "Extras :" row, so "to delete" read 0 for as long as this only
        # looked for one.
        if ($line -match "^\s*Files\s*:\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)") {
            $stats.FilesCopied = [int]$Matches[2]
            $stats.FilesSkipped = [int]$Matches[3]
            $stats.FilesFailed = [int]$Matches[5]
            $stats.FilesDeleted = [int]$Matches[6]
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

function Get-RobocopyListedFile {
    # Parses a robocopy file line — "<size> <path>", the shape /NC leaves —
    # and says whether it is a file to copy (path under the source) or an
    # extra to purge (path under the destination). With /NC robocopy drops
    # the "New File" / "*EXTRA File" labels, so the path root is the only
    # way to tell the two apart; before this, an extra counted as "to copy"
    # and inflated both the file count and the byte total the progress bar
    # divides by. Returns $null for any other line.
    [CmdletBinding()]
    param(
        [string]$Line,
        [Parameter(Mandatory)] [string]$SourceRoot,
        [Parameter(Mandatory)] [string]$DestRoot
    )

    if (-not ($Line -match '^\s+(\d+)\s+(.+)$')) { return $null }
    $size = [long]$Matches[1]
    $path = $Matches[2].Trim()
    $sourcePrefix = $SourceRoot.TrimEnd('\', '/') + '\'
    $destPrefix   = $DestRoot.TrimEnd('\', '/') + '\'
    $kind = if ($path.StartsWith($sourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)) { 'Copy' }
            elseif ($path.StartsWith($destPrefix, [System.StringComparison]::OrdinalIgnoreCase)) { 'Extra' }
            else { 'Copy' }
    return [PSCustomObject]@{ Size = $size; Path = $path; Kind = $kind }
}

# Robocopy error codes that mean the destination itself is gone, not that one
# file is in trouble. Every one of these was observed or documented for an SMB
# share whose host powered off, whose Samba restarted, or whose session reset.
$script:RobocopyDestinationLostCodes = @(51, 53, 58, 59, 64, 67, 121, 1231, 1232)
$script:RobocopyDestinationLostReasons = 'The network path was not found|The specified network name is no longer available|The network name cannot be found|The network location cannot be reached|The semaphore timeout period has expired|An unexpected network error occurred|The remote computer is not available|The specified server cannot perform the requested operation'
$script:RobocopyInUseReasons = 'being used by another process|The process cannot access'

# Canned reason text per code, used when robocopy's own reason line never
# arrives. With /MT the "ERROR n" line and its reason line come from
# different threads and interleave, so a message that waits for the pair
# can lose the file name entirely (the previous parser did exactly that).
$script:RobocopyReasonByCode = @{
    5    = 'Access is denied'
    32   = 'The process cannot access the file because it is being used by another process'
    33   = 'The process cannot access the file because another process has locked a portion of the file'
    51   = 'The remote computer is not available'
    53   = 'The network path was not found'
    58   = 'The specified server cannot perform the requested operation'
    59   = 'An unexpected network error occurred'
    64   = 'The specified network name is no longer available'
    67   = 'The network name cannot be found'
    121  = 'The semaphore timeout period has expired'
    1231 = 'The network location cannot be reached'
    1232 = 'The network location cannot be reached'
}

# How many distinct files must exhaust their retries with a destination-class
# error before the mirror concludes the destination is gone. Three, not one:
# a single file can hit a transient error and robocopy's own /R retries are
# what a short blip needs; three files all failing every attempt is an outage.
$script:MirrorDestinationLostStrikes = 3

function Get-RobocopyErrorClass {
    # Classifies one line of robocopy output. Returns $null for anything that
    # is not an error (file announcements, summary rows, banners); otherwise
    # an object with Class (DestinationLost | InUse | AccessDenied | Other),
    # Code (the numeric error, $null on a bare reason line), Target (what
    # robocopy was doing it to), Operation, IsDirectory (a destination-
    # directory failure, which is never per-file) and IsReason (the follow-up
    # reason line rather than the ERROR line itself).
    [CmdletBinding()]
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }

    if ($Line -match 'ERROR\s+(\d+)\s*\(0x[0-9A-Fa-f]+\)\s*(.*)$') {
        $code = [int]$Matches[1]
        $rest = $Matches[2].Trim()
        $operation = $null
        $target = $rest
        if ($rest -match '^((?:Copying|Deleting Extra|Time-Stamping Destination|Changing File Attributes of|Creating Destination|Accessing Destination|Accessing Source|Scanning Source|Scanning Destination|Getting File System Type of Destination)\s+(?:File|Directory|Dir))\s*(.*)$') {
            $operation = $Matches[1]
            $target = $Matches[2].Trim()
        }
        $class = if ($code -in $script:RobocopyDestinationLostCodes) { 'DestinationLost' }
                 elseif ($code -in @(32, 33)) { 'InUse' }
                 elseif ($code -eq 5) { 'AccessDenied' }
                 else { 'Other' }
        return [PSCustomObject]@{
            Class       = $class
            Code        = $code
            Target      = $target
            Operation   = $operation
            IsDirectory = [bool]($operation -match 'Destination Dir|Source Dir')
            IsReason    = $false
        }
    }

    $reasonClass = if ($Line -match $script:RobocopyDestinationLostReasons) { 'DestinationLost' }
                   elseif ($Line -match $script:RobocopyInUseReasons) { 'InUse' }
                   elseif ($Line -match 'Access is denied') { 'AccessDenied' }
                   else { $null }
    if ($reasonClass) {
        return [PSCustomObject]@{
            Class       = $reasonClass
            Code        = $null
            Target      = $null
            Operation   = $null
            IsDirectory = $false
            IsReason    = $true
        }
    }
    return $null
}

function New-RobocopyErrorState {
    # State bag for Update-RobocopyErrorState. RetryLimit is robocopy's /R
    # value: a file has only truly failed once robocopy has printed more
    # ERROR lines for it than it will retry.
    [CmdletBinding()]
    param([int]$RetryLimit = 0)

    return @{
        RetryLimit     = $RetryLimit
        Pending        = $null                      # last ERROR line awaiting its reason line
        AttemptsByFile = @{}                        # target -> ERROR lines seen
        FinalFiles     = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        ErrorsByClass  = @{}                        # class -> files that failed every attempt (copy side)
        DeleteErrors   = 0                          # extras that could not be removed (every attempt)
        DestLostFiles  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        DestLost       = $false
        DestLostDetail = $null
        Messages       = [System.Collections.Generic.List[object]]::new()
    }
}

function Update-RobocopyErrorState {
    # Feeds one robocopy output line into the error state. Returns $true when
    # the line was an error, reason, wait or retry line (consumed), $false
    # when the caller should parse it as a file announcement or progress.
    # Display text goes into $State.Messages as @{Text; Color} so this stays
    # free of console writes and can be driven by captured output in tests.
    #
    # The abort rule: a file becomes a strike only after robocopy has given
    # up on it (attempts > RetryLimit). Counting first attempts would abort a
    # run that a two-second SMB blip under /MT:16 — sixteen ERROR lines, then
    # "Retrying..." and success — would otherwise have survived. Three
    # distinct strikes, or a single destination-directory failure, means the
    # destination is gone.
    [CmdletBinding()]
    param(
        [string]$Line,
        [Parameter(Mandatory)] [hashtable]$State,
        [switch]$Flush
    )

    if ($Flush) {
        if ($State.Pending) {
            Publish-RobocopyPendingError -State $State -Reason $null
        }
        return $true
    }

    $info = Get-RobocopyErrorClass -Line $Line

    if ($info -and -not $info.IsReason) {
        # A new ERROR line while one is still waiting for its reason means the
        # two came from different /MT threads. Emit the older one with a
        # canned reason rather than silently dropping it.
        if ($State.Pending) {
            Publish-RobocopyPendingError -State $State -Reason $null
        }

        $isDelete = $info.Operation -like 'Deleting*'
        $key = if ($info.Target) { "$($info.Operation)|$($info.Target)" } else { "$($info.Operation)|$($info.Code)" }
        $State.AttemptsByFile[$key] = 1 + $(if ($State.AttemptsByFile.ContainsKey($key)) { $State.AttemptsByFile[$key] } else { 0 })
        $isFinal = $info.IsDirectory -or ($State.AttemptsByFile[$key] -gt $State.RetryLimit)

        if ($isFinal -and $State.FinalFiles.Add($key)) {
            if ($isDelete) {
                $State.DeleteErrors++
            } else {
                $State.ErrorsByClass[$info.Class] = 1 + $(if ($State.ErrorsByClass.ContainsKey($info.Class)) { $State.ErrorsByClass[$info.Class] } else { 0 })
            }
            if ($info.Class -eq 'DestinationLost') {
                [void]$State.DestLostFiles.Add($key)
                $reasonText = if ($script:RobocopyReasonByCode.ContainsKey($info.Code)) { $script:RobocopyReasonByCode[$info.Code] } else { "error $($info.Code)" }
                if ($info.IsDirectory) {
                    $State.DestLost = $true
                    $State.DestLostDetail = "robocopy error $($info.Code) ($reasonText) $($info.Operation.ToLower()) $($info.Target)"
                } elseif ($State.DestLostFiles.Count -ge $script:MirrorDestinationLostStrikes) {
                    $State.DestLost = $true
                    $State.DestLostDetail = "robocopy error $($info.Code) ($reasonText) on $($State.DestLostFiles.Count) files, each after every retry"
                }
            }
        }

        # Delete-permission errors are common on Samba/Kodi shares and were
        # always counted silently; keep that, but still track them so the
        # reason line that follows is consumed rather than shown bare.
        $State.Pending = @{
            Info   = $info
            Key    = $key
            Silent = ($isDelete -and $info.Class -eq 'AccessDenied')
        }
        return $true
    }

    if ($info -and $info.IsReason) {
        if ($State.Pending) {
            Publish-RobocopyPendingError -State $State -Reason $Line.Trim()
        } else {
            # Orphan reason line (its ERROR line was consumed already or
            # interleaved away). Destination-class text is still worth a
            # line; anything else was already reported with its file.
            if ($info.Class -eq 'DestinationLost') {
                $State.Messages.Add(@{ Text = "ERROR: $($Line.Trim())"; Color = 'Red' })
            }
        }
        return $true
    }

    if ($Line -match 'Waiting\s+(\d+)\s+seconds') {
        $State.Messages.Add(@{ Text = "Waiting $($Matches[1])s before retry..."; Color = 'DarkYellow' })
        return $true
    }
    if ($Line -match 'Retrying\.\.\.') {
        $State.Messages.Add(@{ Text = 'Retrying...'; Color = 'DarkYellow' })
        return $true
    }
    if ($Line -match '^ERROR:\s*RETRY LIMIT EXCEEDED') {
        # Robocopy's own "gave up" marker. Attempt counting already decided
        # finality per file (this line is not attributable under /MT).
        return $true
    }
    return $false
}

function Publish-RobocopyPendingError {
    # Turns the ERROR line held in $State.Pending into one display message,
    # using robocopy's reason line when it arrived and the canned text when
    # it did not.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$State,
        [string]$Reason
    )

    $pending = $State.Pending
    $State.Pending = $null
    if ($pending.Silent) { return }

    $info = $pending.Info
    if (-not $Reason) {
        $Reason = if ($script:RobocopyReasonByCode.ContainsKey($info.Code)) { $script:RobocopyReasonByCode[$info.Code] } else { "robocopy error $($info.Code)" }
    }
    $Reason = $Reason.TrimEnd('.')
    $what = if ($info.IsDirectory) {
        "$($info.Operation) $($info.Target)"
    } elseif ($info.Target) {
        Split-Path $info.Target -Leaf
    } else {
        "error $($info.Code)"
    }
    $State.Messages.Add(@{ Text = "ERROR: $what - $Reason."; Color = 'Red' })
}

function Write-MirrorDestinationLostAbort {
    # The three lines the user sees when the mirror gives up on a destination
    # that went away mid-run: what happened, robocopy's own reason, and one
    # probe to say whether the host itself or only the share disappeared —
    # those have different fixes. Pass -HostAnswers when the caller has just
    # probed, so the same 5-second TCP wait is not paid twice.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Dest,
        [string]$Detail,
        [nullable[bool]]$HostAnswers
    )

    Write-Host "  Destination went away mid-run ($Dest) — aborting mirror." -ForegroundColor Red
    if ($Detail) {
        Write-Host "  $Detail" -ForegroundColor DarkYellow
    }
    if ($Dest -match '^\\\\([^\\]+)') {
        $server = $Matches[1]
        if ($null -eq $HostAnswers) { $HostAnswers = Test-MirrorDestAlive -DestRoot $Dest }
        if ($HostAnswers) {
            Write-Host "  $server still answers on port 445, so the share or its drive went away rather than the box — Samba restarted, the drive unmounted, or the SMB session was reset." -ForegroundColor DarkYellow
        } else {
            Write-Host "  $server is not answering on port 445 — powered off, rebooting, or off the network." -ForegroundColor DarkYellow
        }
    }
    Write-Host "  Re-run Status once it is back; the mirror resumes where it stopped." -ForegroundColor DarkGray
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

    # Robocopy's own retry count decides when a failing file has truly failed
    # (see Update-RobocopyErrorState); read it from the flags rather than
    # keeping a second copy of the number.
    $retryLimit = 0
    foreach ($robocopyArg in $robocopyBaseArgs) {
        if ($robocopyArg -match '^/R:(\d+)$') { $retryLimit = [int]$Matches[1] }
    }

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

            # Count files that would be copied. Extras robocopy would purge
            # print in the same shape; they are counted from the summary.
            $listed = Get-RobocopyListedFile -Line $scanLine -SourceRoot $sourcePath -DestRoot $destPath
            if ($listed -and $listed.Kind -eq 'Copy') {
                $filesToCopy++
                $bytesToCopy += $listed.Size
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
        $folderFilesStarted = 0
        $announcedBytes = [long]0
        $errorState = New-RobocopyErrorState -RetryLimit $retryLimit
        $lastProgressUpdate = [DateTime]::MinValue
        $currentFileName = ""

        # Bytes, percent, speed, ETA and the watchdog all derive from
        # robocopy's own write counter (see Get-MirrorProcessWriteBytes for
        # why nothing else is trustworthy), sampled once a second.
        $ioSampleEvery      = 1.0
        $ioLastSample       = [DateTime]::MinValue
        $ioBaseline         = $null
        $effectiveBytes     = [long]0
        $speedSamples       = New-Object 'System.Collections.Generic.Queue[object]'
        $speedWindowSec     = 5.0
        $smoothedSpeed      = 0.0

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

        # The scan phase takes minutes on a large library — long enough for
        # an HTPC to power itself off in between. Probe once before
        # launching the copy so a dead host is one message, not a wall of
        # per-file errors. TCP probe rather than Test-Path: Test-Path on a
        # dead SMB session can hang for minutes. Probe the destination root,
        # not this folder: on a local drive the folder may not exist yet
        # ("WILL BE CREATED" above), and that is not an outage.
        if (-not (Test-MirrorDestAlive -DestRoot $DestDrive)) {
            Write-Host ""
            Write-MirrorDestinationLostAbort -Dest $dest -Detail "The host stopped answering between the scan and the copy." -HostAnswers $false
            $destLost = $true
            break
        }

        # Run robocopy. Its report goes to a /UNILOG file that we tail, not
        # to stdout: redirected stdout is block-buffered (verified — files
        # that finished in the first second were announced seven seconds
        # later, in one burst with the summary), while the log is written
        # line by line as each thread starts a file. UTF-16 so non-ASCII
        # titles survive. Stdout stays redirected and drained so robocopy
        # can never block on a full pipe.
        $outputLines = @()
        $preCopyRow = [Console]::CursorTop  # Save position to clean up robocopy's direct console writes
        $robocopyLog = Join-Path ([IO.Path]::GetTempPath()) "LibraryLint-mirror-$([IO.Path]::GetRandomFileName()).log"
        $logTail = New-MirrorLogTail -Path $robocopyLog
        $copyArgs = $robocopyBaseArgs + @("/UNILOG:$robocopyLog")

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo.FileName = "robocopy"
        $process.StartInfo.Arguments = "`"$source`" `"$dest`" " + (($copyArgs | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' ')
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true
        $process.StartInfo.CreateNoWindow = $true

        $outputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $outputSub = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -MessageData $outputQueue -Action {
            if ($null -ne $EventArgs.Data) {
                $Event.MessageData.Enqueue($EventArgs.Data)
            }
        }

        $process.Start() | Out-Null
        $process.BeginOutputReadLine()
        $currentProcess = $process

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

                # Drain robocopy's report: the log tail carries everything,
                # stdout is kept only so nothing can block there.
                $processedLine = $false
                $newLines = [System.Collections.Generic.List[string]]::new()
                $line = $null
                while ($outputQueue.TryDequeue([ref]$line)) { $newLines.Add($line) }
                foreach ($logLine in (Read-MirrorLogTail -Tail $logTail)) { $newLines.Add($logLine) }
                foreach ($line in $newLines) {
                    $processedLine = $true
                    $outputLines += $line

                    # Errors, retries and waits go through the error state,
                    # which is also what decides that the destination itself
                    # is gone rather than one file being in trouble.
                    if (Update-RobocopyErrorState -Line $line -State $errorState) {
                        foreach ($errorMessage in $errorState.Messages) {
                            Write-Host "`r$(' ' * 120)" -NoNewline
                            Write-Host "`r       $($errorMessage.Text)" -ForegroundColor $errorMessage.Color
                        }
                        $errorState.Messages.Clear()
                        if ($errorState.DestLost) { break }
                        continue
                    }
                    # A file line ("   <size> <path>") is written when a
                    # thread STARTS the file, not when it finishes — with
                    # /MT:16 up to sixteen are in flight at once. So this
                    # counts files started and names the latest one; bytes
                    # come from the I/O counter below, never from here.
                    $listed = Get-RobocopyListedFile -Line $line -SourceRoot $source -DestRoot $dest
                    if ($listed -and $listed.Kind -eq 'Copy') {
                        $folderFilesStarted++
                        $announcedBytes += $listed.Size
                        $currentFileName = $listed.Path
                    }
                }
                # Any line at all means robocopy is alive and working.
                # Listing the extras it is about to purge, for instance,
                # writes nothing to the destination and would otherwise
                # look like a stall to the watchdog.
                if ($processedLine) { $stallSince = $null; $aliveStallStrikes = 0 }

                # Robocopy has failed every retry on enough distinct files (or
                # on the destination directory itself) that this is an outage,
                # not a bad file. Stop now instead of paying /R x /W seconds
                # for each remaining file — tonight that was 1,413 files.
                if ($errorState.DestLost) {
                    Write-Host ""
                    Write-MirrorDestinationLostAbort -Dest $dest -Detail $errorState.DestLostDetail
                    $destLost = $true
                    if (-not $process.HasExited) { $process.Kill() }
                    break
                }

                # Sample robocopy's write counter once a second. In dry-run
                # mode nothing is written, so fall back to the sizes of the
                # files robocopy listed.
                $now = [DateTime]::Now
                if (($now - $ioLastSample).TotalSeconds -ge $ioSampleEvery) {
                    $ioLastSample = $now
                    if ($WhatIf) {
                        $effectiveBytes = $announcedBytes
                    } else {
                        $written = Get-MirrorProcessWriteBytes -ProcessId $process.Id
                        if ($null -ne $written) {
                            if ($null -eq $ioBaseline) { $ioBaseline = $written }
                            $effectiveBytes = $written - $ioBaseline
                            if ($effectiveBytes -lt 0) { $effectiveBytes = [long]0 }
                        }
                    }
                    $smoothedSpeed = Update-MirrorSpeedWindow -Samples $speedSamples -ElapsedSec $folderStopwatch.Elapsed.TotalSeconds -Bytes $effectiveBytes -WindowSec $speedWindowSec
                }

                # Update progress display (throttled to reduce flicker)
                if (($now - $lastProgressUpdate).TotalMilliseconds -ge 200) {
                    $lastProgressUpdate = $now

                    $pctBytes = if ($folderSize -gt 0) { [math]::Min(100, [math]::Round(($effectiveBytes / $folderSize) * 100, 0)) } else { 0 }

                    # Progress bar (25 chars wide)
                    $barFilled = [math]::Floor($pctBytes / 4)
                    $barEmpty = 25 - $barFilled
                    $progressBar = "[" + ("#" * $barFilled) + ("-" * $barEmpty) + "]"

                    $elapsed = $folderStopwatch.Elapsed.TotalSeconds
                    $eta = ""
                    $speedStr = ""
                    if ($elapsed -gt 1) {
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

                    # Latest file a thread picked up (truncated). Not "the
                    # file being copied" — with /MT:16 there are sixteen.
                    $displayName = if ($currentFileName) { Split-Path $currentFileName -Leaf } else { "starting..." }
                    $truncName = if ($displayName.Length -gt 40) { $displayName.Substring(0, 37) + "..." } else { $displayName }

                    # File counter is files STARTED (e.g., "1,247/3,485 started")
                    $fileCounter = if ($folderTotalFiles -gt 0) {
                        "$($folderFilesStarted.ToString('N0'))/$($folderTotalFiles.ToString('N0')) started"
                    } else {
                        "$($folderFilesStarted.ToString('N0')) started"
                    }

                    # Build progress line: [####-----] 42% 1.2 GB (1,247/3,485 started) | 45 MB/s ETA 2m 30s | latest: filename.mkv
                    $progressLine = "       $progressBar $pctBytes% $(Format-MirrorSize $effectiveBytes) ($fileCounter)"
                    if ($speedStr) { $progressLine += " | $speedStr" }
                    if ($eta) { $progressLine += $eta }
                    if ($currentFileName) {
                        $progressLine += " | latest: $truncName"
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
        $lateLines = [System.Collections.Generic.List[string]]::new()
        $drainLine = $null
        while ($outputQueue.TryDequeue([ref]$drainLine)) { $lateLines.Add($drainLine) }
        # The log gets its final flush — the Files:/Bytes: summary — at exit.
        foreach ($logLine in (Read-MirrorLogTail -Tail $logTail)) { $lateLines.Add($logLine) }
        foreach ($drainLine in $lateLines) {
            $outputLines += $drainLine
            # A destination-directory failure exits robocopy in about 20 ms,
            # so its one ERROR line usually lands in this late drain rather
            # than in the loop. Feed it through so a dead destination stops
            # the remaining folders instead of each failing in turn.
            if (-not $destLost -and -not $cancelled) {
                [void](Update-RobocopyErrorState -Line $drainLine -State $errorState)
            }
        }
        [void](Update-RobocopyErrorState -State $errorState -Flush)
        foreach ($errorMessage in $errorState.Messages) {
            Write-Host "       $($errorMessage.Text)" -ForegroundColor $errorMessage.Color
        }
        $errorState.Messages.Clear()
        if ($errorState.DestLost -and -not $destLost -and -not $cancelled) {
            Write-Host ""
            Write-MirrorDestinationLostAbort -Dest $dest -Detail $errorState.DestLostDetail
            $destLost = $true
        }

        $exitCode = $process.ExitCode
        $folderStopwatch.Stop()
        $currentProcess = $null

        # Robocopy's own log is the only per-file record of the run. Keep it
        # when anything went wrong; a clean run does not need it.
        if ($cancelled -or $destLost -or $exitCode -ge 8) {
            Write-Host "       robocopy log kept: $robocopyLog" -ForegroundColor DarkGray
        } else {
            Remove-Item -LiteralPath $robocopyLog -Force -ErrorAction SilentlyContinue
        }

        if ($loopError) {
            Write-Host ""
            Write-Host "  Mirror loop failed: $($loopError.Exception.Message)" -ForegroundColor Red
            Write-Host "  (Reported as an error — this was NOT a user cancel.)" -ForegroundColor DarkYellow
            Write-Host "  At: $($loopError.InvocationInfo.PositionMessage)" -ForegroundColor DarkGray
            if (-not $process.HasExited) { try { $process.Kill() } catch {} }
            # Route through the cancel path below so the run ends the same
            # way a cancel does.
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

            # Files robocopy was still writing are left behind at their full
            # size, because it pre-allocates each destination file when it
            # starts it. Their timestamps do not match the source — robocopy
            # stamps those only after the data — so the next run re-copies
            # them, which is the correct outcome. An earlier version
            # "repaired" those timestamps to avoid the re-copy, judging
            # completeness by size alone; with pre-allocation that judged
            # every in-flight partial complete and hid it from robocopy for
            # good. Completed files were never the problem: robocopy stamps
            # each one as it finishes.
            if (-not $destLost -and $folderFilesStarted -gt 0) {
                Write-Host "  Files still in flight are left partial on the destination; the next run re-copies them." -ForegroundColor DarkGray
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
        # One line per error class actually seen. The old summary called
        # every copy error "in use by another process" — including the
        # network errors that were the whole story when the share dropped.
        $errorClassLabels = [ordered]@{
            InUse           = 'could not be copied (in use by another process)'
            AccessDenied    = 'could not be copied (access denied)'
            DestinationLost = 'failed with a network error'
            Other           = 'could not be copied (see errors above)'
        }
        foreach ($errorClass in $errorClassLabels.Keys) {
            if ($errorState.ErrorsByClass.ContainsKey($errorClass) -and $errorState.ErrorsByClass[$errorClass] -gt 0) {
                Write-Host "       $($errorState.ErrorsByClass[$errorClass]) file(s) $($errorClassLabels[$errorClass])" -ForegroundColor DarkYellow
            }
        }
        if ($errorState.DeleteErrors -gt 0) {
            Write-Host "       $($errorState.DeleteErrors) file(s) could not be deleted from dest" -ForegroundColor DarkYellow
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
    $sawDestinationLost = $false

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
            $lineError = Get-RobocopyErrorClass -Line $line
            if ($lineError -and $lineError.Class -eq 'DestinationLost') { $sawDestinationLost = $true }
            # Same classification Invoke-Mirror uses: a listed file under
            # the source is a copy, under the destination an extra. Capture
            # the path too so we can group by release folder for the
            # dashboard listing.
            $listed = Get-RobocopyListedFile -Line $line -SourceRoot $sourcePath -DestRoot $destPath
            if ($listed -and $listed.Kind -eq 'Copy') {
                $size = $listed.Size
                $path = $listed.Path
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

    # robocopy /L cannot tell an empty destination from an unreachable one:
    # against a host that has gone away it lists every source file as new
    # and exits 1 without a single error line (verified). A share that dies
    # mid-scan would come back as a confident "everything to copy", and the
    # Status flow would start a mirror on the strength of it. One probe after
    # the scan turns that into an honest "unreachable".
    if ($sawDestinationLost -or -not (Test-MirrorDestAlive -DestRoot $DestDrive)) {
        return @{
            Reachable            = $false
            Error                = "destination stopped answering during the scan ($DestDrive)"
            Folders              = @()
            TotalFilesToCopy     = 0
            TotalBytesToCopy     = [long]0
            TotalToDelete        = 0
            ReleaseFoldersToCopy = @()
        }
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
