#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for the mirror's progress sources (modules\Mirror.psm1).

.DESCRIPTION
    A mirror once sat for an hour on "1% 4.75 GB (19/855) | 175 MB/s ETA 34m"
    while robocopy was in fact 100 GB in at 36 MB/s. Every number on that
    line came from a source that lies: robocopy's redirected stdout is
    block-buffered (file lines arrive in flushes that can be an hour apart),
    robocopy pre-allocates each destination file to full size the moment it
    starts it (polling the dest size reads 100% at once), and the NIC
    bytes-sent counter is reported once per NDIS filter driver on the
    adapter (five copies on the machine this was found on). These tests pin
    the replacements: a live tail of robocopy's /UNILOG file, the process's
    own write counter, and a windowed rate over it.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Mirror.Progress.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Mirror.psm1') -Force

    function Get-JoinedSource {
        param([string]$FunctionName)
        $src = (Get-Command $FunctionName).ScriptBlock.ToString()
        return ($src -replace '`\r?\n\s*', ' ')
    }
    $script:mirrorSource = Get-JoinedSource -FunctionName 'Invoke-Mirror'
}

Describe "Read-MirrorLogTail" {
    It "returns nothing while robocopy has not created the log yet" {
        InModuleScope Mirror -Parameters @{ Path = (Join-Path $TestDrive 'missing.log') } {
            param($Path)
            $tail = New-MirrorLogTail -Path $Path
            (Read-MirrorLogTail -Tail $tail).Count | Should -Be 0
        }
    }

    # The log is UTF-16 with a BOM, and a thread can be mid-line when we
    # read. Complete lines come back; the partial one waits for its rest.
    It "reads complete UTF-16 lines as they are appended and holds a partial line back" {
        InModuleScope Mirror -Parameters @{ Path = (Join-Path $TestDrive 'run.log') } {
            param($Path)
            $enc = [Text.Encoding]::Unicode
            $title = "Am$([char]0xE9)lie (2001)"
            $bytes = $enc.GetPreamble() + $enc.GetBytes("header`r`n`t  `t`t  123`tE:\Movies\$title\$title.mkv`r`n`t  `t`t  45")
            [IO.File]::WriteAllBytes($Path, $bytes)

            $tail = New-MirrorLogTail -Path $Path
            $first = Read-MirrorLogTail -Tail $tail
            $first.Count | Should -Be 2
            # Ordinal on purpose: -eq is culture-aware and ignores a leading
            # U+FEFF, so an unskipped BOM would pass a plain Should -Be.
            [string]::Equals($first[0], 'header', [StringComparison]::Ordinal) | Should -BeTrue
            $first[1]    | Should -BeLike "*$title.mkv"

            $stream = [IO.File]::Open($Path, 'Append')
            $more = $enc.GetBytes("6`tE:\b.mkv`r`n`t  `t`t  7`tE:\c.mkv`r`n")
            $stream.Write($more, 0, $more.Length); $stream.Dispose()

            $second = Read-MirrorLogTail -Tail $tail
            $second.Count | Should -Be 2
            $second[0]    | Should -Match '456\tE:\\b\.mkv$'
            $second[1]    | Should -Match '7\tE:\\c\.mkv$'
            (Read-MirrorLogTail -Tail $tail).Count | Should -Be 0
        }
    }

    It "never consumes half a UTF-16 code unit" {
        InModuleScope Mirror -Parameters @{ Path = (Join-Path $TestDrive 'torn.log') } {
            param($Path)
            $enc = [Text.Encoding]::Unicode
            # A complete line, then the first byte of the next character.
            $bytes = $enc.GetPreamble() + $enc.GetBytes("one`r`n") + [byte[]]@(0x74)
            [IO.File]::WriteAllBytes($Path, $bytes)
            $tail = New-MirrorLogTail -Path $Path
            $lines = Read-MirrorLogTail -Tail $tail
            $lines.Count | Should -Be 1
            $lines[0]    | Should -Be 'one'

            $stream = [IO.File]::Open($Path, 'Append')
            $rest = [byte[]]@(0x00) + $enc.GetBytes("wo`r`n")
            $stream.Write($rest, 0, $rest.Length); $stream.Dispose()
            $lines = Read-MirrorLogTail -Tail $tail
            $lines.Count | Should -Be 1
            $lines[0]    | Should -Be 'two'
        }
    }

    It "reads while the writer still holds the file open" {
        InModuleScope Mirror -Parameters @{ Path = (Join-Path $TestDrive 'open.log') } {
            param($Path)
            $enc = [Text.Encoding]::Unicode
            $writer = [IO.File]::Open($Path, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::Read)
            try {
                $bytes = $enc.GetPreamble() + $enc.GetBytes("live`r`n")
                $writer.Write($bytes, 0, $bytes.Length); $writer.Flush()
                $tail = New-MirrorLogTail -Path $Path
                $lines = Read-MirrorLogTail -Tail $tail
                $lines.Count | Should -Be 1
                $lines[0]    | Should -Be 'live'
            } finally {
                $writer.Dispose()
            }
        }
    }
}

Describe "Update-MirrorSpeedWindow" {
    It "needs two samples before it reports a rate" {
        InModuleScope Mirror {
            $q = New-Object 'System.Collections.Generic.Queue[object]'
            Update-MirrorSpeedWindow -Samples $q -ElapsedSec 0 -Bytes 0 | Should -Be 0
        }
    }

    # The old cumulative average kept the ETA optimistic long after the
    # link slowed. The window must follow the current rate.
    It "reports the rate over the window, not the whole run" {
        InModuleScope Mirror {
            $q = New-Object 'System.Collections.Generic.Queue[object]'
            $rate = 0
            foreach ($t in 0..9) { $rate = Update-MirrorSpeedWindow -Samples $q -ElapsedSec $t -Bytes ($t * 100) -WindowSec 5 }
            $rate | Should -Be 100
            $bytes = 900
            foreach ($t in 10..20) { $bytes += 10; $rate = Update-MirrorSpeedWindow -Samples $q -ElapsedSec $t -Bytes $bytes -WindowSec 5 }
            $rate | Should -Be 10
        }
    }

    It "never reports a negative rate after the counter goes backwards" {
        InModuleScope Mirror {
            $q = New-Object 'System.Collections.Generic.Queue[object]'
            [void](Update-MirrorSpeedWindow -Samples $q -ElapsedSec 0 -Bytes 1000)
            Update-MirrorSpeedWindow -Samples $q -ElapsedSec 1 -Bytes 0 | Should -Be 0
        }
    }
}

Describe "Get-MirrorProcessWriteBytes" {
    It "reads the write counter of a live process" {
        InModuleScope Mirror {
            $bytes = Get-MirrorProcessWriteBytes -ProcessId $PID
            $bytes | Should -Not -BeNullOrEmpty
            $bytes | Should -BeGreaterOrEqual 0
        }
    }

    It "returns nothing for a process that is gone" {
        InModuleScope Mirror {
            Get-MirrorProcessWriteBytes -ProcessId 4000000 | Should -BeNullOrEmpty
        }
    }
}

Describe "Get-RobocopyListedFile" {
    # Captured from robocopy /L /NC: an extra on the destination and a new
    # file on the source print in exactly the same shape.
    It "tells an extra to purge from a file to copy by the path root" {
        InModuleScope Mirror {
            $extra = Get-RobocopyListedFile -Line "`t  `t`t      15`tC:\probe\dst\stray.txt" -SourceRoot 'C:\probe\src' -DestRoot 'C:\probe\dst'
            $copy  = Get-RobocopyListedFile -Line "`t  `t`t      15`tC:\probe\src\new.txt"   -SourceRoot 'C:\probe\src' -DestRoot 'C:\probe\dst'
            $extra.Kind | Should -Be 'Extra'
            $copy.Kind  | Should -Be 'Copy'
            $copy.Size  | Should -Be 15
            $copy.Path  | Should -Be 'C:\probe\src\new.txt'
        }
    }

    It "works with a UNC destination and a drive-root source" {
        InModuleScope Mirror {
            $r = Get-RobocopyListedFile -Line "`t  `t`t  8375035319`t\\livingroom\Creighton\Movies\Old (1999)\Old (1999).mkv" -SourceRoot 'E:\Movies' -DestRoot '\\livingroom\Creighton\Movies'
            $r.Kind | Should -Be 'Extra'
            $r = Get-RobocopyListedFile -Line "`t  `t`t  8375035319`tE:\Movies\Backdraft (1991)\Backdraft (1991).mkv" -SourceRoot 'E:\Movies' -DestRoot '\\livingroom\Creighton\Movies'
            $r.Kind | Should -Be 'Copy'
        }
    }

    It "ignores lines that are not file entries" {
        InModuleScope Mirror {
            Get-RobocopyListedFile -Line '   Files :         1         1         0         0         0         1' -SourceRoot 'C:\s' -DestRoot 'C:\d' | Should -BeNullOrEmpty
            Get-RobocopyListedFile -Line '2026/09/09 22:40:01 ERROR 53 (0x00000035) Copying File E:\x.mkv' -SourceRoot 'C:\s' -DestRoot 'C:\d' | Should -BeNullOrEmpty
        }
    }
}

Describe "ConvertFrom-RobocopyOutput" {
    # Captured summary: one file copied, one extra purged.
    It "reads the extras count from the sixth column of the Files row" {
        InModuleScope Mirror {
            $summary = @(
                '               Total    Copied   Skipped  Mismatch    FAILED    Extras'
                '    Dirs :         1         0         1         0         0         0'
                '   Files :         1         1         0         0         0         1'
                '   Bytes :        15        15         0         0         0        15'
            )
            $stats = ConvertFrom-RobocopyOutput $summary
            $stats.FilesCopied  | Should -Be 1
            $stats.FilesDeleted | Should -Be 1
            $stats.BytesCopied  | Should -Be 15
        }
    }
}

Describe "Invoke-Mirror progress sources" {
    It "classifies listed files by root in both the scan and the copy" {
        $script:mirrorSource | Should -Match 'Get-RobocopyListedFile -Line \$scanLine'
        $script:mirrorSource | Should -Match 'Get-RobocopyListedFile -Line \$line -SourceRoot \$source -DestRoot \$dest'
        (Get-Command Get-MirrorPendingChanges).ScriptBlock.ToString() | Should -Match 'Get-RobocopyListedFile'
    }

    It "tails a /UNILOG file instead of trusting buffered stdout" {
        $script:mirrorSource | Should -Match '/UNILOG:'
        $script:mirrorSource | Should -Match 'Read-MirrorLogTail -Tail \$logTail'
    }

    It "takes bytes from robocopy's own write counter" {
        $script:mirrorSource | Should -Match 'Get-MirrorProcessWriteBytes -ProcessId \$process\.Id'
        $script:mirrorSource | Should -Match 'Update-MirrorSpeedWindow'
    }

    It "no longer sums NIC counters or polls the destination file size" {
        $script:mirrorSource | Should -Not -Match 'Get-MirrorNetworkBytesSent'
        $script:mirrorSource | Should -Not -Match 'currentDestFilePath'
    }

    It "counts files as started, since robocopy logs a file when a thread picks it up" {
        $script:mirrorSource | Should -Match 'folderFilesStarted\+\+'
        $script:mirrorSource | Should -Match 'started'
    }

    # Size alone never proves a file is complete once robocopy has
    # pre-allocated it; stamping the source mtime onto a partial hid it
    # from every later run.
    It "no longer stamps source timestamps onto destination files after a cancel" {
        $script:mirrorSource | Should -Not -Match '\.LastWriteTime\s*='
    }

    It "treats any log activity as a sign of life for the watchdog" {
        $script:mirrorSource | Should -Match 'if \(\$processedLine\) \{ \$stallSince = \$null'
    }
}
