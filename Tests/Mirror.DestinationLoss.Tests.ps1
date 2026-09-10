#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for how the mirror reacts when its destination goes away
    mid-run (modules\Mirror.psm1).

.DESCRIPTION
    A mirror against an SMB share that vanished used to fail every file one
    by one: robocopy retried each (/R:2 /W:5), the loop counted each failure
    as "in use by another process", and 1,413 pending files became twenty
    minutes of red. These tests pin the classifier and the abort rule using
    robocopy's real output lines, so the decision "the destination is gone"
    is exercised without a process or a share. The rule is deliberately not
    "abort on the first network error": a short SMB blip under /MT:16 prints
    a burst of first-attempt errors that robocopy's own retries survive.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Mirror.DestinationLoss.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Mirror.psm1') -Force

    # Continuation-joined source, so a one-line pattern can see arguments
    # that wrap onto the next line (same reason as Sync.AlreadyHave.Tests).
    function Get-JoinedSource {
        param([string]$FunctionName)
        $src = (Get-Command $FunctionName).ScriptBlock.ToString()
        return ($src -replace '`\r?\n\s*', ' ')
    }
    $script:mirrorSource  = Get-JoinedSource -FunctionName 'Invoke-Mirror'
    $script:pendingSource = Get-JoinedSource -FunctionName 'Get-MirrorPendingChanges'
}

Describe "Get-RobocopyErrorClass" {
    # Private helpers, so every test runs inside the module scope.
    It "classifies error 53 copying a file as the destination being gone" {
        InModuleScope Mirror {
            $r = Get-RobocopyErrorClass -Line '2026/09/09 22:40:01 ERROR 53 (0x00000035) Copying File E:\Movies\Begin Again (2013)\Begin Again (2013).mkv'
            $r.Class       | Should -Be 'DestinationLost'
            $r.Code        | Should -Be 53
            $r.Operation   | Should -Be 'Copying File'
            $r.Target      | Should -Be 'E:\Movies\Begin Again (2013)\Begin Again (2013).mkv'
            $r.IsDirectory | Should -BeFalse
            $r.IsReason    | Should -BeFalse
        }
    }

    # Captured tonight against a host that does not exist.
    It "flags a destination-directory failure as directory-level" {
        InModuleScope Mirror {
            $r = Get-RobocopyErrorClass -Line '2026/09/09 22:55:14 ERROR 53 (0x00000035) Creating Destination Directory \\no-such-host-llprobe\share\Movies\'
            $r.Class       | Should -Be 'DestinationLost'
            $r.IsDirectory | Should -BeTrue
            $r.Target      | Should -Be '\\no-such-host-llprobe\share\Movies\'
        }
    }

    # Captured tonight against livingroom with a share name that does not exist.
    It "classifies error 67 (share gone while the host is up) the same way" {
        InModuleScope Mirror {
            $r = Get-RobocopyErrorClass -Line '2026/09/09 22:55:15 ERROR 67 (0x00000043) Creating Destination Directory \\livingroom\NoSuchShareLLProbe\Movies\'
            $r.Class | Should -Be 'DestinationLost'
            $r.Code  | Should -Be 67
        }
    }

    It "classifies the other network-loss codes by number alone" {
        InModuleScope Mirror {
            foreach ($code in 51, 58, 59, 64, 121, 1231, 1232) {
                $r = Get-RobocopyErrorClass -Line "2026/09/09 22:40:01 ERROR $code (0x00000000) Copying File E:\x.mkv"
                $r.Class | Should -Be 'DestinationLost' -Because "code $code means the share, not the file, is the problem"
            }
        }
    }

    It "classifies a sharing violation as in-use, never fatal" {
        InModuleScope Mirror {
            $r = Get-RobocopyErrorClass -Line '2026/09/09 22:40:01 ERROR 32 (0x00000020) Copying File E:\Movies\X\x.mkv'
            $r.Class | Should -Be 'InUse'
        }
    }

    It "classifies error 5 as access denied and knows a delete from a copy" {
        InModuleScope Mirror {
            $r = Get-RobocopyErrorClass -Line '2026/09/09 22:40:01 ERROR 5 (0x00000005) Deleting Extra File \\livingroom\Creighton\Movies\Old\old.nfo'
            $r.Class     | Should -Be 'AccessDenied'
            $r.Operation | Should -Be 'Deleting Extra File'
        }
    }

    # With /MT the ERROR line and its reason line come from different
    # threads, so the reason must classify on its own.
    It "recognises a bare reason line without a code" {
        InModuleScope Mirror {
            $r = Get-RobocopyErrorClass -Line 'The network path was not found.'
            $r.Class    | Should -Be 'DestinationLost'
            $r.IsReason | Should -BeTrue
            $r.Code     | Should -BeNullOrEmpty
        }
    }

    It "returns nothing for announcements, summary rows and banners" {
        InModuleScope Mirror {
            Get-RobocopyErrorClass -Line "`t  `t`t  4957495130`tE:\Movies\Deadpool (2016)\Deadpool (2016).mkv" | Should -BeNullOrEmpty
            Get-RobocopyErrorClass -Line '   Files :         2         2         0         0         0         0' | Should -BeNullOrEmpty
            Get-RobocopyErrorClass -Line '   ROBOCOPY     ::     Robust File Copy for Windows' | Should -BeNullOrEmpty
            Get-RobocopyErrorClass -Line '' | Should -BeNullOrEmpty
        }
    }
}

Describe "Update-RobocopyErrorState" {
    # Sixteen threads all hit a two-second blip: sixteen first-attempt
    # errors, then robocopy waits and retries successfully. No file has
    # failed, so nothing may abort.
    It "does not abort on a blip that robocopy's own retries survive" {
        InModuleScope Mirror {
            $s = New-RobocopyErrorState -RetryLimit 2
            foreach ($i in 1..16) {
                foreach ($l in @("2026/09/09 22:40:01 ERROR 53 (0x00000035) Copying File E:\Movies\M$i\M$i.mkv",
                                 'The network path was not found.',
                                 'Waiting 5 seconds... Retrying...')) {
                    [void](Update-RobocopyErrorState -Line $l -State $s)
                }
            }
            $s.DestLost            | Should -BeFalse
            $s.DestLostFiles.Count | Should -Be 0
            $s.ErrorsByClass.Count | Should -Be 0
        }
    }

    It "one file failing every retry is a bad file, not an outage" {
        InModuleScope Mirror {
            $s = New-RobocopyErrorState -RetryLimit 2
            foreach ($attempt in 1..3) {
                [void](Update-RobocopyErrorState -Line '2026/09/09 22:40:01 ERROR 53 (0x00000035) Copying File E:\Movies\A\a.mkv' -State $s)
                [void](Update-RobocopyErrorState -Line 'The network path was not found.' -State $s)
            }
            $s.DestLost                        | Should -BeFalse
            $s.DestLostFiles.Count             | Should -Be 1
            $s.ErrorsByClass['DestinationLost'] | Should -Be 1
        }
    }

    It "aborts once three distinct files have failed every retry" {
        InModuleScope Mirror {
            $s = New-RobocopyErrorState -RetryLimit 2
            $files = 'E:\Movies\A\a.mkv', 'E:\Movies\B\b.mkv', 'E:\Movies\C\c.mkv'
            $seen = @()
            foreach ($file in $files) {
                foreach ($attempt in 1..3) {
                    [void](Update-RobocopyErrorState -Line "2026/09/09 22:40:01 ERROR 53 (0x00000035) Copying File $file" -State $s)
                    [void](Update-RobocopyErrorState -Line 'The network path was not found.' -State $s)
                }
                $seen += $s.DestLost
            }
            $seen[0] | Should -BeFalse
            $seen[1] | Should -BeFalse
            $seen[2] | Should -BeTrue
            $s.DestLostDetail | Should -Match '53'
            $s.DestLostDetail | Should -Match 'network path was not found'
            $s.DestLostDetail | Should -Match '3 files'
        }
    }

    # Robocopy has already retried before it prints this, and a directory
    # it cannot create is never a per-file problem.
    It "aborts at once on a destination-directory failure" {
        InModuleScope Mirror {
            $s = New-RobocopyErrorState -RetryLimit 2
            [void](Update-RobocopyErrorState -Line '2026/09/09 22:55:14 ERROR 53 (0x00000035) Creating Destination Directory \\no-such-host-llprobe\share\Movies\' -State $s)
            $s.DestLost       | Should -BeTrue
            $s.DestLostDetail | Should -Match 'creating destination directory'
            $s.DestLostDetail | Should -Match 'no-such-host-llprobe'
        }
    }

    It "never aborts on in-use or access-denied files, and counts each file once" {
        InModuleScope Mirror {
            $s = New-RobocopyErrorState -RetryLimit 2
            foreach ($attempt in 1..3) {
                [void](Update-RobocopyErrorState -Line '2026/09/09 22:40:01 ERROR 32 (0x00000020) Copying File E:\Movies\A\a.mkv' -State $s)
                [void](Update-RobocopyErrorState -Line 'The process cannot access the file because it is being used by another process.' -State $s)
                [void](Update-RobocopyErrorState -Line '2026/09/09 22:40:01 ERROR 5 (0x00000005) Copying File E:\Movies\B\b.mkv' -State $s)
                [void](Update-RobocopyErrorState -Line 'Access is denied.' -State $s)
            }
            $s.DestLost                     | Should -BeFalse
            $s.ErrorsByClass['InUse']        | Should -Be 1
            $s.ErrorsByClass['AccessDenied'] | Should -Be 1
        }
    }

    # Delete-permission errors are routine on Samba/Kodi shares and were
    # always counted silently; that must survive the rewrite.
    It "counts extras that could not be deleted separately and silently" {
        InModuleScope Mirror {
            $s = New-RobocopyErrorState -RetryLimit 2
            foreach ($attempt in 1..3) {
                [void](Update-RobocopyErrorState -Line '2026/09/09 22:40:01 ERROR 5 (0x00000005) Deleting Extra File \\livingroom\Creighton\Movies\Old\old.nfo' -State $s)
                [void](Update-RobocopyErrorState -Line 'Access is denied.' -State $s)
            }
            $s.DeleteErrors        | Should -Be 1
            $s.ErrorsByClass.Count | Should -Be 0
            $s.Messages.Count      | Should -Be 0
        }
    }

    It "pairs the reason line with its file in one message" {
        InModuleScope Mirror {
            $s = New-RobocopyErrorState -RetryLimit 2
            [void](Update-RobocopyErrorState -Line '2026/09/09 22:40:01 ERROR 53 (0x00000035) Copying File E:\Movies\Begin Again (2013)\Begin Again (2013).mkv' -State $s)
            [void](Update-RobocopyErrorState -Line 'The network path was not found.' -State $s)
            $s.Messages.Count   | Should -Be 1
            $s.Messages[0].Text | Should -Be 'ERROR: Begin Again (2013).mkv - The network path was not found.'
            $s.Messages[0].Color | Should -Be 'Red'
        }
    }

    # Tonight's screenshot: two threads' ERROR lines, then their two reason
    # lines. The old parser dropped the first file's name entirely.
    It "does not lose a file whose reason line was interleaved away" {
        InModuleScope Mirror {
            $s = New-RobocopyErrorState -RetryLimit 2
            foreach ($l in @('2026/09/09 22:40:01 ERROR 53 (0x00000035) Copying File E:\Movies\A\a.mkv',
                             '2026/09/09 22:40:01 ERROR 53 (0x00000035) Copying File E:\Movies\B\b.mkv',
                             'The network path was not found.',
                             'The network path was not found.')) {
                [void](Update-RobocopyErrorState -Line $l -State $s)
            }
            $texts = @($s.Messages | ForEach-Object { $_.Text })
            ($texts -join "`n") | Should -Match 'a\.mkv'
            ($texts -join "`n") | Should -Match 'b\.mkv'
            $texts | Where-Object { $_ -match 'a\.mkv' } | Should -Match 'network path was not found'
        }
    }

    It "consumes wait, retry and retry-limit lines but not file announcements" {
        InModuleScope Mirror {
            $s = New-RobocopyErrorState -RetryLimit 2
            Update-RobocopyErrorState -Line 'Waiting 5 seconds... Retrying...' -State $s | Should -BeTrue
            Update-RobocopyErrorState -Line 'ERROR: RETRY LIMIT EXCEEDED.' -State $s     | Should -BeTrue
            Update-RobocopyErrorState -Line "`t  `t`t  4957495130`tE:\Movies\Deadpool (2016)\Deadpool (2016).mkv" -State $s | Should -BeFalse
            $s.DestLost | Should -BeFalse
        }
    }
}

Describe "Invoke-Mirror destination handling" {
    # The scan phase takes minutes; a host that died in between must be one
    # message before the copy, not a wall of per-file errors after it.
    It "probes the destination before launching the copy" {
        $probeAt = $script:mirrorSource.IndexOf('Test-MirrorDestAlive -DestRoot $dest')
        $startAt = $script:mirrorSource.IndexOf('$process.Start()')
        $probeAt | Should -BeGreaterThan -1
        $startAt | Should -BeGreaterThan -1
        $probeAt | Should -BeLessThan $startAt
    }

    It "routes robocopy errors through the error state and aborts on its verdict" {
        $script:mirrorSource | Should -Match 'Update-RobocopyErrorState -Line \$line -State \$errorState'
        $script:mirrorSource | Should -Match 'Write-MirrorDestinationLostAbort -Dest \$dest -Detail \$errorState\.DestLostDetail'
    }

    It "no longer labels every copy error as in use by another process" {
        $script:mirrorSource | Should -Not -Match '\$copyErrors'
    }
}

Describe "Get-MirrorPendingChanges destination handling" {
    # robocopy /L lists every file as new against a dead host and exits 1
    # with no error line, so the dashboard needs its own probe after the scan.
    It "re-probes the destination after the scan" {
        $script:pendingSource | Should -Match 'Test-MirrorDestAlive -DestRoot \$DestDrive'
    }
}
