#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for the sync's "do I already have this file" decision
    (modules\Sync.psm1).

.DESCRIPTION
    The sync used to refuse any remote file whose parent folder name matched
    a library folder. That answers "do I own this title", not "do I have this
    file", and it turned away every upgrade Radarr fetched — a replacement
    always lands in a folder we already have. Quality is judged once, at
    library-add time. These tests pin the two tiers that remain and pin
    that the sync no longer consults folder names at all.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Sync.AlreadyHave.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Sync.psm1') -Force

    # Calls in this module are wrapped with backtick continuations, so a
    # one-line regex would stop at the line break and miss a -FolderSet on
    # the next line. Joining continuations first makes "the call carries
    # -FolderSet" a question the pattern can actually answer — without this
    # the negative assertion below passed regardless of the code.
    function Get-JoinedSource {
        param([string]$FunctionName)
        $src = (Get-Command $FunctionName).ScriptBlock.ToString()
        return ($src -replace '`\r?\n\s*', ' ')
    }
    $script:syncSource = Get-JoinedSource -FunctionName 'Invoke-SFTPSync'
    $script:reconcileSource = Get-JoinedSource -FunctionName 'Update-SFTPTrackingFromLocal'
}

Describe "Test-RemoteFileAlreadyHave" {
    # Private to the module, so the tests run inside its scope.
    It "matches an identical file by name and size" {
        InModuleScope Sync {
            $index = @{ 'Deadpool (2016).mkv|4957495130' = 'E:\Movies\Deadpool (2016)\Deadpool (2016).mkv' }
            $r = Test-RemoteFileAlreadyHave -FileName 'Deadpool (2016).mkv' -FileSize 4957495130 -FileIndex $index
            $r.Found | Should -BeTrue
            $r.MatchType | Should -Be 'NameSize'
            $r.LocalPath | Should -BeLike '*Deadpool*'
        }
    }

    It "does not match the same name at a different size" {
        InModuleScope Sync {
            $index = @{ 'Deadpool (2016).mkv|4957495130' = 'E:\x' }
            $r = Test-RemoteFileAlreadyHave -FileName 'Deadpool (2016).mkv' -FileSize 1 -FileIndex $index
            $r.Found | Should -BeFalse
        }
    }

    # A REPACK BluRay arriving in a folder called "Deadpool (2016)" is exactly
    # the case the sync must let through.
    It "does not match by folder name when no folder set is supplied" {
        InModuleScope Sync {
            $index = @{ 'Deadpool (2016).mkv|4957495130' = 'E:\x' }
            $r = Test-RemoteFileAlreadyHave -FileName 'Deadpool.2016.REPACK.1080p.BluRay.mkv' -FileSize 9999 `
                -RemoteParentName 'Deadpool (2016)' -FileIndex $index
            $r.Found | Should -BeFalse
        }
    }

    # The folder tier still exists for the explicit reconcile tool, which is
    # a deliberate "trust what is on disk" operation the user runs on purpose.
    It "still matches by folder name when a folder set is explicitly supplied" {
        InModuleScope Sync {
            $index = @{}
            $folders = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            [void]$folders.Add('Deadpool (2016)')
            $r = Test-RemoteFileAlreadyHave -FileName 'Deadpool.2016.REPACK.1080p.BluRay.mkv' -FileSize 9999 `
                -RemoteParentName 'Deadpool (2016)' -FileIndex $index -FolderSet $folders
            $r.Found | Should -BeTrue
            $r.MatchType | Should -Be 'Folder'
        }
    }
}

Describe "Invoke-SFTPSync already-have policy" {
    # The regression that blocked the hardsub replacements: the daily sync
    # consulting library folder names. Pinned at the source so it cannot
    # quietly return.
    It "never passes a folder set to the already-have check" {
        $script:syncSource | Should -Not -Match 'Test-RemoteFileAlreadyHave[^\n]*-FolderSet'
    }

    It "no longer builds a library folder set at all" {
        $script:syncSource | Should -Not -Match 'Build-LocalFolderSet'
    }

    It "no longer reports 'already in library' as a skip reason" {
        $script:syncSource | Should -Not -Match 'already in library'
    }
}

Describe "Update-SFTPTrackingFromLocal reconcile policy" {
    # Reconcile is the one place folder matching is still right: the user is
    # explicitly asking to mark what is on disk as downloaded. This positive
    # match also proves the joined-source pattern can see a continuation-line
    # -FolderSet, which is what gives the negative assertion above its teeth.
    It "still reconciles by folder name" {
        $script:reconcileSource | Should -Match 'Build-LocalFolderSet'
        $script:reconcileSource | Should -Match 'Test-RemoteFileAlreadyHave[^\n]*-FolderSet'
    }
}
