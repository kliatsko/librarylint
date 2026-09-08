#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for subtitle provenance tracking (modules\Subtitles.psm1).

.DESCRIPTION
    Every verified movie folder carries a .subs_ok marker recording WHERE
    its subtitles came from, so "which of these did Whisper write?" stays
    answerable months later. These tests cover the marker's provider
    resolution and its append-only history, including the download-time
    .sub_pending note that carries provenance across the gap between a
    name-matched download and the verification gate that judges it.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Subtitles.Provenance.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Subtitles.psm1') -Force

    function New-TestSubFolder {
        param([string]$Root, [string]$Name)
        $dir = Join-Path $Root $Name
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $dir "$Name.mkv") -Value 'video' -NoNewline
        Set-Content -LiteralPath (Join-Path $dir "$Name.en.srt") -Value '1' -NoNewline
        return $dir
    }

    function Get-Marker {
        param([string]$FolderPath)
        return (Get-Content -LiteralPath (Join-Path $FolderPath '.subs_ok') -Raw | ConvertFrom-Json)
    }

    # Writes the download-time note the way Save-MovieSubtitle does.
    function Write-PendingNote {
        param([string]$FolderPath, [string]$Provider)
        @{ Provider = $Provider; Detail = 'test.en.srt'; Date = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $FolderPath '.sub_pending') -Encoding UTF8
    }
}

Describe "Set-SubtitlesVerified provider resolution" {
    It "records an explicit provider" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Explicit (2001)'
        Set-SubtitlesVerified -FolderPath $folder -Source 'whisper-large-v3 (en -> en, cuda)' -Provider 'whisper' | Should -BeTrue
        (Get-Marker $folder).Provider | Should -Be 'whisper'
    }

    # A sub that was simply present, with nothing recorded about its
    # origin, shipped with the release or was added by hand before
    # provenance tracking existed.
    It "falls back to 'release' when nothing is known" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Unknown Origin (2002)'
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'sync-verified (offset 0.3s)'
        (Get-Marker $folder).Provider | Should -Be 'release'
    }

    It "inherits the provider from a .sub_pending note and consumes it" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'From Pending (2003)'
        Write-PendingNote -FolderPath $folder -Provider 'subdl'
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'sync-verified (offset 0.1s)'
        (Get-Marker $folder).Provider | Should -Be 'subdl'
        Test-Path -LiteralPath (Join-Path $folder '.sub_pending') | Should -BeFalse
    }

    It "prefers an explicit provider over a stale pending note" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Explicit Wins (2004)'
        Write-PendingNote -FolderPath $folder -Provider 'subdl'
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'whisper-large-v3' -Provider 'whisper'
        (Get-Marker $folder).Provider | Should -Be 'whisper'
        Test-Path -LiteralPath (Join-Path $folder '.sub_pending') | Should -BeFalse
    }

    # Regression: without this, a second verification pass over a folder
    # whose sub came from a known provider would relabel it 'release',
    # silently erasing the origin.
    #
    # The stored Source here deliberately does NOT name its provider. That
    # is the real shape of a name-matched download: the gate stamps
    # "sync-verified", and only the Provider field remembers where the file
    # came from. A self-describing source such as "opensubtitles-hashmatch"
    # would let the inference fallback recover the provider anyway, so the
    # test would pass even with this branch removed.
    It "keeps the recorded provider when the folder is verified again" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Re-verified (2005)'
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'sync-verified (offset 0.1s)' -Provider 'subdl'
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'ffsubsync-corrected (offset 2.0s)'
        (Get-Marker $folder).Provider | Should -Be 'subdl'
    }

    It "accepts the hardsub provider for burned-in subtitles" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Burned In (2006)'
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'hardsub-audit (88% coverage)' -Provider 'hardsub'
        (Get-Marker $folder).Provider | Should -Be 'hardsub'
    }

    It "rejects a provider outside the known set" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Bad Provider (2007)'
        { Set-SubtitlesVerified -FolderPath $folder -Source 'x' -Provider 'bittorrent' -ErrorAction Stop } | Should -Throw
    }
}

Describe "Set-SubtitlesVerified history" {
    It "appends each event instead of replacing the last" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Appends (2008)'
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'sync-verified (offset 0.1s)' -Provider 'subdl'
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'ffsubsync-corrected (offset 2.0s)'
        $history = @((Get-Marker $folder).History)
        $history.Count | Should -Be 2
        $history[0].Source | Should -BeLike 'sync-verified*'
        $history[1].Source | Should -BeLike 'ffsubsync-corrected*'
    }

    It "stamps every history entry with its provider and date" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Stamped (2009)'
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'embedded-extraction' -Provider 'embedded'
        $entry = @((Get-Marker $folder).History)[0]
        $entry.Provider | Should -Be 'embedded'
        $entry.Date | Should -Not -BeNullOrEmpty
    }

    # Markers written before provenance existed have a Source string and
    # no History. They must migrate rather than be discarded, so the
    # library's existing verification work survives the upgrade.
    It "migrates a legacy marker into history with an inferred provider" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Legacy (2010)'
        @{ VerifiedDate = '2026-08-01 10:00:00'; Source = 'opensubtitles-hashmatch (Old.Release)'; VerifiedBy = 'LibraryLint' } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $folder '.subs_ok') -Encoding UTF8
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'sync-verified (offset 0.0s)'
        $marker = Get-Marker $folder
        $marker.Provider | Should -Be 'opensubtitles'
        $legacyEntry = @($marker.History)[0]
        $legacyEntry.Provider | Should -Be 'opensubtitles'
        $legacyEntry.Date | Should -Be '2026-08-01 10:00:00'
    }

    It "infers whisper and embedded providers from legacy source strings" {
        $cases = @(
            @{ Source = 'whisper-large-v3 (ja -> en, cuda)'; Expected = 'whisper' }
            @{ Source = 'embedded-extraction';               Expected = 'embedded' }
            @{ Source = 'included';                          Expected = 'release' }
        )
        foreach ($case in $cases) {
            $folder = New-TestSubFolder -Root $TestDrive -Name "Legacy $($case.Expected) (2011)"
            @{ VerifiedDate = '2026-08-01 10:00:00'; Source = $case.Source; VerifiedBy = 'LibraryLint' } |
                ConvertTo-Json | Set-Content -LiteralPath (Join-Path $folder '.subs_ok') -Encoding UTF8
            $null = Set-SubtitlesVerified -FolderPath $folder -Source 'sync-verified (offset 0.0s)'
            (Get-Marker $folder).Provider | Should -Be $case.Expected -Because "source '$($case.Source)' implies $($case.Expected)"
        }
    }
}

Describe "Set-SubtitlePendingProvenance" {
    # Producer and consumer must agree on the note's shape: this drives
    # the real writer (private to the module) and then lets the real
    # reader consume what it wrote.
    It "writes a note the verification path can consume" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Producer Consumer (2012)'
        InModuleScope Subtitles -Parameters @{ Folder = $folder } {
            Set-SubtitlePendingProvenance -FolderPath $Folder -Provider 'subdl' -Detail 'Producer Consumer (2012).en.srt'
        }
        Test-Path -LiteralPath (Join-Path $folder '.sub_pending') | Should -BeTrue
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'sync-verified (offset 0.2s)'
        (Get-Marker $folder).Provider | Should -Be 'subdl'
    }
}

Describe "Test-SubtitlesVerified" {
    It "reports true only once a marker exists" {
        $folder = New-TestSubFolder -Root $TestDrive -Name 'Marker Gate (2013)'
        Test-SubtitlesVerified -FolderPath $folder | Should -BeFalse
        $null = Set-SubtitlesVerified -FolderPath $folder -Source 'manual' -Provider 'manual'
        Test-SubtitlesVerified -FolderPath $folder | Should -BeTrue
    }
}
