#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for the NFO identity check: the runtime comparison, the
    duration cache, the TMDB candidate pick, and the repair walk.

.DESCRIPTION
    "Split (2016)" held Shyamalan's 117-minute film with an NFO for Deborah
    Kampmeier's 150-minute Split. Every name-based check passed: the title
    matched, the year matched, the NFO was well-formed. The video's real
    length is the one fact that could not be fooled, so the health check now
    compares NFO runtime with measured duration (cached per folder, probed
    within a budget) and re-identifies against TMDB by runtime. MediaInfo
    and TMDB are mocked; everything runs offline in TestDrive.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\MovieIdentity.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Quality.psm1') -Force
    Import-Module (Join-Path $repoRoot 'modules\TMDB.psm1') -Force

    $scriptPath = Join-Path $repoRoot 'LibraryLint.ps1'
    $parseErrors = $null
    $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) { throw "LibraryLint.ps1 has parse errors: $($parseErrors[0].Message)" }
    foreach ($name in 'Invoke-MovieIdentityRepair', 'Get-NormalizedTitle') {
        $fn = $scriptAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $fn) { throw "$name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:Config = @{ TMDBApiKey = 'k' }
    function Write-Log { param($Message, $Level) }
    function New-MovieNFOFromTMDB { param($Metadata, $NFOPath) $true }

    $script:shyamalan = [PSCustomObject]@{ Id = 381288; Title = 'Split'; Year = '2017'; Votes = 18606; Runtime = 117; Director = 'M. Night Shyamalan' }
    $script:kampmeier = [PSCustomObject]@{ Id = 358364; Title = 'Split'; Year = '2016'; Votes = 56;    Runtime = 150; Director = 'Deborah Kampmeier' }
}

Describe "Test-NfoRuntimeMismatch" {
    It "flags Split: NFO 150 min against a 117-minute video" {
        Test-NfoRuntimeMismatch -NfoRuntimeMin 150 -VideoDurationSec 7025 | Should -BeTrue
    }

    # Normal gaps: rounding, dropped credits, PAL speed-up.
    It "tolerates the ordinary small differences" {
        Test-NfoRuntimeMismatch -NfoRuntimeMin 117 -VideoDurationSec 7025 | Should -BeFalse   # exact
        Test-NfoRuntimeMismatch -NfoRuntimeMin 100 -VideoDurationSec 5760 | Should -BeFalse   # 96 min, PAL
        Test-NfoRuntimeMismatch -NfoRuntimeMin 90  -VideoDurationSec 5040 | Should -BeFalse   # 84 min, credits
        Test-NfoRuntimeMismatch -NfoRuntimeMin 30  -VideoDurationSec 1560 | Should -BeFalse   # 4 min on a short: under the floor
    }

    # A four-minute gap on a twenty-minute short is 20%, but four minutes is
    # nothing a rip cannot lose; the absolute floor exists for exactly this.
    It "keeps the five-minute floor even when the percentage is large" {
        Test-NfoRuntimeMismatch -NfoRuntimeMin 20 -VideoDurationSec 960 | Should -BeFalse
        Test-NfoRuntimeMismatch -NfoRuntimeMin 20 -VideoDurationSec 900 | Should -BeTrue    # 5 min, 25%
    }

    It "says nothing when either side is unknown" {
        Test-NfoRuntimeMismatch -NfoRuntimeMin 0   -VideoDurationSec 7025 | Should -BeFalse
        Test-NfoRuntimeMismatch -NfoRuntimeMin 117 -VideoDurationSec 0    | Should -BeFalse
    }
}

Describe "Get-CachedVideoDuration" {
    BeforeEach {
        $script:dir = Join-Path $TestDrive "m-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
        New-Item -ItemType Directory -Path $script:dir -Force | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $script:dir 'Movie (2020).mkv'), (New-Object byte[] 4096))
        Set-Content -Path (Join-Path $script:dir 'release-info.json') -Value '{"OriginalFileName":"x.mkv","Resolution":"1080p"}'
        $script:probeCalls = 0
        $script:prober = { param($p) $script:probeCalls++; 7025 }
    }

    It "probes once, caches against the file size, and reads the cache afterwards" {
        $video = Get-Item (Join-Path $script:dir 'Movie (2020).mkv')
        $first = Get-CachedVideoDuration -FolderPath $script:dir -Video $video -Prober $script:prober
        $first.Seconds | Should -Be 7025
        $first.Source  | Should -Be 'probed'
        $second = Get-CachedVideoDuration -FolderPath $script:dir -Video $video -Prober $script:prober
        $second.Seconds | Should -Be 7025
        $second.Source  | Should -Be 'cached'
        $script:probeCalls | Should -Be 1
        $info = Get-Content (Join-Path $script:dir 'release-info.json') -Raw | ConvertFrom-Json
        $info.DetectedDurationSec | Should -Be 7025
        $info.DetectedDurationFileSize | Should -Be 4096
        $info.Resolution | Should -Be '1080p'   # existing fields survive the rewrite
    }

    It "re-probes when the video file changed size" {
        $video = Get-Item (Join-Path $script:dir 'Movie (2020).mkv')
        $null = Get-CachedVideoDuration -FolderPath $script:dir -Video $video -Prober $script:prober
        [IO.File]::WriteAllBytes($video.FullName, (New-Object byte[] 8192))
        $video = Get-Item $video.FullName
        $again = Get-CachedVideoDuration -FolderPath $script:dir -Video $video -Prober $script:prober
        $again.Source | Should -Be 'probed'
        $script:probeCalls | Should -Be 2
    }

    It "reports unmeasured instead of probing when the budget is spent" {
        $video = Get-Item (Join-Path $script:dir 'Movie (2020).mkv')
        $r = Get-CachedVideoDuration -FolderPath $script:dir -Video $video -Prober $script:prober -NoProbe
        $r.Seconds | Should -Be 0
        $r.Source  | Should -Be 'unmeasured'
        $script:probeCalls | Should -Be 0
    }
}

Describe "Select-TMDBCandidateByRuntime" {
    It "picks the film whose runtime fits the video" {
        (Select-TMDBCandidateByRuntime -Candidates @($script:kampmeier, $script:shyamalan) -VideoDurationSec 7025).Id | Should -Be 381288
    }

    It "returns nothing when no candidate is within tolerance (an alternate cut, say)" {
        Select-TMDBCandidateByRuntime -Candidates @($script:kampmeier, $script:shyamalan) -VideoDurationSec 8400 | Should -BeNullOrEmpty
    }
}

Describe "Get-TMDBCandidates" {
    It "returns same-title hits within a year, best-known first, each with its runtime" {
        InModuleScope TMDB {
            Mock Invoke-RestMethod {
                [PSCustomObject]@{ results = @(
                    [PSCustomObject]@{ id = 381288; title = 'Split'; original_title = 'Split'; release_date = '2017-01-19'; vote_count = 18606 }
                    [PSCustomObject]@{ id = 358364; title = 'Split'; original_title = 'Split'; release_date = '2016-04-07'; vote_count = 56 }
                    [PSCustomObject]@{ id = 53373;  title = 'Split Estate'; original_title = 'Split Estate'; release_date = '2009-08-07'; vote_count = 0 }
                    [PSCustomObject]@{ id = 729927; title = 'Split'; original_title = 'Split'; release_date = '2018-10-12'; vote_count = 1 }
                ) }
            } -ParameterFilter { $Uri -like '*search/movie*' }
            Mock Get-TMDBMovieDetails { param($MovieId, $ApiKey) @{ Runtime = $(if ($MovieId -eq 381288) { 117 } else { 150 }); Directors = @('Someone') } }
            $c = @(Get-TMDBCandidates -Title 'Split' -Year '2016' -ApiKey 'k')
            @($c | ForEach-Object { $_.Id }) | Should -Be @(381288, 358364)
            ($c | Where-Object { $_.Id -eq 381288 }).Runtime | Should -Be 117
        }
    }
}

Describe "Invoke-MovieIdentityRepair" {
    BeforeEach {
        Mock Write-Host { }
        Mock Get-TMDBCandidates { @($script:kampmeier, $script:shyamalan) }
        Mock Get-TMDBMovieDetails { @{ Id = 381288; TMDBID = 381288; Title = 'Split'; Year = '2017'; Runtime = 117 } }
        Mock New-MovieNFOFromTMDB { $true }
        Mock Set-MovieIdentityConfirmed { $true }
        $script:splitItem = @{
            Folder = 'Split (2016)'; FolderPath = 'E:\Movies\Split (2016)'; NfoPath = 'E:\Movies\Split (2016)\Split (2016).nfo'
            NfoTitle = 'Split'; NfoYear = '2016'; NfoRuntime = 150; NfoTmdbId = '358364'; VideoMinutes = 117; VideoSeconds = 7025
        }
    }

    It "regenerates the NFO from the runtime-matched film on Y" {
        Mock Read-Host { 'Y' }
        $left = @(Invoke-MovieIdentityRepair -Items @($script:splitItem))
        $left.Count | Should -Be 0
        Should -Invoke New-MovieNFOFromTMDB -Times 1 -Exactly -ParameterFilter { $NFOPath -eq 'E:\Movies\Split (2016)\Split (2016).nfo' -and $Metadata.TMDBID -eq 381288 }
        Should -Invoke Get-TMDBMovieDetails -Times 1 -ParameterFilter { $MovieId -eq 381288 }
    }

    It "records the identity as confirmed on S and writes no NFO" {
        Mock Read-Host { 'S' }
        $left = @(Invoke-MovieIdentityRepair -Items @($script:splitItem))
        $left.Count | Should -Be 0
        Should -Invoke Set-MovieIdentityConfirmed -Times 1 -ParameterFilter { $FolderPath -eq 'E:\Movies\Split (2016)' }
        Should -Invoke New-MovieNFOFromTMDB -Times 0
    }

    It "leaves the item unresolved on N and touches nothing" {
        Mock Read-Host { 'N' }
        $left = @(Invoke-MovieIdentityRepair -Items @($script:splitItem))
        $left.Count | Should -Be 1
        Should -Invoke New-MovieNFOFromTMDB -Times 0
        Should -Invoke Set-MovieIdentityConfirmed -Times 0
    }

    It "offers only a confirmation when no candidate fits the runtime" {
        Mock Get-TMDBCandidates { @($script:kampmeier) }
        Mock Read-Host { 'N' }
        $left = @(Invoke-MovieIdentityRepair -Items @($script:splitItem))
        $left.Count | Should -Be 1
        Should -Invoke New-MovieNFOFromTMDB -Times 0
        Should -Invoke Read-Host -Times 1 -ParameterFilter { $Prompt -like '*Mark this identity as confirmed*' }
    }

    It "writes nothing under -WhatIf" {
        Mock Read-Host { 'Y' }
        $left = @(Invoke-MovieIdentityRepair -Items @($script:splitItem) -WhatIf)
        $left.Count | Should -Be 1
        Should -Invoke New-MovieNFOFromTMDB -Times 0
    }
}
