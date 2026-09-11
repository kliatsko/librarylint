#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for Invoke-SubtitleProcessing (LibraryLint.ps1), the
    subtitle step of initial processing.

.DESCRIPTION
    Initial processing moves a release's subtitles out of its Subs folder to
    sit beside the video. It used to leave the folder behind with whatever
    else the release put there, a checksum as a rule, and no later step
    removed it. The step now sweeps spent subtitle folders after the moves.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\SubtitleProcessing.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Subtitles.psm1') -Force
    $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'LibraryLint.ps1'), [ref]$null, [ref]$null)
    $fn = $scriptAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-SubtitleProcessing'
    }, $true)
    if (-not $fn) { throw 'Invoke-SubtitleProcessing not found in LibraryLint.ps1' }
    . ([scriptblock]::Create($fn.Extent.Text))

    function Write-Log { param($Message, $Level) }
    Mock Write-Host {} -ModuleName Subtitles

    function New-Release {
        param([string]$Root, [string[]]$SubsFiles)
        $movie = Join-Path $Root 'Movie (2009)'
        New-Item -ItemType Directory -Path (Join-Path $movie 'Subs') -Force | Out-Null
        Set-Content -Path (Join-Path $movie 'Movie (2009).mkv') -Value 'x'
        foreach ($name in $SubsFiles) { Set-Content -Path (Join-Path $movie "Subs\$name") -Value 'x' }
        return $movie
    }
}

Describe "Invoke-SubtitleProcessing" {
    BeforeEach {
        $script:lib = Join-Path $TestDrive "proc-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
        New-Item -ItemType Directory -Path $script:lib -Force | Out-Null
        $script:Config = @{
            VideoExtensions = @('.mkv', '.mp4'); SubtitleExtensions = @('.srt', '.sub', '.idx')
            PreferredSubtitleLanguages = @('eng', 'en', 'english'); KeepSubtitles = $true; DryRun = $false
        }
        $script:Stats = @{ SubtitlesProcessed = 0; SubtitlesDeleted = 0; FilesDeleted = 0; BytesDeleted = [long]0 }
        Mock Write-Host {}
    }

    It "moves the subtitle beside the video and removes the spent Subs folder with its checksum" {
        $movie = New-Release -Root $script:lib -SubsFiles 'movie.2009.eng.srt', 'movie.2009.sfv'
        Invoke-SubtitleProcessing -Path $script:lib
        Test-Path -LiteralPath (Join-Path $movie 'Movie (2009).en.srt') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $movie 'Subs') | Should -BeFalse
        $script:Stats.FilesDeleted | Should -Be 1
    }

    # The old code returned early when the release had no subtitle files at
    # all, which is exactly the state of a folder holding only a checksum.
    It "sweeps a Subs folder that never had a subtitle in it" {
        $movie = New-Release -Root $script:lib -SubsFiles 'movie.2009.sfv'
        Invoke-SubtitleProcessing -Path $script:lib
        Test-Path -LiteralPath (Join-Path $movie 'Subs') | Should -BeFalse
    }

    It "touches nothing in dry-run mode" {
        $script:Config.DryRun = $true
        $movie = New-Release -Root $script:lib -SubsFiles 'movie.2009.eng.srt', 'movie.2009.sfv'
        Invoke-SubtitleProcessing -Path $script:lib
        Test-Path -LiteralPath (Join-Path $movie 'Subs\movie.2009.eng.srt') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $movie 'Subs\movie.2009.sfv') | Should -BeTrue
        $script:Stats.FilesDeleted | Should -Be 0
    }

    It "keeps a Subs folder that still holds a subtitle it could not place" {
        $movie = New-Release -Root $script:lib -SubsFiles 'movie.2009.eng.srt', 'movie.2009.sfv'
        Set-Content -Path (Join-Path $movie 'Movie (2009).en.srt') -Value 'x'   # destination taken
        Invoke-SubtitleProcessing -Path $script:lib
        Test-Path -LiteralPath (Join-Path $movie 'Subs\movie.2009.eng.srt') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $movie 'Subs') | Should -BeTrue
    }
}
