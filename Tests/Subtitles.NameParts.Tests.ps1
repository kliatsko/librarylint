#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for Get-SubtitleNameParts and the orphan matching in
    Repair-OrphanedSubtitles (modules\Subtitles.psm1).

.DESCRIPTION
    The health check and the orphan repair used to decide "does this
    subtitle belong to that video" with two different language lists. A
    Danish "500 Days of Summer (2009).da.idx" was orphaned by one and fine
    by the other, so the health check flagged it every run and the repair
    it offered did nothing. Both now go through Get-SubtitleNameParts;
    these tests pin what it strips and what the repair does with it.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Subtitles.NameParts.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Subtitles.psm1') -Force
}

Describe "Get-SubtitleNameParts" {
    It "splits '<Name>' into '<Base>' plus '<Suffix>'" -ForEach @(
        @{ Name = '500 Days of Summer (2009).da.idx'; Base = '500 Days of Summer (2009)'; Suffix = '.da' }
        @{ Name = 'Movie (2009).en.forced.srt';       Base = 'Movie (2009)';              Suffix = '.en.forced' }
        @{ Name = 'Movie (2009).forced.en.srt';       Base = 'Movie (2009)';              Suffix = '.forced.en' }
        @{ Name = 'Movie (2009).eng.sdh.srt';         Base = 'Movie (2009)';              Suffix = '.eng.sdh' }
        @{ Name = 'Movie (2009).english.srt';         Base = 'Movie (2009)';              Suffix = '.english' }
        @{ Name = 'Movie (2009).srt';                 Base = 'Movie (2009)';              Suffix = '' }
        @{ Name = 'Movie.Name.srt';                   Base = 'Movie.Name';                Suffix = '' }
        @{ Name = 'Show - S01E01.srt';                Base = 'Show - S01E01';             Suffix = '' }
    ) {
        $parts = Get-SubtitleNameParts -FileName $Name
        $parts.BaseName | Should -Be $Base
        $parts.Suffix | Should -Be $Suffix
    }

    It "reports the language it peeled" {
        (Get-SubtitleNameParts -FileName 'Movie (2009).da.idx').Language | Should -Be 'da'
        (Get-SubtitleNameParts -FileName 'Movie (2009).srt').Language | Should -Be ''
    }
}

Describe "Repair-OrphanedSubtitles matching" {
    BeforeAll {
        Mock Write-Host {} -ModuleName Subtitles
    }

    BeforeEach {
        $script:lib = Join-Path $TestDrive "subs-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
        $script:movie = Join-Path $script:lib 'Movie (2009)'
        New-Item -ItemType Directory -Path $script:movie -Force | Out-Null
        Set-Content -Path (Join-Path $script:movie 'Movie (2009).mkv') -Value 'x'
    }

    It "does not treat a Danish VobSub index named for the video as an orphan" {
        Set-Content -Path (Join-Path $script:movie 'Movie (2009).da.idx') -Value 'x'
        (Repair-OrphanedSubtitles -Path $script:lib -WhatIf).Total | Should -Be 0
    }

    It "keeps a forced English track that carries two suffixes" {
        Set-Content -Path (Join-Path $script:movie 'Movie (2009).en.forced.srt') -Value 'x'
        (Repair-OrphanedSubtitles -Path $script:lib -WhatIf).Total | Should -Be 0
    }

    It "renames a stray subtitle onto the video and keeps its language" {
        Set-Content -Path (Join-Path $script:movie 'release.name.da.srt') -Value 'x'
        $stats = Repair-OrphanedSubtitles -Path $script:lib
        $stats.Renamed | Should -Be 1
        Test-Path -LiteralPath (Join-Path $script:movie 'Movie (2009).da.srt') | Should -BeTrue
    }

    # Two capital letters at the end of a release name look like a language
    # code to the name splitter; equality with the video settles it first.
    It "never touches a subtitle named exactly like the video, whatever its last segment looks like" {
        Remove-Item -LiteralPath (Join-Path $script:movie 'Movie (2009).mkv')
        Set-Content -Path (Join-Path $script:movie 'Movie.2009.YTS.AG.mkv') -Value 'x'
        Set-Content -Path (Join-Path $script:movie 'Movie.2009.YTS.AG.srt') -Value 'x'
        (Repair-OrphanedSubtitles -Path $script:lib -WhatIf).Total | Should -Be 0
    }
}
