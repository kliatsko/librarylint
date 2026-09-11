#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for spent subtitle folders (modules\Subtitles.psm1):
    Test-SpentSubtitleFolder, Get-SpentSubtitleFolders,
    Remove-SpentSubtitleFolders, and the placement repair's sweep.

.DESCRIPTION
    A release's Subs folder is spent once its subtitles have been moved up
    beside the video. Seven of them sat in a library holding one checksum
    or readme each: subtitle processing moved subtitles only, Clean
    Unnecessary Files matched names, Remove Empty Folders wanted nothing
    inside. These tests pin what "spent" means and that the removal
    re-checks the folder at the moment of deletion.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Subtitles.SpentFolders.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Subtitles.psm1') -Force

    function New-Lib {
        $lib = Join-Path $TestDrive "spent-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
        New-Item -ItemType Directory -Path $lib -Force | Out-Null
        return $lib
    }
    function Add-File {
        param([string]$Root, [string]$RelativePath)
        $full = Join-Path $Root $RelativePath
        New-Item -ItemType Directory -Path (Split-Path $full -Parent) -Force | Out-Null
        Set-Content -Path $full -Value 'x'
        return $full
    }
}

Describe "Test-SpentSubtitleFolder" {
    BeforeEach { $script:lib = New-Lib }

    It "calls a Subs folder holding only a checksum spent" {
        $null = Add-File $script:lib 'Movie (2009)\Subs\release.sfv'
        $state = Test-SpentSubtitleFolder -FolderPath (Join-Path $script:lib 'Movie (2009)\Subs')
        $state.Spent | Should -BeTrue
        @($state.Files).Count | Should -Be 1
    }

    It "calls an empty Subs folder spent" {
        New-Item -ItemType Directory -Path (Join-Path $script:lib 'Movie (2009)\Subs') -Force | Out-Null
        (Test-SpentSubtitleFolder -FolderPath (Join-Path $script:lib 'Movie (2009)\Subs')).Spent | Should -BeTrue
    }

    It "keeps a Subs folder that still holds a subtitle, even nested" {
        $null = Add-File $script:lib 'Movie (2009)\Subs\release.sfv'
        $null = Add-File $script:lib 'Movie (2009)\Subs\more\extra.srt'
        (Test-SpentSubtitleFolder -FolderPath (Join-Path $script:lib 'Movie (2009)\Subs')).Spent | Should -BeFalse
    }

    It "keeps a Subs folder that holds a video" {
        $null = Add-File $script:lib 'Movie (2009)\Subs\sample.mkv'
        (Test-SpentSubtitleFolder -FolderPath (Join-Path $script:lib 'Movie (2009)\Subs')).Spent | Should -BeFalse
    }

    # Only folders named for subtitles qualify; a folder of anything else
    # holding a text file is not this rule's business.
    It "ignores a folder that is not a subtitle folder by name" {
        $null = Add-File $script:lib 'Movie (2009)\Extras\notes.txt'
        (Test-SpentSubtitleFolder -FolderPath (Join-Path $script:lib 'Movie (2009)\Extras')).Spent | Should -BeFalse
    }
}

Describe "Get-SpentSubtitleFolders" {
    It "lists the spent ones with their parent and litter, and skips the rest" {
        $lib = New-Lib
        $null = Add-File $lib 'A (2001)\Subs\a.sfv'
        $null = Add-File $lib 'B (2002)\Subs\b.srt'
        $null = Add-File $lib 'C (2003)\Subtitles\readme.txt'
        $null = Add-File $lib 'D (2004)\Extras\notes.txt'
        $found = @(Get-SpentSubtitleFolders -Path $lib)
        @($found | ForEach-Object { $_.Parent }) | Sort-Object | Should -Be @('A (2001)', 'C (2003)')
        ($found | Where-Object { $_.Parent -eq 'A (2001)' }).Files | Should -Be @('a.sfv')
    }
}

Describe "Remove-SpentSubtitleFolders" {
    BeforeAll { Mock Write-Host {} -ModuleName Subtitles }
    BeforeEach {
        $script:lib = New-Lib
        $script:subs = Join-Path $script:lib 'Movie (2009)\Subs'
        $null = Add-File $script:lib 'Movie (2009)\Subs\release.sfv'
    }

    It "removes nothing under -WhatIf but reports the plan" {
        $stats = Remove-SpentSubtitleFolders -Folders @($script:subs) -WhatIf
        $stats.Removed | Should -Be 1
        $stats.FilesRemoved | Should -Be 1
        Test-Path -LiteralPath $script:subs | Should -BeTrue
    }

    It "removes the folder with its litter and counts it" {
        $stats = Remove-SpentSubtitleFolders -Folders @($script:subs)
        $stats.Removed | Should -Be 1
        $stats.FilesRemoved | Should -Be 1
        $stats.BytesFreed | Should -BeGreaterThan 0
        Test-Path -LiteralPath $script:subs | Should -BeFalse
    }

    # The list may be minutes old; a subtitle that landed since makes the
    # folder worth keeping.
    It "re-checks at deletion time and skips a folder that gained a subtitle" {
        $null = Add-File $script:lib 'Movie (2009)\Subs\late.srt'
        $stats = Remove-SpentSubtitleFolders -Folders @($script:subs)
        $stats.Removed | Should -Be 0
        $stats.Skipped | Should -Be 1
        Test-Path -LiteralPath $script:subs | Should -BeTrue
    }

    It "passes silently over a folder that is already gone" {
        Remove-Item -LiteralPath $script:subs -Recurse -Force
        $stats = Remove-SpentSubtitleFolders -Folders @($script:subs)
        $stats.Removed | Should -Be 0
        $stats.Errors | Should -Be 0
    }
}

Describe "Repair-SubtitlePlacement sweeps the folder it empties" {
    BeforeAll { Mock Write-Host {} -ModuleName Subtitles }
    BeforeEach {
        $script:lib = New-Lib
        $null = Add-File $script:lib 'Movie (2009)\Movie (2009).mkv'
        $null = Add-File $script:lib 'Movie (2009)\Subs\movie.2009.eng.srt'
        $null = Add-File $script:lib 'Movie (2009)\Subs\movie.2009.sfv'
    }

    It "plans the removal in a dry run and touches nothing" {
        $stats = Repair-SubtitlePlacement -Path $script:lib -WhatIf
        $stats.SubsMoved | Should -Be 1
        $stats.FoldersRemoved | Should -Be 1
        Test-Path -LiteralPath (Join-Path $script:lib 'Movie (2009)\Subs\movie.2009.eng.srt') | Should -BeTrue
    }

    It "moves the subtitle up and removes the folder with the checksum" {
        $stats = Repair-SubtitlePlacement -Path $script:lib
        $stats.SubsMoved | Should -Be 1
        $stats.FoldersRemoved | Should -Be 1
        Test-Path -LiteralPath (Join-Path $script:lib 'Movie (2009)\Movie (2009).en.srt') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:lib 'Movie (2009)\Subs') | Should -BeFalse
    }

    # A subtitle the pass chose to leave (its destination exists already)
    # keeps the folder, litter and all.
    It "keeps the folder when a subtitle stays behind" {
        $null = Add-File $script:lib 'Movie (2009)\Movie (2009).en.srt'
        $stats = Repair-SubtitlePlacement -Path $script:lib
        $stats.FoldersRemoved | Should -Be 0
        Test-Path -LiteralPath (Join-Path $script:lib 'Movie (2009)\Subs\movie.2009.eng.srt') | Should -BeTrue
    }
}
