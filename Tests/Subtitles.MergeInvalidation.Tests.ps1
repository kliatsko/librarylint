#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for subtitle invalidation during a folder merge
    (Rename-OrMergeFolder in LibraryLint.ps1).

.DESCRIPTION
    When two library folders for one film collapse and the incoming copy
    scores higher, the target's video is deleted and the source's moved in
    while every other file in the target survives. Anything describing the
    deleted video has to go with it: its subtitles were cut for a file that
    will not exist, and its .subs_ok claims those subtitles are verified.

    The load-bearing test here is the NEGATIVE one: when the target's video
    WINS, its subtitles and marker must survive untouched. Getting that
    backwards would make this change destructive.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Subtitles.MergeInvalidation.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # Same AST-extraction idiom as Tests\ArrStatus.Tests.ps1: pull the real
    # function out of the main script so the test can never drift from it.
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $scriptPath = Join-Path $repoRoot 'LibraryLint.ps1'
    $parseTokens = $null
    $parseErrors = $null
    $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$parseTokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "LibraryLint.ps1 has $($parseErrors.Count) parse error(s); first: $($parseErrors[0].Message)"
    }
    $fn = $scriptAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Rename-OrMergeFolder'
    }, $true)
    if (-not $fn) { throw "Rename-OrMergeFolder not found in LibraryLint.ps1" }
    . ([scriptblock]::Create($fn.Extent.Text))

    Import-Module (Join-Path $repoRoot 'modules\Subtitles.psm1') -Force

    $script:Config = @{
        VideoExtensions    = @('.mkv', '.mp4', '.avi', '.m4v')
        SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt')
    }
    # Collaborators the function calls that are not under test here.
    function Add-UndoOperation { param($OperationType, $SourcePath, $DestinationPath) }
    function Write-Log { param($Message, $Level) }

    # Quality is decided by these stubs so the fixtures don't need real media.
    # Score comes from the filename tag, mirroring how the production scorer
    # ranks 1080p above 720p.
    function Invoke-QualityScore {
        param([string]$FileName, [string]$FilePath)
        $score = if ($FileName -match '1080p') { 80 } elseif ($FileName -match '720p') { 60 } else { 40 }
        return [PSCustomObject]@{ Score = $score; Resolution = 'x' }
    }

    function New-MergeFixture {
        param(
            [string]$Name,
            [string]$TargetVideo,
            [string]$SourceVideo,
            [string[]]$TargetExtras = @(),
            [string[]]$SourceExtras = @(),
            [switch]$TargetVerified
        )
        $root = Join-Path $TestDrive $Name
        $target = Join-Path $root 'Movie (2020)'
        $source = Join-Path $root 'Movie 2020 RELEASE-TAGS'
        New-Item -Path $target, $source -ItemType Directory -Force | Out-Null

        Set-Content -LiteralPath (Join-Path $target $TargetVideo) -Value 'target-video' -NoNewline
        Set-Content -LiteralPath (Join-Path $source $SourceVideo) -Value 'source-video' -NoNewline
        foreach ($e in $TargetExtras) { Set-Content -LiteralPath (Join-Path $target $e) -Value 'x' -NoNewline }
        foreach ($e in $SourceExtras) { Set-Content -LiteralPath (Join-Path $source $e) -Value 'x' -NoNewline }
        if ($TargetVerified) {
            $null = Set-SubtitlesVerified -FolderPath $target -Source 'opensubtitles-hashmatch (Old.Release)' -Provider 'opensubtitles'
        }
        return [PSCustomObject]@{ Root = $root; Target = $target; Source = $source }
    }
}

Describe "Rename-OrMergeFolder when the incoming video wins" {
    BeforeEach {
        # Target holds a 720p video with a matching subtitle and a marker;
        # source holds a 1080p video, so the source replaces the video.
        $script:fx = New-MergeFixture -Name "SourceWins_$(Get-Random)" `
            -TargetVideo 'Movie (2020) 720p.mkv' `
            -SourceVideo 'Movie 2020 1080p.mkv' `
            -TargetExtras @('Movie (2020) 720p.en.srt', 'poster.jpg') `
            -TargetVerified
        $null = Rename-OrMergeFolder -SourceFolder $script:fx.Source -NewName 'Movie (2020)'
    }

    It "brings the better video across" {
        Test-Path -LiteralPath (Join-Path $script:fx.Target 'Movie 2020 1080p.mkv') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:fx.Target 'Movie (2020) 720p.mkv') | Should -BeFalse
    }

    # The claim "these subtitles are verified" was earned by the video that
    # was just deleted. Left in place it vouches for a different release.
    It "clears the target's subtitle verification" {
        Test-Path -LiteralPath (Join-Path $script:fx.Target '.subs_ok') | Should -BeFalse
    }

    # Left behind, Repair-OrphanedSubtitles would later rename this onto the
    # incoming video, hiding the fact that it was cut for another release.
    It "removes the subtitle that belonged to the replaced video" {
        Test-Path -LiteralPath (Join-Path $script:fx.Target 'Movie (2020) 720p.en.srt') | Should -BeFalse
    }

    It "leaves unrelated target files alone" {
        Test-Path -LiteralPath (Join-Path $script:fx.Target 'poster.jpg') | Should -BeTrue
    }
}

Describe "Rename-OrMergeFolder when the target video wins" {
    BeforeEach {
        # Reversed: the target already holds the better copy, so nothing about
        # it changes and its verification must stand.
        $script:fx = New-MergeFixture -Name "TargetWins_$(Get-Random)" `
            -TargetVideo 'Movie (2020) 1080p.mkv' `
            -SourceVideo 'Movie 2020 720p.mkv' `
            -TargetExtras @('Movie (2020) 1080p.en.srt') `
            -TargetVerified
        $null = Rename-OrMergeFolder -SourceFolder $script:fx.Source -NewName 'Movie (2020)'
    }

    It "keeps the target video" {
        Test-Path -LiteralPath (Join-Path $script:fx.Target 'Movie (2020) 1080p.mkv') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:fx.Target 'Movie 2020 720p.mkv') | Should -BeFalse
    }

    # This is the destructive-regression guard. The video did not change, so
    # the subtitle and the verification are both still accurate.
    It "keeps the target's subtitle verification" {
        Test-Path -LiteralPath (Join-Path $script:fx.Target '.subs_ok') | Should -BeTrue
    }

    It "keeps the target's subtitle" {
        Test-Path -LiteralPath (Join-Path $script:fx.Target 'Movie (2020) 1080p.en.srt') | Should -BeTrue
    }
}

Describe "Rename-OrMergeFolder subtitle scoping" {
    # Only subtitles named for a DELETED video go. One named for a video that
    # survives is still valid and must be left alone.
    It "spares a subtitle belonging to a video that is kept" {
        $root = Join-Path $TestDrive "Scoped_$(Get-Random)"
        $target = Join-Path $root 'Movie (2020)'
        $source = Join-Path $root 'Movie 2020 TAGS'
        New-Item -Path $target, $source -ItemType Directory -Force | Out-Null
        # Two videos in the target: the 720p one loses, the extras clip is not
        # a video by extension so it is untouched either way.
        Set-Content -LiteralPath (Join-Path $target 'Movie (2020) 720p.mkv') -Value 'v' -NoNewline
        Set-Content -LiteralPath (Join-Path $target 'Movie (2020) 720p.en.srt') -Value 's' -NoNewline
        Set-Content -LiteralPath (Join-Path $target 'Bonus Feature.txt') -Value 't' -NoNewline
        Set-Content -LiteralPath (Join-Path $source 'Movie 2020 1080p.mkv') -Value 'v' -NoNewline

        $null = Rename-OrMergeFolder -SourceFolder $source -NewName 'Movie (2020)'

        Test-Path -LiteralPath (Join-Path $target 'Movie (2020) 720p.en.srt') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $target 'Bonus Feature.txt') | Should -BeTrue
    }

    It "brings the source's own subtitle across" {
        $fx = New-MergeFixture -Name "SourceSubs_$(Get-Random)" `
            -TargetVideo 'Movie (2020) 720p.mkv' `
            -SourceVideo 'Movie 2020 1080p.mkv' `
            -SourceExtras @('Movie 2020 1080p.en.srt')
        $null = Rename-OrMergeFolder -SourceFolder $fx.Source -NewName 'Movie (2020)'
        Test-Path -LiteralPath (Join-Path $fx.Target 'Movie 2020 1080p.en.srt') | Should -BeTrue
    }

    It "does nothing to verification when the target has no video to replace" {
        $root = Join-Path $TestDrive "NoTargetVideo_$(Get-Random)"
        $target = Join-Path $root 'Movie (2020)'
        $source = Join-Path $root 'Movie 2020 TAGS'
        New-Item -Path $target, $source -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $target 'poster.jpg') -Value 'x' -NoNewline
        Set-Content -LiteralPath (Join-Path $source 'Movie 2020 1080p.mkv') -Value 'v' -NoNewline
        $null = Set-SubtitlesVerified -FolderPath $target -Source 'manual' -Provider 'manual'

        $null = Rename-OrMergeFolder -SourceFolder $source -NewName 'Movie (2020)'

        # No target video was deleted, so nothing was invalidated.
        Test-Path -LiteralPath (Join-Path $target '.subs_ok') | Should -BeTrue
    }
}

Describe "Rename-OrMergeFolder simple rename path" {
    It "renames without touching verification when no target exists" {
        $root = Join-Path $TestDrive "SimpleRename_$(Get-Random)"
        $source = Join-Path $root 'Movie 2020 TAGS'
        New-Item -Path $source -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $source 'Movie 2020 1080p.mkv') -Value 'v' -NoNewline
        $null = Set-SubtitlesVerified -FolderPath $source -Source 'manual' -Provider 'manual'

        $result = Rename-OrMergeFolder -SourceFolder $source -NewName 'Movie (2020)'

        $result | Should -Be (Join-Path $root 'Movie (2020)')
        Test-Path -LiteralPath (Join-Path $result '.subs_ok') | Should -BeTrue
    }
}
