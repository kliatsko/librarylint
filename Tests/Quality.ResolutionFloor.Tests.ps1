#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for Get-ResolutionFloorReport (modules\Quality.psm1).

.DESCRIPTION
    The resolution floor is the Status pipeline's answer to "which movies
    are below my standard?". These tests cover the source-preference
    ladder (cached height, then release name, then a budgeted probe), the
    at-or-below boundary, and the probe cache write.

    No MediaInfo dependency: every case here supplies resolution through
    release-info.json, and the probe path is exercised only for its
    budget behaviour.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Quality.ResolutionFloor.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Quality.psm1') -Force

    function New-ResTestLibrary {
        param([string]$Name, [hashtable[]]$Movies)
        $root = Join-Path $TestDrive $Name
        New-Item -Path $root -ItemType Directory -Force | Out-Null
        foreach ($m in $Movies) {
            $dir = Join-Path $root $m.Folder
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $dir "$($m.Folder).mkv") -Value 'video' -NoNewline
            $info = [ordered]@{ OriginalFileName = "$($m.Folder).mkv" }
            if ($m.ContainsKey('Resolution'))     { $info['Resolution'] = $m.Resolution }
            if ($m.ContainsKey('DetectedResolution')) { $info['DetectedResolution'] = $m.DetectedResolution }
            if ($m.ContainsKey('NoInfoFile') -and $m.NoInfoFile) { continue }
            $info | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $dir 'release-info.json') -Encoding UTF8
        }
        return $root
    }
}

Describe "Get-ResolutionTier" {
    # Regression: comparing raw pixel HEIGHT against the floor reported
    # Inception (1280x528, 2.39:1) as "528p" and Life of Brian (1280x696)
    # as "696p". Both are 720p releases. Bucketing by the larger dimension
    # is what keeps a scope release at its mastered tier.
    It "tiers wide-aspect releases by their larger dimension" {
        Get-ResolutionTier -Width 1280 -Height 528 | Should -Be '720p'
        Get-ResolutionTier -Width 1280 -Height 696 | Should -Be '720p'
        Get-ResolutionTier -Width 1920 -Height 696 | Should -Be '1080p'
        Get-ResolutionTier -Width 3840 -Height 1600 | Should -Be '2160p'
    }

    It "tiers tall aspects by height" {
        Get-ResolutionTier -Width 1440 -Height 1080 | Should -Be '1080p'
        Get-ResolutionTier -Width 640 -Height 480 | Should -Be '480p'
    }

    It "tiers standard 16:9 dimensions" {
        Get-ResolutionTier -Width 3840 -Height 2160 | Should -Be '2160p'
        Get-ResolutionTier -Width 1920 -Height 1080 | Should -Be '1080p'
        Get-ResolutionTier -Width 1280 -Height 720 | Should -Be '720p'
        Get-ResolutionTier -Width 720 -Height 480 | Should -Be '480p'
    }

    It "falls back to raw height below every tier" {
        Get-ResolutionTier -Width 480 -Height 368 | Should -Be '368p'
    }

    It "returns null when no dimensions are usable" {
        Get-ResolutionTier -Width 0 -Height 0 | Should -BeNullOrEmpty
    }
}

Describe "Get-ResolutionFloorReport disabled state" {
    It "returns nothing when the floor is zero" {
        $root = New-ResTestLibrary -Name 'DisabledLib' -Movies @(@{ Folder = 'Low One (2001)'; Resolution = '480p' })
        $report = Get-ResolutionFloorReport -Path $root -FloorHeight 0 -VideoExtensions @('.mkv')
        @($report.Flagged).Count | Should -Be 0
        $report.Total | Should -Be 0
    }

    It "returns nothing for a library path that does not exist" {
        $report = Get-ResolutionFloorReport -Path (Join-Path $TestDrive 'nowhere') -FloorHeight 720 -VideoExtensions @('.mkv')
        @($report.Flagged).Count | Should -Be 0
    }
}

Describe "Get-ResolutionFloorReport boundary" {
    BeforeAll {
        $boundaryRoot = New-ResTestLibrary -Name 'BoundaryLib' -Movies @(
            @{ Folder = 'SD Movie (2001)';    Resolution = '480p' }
            @{ Folder = 'HD Movie (2002)';    Resolution = '720p' }
            @{ Folder = 'FullHD Movie (2003)'; Resolution = '1080p' }
            @{ Folder = 'UHD Movie (2004)';   Resolution = '2160p' }
        )
    }

    # The key semantic: a floor of 720 flags 720p itself, which is what
    # the config comment promises and why the parameter is not called a
    # minimum. Getting this backwards would either miss the files the
    # user cares about or flag their whole library.
    It "flags the floor value itself and everything under it" {
        $report = Get-ResolutionFloorReport -Path $boundaryRoot -FloorHeight 720 -VideoExtensions @('.mkv')
        $flagged = @($report.Flagged | ForEach-Object { $_.Folder })
        $flagged | Should -Contain 'SD Movie (2001)'
        $flagged | Should -Contain 'HD Movie (2002)'
        $flagged | Should -Not -Contain 'FullHD Movie (2003)'
        $flagged | Should -Not -Contain 'UHD Movie (2004)'
    }

    It "flags only sub-HD at a floor of 480" {
        $report = Get-ResolutionFloorReport -Path $boundaryRoot -FloorHeight 480 -VideoExtensions @('.mkv')
        @($report.Flagged).Count | Should -Be 1
        $report.Flagged[0].Folder | Should -Be 'SD Movie (2001)'
    }

    It "flags everything below 4K at a floor of 1080" {
        $report = Get-ResolutionFloorReport -Path $boundaryRoot -FloorHeight 1080 -VideoExtensions @('.mkv')
        @($report.Flagged).Count | Should -Be 3
    }

    It "counts every movie it could resolve" {
        $report = Get-ResolutionFloorReport -Path $boundaryRoot -FloorHeight 720 -VideoExtensions @('.mkv')
        $report.Total | Should -Be 4
        $report.Unknown | Should -Be 0
    }

    It "orders the flagged list worst first" {
        $report = Get-ResolutionFloorReport -Path $boundaryRoot -FloorHeight 1080 -VideoExtensions @('.mkv')
        $report.Flagged[0].Height | Should -Be 480
        $report.Flagged[-1].Height | Should -Be 1080
    }
}

Describe "Get-ResolutionFloorReport source preference" {
    It "prefers a cached DetectedResolution over the release name" {
        # Release name claims 1080p, the verified probe says 720. A
        # mislabelled release is exactly the case worth catching, so the
        # measured value has to win.
        $root = New-ResTestLibrary -Name 'PreferVerified' -Movies @(
            @{ Folder = 'Mislabelled (2005)'; Resolution = '1080p'; DetectedResolution = '720p' }
        )
        $report = Get-ResolutionFloorReport -Path $root -FloorHeight 720 -VideoExtensions @('.mkv')
        @($report.Flagged).Count | Should -Be 1
        $report.Flagged[0].Height | Should -Be 720
        $report.Flagged[0].Source | Should -Be 'verified'
    }

    It "falls back to the release name and labels the source" {
        $root = New-ResTestLibrary -Name 'ReleaseNameOnly' -Movies @(
            @{ Folder = 'From Name (2006)'; Resolution = '720p' }
        )
        $report = Get-ResolutionFloorReport -Path $root -FloorHeight 720 -VideoExtensions @('.mkv')
        $report.Flagged[0].Source | Should -Be 'release name'
    }

    It "ignores an unparseable release-name resolution" {
        $root = New-ResTestLibrary -Name 'BadResString' -Movies @(
            @{ Folder = 'Odd Res (2007)'; Resolution = 'SuperHD' }
        )
        # No probe budget, so nothing can resolve it: counted as unknown.
        $report = Get-ResolutionFloorReport -Path $root -FloorHeight 720 -VideoExtensions @('.mkv') -ProbeLimit 0
        @($report.Flagged).Count | Should -Be 0
        $report.Unknown | Should -Be 1
    }

    It "counts folders it cannot resolve as unknown rather than flagging them" {
        $root = New-ResTestLibrary -Name 'NoData' -Movies @(
            @{ Folder = 'No Info (2008)'; NoInfoFile = $true }
        )
        $report = Get-ResolutionFloorReport -Path $root -FloorHeight 720 -VideoExtensions @('.mkv') -ProbeLimit 0
        @($report.Flagged).Count | Should -Be 0
        $report.Unknown | Should -Be 1
        $report.Probed | Should -Be 0
    }

    It "respects the probe budget" {
        $root = New-ResTestLibrary -Name 'ProbeBudget' -Movies @(
            @{ Folder = 'Needs Probe A (2009)'; NoInfoFile = $true }
            @{ Folder = 'Needs Probe B (2010)'; NoInfoFile = $true }
            @{ Folder = 'Needs Probe C (2011)'; NoInfoFile = $true }
        )
        $report = Get-ResolutionFloorReport -Path $root -FloorHeight 720 -VideoExtensions @('.mkv') -ProbeLimit 2
        $report.Probed | Should -BeLessOrEqual 2
    }

    It "skips working folders prefixed with an underscore" {
        $root = New-ResTestLibrary -Name 'SkipsWorking' -Movies @(
            @{ Folder = '_Trailers'; Resolution = '480p' }
            @{ Folder = 'Real Movie (2012)'; Resolution = '480p' }
        )
        $report = Get-ResolutionFloorReport -Path $root -FloorHeight 720 -VideoExtensions @('.mkv')
        @($report.Flagged).Count | Should -Be 1
        $report.Flagged[0].Folder | Should -Be 'Real Movie (2012)'
    }
}
