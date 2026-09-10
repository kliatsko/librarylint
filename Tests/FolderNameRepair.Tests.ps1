#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for two Library Tools repairs: Fix Folder Names catching a
    well-formed but wrong year, and Fix File Names leaving subtitle
    qualifiers alone.

.DESCRIPTION
    "The Town (2009)" sat in the library for months with an NFO, a TMDB id
    and a release-info original filename that all said 2010. Fix Folder
    Names never revisited it because its skip looked only at the shape of
    the name. Separately, Fix File Names' double-name cleanup knew one
    2–3-letter language code and treated everything else between basename
    and extension as duplicated-title junk, so "Title (Year).en.forced.srt"
    became "Title (Year).srt". Both functions are extracted from the main
    script by AST; TMDB and NFO helpers are mocked, so this runs offline.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\FolderNameRepair.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $scriptPath = Join-Path $repoRoot 'LibraryLint.ps1'
    $parseTokens = $null
    $parseErrors = $null
    $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$parseTokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "LibraryLint.ps1 has $($parseErrors.Count) parse error(s); first: $($parseErrors[0].Message)"
    }
    foreach ($name in 'Repair-MovieFolderYears', 'Rename-VideoToMatchFolder', 'Get-SafeFileName') {
        $fn = $scriptAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
        }, $true)
        if (-not $fn) { throw "$name not found in LibraryLint.ps1" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    $script:Config = @{
        TMDBApiKey      = 'test-key'
        VideoExtensions = @('.mkv', '.mp4', '.avi', '.m4v')
        DryRun          = $false
    }
    # Helpers the two functions call; stubbed here so they can be mocked.
    function Write-Log { param($Message, $Level) }
    function Test-FolderHasValidNFO { param($FolderPath) @{ Valid = $true; NfoPath = $null; Metadata = @{ Title = 'x' }; Reason = 'OK' } }
    function Read-NFOFile { param($NfoPath) @{ Title = $null; Year = $null } }
    function Get-NormalizedTitle { param($Name, [switch]$Strict) @{ NormalizedTitle = ($Name -replace '\s*\(\d{4}\)\s*$', ''); Year = $null } }
    function Search-TMDBMovie { param($Title, $Year, $ApiKey) $null }
    function Read-ReleaseInfo { param($FolderPath) $null }
    function Save-ReleaseInfo { param($FolderPath, $OriginalFileName) }

    # A movie folder with a small video and an NFO whose parsed content the
    # tests decide through the Read-NFOFile mock.
    function New-MovieFixture {
        param([string]$Root, [string]$FolderName)
        $dir = Join-Path $Root $FolderName
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $dir "$FolderName.mkv"), (New-Object byte[] 2048))
        Set-Content -Path (Join-Path $dir "$FolderName.nfo") -Value '<movie/>'
        return $dir
    }
}

Describe "Repair-MovieFolderYears with a well-formed but wrong year" {
    BeforeEach {
        $script:lib = Join-Path $TestDrive "lib-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
        New-Item -ItemType Directory -Path $script:lib -Force | Out-Null
        $script:hostLines = [System.Collections.Generic.List[string]]::new()
        Mock Write-Host { $script:hostLines.Add(($Object -join ' ')) }
        # NFO says what the folder should have said; TMDB agrees with the NFO.
        Mock Read-NFOFile {
            if ($NfoPath -like '*The Town*') { @{ Title = 'The Town'; Year = '2010' } }
            elseif ($NfoPath -like '*Candy*') { @{ Title = 'Candy'; Year = '2006' } }
            else { @{ Title = $null; Year = $null } }
        }
        Mock Search-TMDBMovie { @{ Title = $Title; Year = $(if ($Year) { $Year } else { '2010' }); TMDBID = 23168 } }
    }

    It "proposes renaming The Town (2009) to (2010) when the NFO says 2010 and TMDB agrees" {
        $null = New-MovieFixture -Root $script:lib -FolderName 'The Town (2009)'
        Repair-MovieFolderYears -Path $script:lib -WhatIf
        $text = $script:hostLines -join "`n"
        $text | Should -Match 'The Town \(2010\)'
        $text | Should -Match 'year \(2009 -> 2010\)'
        $text | Should -Match 'NFO \+ TMDB \(confirmed\)'
    }

    It "still leaves a folder alone when its year matches the NFO" {
        $null = New-MovieFixture -Root $script:lib -FolderName 'Candy (2006)'
        Repair-MovieFolderYears -Path $script:lib -WhatIf
        ($script:hostLines -join "`n") | Should -Match 'No folders need fixing'
    }

    # The NFO can be the wrong one. When TMDB sides with the folder, the
    # proposed name equals the current name and is filtered out as a no-op.
    It "does not rename when TMDB sides with the folder's year over the NFO" {
        $null = New-MovieFixture -Root $script:lib -FolderName 'The Town (2009)'
        Mock Search-TMDBMovie { @{ Title = 'The Town'; Year = '2009'; TMDBID = 1 } }
        Repair-MovieFolderYears -Path $script:lib -WhatIf
        ($script:hostLines -join "`n") | Should -Match 'No folders need fixing'
    }

    It "renames nothing under -WhatIf" {
        $dir = New-MovieFixture -Root $script:lib -FolderName 'The Town (2009)'
        Repair-MovieFolderYears -Path $script:lib -WhatIf
        Test-Path -LiteralPath $dir | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:lib 'The Town (2010)') | Should -BeFalse
    }
}

Describe "Rename-VideoToMatchFolder double-name cleanup and subtitle qualifiers" {
    BeforeEach {
        $script:lib = Join-Path $TestDrive "files-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
        New-Item -ItemType Directory -Path $script:lib -Force | Out-Null
        Mock Write-Host { }
        Mock Test-FolderHasValidNFO {
            $title = (Split-Path $FolderPath -Leaf) -replace '\s*\(\d{4}\)\s*$', ''
            @{ Valid = $true; NfoPath = $null; Metadata = @{ Title = $title }; Reason = 'OK' }
        }
    }

    # The renames from the screenshot: full subtitle tracks that lost their
    # ".en" (and ".forced") because the parser did not know "forced".
    It "leaves subtitle files with language and flag qualifiers untouched" {
        $dir = New-MovieFixture -Root $script:lib -FolderName 'Movie (2024)'
        foreach ($sub in 'Movie (2024).en.forced.srt', 'Movie (2024).en.srt', 'Movie (2024).pt-BR.sdh.srt', 'Movie (2024).eng.hi.srt', 'Movie (2024).en.cc.vtt') {
            Set-Content -Path (Join-Path $dir $sub) -Value '1'
        }
        $null = Rename-VideoToMatchFolder -Path $script:lib
        foreach ($sub in 'Movie (2024).en.forced.srt', 'Movie (2024).en.srt', 'Movie (2024).pt-BR.sdh.srt', 'Movie (2024).eng.hi.srt', 'Movie (2024).en.cc.vtt') {
            Test-Path -LiteralPath (Join-Path $dir $sub) | Should -BeTrue -Because "$sub must survive"
        }
        Test-Path -LiteralPath (Join-Path $dir 'Movie (2024).srt') | Should -BeFalse
    }

    It "still collapses a doubled basename, keeping the qualifiers it carried" {
        $dir = New-MovieFixture -Root $script:lib -FolderName 'Other (2020)'
        Set-Content -Path (Join-Path $dir 'Other (2020) Other (2020).en.forced.srt') -Value '1'
        Set-Content -Path (Join-Path $dir 'Other (2020) Other (2020).nfo') -Value '<movie/>'
        $null = Rename-VideoToMatchFolder -Path $script:lib
        Test-Path -LiteralPath (Join-Path $dir 'Other (2020).en.forced.srt')             | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $dir 'Other (2020) Other (2020).en.forced.srt') | Should -BeFalse
        # The doubled NFO duplicates one that already exists, so it is removed.
        Test-Path -LiteralPath (Join-Path $dir 'Other (2020) Other (2020).nfo') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $dir 'Other (2020).nfo')              | Should -BeTrue
    }

    # An unknown tag with no "(Year)" in it is not duplicated-title junk.
    It "does not guess at a tag it does not recognise" {
        $dir = New-MovieFixture -Root $script:lib -FolderName 'Plain (2019)'
        Set-Content -Path (Join-Path $dir 'Plain (2019).commentary.srt') -Value '1'
        $null = Rename-VideoToMatchFolder -Path $script:lib
        Test-Path -LiteralPath (Join-Path $dir 'Plain (2019).commentary.srt') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $dir 'Plain (2019).srt')            | Should -BeFalse
    }
}
