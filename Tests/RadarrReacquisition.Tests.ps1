#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for Invoke-RadarrReacquisition, the shared engine that
    adds, re-monitors and searches movies in Radarr.

.DESCRIPTION
    All HTTP is mocked, so these run offline against no Radarr instance.
    This function writes to a live *arr install, so the behaviour worth
    pinning is what it sends: a movie Radarr already knows gets
    re-monitored and searched rather than duplicated, an unknown movie
    gets added with a search, and -WhatIf sends nothing at all.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\RadarrReacquisition.Tests.ps1 -Output Detailed
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
    foreach ($name in 'Invoke-RadarrReacquisition', 'Initialize-RadarrConnection') {
        $fn = $scriptAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
        }, $true)
        if (-not $fn) { throw "$name not found in LibraryLint.ps1" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    # Production reads these at call time through $script: scope.
    $script:Config = @{
        RadarrUrl              = 'http://radarr.test:7878'
        RadarrApiKey           = 'test-key'
        RadarrQualityProfileId = 4
        RadarrRootFolder       = '/movies'
    }
    function Write-Log { param($Message, $Level) }
    function Export-Configuration { }
}

Describe "Invoke-RadarrReacquisition guards" {
    It "does nothing for an empty movie list" {
        Mock Invoke-RestMethod { throw 'should not be called' }
        $stats = Invoke-RadarrReacquisition -Movies @() -SkipConfirm
        $stats.Added | Should -Be 0
        Should -Invoke Invoke-RestMethod -Times 0
    }

    It "skips entries with no title" {
        Mock Invoke-RestMethod { throw 'should not be called' }
        $stats = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Year = 2020 }) -SkipConfirm
        $stats.Added | Should -Be 0
        Should -Invoke Invoke-RestMethod -Times 0
    }

    It "reports cancelled when the connection cannot be established" {
        Mock Initialize-RadarrConnection { @{ Ok = $false; Error = 'nope' } }
        $stats = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'X'; Year = 2020 }) -SkipConfirm
        $stats.Cancelled | Should -BeTrue
    }
}

Describe "Invoke-RadarrReacquisition adding a movie Radarr has never seen" {
    BeforeEach {
        Mock Initialize-RadarrConnection { @{ Ok = $true; Url = 'http://radarr.test:7878'; Headers = @{ 'X-Api-Key' = 'test-key' } } }
        # Empty existing library.
        Mock Invoke-RestMethod { @() } -ParameterFilter { $Uri -like '*/api/v3/movie' -and $Method -ne 'Post' }
        Mock Invoke-RestMethod {
            @([PSCustomObject]@{ title = 'Deadpool'; year = 2016; tmdbId = 293660 })
        } -ParameterFilter { $Uri -like '*movie/lookup?term=*' }
        Mock Invoke-RestMethod { $null } -ParameterFilter { $Method -eq 'Post' }
    }

    It "posts an add payload and counts it as added" {
        $stats = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm
        $stats.Added | Should -Be 1
        $stats.Remonitored | Should -Be 0
        Should -Invoke Invoke-RestMethod -Times 1 -ParameterFilter { $Method -eq 'Post' -and $Uri -like '*/api/v3/movie' }
    }

    It "requests monitoring, the configured profile, and a search on add" {
        $script:capturedBody = $null
        Mock Invoke-RestMethod { $script:capturedBody = $Body; $null } -ParameterFilter { $Method -eq 'Post' -and $Uri -like '*/api/v3/movie' }
        $null = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm
        $payload = $script:capturedBody | ConvertFrom-Json
        $payload.monitored | Should -BeTrue
        $payload.qualityProfileId | Should -Be 4
        $payload.rootFolderPath | Should -Be '/movies'
        $payload.addOptions.searchForMovie | Should -BeTrue
        $payload.tmdbId | Should -Be 293660
    }

    It "adds without monitoring or searching under -Unmonitored" {
        $script:capturedBody = $null
        Mock Invoke-RestMethod { $script:capturedBody = $Body; $null } -ParameterFilter { $Method -eq 'Post' -and $Uri -like '*/api/v3/movie' }
        $null = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm -Unmonitored
        $payload = $script:capturedBody | ConvertFrom-Json
        $payload.monitored | Should -BeFalse
        $payload.addOptions.searchForMovie | Should -BeFalse
    }

    It "counts a movie Radarr cannot identify as not found" {
        Mock Invoke-RestMethod { @() } -ParameterFilter { $Uri -like '*movie/lookup?term=*' }
        $stats = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Nonexistent Film'; Year = 1999 }) -SkipConfirm
        $stats.NotFound | Should -Be 1
        $stats.Added | Should -Be 0
    }

    It "uses a supplied TmdbId instead of a title search" {
        Mock Invoke-RestMethod {
            [PSCustomObject]@{ title = 'Split'; year = 2017; tmdbId = 381288 }
        } -ParameterFilter { $Uri -like '*movie/lookup/tmdb?tmdbId=381288*' }
        $stats = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Split'; Year = 2016; TmdbId = 381288 }) -SkipConfirm
        $stats.Added | Should -Be 1
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Uri -like '*movie/lookup?term=*' }
    }
}

Describe "Invoke-RadarrReacquisition on a movie Radarr already tracks" {
    BeforeEach {
        Mock Initialize-RadarrConnection { @{ Ok = $true; Url = 'http://radarr.test:7878'; Headers = @{ 'X-Api-Key' = 'test-key' } } }
        Mock Invoke-RestMethod {
            @([PSCustomObject]@{ id = 77; tmdbId = 293660; title = 'Deadpool'; year = 2016; monitored = $false; qualityProfileId = 1; hasFile = $true })
        } -ParameterFilter { $Uri -like '*/api/v3/movie' -and $Method -ne 'Post' }
        Mock Invoke-RestMethod {
            @([PSCustomObject]@{ title = 'Deadpool'; year = 2016; tmdbId = 293660 })
        } -ParameterFilter { $Uri -like '*movie/lookup?term=*' }
        Mock Invoke-RestMethod { $null } -ParameterFilter { $Method -in @('Post', 'Put') }
    }

    # The whole point of re-acquisition: never create a second entry for a
    # movie already in Radarr, just re-monitor it and search for better.
    It "re-monitors rather than adding a duplicate" {
        $stats = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm
        $stats.Remonitored | Should -Be 1
        $stats.Added | Should -Be 0
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Method -eq 'Post' -and $Uri -like '*/api/v3/movie' }
    }

    It "flips monitored on and aligns the quality profile" {
        $script:putBody = $null
        Mock Invoke-RestMethod { $script:putBody = $Body; $null } -ParameterFilter { $Method -eq 'Put' }
        $null = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm
        $sent = $script:putBody | ConvertFrom-Json
        $sent.monitored | Should -BeTrue
        $sent.qualityProfileId | Should -Be 4
    }

    It "posts a MoviesSearch command for the existing movie id" {
        $script:cmdBody = $null
        Mock Invoke-RestMethod { $script:cmdBody = $Body; $null } -ParameterFilter { $Uri -like '*/api/v3/command' }
        $null = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm
        $cmd = $script:cmdBody | ConvertFrom-Json
        $cmd.name | Should -Be 'MoviesSearch'
        @($cmd.movieIds)[0] | Should -Be 77
    }

    It "does not search under -Unmonitored" {
        $null = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm -Unmonitored
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Uri -like '*/api/v3/command' }
    }
}

Describe "Invoke-RadarrReacquisition resilience" {
    BeforeEach {
        Mock Initialize-RadarrConnection { @{ Ok = $true; Url = 'http://radarr.test:7878'; Headers = @{ 'X-Api-Key' = 'test-key' } } }
        Mock Invoke-RestMethod { @() } -ParameterFilter { $Uri -like '*/api/v3/movie' -and $Method -ne 'Post' }
        Mock Invoke-RestMethod {
            @([PSCustomObject]@{ title = 'Deadpool'; year = 2016; tmdbId = 293660 })
        } -ParameterFilter { $Uri -like '*movie/lookup?term=*' }
    }

    # Radarr's own "already exists" rejection is a success for our purposes:
    # the movie is in Radarr, which is what the caller wanted.
    It "treats an already-exists rejection as re-monitored, not failed" {
        Mock Invoke-RestMethod { throw 'This movie has already been added' } -ParameterFilter { $Method -eq 'Post' }
        $stats = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm
        $stats.Remonitored | Should -Be 1
        $stats.Failed | Should -Be 0
    }

    It "counts a genuine error as failed without aborting the run" {
        Mock Invoke-RestMethod { throw 'Invalid quality profile' } -ParameterFilter { $Method -eq 'Post' }
        $stats = Invoke-RadarrReacquisition -Movies @(
            [PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }
            [PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }
        ) -SkipConfirm
        $stats.Failed | Should -Be 2
    }

    # Radarr's SQLite is single-writer, so a burst can collide with its
    # background tasks. One retry turns a transient lock into a success.
    It "retries once through a transient database lock" {
        $script:postAttempts = 0
        Mock Invoke-RestMethod {
            $script:postAttempts++
            if ($script:postAttempts -eq 1) { throw 'database is locked' }
            $null
        } -ParameterFilter { $Method -eq 'Post' }
        $stats = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm
        $script:postAttempts | Should -Be 2
        $stats.Added | Should -Be 1
        $stats.Failed | Should -Be 0
    }

    It "writes nothing under -WhatIf" {
        Mock Invoke-RestMethod { $null } -ParameterFilter { $Method -eq 'Post' }
        $stats = Invoke-RadarrReacquisition -Movies @([PSCustomObject]@{ Title = 'Deadpool'; Year = 2016 }) -SkipConfirm -WhatIf
        $stats.Cancelled | Should -BeTrue
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Method -eq 'Post' }
    }
}
