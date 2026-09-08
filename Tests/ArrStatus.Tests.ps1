#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for Get-ArrStatusSummary, the Radarr/Sonarr dashboard feed.

.DESCRIPTION
    All HTTP is mocked, so these run offline with no *arr instance. The
    behaviour worth locking down is defensive: each endpoint is isolated so
    a quirk in one cannot discard data already collected from another, and
    the renderer must not be handed empty rows.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\ArrStatus.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # Same idiom as LibraryLint.Tests.ps1: extract the real function from
    # the main script by AST so the tests can never drift from production.
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $scriptPath = Join-Path $repoRoot 'LibraryLint.ps1'
    $parseTokens = $null
    $parseErrors = $null
    $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$parseTokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "LibraryLint.ps1 has $($parseErrors.Count) parse error(s); first: $($parseErrors[0].Message)"
    }
    $functionAst = $scriptAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-ArrStatusSummary'
    }, $true)
    if (-not $functionAst) { throw "Get-ArrStatusSummary not found in LibraryLint.ps1" }
    . ([scriptblock]::Create($functionAst.Extent.Text))

    $testUrl = 'http://arr.test:8989'
    $testKey = 'test-api-key'

    # Canned payloads, shaped like the real API responses.
    $okStatus = [PSCustomObject]@{ version = '4.0.19.2979' }

    $sonarrQueueRecord = [PSCustomObject]@{
        title                 = '[Group] Solo Leveling - S02E02 - I Suppose You Are Not Aware - 1080p'
        size                  = 1000
        sizeleft              = 0
        trackedDownloadStatus = 'warning'
        trackedDownloadState  = 'importPending'
        series                = [PSCustomObject]@{ title = 'Solo Leveling' }
        episode               = [PSCustomObject]@{ seasonNumber = 2; episodeNumber = 2 }
        statusMessages        = @(
            [PSCustomObject]@{ title = 'release'; messages = @('Not a quality revision upgrade for existing episode file(s)') }
        )
        errorMessage          = $null
    }

    $radarrQueueRecord = [PSCustomObject]@{
        title                 = 'Deadpool.2016.1080p.BluRay.x264-GRP'
        size                  = 1000
        sizeleft              = 250
        trackedDownloadStatus = 'ok'
        trackedDownloadState  = 'downloading'
        movie                 = [PSCustomObject]@{ title = 'Deadpool'; year = 2016 }
        statusMessages        = @()
        errorMessage          = $null
    }

    $bareQueueRecord = [PSCustomObject]@{
        title                 = 'Some.Unmatched.Release.2024.1080p'
        size                  = 0
        sizeleft              = 0
        trackedDownloadStatus = 'ok'
        trackedDownloadState  = 'downloading'
        statusMessages        = @()
        errorMessage          = $null
    }
}

Describe "Get-ArrStatusSummary configuration guard" {
    It "reports not configured without a URL or key" {
        $result = Get-ArrStatusSummary -Url '' -ApiKey ''
        $result.IsConfigured | Should -BeFalse
        $result.Error | Should -BeNullOrEmpty
    }
}

Describe "Get-ArrStatusSummary health messages" {
    # Regression: an empty health array arrives as $null, and @($null) is a
    # ONE-element array, so a perfectly healthy app rendered a single bare
    # "!" row on the dashboard.
    It "returns no messages when the API reports an empty health array" {
        Mock Invoke-RestMethod { $okStatus } -ParameterFilter { $Uri -like '*system/status*' }
        Mock Invoke-RestMethod { $null } -ParameterFilter { $Uri -like '*/health*' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 0; records = @() } } -ParameterFilter { $Uri -like '*/queue*' }
        Mock Invoke-RestMethod { @() } -ParameterFilter { $Uri -like '*diskspace*' }

        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        $result.IsConfigured | Should -BeTrue
        $result.Version | Should -Be '4.0.19.2979'
        @($result.HealthMessages).Count | Should -Be 0
    }

    It "maps type, source and message for real warnings" {
        Mock Invoke-RestMethod { $okStatus } -ParameterFilter { $Uri -like '*system/status*' }
        Mock Invoke-RestMethod {
            @(
                [PSCustomObject]@{ type = 'warning'; source = 'ImportListStatusCheck'; message = "Lists unavailable due to failures: Phil's List" }
                [PSCustomObject]@{ type = 'error';   source = 'DownloadClientCheck';   message = 'Unable to communicate with rTorrent' }
            )
        } -ParameterFilter { $Uri -like '*/health*' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 0; records = @() } } -ParameterFilter { $Uri -like '*/queue*' }
        Mock Invoke-RestMethod { @() } -ParameterFilter { $Uri -like '*diskspace*' }

        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        @($result.HealthMessages).Count | Should -Be 2
        $result.HealthMessages[0].Type | Should -Be 'warning'
        $result.HealthMessages[0].Message | Should -BeLike "*Phil's List*"
        $result.HealthMessages[1].Type | Should -Be 'error'
    }

    It "drops health entries that carry no message" {
        Mock Invoke-RestMethod { $okStatus } -ParameterFilter { $Uri -like '*system/status*' }
        Mock Invoke-RestMethod {
            @(
                [PSCustomObject]@{ type = 'warning'; source = 'Real'; message = 'A genuine warning' }
                [PSCustomObject]@{ type = 'warning'; source = 'Empty'; message = $null }
            )
        } -ParameterFilter { $Uri -like '*/health*' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 0; records = @() } } -ParameterFilter { $Uri -like '*/queue*' }
        Mock Invoke-RestMethod { @() } -ParameterFilter { $Uri -like '*diskspace*' }

        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        @($result.HealthMessages).Count | Should -Be 1
        $result.HealthMessages[0].Message | Should -Be 'A genuine warning'
    }
}

Describe "Get-ArrStatusSummary queue items" {
    BeforeEach {
        Mock Invoke-RestMethod { $okStatus } -ParameterFilter { $Uri -like '*system/status*' }
        Mock Invoke-RestMethod { $null } -ParameterFilter { $Uri -like '*/health*' }
        Mock Invoke-RestMethod { @() } -ParameterFilter { $Uri -like '*diskspace*' }
    }

    It "names a Sonarr item as series plus season and episode" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 1; records = @($sonarrQueueRecord) } } -ParameterFilter { $Uri -like '*/queue*' }
        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        $result.QueueCount | Should -Be 1
        $result.QueueItems[0].Name | Should -Be 'Solo Leveling S02E02'
        $result.QueueItems[0].State | Should -Be 'importPending'
        $result.QueueItems[0].Status | Should -Be 'warning'
    }

    It "surfaces the app's first status message as the item detail" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 1; records = @($sonarrQueueRecord) } } -ParameterFilter { $Uri -like '*/queue*' }
        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        $result.QueueItems[0].Detail | Should -Be 'Not a quality revision upgrade for existing episode file(s)'
    }

    It "names a Radarr item as movie plus year and computes progress" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 1; records = @($radarrQueueRecord) } } -ParameterFilter { $Uri -like '*/queue*' }
        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        $result.QueueItems[0].Name | Should -Be 'Deadpool (2016)'
        $result.QueueItems[0].ProgressPct | Should -Be 75
    }

    It "falls back to the release title when no media object is attached" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 1; records = @($bareQueueRecord) } } -ParameterFilter { $Uri -like '*/queue*' }
        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        $result.QueueItems[0].Name | Should -Be 'Some.Unmatched.Release.2024.1080p'
        $result.QueueItems[0].ProgressPct | Should -Be 0
    }

    It "reports an empty queue without inventing an item" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 0; records = $null } } -ParameterFilter { $Uri -like '*/queue*' }
        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        $result.QueueCount | Should -Be 0
        @($result.QueueItems).Count | Should -Be 0
    }
}

Describe "Get-ArrStatusSummary resilience" {
    It "marks the app unreachable when system/status fails" {
        Mock Invoke-RestMethod { throw 'Connection refused' }
        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        $result.IsConfigured | Should -BeTrue
        $result.Error | Should -Not -BeNullOrEmpty
    }

    # Regression: seedbox mounts report diskspace fields as ARRAYS, which
    # crashed the percentage maths. One shared try/catch meant that crash
    # discarded the health and queue data already fetched, silently hiding
    # *arr warnings from the dashboard. Each endpoint is isolated now.
    It "keeps health and queue data when diskspace returns array-valued fields" {
        Mock Invoke-RestMethod { $okStatus } -ParameterFilter { $Uri -like '*system/status*' }
        Mock Invoke-RestMethod {
            @([PSCustomObject]@{ type = 'warning'; source = 'Check'; message = 'Must survive the diskspace quirk' })
        } -ParameterFilter { $Uri -like '*/health*' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 1; records = @($sonarrQueueRecord) } } -ParameterFilter { $Uri -like '*/queue*' }
        Mock Invoke-RestMethod {
            @([PSCustomObject]@{ path = '/home16'; freeSpace = @(702000000000, 0); totalSpace = @(1820000000000, 0) })
        } -ParameterFilter { $Uri -like '*diskspace*' }

        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        @($result.HealthMessages).Count | Should -Be 1
        $result.HealthMessages[0].Message | Should -Be 'Must survive the diskspace quirk'
        $result.QueueCount | Should -Be 1
    }

    It "keeps health data when the queue endpoint fails" {
        Mock Invoke-RestMethod { $okStatus } -ParameterFilter { $Uri -like '*system/status*' }
        Mock Invoke-RestMethod {
            @([PSCustomObject]@{ type = 'error'; source = 'Check'; message = 'Health still reported' })
        } -ParameterFilter { $Uri -like '*/health*' }
        Mock Invoke-RestMethod { throw 'queue exploded' } -ParameterFilter { $Uri -like '*/queue*' }
        Mock Invoke-RestMethod { @() } -ParameterFilter { $Uri -like '*diskspace*' }

        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey $testKey
        $result.Error | Should -BeNullOrEmpty
        @($result.HealthMessages).Count | Should -Be 1
        $result.QueueCount | Should -Be 0
    }
}
