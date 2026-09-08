#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for Invoke-ArrRequest, the shared Radarr/Sonarr caller.

.DESCRIPTION
    Every *arr request in the codebase goes through this, so its retry
    classification is worth pinning precisely: retrying a genuine error
    would double-apply writes, and failing to retry a transient SQLite
    lock is the bug it exists to prevent.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\ArrRequest.Tests.ps1 -Output Detailed
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
    foreach ($name in 'Invoke-ArrRequest', 'Initialize-SonarrConnection') {
        $fn = $scriptAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
        }, $true)
        if (-not $fn) { throw "$name not found in LibraryLint.ps1" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $testHeaders = @{ 'X-Api-Key' = 'k' }
}

Describe "Invoke-ArrRequest happy path" {
    It "returns whatever the endpoint returned" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ version = '5.0' } }
        (Invoke-ArrRequest -Uri 'http://a/api/v3/system/status' -Headers $testHeaders).version | Should -Be '5.0'
        Should -Invoke Invoke-RestMethod -Times 1
    }

    # Pester surfaces a mocked command's bound parameters as named
    # variables inside the mock body, so capture those rather than
    # $PSBoundParameters, which belongs to the mock scriptblock itself.
    It "defaults to a GET with no body" {
        $script:seenMethod = 'unset'; $script:seenBody = 'unset'; $script:seenType = 'unset'
        Mock Invoke-RestMethod { $script:seenMethod = $Method; $script:seenBody = $Body; $script:seenType = $ContentType; $null }
        $null = Invoke-ArrRequest -Uri 'http://a/api/v3/movie' -Headers $testHeaders
        $script:seenMethod | Should -Be 'Get'
        $script:seenBody | Should -BeNullOrEmpty
        $script:seenType | Should -BeNullOrEmpty
    }

    It "sends JSON content type whenever a body is supplied" {
        $script:seenMethod = $null; $script:seenBody = $null; $script:seenType = $null
        Mock Invoke-RestMethod { $script:seenMethod = $Method; $script:seenBody = $Body; $script:seenType = $ContentType; $null }
        $null = Invoke-ArrRequest -Uri 'http://a/api/v3/command' -Headers $testHeaders -Method Post -Body '{"name":"MoviesSearch"}'
        $script:seenMethod | Should -Be 'Post'
        $script:seenBody | Should -Be '{"name":"MoviesSearch"}'
        $script:seenType | Should -Be 'application/json'
    }

    It "rejects a method outside the known verbs" {
        { Invoke-ArrRequest -Uri 'http://a' -Headers $testHeaders -Method 'Patch' -ErrorAction Stop } | Should -Throw
    }
}

Describe "Invoke-ArrRequest retry classification" {
    # These are the conditions *arr apps produce under their own load, where
    # one retry succeeds. Retrying anything else risks double-applying a
    # write, which is why the list is deliberately narrow.
    It "retries once on a transient condition and succeeds" -ForEach @(
        @{ Failure = 'database is locked' }
        @{ Failure = 'The operation has timed out' }
        @{ Failure = 'Service temporarily unavailable' }
        @{ Failure = 'connection reset by peer' }
    ) {
        $script:attempts = 0
        Mock Invoke-RestMethod {
            $script:attempts++
            if ($script:attempts -eq 1) { throw $Failure }
            [PSCustomObject]@{ ok = $true }
        }
        (Invoke-ArrRequest -Uri 'http://a' -Headers $testHeaders).ok | Should -BeTrue
        $script:attempts | Should -Be 2
    }

    It "does not retry a genuine error, so a write is never double-applied" {
        $script:attempts = 0
        Mock Invoke-RestMethod { $script:attempts++; throw 'Invalid quality profile' }
        { Invoke-ArrRequest -Uri 'http://a' -Headers $testHeaders } | Should -Throw
        $script:attempts | Should -Be 1
    }

    # Radarr rejects a duplicate add with its own message. Callers classify
    # that as success, so it must reach them unretried and unswallowed.
    It "propagates an already-exists rejection to the caller" {
        Mock Invoke-RestMethod { throw 'This movie has already been added' }
        { Invoke-ArrRequest -Uri 'http://a' -Headers $testHeaders -Method Post -Body '{}' } |
            Should -Throw -ExpectedMessage '*already been added*'
    }

    It "rethrows when the retry also fails" {
        $script:attempts = 0
        Mock Invoke-RestMethod { $script:attempts++; throw 'database is locked' }
        { Invoke-ArrRequest -Uri 'http://a' -Headers $testHeaders } | Should -Throw
        $script:attempts | Should -Be 2
    }
}

Describe "Invoke-ArrRequest SuppressErrors" {
    # A TMDB-id lookup that finds nothing is a normal answer, and the
    # caller's next move is a title search. Returning null lets it fall
    # through without wrapping every lookup in its own try/catch.
    It "returns null instead of throwing when a lookup fails" {
        Mock Invoke-RestMethod { throw 'NotFound' }
        $result = Invoke-ArrRequest -Uri 'http://a/api/v3/movie/lookup/tmdb?tmdbId=1' -Headers $testHeaders -SuppressErrors
        $result | Should -BeNullOrEmpty
    }

    It "returns null when the retry also fails" {
        Mock Invoke-RestMethod { throw 'database is locked' }
        $result = Invoke-ArrRequest -Uri 'http://a/api/v3/movie/lookup/tmdb?tmdbId=1' -Headers $testHeaders -SuppressErrors
        $result | Should -BeNullOrEmpty
    }

    It "throws by default so real failures are not swallowed" {
        Mock Invoke-RestMethod { throw 'NotFound' }
        { Invoke-ArrRequest -Uri 'http://a/api/v3/movie' -Headers $testHeaders } | Should -Throw
    }
}

Describe "Initialize-SonarrConnection" {
    BeforeEach {
        $script:Config = @{ SonarrUrl = 'http://sonarr.test:8989/'; SonarrApiKey = 'sk' }
    }

    It "reports not configured when the URL or key is missing" {
        $script:Config = @{ SonarrUrl = $null; SonarrApiKey = $null }
        $conn = Initialize-SonarrConnection
        $conn.Ok | Should -BeFalse
        $conn.Error | Should -Be 'not configured'
    }

    It "returns a trimmed URL and usable headers on success" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ version = '4.0.19.2979' } }
        $conn = Initialize-SonarrConnection
        $conn.Ok | Should -BeTrue
        $conn.Url | Should -Be 'http://sonarr.test:8989'
        $conn.Headers['X-Api-Key'] | Should -Be 'sk'
        $conn.Version | Should -Be '4.0.19.2979'
    }

    It "reports unreachable rather than throwing when the status call fails" {
        Mock Invoke-RestMethod { throw 'Connection refused' }
        $conn = Initialize-SonarrConnection
        $conn.Ok | Should -BeFalse
        $conn.Error | Should -Not -BeNullOrEmpty
    }
}
