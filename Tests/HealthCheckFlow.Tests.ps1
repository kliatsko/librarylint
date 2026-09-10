#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for the library health check's category ordering and fix walk
    (Invoke-LibraryHealthCheck in LibraryLint.ps1).

.DESCRIPTION
    The report and the fix steps used to be two hand-maintained lists in
    different orders, so the screen showed orphaned subtitles second while the
    menu offered them last. They are now one ordered definition driving both.

    These tests pin that: a single definition exists, its order is the real
    dependency chain, and every fixable entry is complete enough to prompt
    with. Extracting and evaluating the actual array means a second list
    reappearing, or the order being shuffled, fails here.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\HealthCheckFlow.Tests.ps1 -Output Detailed
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

    $script:healthFn = $scriptAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-LibraryHealthCheck'
    }, $true)
    if (-not $script:healthFn) { throw 'Invoke-LibraryHealthCheck not found' }
    $script:healthSource = $script:healthFn.Extent.Text

    # Pull out the category table and evaluate it, so these assertions run
    # against the real data rather than against a copy in the test.
    $script:assignments = @($script:healthFn.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $node.Left.Extent.Text -eq '$healthCategories'
    }, $true))

    # The formatter for small videos calls this; defined so evaluating the
    # scriptblocks doesn't need the whole script loaded.
    function Format-FileSize { param($Size) "$Size B" }

    if ($script:assignments.Count -eq 1) {
        $script:categories = @(Invoke-Expression $script:assignments[0].Right.Extent.Text)
    } else {
        $script:categories = @()
    }
}

Describe "Health check category definition" {
    # The whole point: one list, not two that can disagree.
    It "defines the category order exactly once" {
        $script:assignments.Count | Should -Be 1
    }

    It "drives the report from that definition" {
        $script:healthSource | Should -Match 'foreach \(\$cat in \$healthCategories\)'
    }

    It "drives the fix walk from that same definition" {
        $script:healthSource | Should -Match '\$healthCategories \| Where-Object \{ \$_\.Fixable'
    }

    # A per-category display block reappearing is how the two lists drifted
    # the first time.
    It "has no hand-rolled per-category report blocks left" {
        foreach ($key in 'EmptyFolders', 'OrphanedSubtitles', 'NamingIssues', 'MismatchedFiles') {
            $script:healthSource | Should -Not -Match "if \(\`$issues\.$key\.Count -gt 0\) \{\s*\r?\n\s*Write-Host"
        }
    }
}

Describe "Health check category order" {
    It "lists every category the health check collects" {
        $keys = @($script:categories | ForEach-Object { $_.Key })
        foreach ($expected in 'EmptyFolders', 'NoVideoFiles', 'ZeroByteFiles', 'SmallVideos',
                              'NfoIdentity', 'NamingIssues', 'MismatchedFiles', 'MismatchedTrailers',
                              'OrphanedSubtitles', 'CodecSidecars') {
            $keys | Should -Contain $expected
        }
    }

    # Re-identifying an NFO can change its year; the folder rename that
    # follows must see the corrected one.
    It "settles NFO identity before fixing folder names" {
        $keys = @($script:categories | ForEach-Object { $_.Key })
        $keys.IndexOf('NfoIdentity') | Should -BeLessThan $keys.IndexOf('NamingIssues')
    }

    # These three are a genuine dependency chain, not a preference. Files are
    # renamed to match their folder, and subtitles to match their video, so
    # running them out of order re-breaks what the previous step just fixed.
    It "fixes folder names before renaming files to match them" {
        $keys = @($script:categories | ForEach-Object { $_.Key })
        $keys.IndexOf('NamingIssues') | Should -BeLessThan $keys.IndexOf('MismatchedFiles')
    }

    It "fixes file names before matching subtitles to them" {
        $keys = @($script:categories | ForEach-Object { $_.Key })
        $keys.IndexOf('MismatchedFiles') | Should -BeLessThan $keys.IndexOf('OrphanedSubtitles')
    }

    It "removes junk before reasoning about what is left" {
        $keys = @($script:categories | ForEach-Object { $_.Key })
        $keys.IndexOf('EmptyFolders') | Should -BeLessThan $keys.IndexOf('NamingIssues')
        $keys.IndexOf('ZeroByteFiles') | Should -BeLessThan $keys.IndexOf('NamingIssues')
    }
}

Describe "Health check category completeness" {
    It "gives every category a label, colour and formatter" {
        foreach ($cat in $script:categories) {
            $cat.Label | Should -Not -BeNullOrEmpty -Because "$($cat.Key) needs a heading"
            $cat.Color | Should -Not -BeNullOrEmpty -Because "$($cat.Key) needs a colour"
            $cat.Format | Should -BeOfType [scriptblock] -Because "$($cat.Key) needs a formatter"
        }
    }

    # A fixable category with no prompt would ask the user a blank question.
    It "gives every fixable category a prompt with a count placeholder" {
        foreach ($cat in ($script:categories | Where-Object { $_.Fixable })) {
            $cat.Prompt | Should -Not -BeNullOrEmpty -Because "$($cat.Key) is fixable so it must be able to ask"
            $cat.Prompt | Should -Match '\{0\}' -Because "$($cat.Key)'s prompt should state how many"
        }
    }

    It "marks the informational category as not fixable" {
        ($script:categories | Where-Object Key -eq 'NoVideoFiles').Fixable | Should -BeFalse
    }

    It "produces a printable line from each formatter" {
        # Objects shaped like what the health check collects.
        $samples = @{
            EmptyFolders       = [PSCustomObject]@{ FullName = 'C:\x\Empty'; Name = 'Empty' }
            NoVideoFiles       = [PSCustomObject]@{ Name = 'No Video (2020)' }
            ZeroByteFiles      = [PSCustomObject]@{ FullName = 'C:\x\zero.mkv'; Name = 'zero.mkv' }
            SmallVideos        = [PSCustomObject]@{ Name = 'small.mkv'; Length = 1024 }
            NfoIdentity        = [PSCustomObject]@{ Folder = 'Split (2016)'; NfoTitle = 'Split'; NfoYear = '2016'; NfoRuntime = 150; VideoMinutes = 117 }
            NamingIssues       = [PSCustomObject]@{ Path = 'C:\x\bad'; Issue = 'no year' }
            MismatchedFiles    = [PSCustomObject]@{ Folder = 'M (2020)'; CurrentFile = 'a.mkv'; ExpectedName = 'M (2020).mkv' }
            MismatchedTrailers = [PSCustomObject]@{ Folder = 'M (2020)'; CurrentTrailer = 'a-trailer.mp4'; ExpectedTrailer = 'M (2020)-trailer.mp4' }
            OrphanedSubtitles  = [PSCustomObject]@{ Name = 'orphan.en.srt' }
            CodecSidecars      = [PSCustomObject]@{ FullName = 'C:\x\codec-info.json' }
        }
        foreach ($cat in $script:categories) {
            $line = & $cat.Format $samples[$cat.Key]
            $line | Should -Not -BeNullOrEmpty -Because "$($cat.Key) must render its items"
        }
    }
}

Describe "Health check against a real folder" {
    # The real function, run against a temp library, with the interactive
    # phases and the helpers it calls stubbed. What it must not do: report a
    # category that found nothing as one item. A scan that captured an empty
    # pipeline held $null, @($null) counts as one, and the walk then offered
    # "Delete 1 empty folder(s)?" for a library with none — and the delete
    # failed binding a null path.
    BeforeAll {
        Import-Module (Join-Path $repoRoot 'modules\Quality.psm1') -Force
        foreach ($name in 'Invoke-LibraryHealthCheck', 'Test-FolderNameClean', 'Read-NFOFile') {
            $fn = $scriptAst.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
            }, $true)
            if (-not $fn) { throw "$name not found in LibraryLint.ps1" }
            . ([scriptblock]::Create($fn.Extent.Text))
        }
        function Write-Log { param($Message, $Level) }
        function Remove-CodecSidecarFiles { param($Path, $WhatIf) 0 }
        function Get-QuarantineRoot { param($LibraryRoot, [switch]$PromptIfMissing) $null }
        function Test-QuarantineContents { param($QuarantineRoot) @{ Exists = $false; FolderCount = 0 } }
        function Invoke-HealthCheckNfoPhase { param($Path) 'Continue' }
        function Invoke-HealthCheckDuplicatePhase { param($Path) 'Continue' }
        $script:Stats = @{ NFOFilesRead = 0 }
        $script:Config = @{ VideoExtensions = @('.mkv', '.mp4'); SubtitleExtensions = @('.srt'); DryRun = $false; MediaInfoPath = 'no-such-mediainfo' }
    }

    BeforeEach {
        $script:lib = Join-Path $TestDrive "hc-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
        New-Item -ItemType Directory -Path (Join-Path $script:lib 'Movie (2020)') -Force | Out-Null
        # Above the 50 MB "suspiciously small" line, so the only findings are
        # the ones each test plants.
        $stream = [IO.File]::Create((Join-Path $script:lib 'Movie (2020)\Movie (2020).mkv'))
        try { $stream.SetLength(51MB) } finally { $stream.Dispose() }
        Set-Content (Join-Path $script:lib 'Movie (2020)\Movie (2020).nfo') -Value '<movie><title>Movie</title><year>2020</year><uniqueid type="tmdb">1</uniqueid></movie>'
        $script:hostLines = [System.Collections.Generic.List[string]]::new()
        $script:prompts = [System.Collections.Generic.List[string]]::new()
        Mock Write-Host { $script:hostLines.Add(($Object -join ' ')) }
    }

    It "does not report categories that found nothing, nor prompt to fix them" {
        Mock Read-Host { $script:prompts.Add($Prompt); 'q' }
        $null = Invoke-LibraryHealthCheck -Path $script:lib -MediaType Movies
        $text = $script:hostLines -join "`n"
        $text | Should -Not -Match 'Empty Folders \(1\)'
        $text | Should -Not -Match 'Zero-Byte Files \(1\)'
        $text | Should -Not -Match 'Suspiciously Small Videos \(1\)'
        $text | Should -Not -Match 'Error during health check'
        ($script:prompts -join "`n") | Should -Not -Match 'Delete'
    }

    It "deletes a real empty folder when asked" {
        New-Item -ItemType Directory -Path (Join-Path $script:lib 'Empty') -Force | Out-Null
        Mock Read-Host { $script:prompts.Add($Prompt); if ($Prompt -like '*empty folder*') { 'y' } else { 'q' } }
        $null = Invoke-LibraryHealthCheck -Path $script:lib -MediaType Movies
        Test-Path (Join-Path $script:lib 'Empty') | Should -BeFalse
        ($script:hostLines -join "`n") | Should -Not -Match 'Error during health check'
    }
}

Describe "Health check fix walk" {
    It "offers a stop that abandons the remaining steps" {
        $script:healthSource | Should -Match "\(Y/N/Q\)"
        $script:healthSource | Should -Match "stepAnswer -match '\^\[Qq\]'"
    }

    It "defaults each step to yes so Enter walks straight through" {
        $script:healthSource | Should -Match "\(Y/N/Q\) \[Y\]"
    }

    # The numbered menu and its "fix all" option are what the walk replaced.
    It "no longer builds a numbered option list" {
        $script:healthSource | Should -Not -Match 'Fix all of the above'
        $script:healthSource | Should -Not -Match '\$actionOptions'
    }

    # Later counts are stale once an earlier fix lands, so the walk has to say
    # so rather than leaving wrong numbers on screen.
    It "offers a re-scan after anything is fixed" {
        $script:healthSource | Should -Match 'now stale'
        $script:healthSource | Should -Match 'Re-scan the library'
    }
}
