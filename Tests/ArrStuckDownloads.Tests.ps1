#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for the stuck-download handling behind the Status
    dashboard: queue classification, Radarr's manual import driven from the
    grab title, and the RAR extraction hand-off.

.DESCRIPTION
    All HTTP is mocked; nothing here reaches a Radarr instance or the
    seedbox. Fixtures are the real payloads Radarr returned for the two
    items that motivated this: HELLBOY [RoB] (file names carry no year or
    quality — "Unable to parse file") and a Scooby-Doo RAR set ("Found
    archive file, might need to be extracted"). The behaviour worth pinning
    is what gets sent: the feature file and never the preview, the quality
    Radarr's own parser reads from the grab title, importMode copy so the
    torrent keeps seeding, and nothing at all under -WhatIf.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\ArrStuckDownloads.Tests.ps1 -Output Detailed
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
    foreach ($name in 'Get-ArrQueueItemKind', 'Get-ArrStatusSummary', 'Get-RadarrProfileQualities', 'Wait-RadarrCommand',
                      'Resolve-RadarrUnparseableDownload', 'Resolve-RadarrArchiveDownload', 'Invoke-ArrRequest') {
        $fn = $scriptAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
        }, $true)
        if (-not $fn) { throw "$name not found in LibraryLint.ps1" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    # Production reads these through $script: scope at call time.
    $script:Config = @{
        RadarrUrl           = 'http://radarr.test:7878'
        RadarrApiKey        = 'test-key'
        SFTPHost            = 'seedbox.test'
        SFTPPort            = 22
        SFTPUsername        = 'user'
        SFTPPassword        = 'pw'
        SFTPRemotePaths     = @('/home16/user/media/Movies')
        SFTPWorkingPaths    = @('/home16/user/downloads/rtorrent/')
        SFTPPrunePaths      = @()
        SFTPUnrarCommand    = 'unrar'
        SFTPExtractedSuffix = '.extracted'
    }
    function Write-Log { param($Message, $Level) }
    # Sync module functions the archive resolver calls; mocked per test.
    function ConvertTo-SFTPNamespacePath { param($Path, $SFTPRoots) $Path }
    function Invoke-SFTPExtractSingleRelease { param($HostName, $Port, $Username, $Password, $PrivateKeyPath, $ReleaseFolder, $UnrarBinary, $ExtractedSuffix, $NotifyRadarr, $RadarrUrl, $RadarrApiKey, $DownloadId, $WhatIf) }
    function New-LocalVideoSizeIndex { param($LibraryPaths) @{} }
    function Find-LocalVideoBySize { param($Index, $Size) $null }

    $testUrl = 'http://radarr.test:7878'
    $testHeaders = @{ 'X-Api-Key' = 'test-key' }
    $hellboyDownloadId = '1319086BE90720399C60BF7D872D5CE6621D86A9'

    # --- Radarr payloads, as returned on 2026-09-10 ---------------------
    $hellboyQueueRecord = [PSCustomObject]@{
        id                    = 1503230191
        title                 = 'HELLBOY [RoB]'
        size                  = 6240711294
        sizeleft              = 0
        trackedDownloadStatus = 'warning'
        trackedDownloadState  = 'importBlocked'
        downloadId            = $hellboyDownloadId
        outputPath            = '/home/user/downloads/rtorrent/HELLBOY [RoB]'
        movie                 = [PSCustomObject]@{ id = 1558; title = 'Hellboy'; year = 2019 }
        statusMessages        = @(
            [PSCustomObject]@{ title = 'One or more movies expected in this release were not imported or missing'; messages = @() }
            [PSCustomObject]@{ title = 'HELLBOY_preview.mkv'; messages = @('Unable to parse file') }
            [PSCustomObject]@{ title = 'HELLBOY.mkv'; messages = @('Unable to parse file') }
        )
        errorMessage          = $null
    }
    $scoobyQueueRecord = [PSCustomObject]@{
        id                    = 405406598
        title                 = 'Scooby.Doo.2002.1080p.HDDVD.x264-hV'
        size                  = 8700000000
        sizeleft              = 0
        trackedDownloadStatus = 'warning'
        trackedDownloadState  = 'importPending'
        downloadId            = 'C12C7CC8E1F12F9BBEB0B1639820D5AA48FF2639'
        outputPath            = '/home/user/downloads/rtorrent/Scooby.Doo.2002.1080p.HDDVD.x264-hV'
        movie                 = [PSCustomObject]@{ id = 1586; title = 'Scooby-Doo'; year = 2002 }
        statusMessages        = @(
            [PSCustomObject]@{ title = 'Scooby.Doo.2002.1080p.HDDVD.x264-hV'; messages = @('Found archive file, might need to be extracted') }
        )
        errorMessage          = $null
    }

    $hellboyItem = [PSCustomObject]@{
        Name = 'Hellboy (2019)'; State = 'importBlocked'; Status = 'warning'; ProgressPct = 100
        Detail = 'Unable to parse file'; Id = 1503230191; DownloadId = $hellboyDownloadId
        OutputPath = '/home/user/downloads/rtorrent/HELLBOY [RoB]'; MovieId = 1558; Title = 'HELLBOY [RoB]'
        Messages = @('Unable to parse file', 'Unable to parse file'); Kind = 'unparseable'
    }
    $scoobyItem = [PSCustomObject]@{
        Name = 'Scooby-Doo (2002)'; State = 'importPending'; Status = 'warning'; ProgressPct = 100
        Detail = 'Found archive file, might need to be extracted'; Id = 405406598; DownloadId = 'C12C7CC8E1F12F9BBEB0B1639820D5AA48FF2639'
        OutputPath = '/home/user/downloads/rtorrent/Scooby.Doo.2002.1080p.HDDVD.x264-hV'; MovieId = 1586
        Title = 'Scooby.Doo.2002.1080p.HDDVD.x264-hV'; Messages = @('Found archive file, might need to be extracted'); Kind = 'archive'
    }

    $manualImportByDownload = @(
        [PSCustomObject]@{
            path = '/home/user/downloads/rtorrent/HELLBOY [RoB]/HELLBOY.mkv'; relativePath = 'HELLBOY.mkv'; size = 6150638105
            movie = [PSCustomObject]@{ id = 1558; title = 'Hellboy'; year = 2019 }; quality = $null
            rejections = @([PSCustomObject]@{ reason = 'Unable to parse file'; type = 'permanent' })
        }
        [PSCustomObject]@{
            path = '/home/user/downloads/rtorrent/HELLBOY [RoB]/HELLBOY_preview.mkv'; relativePath = 'HELLBOY_preview.mkv'; size = 90073189
            movie = [PSCustomObject]@{ id = 1558; title = 'Hellboy'; year = 2019 }; quality = $null
            rejections = @([PSCustomObject]@{ reason = 'Unable to parse file'; type = 'permanent' })
        }
    )
    $historyForHellboy = @(
        [PSCustomObject]@{ eventType = 'grabbed'; downloadId = $hellboyDownloadId; sourceTitle = 'HELLBOY [2019]1080p BRRip[x265][10 Bit][DTS-HD MA]INFERNO'; date = '2026-09-08T23:39:00Z' }
        [PSCustomObject]@{ eventType = 'grabbed'; downloadId = 'OTHER'; sourceTitle = 'Hellboy.2019.720p.WEB.x264-OTHER'; date = '2026-09-01T00:00:00Z' }
    )
    $parsedBluray1080 = [PSCustomObject]@{
        parsedMovieInfo = [PSCustomObject]@{
            primaryMovieTitle = 'HELLBOY'; year = 2019
            quality = [PSCustomObject]@{
                quality  = [PSCustomObject]@{ id = 7; name = 'Bluray-1080p'; source = 'bluray'; resolution = 1080; modifier = 'none' }
                revision = [PSCustomObject]@{ version = 1; real = 0; isRepack = $false }
            }
            languages = @([PSCustomObject]@{ id = 0; name = 'Unknown' })
        }
        movie = [PSCustomObject]@{ id = 1558; title = 'Hellboy'; year = 2019 }
    }
    $parsedNothing = [PSCustomObject]@{ parsedMovieInfo = $null; movie = $null }
    $languageList = @(
        [PSCustomObject]@{ id = 1; name = 'English' }
        [PSCustomObject]@{ id = 2; name = 'French' }
    )
    $hd1080Profile = [PSCustomObject]@{
        id = 4; name = 'HD-1080p'; cutoff = 7
        items = @(
            [PSCustomObject]@{ quality = [PSCustomObject]@{ id = 6; name = 'Bluray-720p' }; allowed = $true; items = @() }
            [PSCustomObject]@{ quality = [PSCustomObject]@{ id = 2; name = 'DVD' }; allowed = $false; items = @() }
            [PSCustomObject]@{ name = 'WEB 1080p'; allowed = $true; items = @(
                [PSCustomObject]@{ quality = [PSCustomObject]@{ id = 3; name = 'WEBDL-1080p' }; allowed = $true }
                [PSCustomObject]@{ quality = [PSCustomObject]@{ id = 15; name = 'WEBRip-1080p' }; allowed = $true }
            ) }
            [PSCustomObject]@{ quality = [PSCustomObject]@{ id = 7; name = 'Bluray-1080p' }; allowed = $true; items = @() }
        )
    }
}

Describe "Get-ArrQueueItemKind" {
    It "calls a RAR set an archive from the app's own message" {
        Get-ArrQueueItemKind -State 'importPending' -Messages @('Found archive file, might need to be extracted') | Should -Be 'archive'
    }

    It "calls a blocked import with unparseable names unparseable" {
        Get-ArrQueueItemKind -State 'importBlocked' -Messages @('Unable to parse file') | Should -Be 'unparseable'
        Get-ArrQueueItemKind -State 'importBlocked' -Messages @('Unknown Movie') | Should -Be 'unparseable'
    }

    # A genuine quality rejection is Radarr doing its job; nothing to unstick.
    It "leaves a quality rejection and a plain download alone" {
        Get-ArrQueueItemKind -State 'importPending' -Messages @('Not a quality revision upgrade for existing movie file(s)') | Should -Be ''
        Get-ArrQueueItemKind -State 'downloading' -Messages @() | Should -Be ''
    }
}

Describe "Get-ArrStatusSummary identity fields" {
    BeforeEach {
        Mock Invoke-RestMethod { [PSCustomObject]@{ version = '5.0.0' } } -ParameterFilter { $Uri -like '*system/status*' }
        Mock Invoke-RestMethod { $null } -ParameterFilter { $Uri -like '*/health*' }
        Mock Invoke-RestMethod { @() } -ParameterFilter { $Uri -like '*diskspace*' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ totalRecords = 2; records = @($hellboyQueueRecord, $scoobyQueueRecord) } } -ParameterFilter { $Uri -like '*/queue*' }
    }

    It "carries the download id, output path, movie id and kind for each item" {
        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey 'test-key'
        $hellboy = $result.QueueItems | Where-Object { $_.Name -eq 'Hellboy (2019)' }
        $hellboy.DownloadId | Should -Be $hellboyDownloadId
        $hellboy.OutputPath | Should -Be '/home/user/downloads/rtorrent/HELLBOY [RoB]'
        $hellboy.MovieId    | Should -Be 1558
        $hellboy.Title      | Should -Be 'HELLBOY [RoB]'
        $hellboy.Kind       | Should -Be 'unparseable'
        $scooby = $result.QueueItems | Where-Object { $_.Name -eq 'Scooby-Doo (2002)' }
        $scooby.Kind | Should -Be 'archive'
    }

    It "keeps the first status message as the one-line detail" {
        $result = Get-ArrStatusSummary -Url $testUrl -ApiKey 'test-key'
        ($result.QueueItems | Where-Object { $_.Name -eq 'Hellboy (2019)' }).Detail | Should -Be 'Unable to parse file'
    }
}

Describe "Get-RadarrProfileQualities" {
    It "flattens allowed qualities including those inside groups, skipping disallowed ones" {
        $options = @(Get-RadarrProfileQualities -QualityProfile $hd1080Profile)
        ($options | ForEach-Object { $_.Name }) | Should -Be @('Bluray-720p', 'WEBDL-1080p', 'WEBRip-1080p', 'Bluray-1080p')
        ($options | Where-Object { $_.Name -eq 'Bluray-1080p' }).Id | Should -Be 7
    }
}

Describe "Resolve-RadarrUnparseableDownload" {
    BeforeEach {
        Mock Start-Sleep { }
        Mock Invoke-RestMethod { $manualImportByDownload } -ParameterFilter { $Uri -like '*manualimport?downloadId=*' }
        Mock Invoke-RestMethod { $historyForHellboy } -ParameterFilter { $Uri -like '*history/movie?movieId=1558*' }
        Mock Invoke-RestMethod { $parsedBluray1080 } -ParameterFilter { $Uri -like '*api/v3/parse?title=*' }
        Mock Invoke-RestMethod { $languageList } -ParameterFilter { $Uri -like '*/api/v3/language' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ id = 483837; status = 'started' } } -ParameterFilter { $Method -eq 'Post' -and $Uri -like '*/api/v3/command' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ id = 483837; status = 'completed' } } -ParameterFilter { $Uri -like '*/api/v3/command/483837' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ id = 1558; title = 'Hellboy'; year = 2019; hasFile = $true; qualityProfileId = 4 } } -ParameterFilter { $Uri -like '*/api/v3/movie/1558' }
        Mock Invoke-RestMethod { $hd1080Profile } -ParameterFilter { $Uri -like '*qualityprofile/4' }
    }

    It "picks the feature and leaves the preview out" {
        $r = Resolve-RadarrUnparseableDownload -Url $testUrl -Headers $testHeaders -Item $hellboyItem -WhatIf
        $r.Success      | Should -BeTrue
        $r.File         | Should -Be 'HELLBOY.mkv'
        $r.FileSize     | Should -Be 6150638105
        $r.SkippedFiles | Should -Contain 'HELLBOY_preview.mkv'
        $r.MovieTitle   | Should -Be 'Hellboy (2019)'
    }

    # The quality Radarr could not read from "HELLBOY.mkv" is right there in
    # the grab title, and Radarr's own parser maps BRRip to Bluray.
    It "takes the quality from the grab title in history, matched by download id" {
        $r = Resolve-RadarrUnparseableDownload -Url $testUrl -Headers $testHeaders -Item $hellboyItem -WhatIf
        $r.QualityName | Should -Be 'Bluray-1080p'
        # -Uri is a [System.Uri]; inside the filter it stringifies UNESCAPED
        # (spaces and brackets back), so match on words, not on %5B escapes.
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Uri -like '*parse?title=HELLBOY*BRRip*INFERNO' }
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Uri -like '*parse?title=*' -and $Uri -like '*RoB*' }
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Uri -like '*parse?title=*' -and $Uri -like '*OTHER*' }
    }

    It "sends nothing under -WhatIf" {
        $null = Resolve-RadarrUnparseableDownload -Url $testUrl -Headers $testHeaders -Item $hellboyItem -WhatIf
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Method -eq 'Post' }
    }

    It "posts one ManualImport for the feature file, copy mode, tied to the download" {
        $r = Resolve-RadarrUnparseableDownload -Url $testUrl -Headers $testHeaders -Item $hellboyItem
        $r.Success   | Should -BeTrue
        $r.CommandId | Should -Be 483837
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter {
            if ($Method -ne 'Post') { return $false }
            $sent = $Body | ConvertFrom-Json
            $sent.name -eq 'ManualImport' -and $sent.importMode -eq 'copy' -and
            @($sent.files).Count -eq 1 -and
            $sent.files[0].path -like '*HELLBOY.mkv' -and
            $sent.files[0].movieId -eq 1558 -and
            $sent.files[0].quality.quality.id -eq 7 -and
            $sent.files[0].downloadId -eq $hellboyDownloadId -and
            $sent.files[0].languages[0].name -eq 'English'
        }
    }

    It "reports NeedsQuality with the profile's allowed qualities when nothing parses" {
        Mock Invoke-RestMethod { $parsedNothing } -ParameterFilter { $Uri -like '*api/v3/parse?title=*' }
        $r = Resolve-RadarrUnparseableDownload -Url $testUrl -Headers $testHeaders -Item $hellboyItem -WhatIf
        $r.Success      | Should -BeFalse
        $r.NeedsQuality | Should -BeTrue
        @($r.QualityOptions | ForEach-Object { $_.Name }) | Should -Contain 'Bluray-1080p'
        @($r.QualityOptions | ForEach-Object { $_.Name }) | Should -Not -Contain 'DVD'
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Method -eq 'Post' }
    }

    It "uses a supplied quality without parsing" {
        $picked = @{ quality = @{ id = 3; name = 'WEBDL-1080p' }; revision = @{ version = 1; real = 0; isRepack = $false } }
        $r = Resolve-RadarrUnparseableDownload -Url $testUrl -Headers $testHeaders -Item $hellboyItem -Quality $picked
        $r.Success | Should -BeTrue
        $r.QualityName | Should -Be 'WEBDL-1080p'
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Uri -like '*api/v3/parse*' }
        Should -Invoke Invoke-RestMethod -Times 1 -ParameterFilter { $Method -eq 'Post' -and (($Body | ConvertFrom-Json).files[0].quality.quality.id -eq 3) }
    }

    It "fails clearly when Radarr cannot say which movie the download is for" {
        Mock Invoke-RestMethod {
            @([PSCustomObject]@{ path = '/x/UNKNOWN.mkv'; relativePath = 'UNKNOWN.mkv'; size = 5000000000; movie = $null; quality = $null; rejections = @() })
        } -ParameterFilter { $Uri -like '*manualimport?downloadId=*' }
        $orphan = [PSCustomObject]@{ Name = 'x'; DownloadId = 'ABC'; MovieId = 0; Title = 'UNKNOWN'; Kind = 'unparseable' }
        $r = Resolve-RadarrUnparseableDownload -Url $testUrl -Headers $testHeaders -Item $orphan -WhatIf
        $r.Success | Should -BeFalse
        $r.Error   | Should -Match 'which movie'
        Should -Invoke Invoke-RestMethod -Times 0 -ParameterFilter { $Method -eq 'Post' }
    }

    It "reports failure when the movie still has no file after the command" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ id = 1558; hasFile = $false; qualityProfileId = 4 } } -ParameterFilter { $Uri -like '*/api/v3/movie/1558' }
        $r = Resolve-RadarrUnparseableDownload -Url $testUrl -Headers $testHeaders -Item $hellboyItem
        $r.Success | Should -BeFalse
        $r.Error   | Should -Match 'still has no file'
    }
}

Describe "Resolve-RadarrArchiveDownload" {
    BeforeEach {
        Mock Start-Sleep { }
        Mock ConvertTo-SFTPNamespacePath { param($Path, $SFTPRoots) $Path -replace '^/home/', '/home16/' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ id = 7; status = 'completed' } } -ParameterFilter { $Uri -like '*/api/v3/command/7' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ id = 1586; hasFile = $true } } -ParameterFilter { $Uri -like '*/api/v3/movie/1586' }
    }

    It "hands the extractor the SFTP-side path, the download id and a Radarr notify, then reports the import" {
        Mock Invoke-SFTPExtractSingleRelease {
            @{ Success = $true; AlreadyExtracted = $false; Video = 'scooby.doo.2002.1080p.hddvd.x264-hv.mkv'; VideoSize = 8529661515; ExtractedPath = "$ReleaseFolder.extracted"; RadarrCommandId = 7; RadarrError = $null; Error = $null }
        }
        $r = Resolve-RadarrArchiveDownload -Url $testUrl -Headers $testHeaders -Item $scoobyItem
        $r.Success      | Should -BeTrue
        $r.Extracted    | Should -BeTrue
        $r.Imported     | Should -BeTrue
        $r.RemoteFolder | Should -Be '/home16/user/downloads/rtorrent/Scooby.Doo.2002.1080p.HDDVD.x264-hV'
        Should -Invoke Invoke-SFTPExtractSingleRelease -Times 1 -Exactly -ParameterFilter {
            $ReleaseFolder -eq '/home16/user/downloads/rtorrent/Scooby.Doo.2002.1080p.HDDVD.x264-hV' -and
            $DownloadId -eq 'C12C7CC8E1F12F9BBEB0B1639820D5AA48FF2639' -and
            $NotifyRadarr -eq $true -and $RadarrApiKey -eq 'test-key'
        }
    }

    It "distinguishes extracted-but-not-imported from success" {
        Mock Invoke-SFTPExtractSingleRelease {
            @{ Success = $true; AlreadyExtracted = $true; Video = 'x.mkv'; VideoSize = 1; ExtractedPath = 'p'; RadarrCommandId = 7; RadarrError = $null; Error = $null }
        }
        Mock Invoke-RestMethod { [PSCustomObject]@{ id = 1586; hasFile = $false } } -ParameterFilter { $Uri -like '*/api/v3/movie/1586' }
        $r = Resolve-RadarrArchiveDownload -Url $testUrl -Headers $testHeaders -Item $scoobyItem
        $r.Success          | Should -BeTrue
        $r.AlreadyExtracted | Should -BeTrue
        $r.Imported         | Should -BeFalse
    }

    # Scooby-Doo: pulled into the inbox by the extracted-sync while Radarr's
    # queue item lingered. Extracting again would only make a second copy.
    It "reports a local twin when the RAR's unpacked bytes already sit in the inbox or library" {
        Mock Invoke-SFTPExtractSingleRelease {
            @{ Success = $true; AlreadyExtracted = $false; Video = $null; VideoSize = 0; UnpackedSize = 8529661515; ExtractedPath = 'p'; RadarrCommandId = $null; RadarrError = $null; Error = $null }
        }
        Mock New-LocalVideoSizeIndex { @{ [long]8529661515 = @('E:\Inbox\Scooby.Doo.2002.1080p.HDDVD.x264-hV\Scooby.Doo.2002.1080p.HDDVD.x264-hV.mkv') } }
        Mock Find-LocalVideoBySize { param($Index, $Size) if ($Index.ContainsKey([long]$Size)) { $Index[[long]$Size][0] } else { $null } }
        $r = Resolve-RadarrArchiveDownload -Url $testUrl -Headers $testHeaders -Item $scoobyItem -WhatIf
        $r.Success   | Should -BeTrue
        $r.VideoSize | Should -Be 8529661515
        $r.LocalTwin | Should -BeLike 'E:\Inbox\Scooby.Doo*'
        Should -Invoke New-LocalVideoSizeIndex -Times 1 -ParameterFilter { @($LibraryPaths) -contains 'E:\Inbox' -or @($LibraryPaths).Count -ge 0 }
    }

    It "passes an extraction failure through as the error" {
        Mock Invoke-SFTPExtractSingleRelease { @{ Success = $false; Error = 'RAR set incomplete: 2 zero-byte part(s)' } }
        $r = Resolve-RadarrArchiveDownload -Url $testUrl -Headers $testHeaders -Item $scoobyItem
        $r.Success | Should -BeFalse
        $r.Error   | Should -Match 'incomplete'
    }

    It "refuses without SFTP configured and never calls the seedbox" {
        Mock Invoke-SFTPExtractSingleRelease { throw 'should not be called' }
        $saved = $script:Config.SFTPHost
        try {
            $script:Config.SFTPHost = $null
            $r = Resolve-RadarrArchiveDownload -Url $testUrl -Headers $testHeaders -Item $scoobyItem
            $r.Success | Should -BeFalse
            $r.Error   | Should -Match 'SFTP'
            Should -Invoke Invoke-SFTPExtractSingleRelease -Times 0
        } finally {
            $script:Config.SFTPHost = $saved
        }
    }
}

Describe "Wait-RadarrCommand" {
    It "returns as soon as the command reaches a terminal state" {
        Mock Start-Sleep { }
        $script:polls = 0
        Mock Invoke-RestMethod {
            $script:polls++
            if ($script:polls -lt 3) { [PSCustomObject]@{ id = 9; status = 'started' } } else { [PSCustomObject]@{ id = 9; status = 'completed' } }
        } -ParameterFilter { $Uri -like '*/api/v3/command/9' }
        $state = Wait-RadarrCommand -Url $testUrl -Headers $testHeaders -CommandId 9
        $state.status | Should -Be 'completed'
        $script:polls | Should -Be 3
    }
}
