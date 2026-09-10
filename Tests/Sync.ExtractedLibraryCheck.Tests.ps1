#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for the seedbox extraction tools' "already in the library"
    check and the single-release extractor (modules\Sync.psm1).

.DESCRIPTION
    The extracted-sync tool once pulled eight movies the library had held
    for months: their <Release>.extracted/ folders were leftovers of
    extractions Radarr had imported long ago, and the tool only asked "is
    it in the inbox?". The extractor's upgrade-tag bypass (REMASTERED,
    PROPER, RERIP...) cannot tell "this release IS the library copy" from
    "this might upgrade it" by name either. An exact byte-size match against
    the library can — the extracted video is the very file the library came
    from. These tests pin that check, the /home vs /home<digits> namespace
    translation Radarr-reported paths need, the unrar listing parser that
    reads a RAR's unpacked size before extraction, and the single-release
    extractor's contract with Radarr.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Sync.ExtractedLibraryCheck.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Sync.psm1') -Force

    function Get-JoinedSource {
        param([string]$FunctionName)
        $src = (Get-Command $FunctionName).ScriptBlock.ToString()
        return ($src -replace '`\r?\n\s*', ' ')
    }
    $script:extractedSyncSource = Get-JoinedSource -FunctionName 'Invoke-SFTPExtractedSync'
    $script:extractorSource     = Get-JoinedSource -FunctionName 'Invoke-SFTPExtractRarReleases'
}

Describe "Namespace translation" {
    # Radarr reports /home/<user>/...; the SFTP session sees /home16/<user>/...
    It "rewrites a chroot path to the numbered SFTP projection the configured roots reveal" {
        InModuleScope Sync {
            ConvertTo-SFTPNamespacePath -Path '/home/user/downloads/rtorrent/X' -SFTPRoots @('/home16/user/media/Movies', '/home16/user/downloads/rtorrent/') |
                Should -Be '/home16/user/downloads/rtorrent/X'
        }
    }

    It "leaves a path alone when no root carries a numbered prefix" {
        InModuleScope Sync {
            ConvertTo-SFTPNamespacePath -Path '/home/user/x' -SFTPRoots @('/data/media') | Should -Be '/home/user/x'
            ConvertTo-SFTPNamespacePath -Path '/mnt/x' -SFTPRoots @('/home16/user/media') | Should -Be '/mnt/x'
        }
    }

    It "rewrites the SFTP projection back to the chroot view for Radarr" {
        InModuleScope Sync {
            ConvertTo-ChrootNamespacePath -Path '/home16/user/downloads/rtorrent/X.extracted' | Should -Be '/home/user/downloads/rtorrent/X.extracted'
            ConvertTo-ChrootNamespacePath -Path '/home/user/x' | Should -Be '/home/user/x'
        }
    }
}

Describe "ConvertFrom-UnrarListing" {
    # Captured from `unrar l` on the seedbox (UNRAR 7.12).
    It "reads the unpacked name and size and ignores banner, rule and totals" {
        InModuleScope Sync {
            $lines = @(
                ''
                'UNRAR 7.12 freeware      Copyright (c) 1993-2025 Alexander Roshal'
                ''
                'Archive: /home16/user/downloads/rtorrent/Scooby.Doo.2002.1080p.HDDVD.x264-hV/scooby.doo.2002.1080p.hddvd.x264-hv.rar'
                'Details: RAR 1.5, volume, recovery record'
                ''
                ' Attributes      Size     Date    Time   Name'
                '----------- ---------  ---------- -----  ----'
                '    ..A.... 8529661515  2008-10-10 17:02  scooby.doo.2002.1080p.hddvd.x264-hv.mkv'
                '----------- ---------  ---------- -----  ----'
                '                    0  volume 87         0'
                '           8529661515                    1'
            )
            $entries = @(ConvertFrom-UnrarListing -Lines $lines)
            $entries.Count   | Should -Be 1
            $entries[0].Name | Should -Be 'scooby.doo.2002.1080p.hddvd.x264-hv.mkv'
            $entries[0].Size | Should -Be 8529661515
        }
    }

    It "reports the largest video's size from a listing, or nothing" {
        InModuleScope Sync {
            $session = [PSCustomObject]@{}
            $session | Add-Member -MemberType ScriptMethod -Name ExecuteCommand -Value {
                param($cmd)
                [PSCustomObject]@{ Output = "    ..A....       1234  2020-01-01 00:00  a.nfo`n    ..A.... 8529661515  2008-10-10 17:02  movie.mkv`n    ..A....    5000000  2008-10-10 17:02  sample.mkv" }
            }
            Get-RarArchiveUnpackedSize -Session $session -FirstArchive '/x/y.rar' | Should -Be 8529661515
            $empty = [PSCustomObject]@{}
            $empty | Add-Member -MemberType ScriptMethod -Name ExecuteCommand -Value { param($cmd) [PSCustomObject]@{ Output = '' } }
            Get-RarArchiveUnpackedSize -Session $empty -FirstArchive '/x/y.rar' | Should -BeNullOrEmpty
        }
    }
}

Describe "Library size index" {
    BeforeAll {
        $script:lib = Join-Path $TestDrive 'Movies'
        New-Item -ItemType Directory -Path (Join-Path $script:lib 'Batman (1989)\extras') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $script:lib 'Candy (2006)') -Force | Out-Null
        # Sizes are what matters; content is filler.
        [IO.File]::WriteAllBytes((Join-Path $script:lib 'Batman (1989)\Batman (1989).mkv'), (New-Object byte[] 4096))
        [IO.File]::WriteAllBytes((Join-Path $script:lib 'Batman (1989)\extras\featurette.mkv'), (New-Object byte[] 5000))
        [IO.File]::WriteAllBytes((Join-Path $script:lib 'Batman (1989)\poster.jpg'), (New-Object byte[] 4096))
        [IO.File]::WriteAllBytes((Join-Path $script:lib 'Candy (2006)\Candy (2006).mp4'), (New-Object byte[] 7000))
    }

    It "finds a library video by exact size regardless of its name" {
        InModuleScope Sync -Parameters @{ Lib = $script:lib } {
            param($Lib)
            $index = New-LocalVideoSizeIndex -LibraryPaths @($Lib) -MinBytes 1KB
            Find-LocalVideoBySize -Index $index -Size 4096 | Should -BeLike '*Batman (1989).mkv'
            Find-LocalVideoBySize -Index $index -Size 7000 | Should -BeLike '*Candy (2006).mp4'
        }
    }

    It "does not match a size one byte off, a non-video, or a video buried in extras" {
        InModuleScope Sync -Parameters @{ Lib = $script:lib } {
            param($Lib)
            $index = New-LocalVideoSizeIndex -LibraryPaths @($Lib) -MinBytes 1KB
            Find-LocalVideoBySize -Index $index -Size 4097 | Should -BeNullOrEmpty
            Find-LocalVideoBySize -Index $index -Size 5000 | Should -BeNullOrEmpty
            # poster.jpg is also 4096 bytes; only the video may answer.
            @($index[[long]4096]).Count | Should -Be 1
        }
    }

    It "tolerates missing roots and an empty index" {
        InModuleScope Sync {
            $index = New-LocalVideoSizeIndex -LibraryPaths @('Z:\does\not\exist', $null)
            $index.Count | Should -Be 0
            Find-LocalVideoBySize -Index $index -Size 4096 | Should -BeNullOrEmpty
            Find-LocalVideoBySize -Index $null -Size 4096 | Should -BeNullOrEmpty
        }
    }
}

Describe "Extraction tools consult the library" {
    # Pinned at the source: the extracted-sync must ask the library before
    # the inbox, and the extractor must read the unpacked size in its
    # upgrade-tag branch. Both are what turned eight leftovers into 64 GB
    # of duplicate downloads.
    It "extracted-sync checks the library size index before the inbox" {
        $libraryAt = $script:extractedSyncSource.IndexOf('Find-LocalVideoBySize -Index $librarySizeIndex -Size $mainVideo.Size')
        $inboxAt   = $script:extractedSyncSource.IndexOf('skip: already in inbox')
        $libraryAt | Should -BeGreaterThan -1
        $libraryAt | Should -BeLessThan $inboxAt
        $script:extractedSyncSource | Should -Match 'LocalLibraryPaths'
    }

    It "the extractor reads a tagged release's unpacked size and treats an identical library video as already owned" {
        $script:extractorSource | Should -Match 'Get-RarArchiveUnpackedSize -Session \$session -FirstArchive \$r\.FirstArchive'
        $script:extractorSource | Should -Match 'Identical = \$true'
    }

    It "the extractor tells Radarr the chroot path, not the SFTP one" {
        $script:extractorSource | Should -Match 'Invoke-RadarrDownloadedScan[^\n]*-Path \(ConvertTo-ChrootNamespacePath -Path \$destDir\)'
    }
}

Describe "Invoke-RadarrDownloadedScan" {
    It "ties the scan to the queue item when a download id is given, and omits the field otherwise" {
        InModuleScope Sync {
            Mock Invoke-RestMethod { [PSCustomObject]@{ id = 42; status = 'queued' } }
            $r = Invoke-RadarrDownloadedScan -RadarrUrl 'http://radarr.test' -ApiKey 'k' -Path '/home/user/x.extracted' -ImportMode Copy -DownloadClientId 'ABC123'
            $r.Success   | Should -BeTrue
            $r.CommandId | Should -Be 42
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter {
                $sent = $Body | ConvertFrom-Json
                $sent.name -eq 'DownloadedMoviesScan' -and $sent.path -eq '/home/user/x.extracted' -and $sent.importMode -eq 'Copy' -and $sent.downloadClientId -eq 'ABC123'
            }
            $null = Invoke-RadarrDownloadedScan -RadarrUrl 'http://radarr.test' -ApiKey 'k' -Path '/p'
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { -not (($Body | ConvertFrom-Json).PSObject.Properties.Name -contains 'downloadClientId') }
        }
    }
}

Describe "Get-RemoteFolderRarRelease" {
    It "recognises a complete RAR chain from one directory listing" {
        InModuleScope Sync {
            $session = [PSCustomObject]@{}
            $session | Add-Member -MemberType ScriptMethod -Name ListDirectory -Value {
                param($path)
                [PSCustomObject]@{ Files = @(
                    [PSCustomObject]@{ Name = '.'; Length = 0; IsDirectory = $true }
                    [PSCustomObject]@{ Name = 'Sample'; Length = 0; IsDirectory = $true }
                    [PSCustomObject]@{ Name = 'x.nfo'; Length = 759; IsDirectory = $false }
                    [PSCustomObject]@{ Name = 'x.rar'; Length = 100000000; IsDirectory = $false }
                    [PSCustomObject]@{ Name = 'x.r00'; Length = 100000000; IsDirectory = $false }
                    [PSCustomObject]@{ Name = 'x.r01'; Length = 100000000; IsDirectory = $false }
                ) }
            }
            $release = Get-RemoteFolderRarRelease -Session $session -Folder '/home16/user/downloads/rtorrent/X/'
            $release              | Should -Not -BeNullOrEmpty
            $release.Folder       | Should -Be '/home16/user/downloads/rtorrent/X'
            $release.FirstArchive | Should -Be '/home16/user/downloads/rtorrent/X/x.rar'
            $release.PartCount    | Should -Be 3
            $release.Complete     | Should -BeTrue
        }
    }

    It "returns nothing for a folder that already holds a playable video" {
        InModuleScope Sync {
            $session = [PSCustomObject]@{}
            $session | Add-Member -MemberType ScriptMethod -Name ListDirectory -Value {
                param($path)
                [PSCustomObject]@{ Files = @([PSCustomObject]@{ Name = 'movie.mkv'; Length = 6000000000; IsDirectory = $false }) }
            }
            Get-RemoteFolderRarRelease -Session $session -Folder '/x' | Should -BeNullOrEmpty
        }
    }
}

Describe "Invoke-SFTPExtractSingleRelease" {
    BeforeEach {
        # A fake WinSCP session: the release folder holds a complete chain;
        # the .extracted/ sibling is empty until "unrar" runs. The flag is
        # global because the unrar mock body runs inside the module's scope,
        # where $script: would be a different variable.
        $global:LLTestExtractedHasVideo = $false
        $script:fakeSession = [PSCustomObject]@{}
        $script:fakeSession | Add-Member -MemberType ScriptMethod -Name ListDirectory -Value {
            param($path)
            if ($path -like '*.extracted') {
                if ($global:LLTestExtractedHasVideo) { return [PSCustomObject]@{ Files = @([PSCustomObject]@{ Name = 'x.mkv'; Length = 8529661515; IsDirectory = $false }) } }
                throw 'no such directory'
            }
            [PSCustomObject]@{ Files = @(
                [PSCustomObject]@{ Name = 'x.rar'; Length = 100000000; IsDirectory = $false }
                [PSCustomObject]@{ Name = 'x.r00'; Length = 100000000; IsDirectory = $false }
            ) }
        }
        $script:fakeSession | Add-Member -MemberType ScriptMethod -Name ExecuteCommand -Value {
            param($cmd)
            # `command -v unrar` resolves; `unrar l <archive>` lists the chain's video.
            if ($cmd -match '\bl\b') { return [PSCustomObject]@{ Output = "    ..A.... 8529661515  2008-10-10 17:02  x.mkv"; ExitStatus = 0; ErrorOutput = '' } }
            [PSCustomObject]@{ Output = '/usr/bin/unrar'; ExitStatus = 0; ErrorOutput = '' }
        }
        $script:fakeSession | Add-Member -MemberType ScriptMethod -Name Dispose -Value { }
    }

    It "extracts, records tracking, and asks Radarr to scan the chroot path tied to the download" {
        InModuleScope Sync -Parameters @{ Session = $script:fakeSession } {
            param($Session)
            Mock Test-WinSCPInstalled { 'C:\fake\WinSCPnet.dll' }
            Mock Connect-SFTPSession { $Session }
            Mock Invoke-SFTPRemoteUnrar { $global:LLTestExtractedHasVideo = $true; @{ Success = $true; ExitStatus = 0; Output = ''; ErrorOutput = '' } }
            Mock Read-RarExtractionTracking { @{ version = 1; extractions = @{} } }
            Mock Save-RarExtractionTracking { }
            Mock Invoke-RadarrDownloadedScan { @{ Success = $true; CommandId = 77 } }

            $r = Invoke-SFTPExtractSingleRelease -HostName h -Username u -Password p -ReleaseFolder '/home16/user/downloads/rtorrent/X' `
                -NotifyRadarr -RadarrUrl 'http://radarr.test' -RadarrApiKey 'k' -DownloadId 'DL1'
            $r.Success          | Should -BeTrue
            $r.AlreadyExtracted | Should -BeFalse
            $r.Video            | Should -Be 'x.mkv'
            $r.ExtractedPath    | Should -Be '/home16/user/downloads/rtorrent/X.extracted'
            $r.RadarrCommandId  | Should -Be 77
            Should -Invoke Invoke-SFTPRemoteUnrar -Times 1 -Exactly -ParameterFilter { $FirstArchive -eq '/home16/user/downloads/rtorrent/X/x.rar' -and $DestinationDir -eq '/home16/user/downloads/rtorrent/X.extracted' }
            Should -Invoke Save-RarExtractionTracking -Times 1
            Should -Invoke Invoke-RadarrDownloadedScan -Times 1 -Exactly -ParameterFilter { $Path -eq '/home/user/downloads/rtorrent/X.extracted' -and $DownloadClientId -eq 'DL1' -and $ImportMode -eq 'Copy' }
        }
    }

    It "reuses an existing extraction instead of running unrar again" {
        $global:LLTestExtractedHasVideo = $true
        InModuleScope Sync -Parameters @{ Session = $script:fakeSession } {
            param($Session)
            Mock Test-WinSCPInstalled { 'C:\fake\WinSCPnet.dll' }
            Mock Connect-SFTPSession { $Session }
            Mock Invoke-SFTPRemoteUnrar { throw 'should not run' }
            Mock Invoke-RadarrDownloadedScan { @{ Success = $true; CommandId = 78 } }
            $r = Invoke-SFTPExtractSingleRelease -HostName h -Username u -Password p -ReleaseFolder '/home16/user/downloads/rtorrent/X' -NotifyRadarr -RadarrUrl 'http://radarr.test' -RadarrApiKey 'k'
            $r.Success          | Should -BeTrue
            $r.AlreadyExtracted | Should -BeTrue
            Should -Invoke Invoke-SFTPRemoteUnrar -Times 0
            Should -Invoke Invoke-RadarrDownloadedScan -Times 1
        }
    }

    It "runs nothing under -WhatIf" {
        InModuleScope Sync -Parameters @{ Session = $script:fakeSession } {
            param($Session)
            Mock Test-WinSCPInstalled { 'C:\fake\WinSCPnet.dll' }
            Mock Connect-SFTPSession { $Session }
            Mock Invoke-SFTPRemoteUnrar { throw 'should not run' }
            Mock Invoke-RadarrDownloadedScan { throw 'should not run' }
            $r = Invoke-SFTPExtractSingleRelease -HostName h -Username u -Password p -ReleaseFolder '/home16/user/downloads/rtorrent/X' -NotifyRadarr -RadarrUrl 'http://radarr.test' -RadarrApiKey 'k' -WhatIf
            $r.Success      | Should -BeTrue
            # The preview still reads the unpacked size from the RAR headers.
            $r.UnpackedSize | Should -Be 8529661515
            Should -Invoke Invoke-SFTPRemoteUnrar -Times 0
            Should -Invoke Invoke-RadarrDownloadedScan -Times 0
        }
    }

    It "refuses an incomplete RAR set before touching unrar" {
        $broken = [PSCustomObject]@{}
        $broken | Add-Member -MemberType ScriptMethod -Name ListDirectory -Value {
            param($path)
            [PSCustomObject]@{ Files = @(
                [PSCustomObject]@{ Name = 'x.rar'; Length = 100000000; IsDirectory = $false }
                [PSCustomObject]@{ Name = 'x.r00'; Length = 0; IsDirectory = $false }
            ) }
        }
        $broken | Add-Member -MemberType ScriptMethod -Name Dispose -Value { }
        InModuleScope Sync -Parameters @{ Session = $broken } {
            param($Session)
            Mock Test-WinSCPInstalled { 'C:\fake\WinSCPnet.dll' }
            Mock Connect-SFTPSession { $Session }
            Mock Invoke-SFTPRemoteUnrar { throw 'should not run' }
            $r = Invoke-SFTPExtractSingleRelease -HostName h -Username u -Password p -ReleaseFolder '/x'
            $r.Success | Should -BeFalse
            $r.Error   | Should -Match 'incomplete'
            Should -Invoke Invoke-SFTPRemoteUnrar -Times 0
        }
    }
}
