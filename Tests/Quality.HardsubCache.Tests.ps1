#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for the hardsub audit's verdict cache (modules\Quality.psm1).

.DESCRIPTION
    Covers the machinery that stops the audit re-sampling movies it has
    already judged: content fingerprinting, the per-folder record in
    release-info.json, human review, and the backlog count.

    Scope note: these exercise the CACHE, not the OCR. Frame sampling needs
    ffmpeg, tesseract and a burned-in fixture (minutes per run), so it stays
    out of the fast suite; the cache is where the regressions actually live.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Quality.HardsubCache.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Quality.psm1') -Force

    # A movie folder holding one video file, optionally with a
    # pre-existing release-info.json.
    function New-TestMovieFolder {
        param(
            [string]$Root,
            [string]$Name,
            [string]$VideoContent = 'video-bytes',
            [hashtable]$ReleaseInfo
        )
        $dir = Join-Path $Root $Name
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $dir "$Name.mkv") -Value $VideoContent -NoNewline
        if ($ReleaseInfo) {
            $ReleaseInfo | ConvertTo-Json |
                Set-Content -LiteralPath (Join-Path $dir 'release-info.json') -Encoding UTF8
        }
        return $dir
    }

    function Get-TestVideoPath {
        param([string]$FolderPath)
        return (Get-ChildItem -LiteralPath $FolderPath -Filter '*.mkv' -File | Select-Object -First 1).FullName
    }
}

Describe "Get-VideoFingerprint" {
    It "returns a stable value for unchanged content" {
        $file = Join-Path $TestDrive 'stable.mkv'
        Set-Content -LiteralPath $file -Value 'the same bytes' -NoNewline
        $first = Get-VideoFingerprint -Path $file
        $second = Get-VideoFingerprint -Path $file
        $first | Should -Not -BeNullOrEmpty
        $second | Should -Be $first
    }

    It "changes when the content changes" {
        $file = Join-Path $TestDrive 'mutable.mkv'
        Set-Content -LiteralPath $file -Value 'original bytes' -NoNewline
        $before = Get-VideoFingerprint -Path $file
        Set-Content -LiteralPath $file -Value 'different bytes!' -NoNewline
        Get-VideoFingerprint -Path $file | Should -Not -Be $before
    }

    It "leads with the file size so the key is human-readable" {
        $file = Join-Path $TestDrive 'sized.mkv'
        Set-Content -LiteralPath $file -Value 'abcdefghij' -NoNewline
        $size = (Get-Item -LiteralPath $file).Length
        Get-VideoFingerprint -Path $file | Should -BeLike "$size-*"
    }

    It "returns null for a missing file instead of throwing" {
        { Get-VideoFingerprint -Path (Join-Path $TestDrive 'not-here.mkv') } | Should -Not -Throw
        Get-VideoFingerprint -Path (Join-Path $TestDrive 'not-here.mkv') | Should -BeNullOrEmpty
    }

    # Regression: [Math]::Min(65536, $size) bound the Int32 overload from
    # the integer literal, so any file over 2.1GB threw inside the
    # function's try/catch and returned null. Every full-size movie then
    # got a record with no cache key and re-audited on every run, while
    # sub-2GB fixtures passed happily. Sparse file keeps this test instant.
    It "fingerprints a file larger than Int32 max" {
        $big = Join-Path $TestDrive 'huge.mkv'
        $stream = [System.IO.File]::Open($big, 'CreateNew', 'Write')
        try {
            $stream.SetLength(2.5GB)
        } finally {
            $stream.Dispose()
        }
        $fingerprint = Get-VideoFingerprint -Path $big
        $fingerprint | Should -Not -BeNullOrEmpty
        $fingerprint | Should -BeLike '2684354560-*'
    }
}

Describe "Hardsub audit records in release-info.json" {
    It "round-trips a record" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'Round Trip (2001)'
        $written = Set-HardsubAuditRecord -FolderPath $folder -Record ([ordered]@{
            Verdict = 'CLEAN'; CoveragePct = 0; SamplesTaken = 48; Fingerprint = 'abc-123'
        })
        $written | Should -BeTrue
        $record = Get-HardsubAuditRecord -FolderPath $folder
        $record.Verdict | Should -Be 'CLEAN'
        $record.SamplesTaken | Should -Be 48
        $record.Fingerprint | Should -Be 'abc-123'
    }

    # The audit writes into a file other features own: the orphan-rename
    # logic reads OriginalFileName, the tag pre-pass reads release names.
    It "preserves the file's other keys" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'Keeps Keys (2002)' -ReleaseInfo @{
            OriginalFileName = 'Keeps.Keys.2002.1080p.BluRay.x264-GRP.mkv'
            Source           = 'BluRay'
            Tags             = @('1080p', 'x264')
        }
        $null = Set-HardsubAuditRecord -FolderPath $folder -Record ([ordered]@{ Verdict = 'FULL' })
        $info = Get-Content -LiteralPath (Join-Path $folder 'release-info.json') -Raw | ConvertFrom-Json
        $info.OriginalFileName | Should -Be 'Keeps.Keys.2002.1080p.BluRay.x264-GRP.mkv'
        $info.Source | Should -Be 'BluRay'
        @($info.Tags).Count | Should -Be 2
        $info.HardsubAudit.Verdict | Should -Be 'FULL'
    }

    It "creates release-info.json when the folder has none" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'No Info Yet (2003)'
        Test-Path -LiteralPath (Join-Path $folder 'release-info.json') | Should -BeFalse
        $null = Set-HardsubAuditRecord -FolderPath $folder -Record ([ordered]@{ Verdict = 'SUSPECT' })
        (Get-HardsubAuditRecord -FolderPath $folder).Verdict | Should -Be 'SUSPECT'
    }

    It "returns null for a folder with no record" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'Unaudited (2004)'
        Get-HardsubAuditRecord -FolderPath $folder | Should -BeNullOrEmpty
    }

    It "returns null when release-info.json is corrupt rather than throwing" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'Corrupt Info (2005)'
        Set-Content -LiteralPath (Join-Path $folder 'release-info.json') -Value '{not json'
        { Get-HardsubAuditRecord -FolderPath $folder } | Should -Not -Throw
        Get-HardsubAuditRecord -FolderPath $folder | Should -BeNullOrEmpty
    }
}

Describe "Set-HardsubReview" {
    It "records the verdict with a timestamp" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'Reviewed Hardsub (2006)'
        $null = Set-HardsubReview -FolderPath $folder -Verdict 'hardsub' -VideoPath (Get-TestVideoPath $folder)
        $record = Get-HardsubAuditRecord -FolderPath $folder
        $record.Reviewed | Should -Be 'hardsub'
        $record.ReviewedAt | Should -Not -BeNullOrEmpty
    }

    It "keeps the machine verdict fields already on the record" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'Keeps Verdict (2007)'
        $null = Set-HardsubAuditRecord -FolderPath $folder -Record ([ordered]@{
            Verdict = 'SUSPECT'; CoveragePct = 31; SamplesTaken = 48; Fingerprint = 'keep-me'
        })
        $null = Set-HardsubReview -FolderPath $folder -Verdict 'clean' -VideoPath (Get-TestVideoPath $folder)
        $record = Get-HardsubAuditRecord -FolderPath $folder
        $record.Reviewed | Should -Be 'clean'
        $record.Verdict | Should -Be 'SUSPECT'
        $record.CoveragePct | Should -Be 31
        $record.Fingerprint | Should -Be 'keep-me'
    }

    # A review reached from a report that predates fingerprinting still
    # needs a cache key, or the audit can't tell the review applies.
    It "fills in the fingerprint when the record has none" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'Needs Fingerprint (2008)'
        $null = Set-HardsubReview -FolderPath $folder -Verdict 'clean' -VideoPath (Get-TestVideoPath $folder)
        (Get-HardsubAuditRecord -FolderPath $folder).Fingerprint | Should -Not -BeNullOrEmpty
    }

    It "rejects a verdict outside clean/hardsub" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'Bad Verdict (2009)'
        { Set-HardsubReview -FolderPath $folder -Verdict 'maybe' -ErrorAction Stop } | Should -Throw
    }

    # House convention is that anything which writes supports -WhatIf, so
    # this is worth a real behavioural check. PowerShell writes the "What
    # if:" notice straight to the host, past every redirectable stream, so
    # one such line in the suite output is expected and not a failure.
    It "writes nothing under -WhatIf" {
        $folder = New-TestMovieFolder -Root $TestDrive -Name 'WhatIf Review (2010)'
        $null = Set-HardsubReview -FolderPath $folder -Verdict 'hardsub' -VideoPath (Get-TestVideoPath $folder) -WhatIf
        Get-HardsubAuditRecord -FolderPath $folder | Should -BeNullOrEmpty
    }
}

Describe "Get-HardsubAuditBacklog" {
    BeforeAll {
        $backlogRoot = Join-Path $TestDrive 'BacklogLibrary'
        New-Item -Path $backlogRoot -ItemType Directory -Force | Out-Null

        # Two unaudited movies.
        $null = New-TestMovieFolder -Root $backlogRoot -Name 'Unaudited One (2011)'
        $null = New-TestMovieFolder -Root $backlogRoot -Name 'Unaudited Two (2012)'

        # One already audited.
        $done = New-TestMovieFolder -Root $backlogRoot -Name 'Already Done (2013)'
        $null = Set-HardsubAuditRecord -FolderPath $done -Record ([ordered]@{ Verdict = 'CLEAN'; Fingerprint = 'x' })

        # Working folders and a video-less folder must not count.
        New-Item -Path (Join-Path $backlogRoot '_Trailers') -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $backlogRoot '_Trailers\clip.mkv') -Value 'x' -NoNewline
        New-Item -Path (Join-Path $backlogRoot 'Artwork Only (2014)') -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $backlogRoot 'Artwork Only (2014)\poster.jpg') -Value 'x' -NoNewline
    }

    It "counts only movies with a video and no verdict" {
        Get-HardsubAuditBacklog -Path $backlogRoot -VideoExtensions @('.mkv') | Should -Be 2
    }

    It "drops to zero once every movie has a record" {
        $extraRoot = Join-Path $TestDrive 'FullyAudited'
        New-Item -Path $extraRoot -ItemType Directory -Force | Out-Null
        $folder = New-TestMovieFolder -Root $extraRoot -Name 'Only Movie (2015)'
        $null = Set-HardsubAuditRecord -FolderPath $folder -Record ([ordered]@{ Verdict = 'CLEAN'; Fingerprint = 'y' })
        Get-HardsubAuditBacklog -Path $extraRoot -VideoExtensions @('.mkv') | Should -Be 0
    }

    It "returns zero for a library path that does not exist" {
        Get-HardsubAuditBacklog -Path (Join-Path $TestDrive 'nope') -VideoExtensions @('.mkv') | Should -Be 0
    }
}
