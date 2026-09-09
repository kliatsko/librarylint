#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for hardsub self-confirmation and re-acquisition tracking
    (modules\Quality.psm1).

.DESCRIPTION
    Two behaviours, both aimed at the same failure: a rip whose own filename
    admits it is hardsubbed was being flagged every run and never acted on.

    A release-name tag confirms itself, because there is nothing for a human
    to adjudicate about a label they can read. OCR verdicts still wait for
    review, since those genuinely produce false positives.

    A confirmed hardsub then carries a re-acquisition stamp, so it is asked
    about once and tracked afterwards rather than re-prompted, and a request
    that never landed surfaces on its own instead of hiding.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\Quality.HardsubReacquisition.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\Quality.psm1') -Force

    function New-HardsubFixture {
        param(
            [string]$Root,
            [string]$Name,
            [string]$VideoName,
            [string]$OriginalFileName,
            [string]$Reviewed,
            [string]$RequestedAt
        )
        $dir = Join-Path $Root $Name
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $dir $VideoName) -Value 'video-bytes' -NoNewline
        $info = [ordered]@{ OriginalFileName = $OriginalFileName }
        if ($Reviewed -or $RequestedAt) {
            $info['HardsubAudit'] = [ordered]@{
                Verdict = 'TAGGED'; CoveragePct = 100; SamplesTaken = 0; TextFrames = 0
                Fingerprint = 'fp-placeholder'; AuditedAt = '2026-09-01 10:00:00'
                Reviewed = $Reviewed
                ReviewedAt = $(if ($Reviewed) { '2026-09-01 10:00:00' } else { $null })
                ReacquisitionRequestedAt = $RequestedAt
            }
        }
        $info | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $dir 'release-info.json') -Encoding UTF8
        return $dir
    }
}

Describe "Release-tagged rips confirm themselves" {
    BeforeAll {
        $script:tagRoot = Join-Path $TestDrive 'TagLib'
        # An HC-tagged release. The audit's pre-pass reads the tag from
        # release-info.json's OriginalFileName, so no OCR tooling is needed.
        $null = New-HardsubFixture -Root $script:tagRoot -Name 'Tagged Movie (2016)' `
            -VideoName 'Tagged Movie (2016).mkv' `
            -OriginalFileName 'Tagged.Movie.2016.1080p.HC.HDRip.X264.AC3-EVO.mkv'
        $script:tagAudit = Invoke-HardsubAudit -Path $script:tagRoot -VideoExtensions @('.mkv') -SampleCount 4
    }

    It "classifies the tag without sampling a single frame" {
        $r = $script:tagAudit.Results | Where-Object Folder -eq 'Tagged Movie (2016)'
        $r.Classification | Should -Be 'TAGGED'
        $r.SamplesTaken | Should -Be 0
        $script:tagAudit.TaggedByName | Should -Be 1
    }

    # The point of the change: a label anyone can read in the filename is not
    # a judgement call, so it must not queue up behind a human.
    It "records the verdict as already reviewed" {
        $rec = Get-HardsubAuditRecord -FolderPath (Join-Path $script:tagRoot 'Tagged Movie (2016)')
        $rec.Reviewed | Should -Be 'hardsub'
        $rec.ReviewedAt | Should -Not -BeNullOrEmpty
    }

    It "attributes the confirmation to the release tag, not a person" {
        $rec = Get-HardsubAuditRecord -FolderPath (Join-Path $script:tagRoot 'Tagged Movie (2016)')
        $rec.ReviewedBy | Should -Be 'release-tag'
    }

    It "starts with no re-acquisition request recorded" {
        $rec = Get-HardsubAuditRecord -FolderPath (Join-Path $script:tagRoot 'Tagged Movie (2016)')
        $rec.ReacquisitionRequestedAt | Should -BeNullOrEmpty
    }

    # OCR is fallible in a way a filename is not, so its verdicts still wait.
    It "leaves a sampled verdict unreviewed" {
        $sampledRoot = Join-Path $TestDrive 'SampledLib'
        $folder = New-HardsubFixture -Root $sampledRoot -Name 'Untagged Movie (2016)' `
            -VideoName 'Untagged Movie (2016).mkv' `
            -OriginalFileName 'Untagged.Movie.2016.1080p.BluRay.x264-GRP.mkv'
        # A 5-byte fake video yields no duration, so the audit records FAILED
        # rather than a coverage verdict — either way it must not self-review.
        $null = Invoke-HardsubAudit -Path $sampledRoot -VideoExtensions @('.mkv') -SampleCount 4
        (Get-HardsubAuditRecord -FolderPath $folder).Reviewed | Should -BeNullOrEmpty
    }
}

Describe "Set-HardsubReacquisitionRequested" {
    It "stamps the request and preserves the rest of the record" {
        $root = Join-Path $TestDrive 'StampLib'
        $folder = New-HardsubFixture -Root $root -Name 'Stamp Me (2016)' `
            -VideoName 'Stamp Me (2016).mkv' -OriginalFileName 'x.HC.mkv' -Reviewed 'hardsub'
        $null = Set-HardsubReacquisitionRequested -FolderPath $folder
        $rec = Get-HardsubAuditRecord -FolderPath $folder
        $rec.ReacquisitionRequestedAt | Should -Not -BeNullOrEmpty
        $rec.Reviewed | Should -Be 'hardsub'
        $rec.Verdict | Should -Be 'TAGGED'
    }

    It "preserves other release-info keys" {
        $root = Join-Path $TestDrive 'StampKeysLib'
        $folder = New-HardsubFixture -Root $root -Name 'Keep Keys (2016)' `
            -VideoName 'Keep Keys (2016).mkv' -OriginalFileName 'original.HC.mkv' -Reviewed 'hardsub'
        $null = Set-HardsubReacquisitionRequested -FolderPath $folder
        $info = Get-Content -LiteralPath (Join-Path $folder 'release-info.json') -Raw | ConvertFrom-Json
        $info.OriginalFileName | Should -Be 'original.HC.mkv'
    }

    It "writes nothing under -WhatIf" {
        $root = Join-Path $TestDrive 'StampWhatIfLib'
        $folder = New-HardsubFixture -Root $root -Name 'WhatIf Stamp (2016)' `
            -VideoName 'WhatIf Stamp (2016).mkv' -OriginalFileName 'x.HC.mkv' -Reviewed 'hardsub'
        $null = Set-HardsubReacquisitionRequested -FolderPath $folder -WhatIf
        (Get-HardsubAuditRecord -FolderPath $folder).ReacquisitionRequestedAt | Should -BeNullOrEmpty
    }
}

Describe "Get-HardsubReacquisitionStatus" {
    BeforeAll {
        $script:statusRoot = Join-Path $TestDrive 'StatusLib'
        New-Item -Path $script:statusRoot -ItemType Directory -Force | Out-Null
        # Confirmed, never sent.
        $null = New-HardsubFixture -Root $script:statusRoot -Name 'Never Sent (2016)' `
            -VideoName 'a.mkv' -OriginalFileName 'a.HC.mkv' -Reviewed 'hardsub'
        # Sent recently: waiting, not yet suspicious.
        $null = New-HardsubFixture -Root $script:statusRoot -Name 'Sent Recently (2017)' `
            -VideoName 'b.mkv' -OriginalFileName 'b.HC.mkv' -Reviewed 'hardsub' `
            -RequestedAt (Get-Date).AddDays(-3).ToString('yyyy-MM-dd HH:mm:ss')
        # Sent long ago and still here: the grab never landed.
        $null = New-HardsubFixture -Root $script:statusRoot -Name 'Sent Long Ago (2018)' `
            -VideoName 'c.mkv' -OriginalFileName 'c.HC.mkv' -Reviewed 'hardsub' `
            -RequestedAt (Get-Date).AddDays(-45).ToString('yyyy-MM-dd HH:mm:ss')
        # Reviewed clean: not a hardsub, must never appear.
        $null = New-HardsubFixture -Root $script:statusRoot -Name 'False Positive (2019)' `
            -VideoName 'd.mkv' -OriginalFileName 'd.BluRay.mkv' -Reviewed 'clean'
        # No audit record at all.
        $null = New-HardsubFixture -Root $script:statusRoot -Name 'Unaudited (2020)' `
            -VideoName 'e.mkv' -OriginalFileName 'e.BluRay.mkv'
        $script:status = Get-HardsubReacquisitionStatus -Path $script:statusRoot -VideoExtensions @('.mkv')
    }

    It "counts only confirmed hardsubs" {
        $script:status.ConfirmedTotal | Should -Be 3
    }

    It "lists a confirmed hardsub that was never sent as pending" {
        @($script:status.Pending).Count | Should -Be 1
        $script:status.Pending[0].Folder | Should -Be 'Never Sent (2016)'
    }

    It "lists a recent request as awaiting, with the wait in days" {
        @($script:status.Awaiting).Count | Should -Be 1
        $script:status.Awaiting[0].Folder | Should -Be 'Sent Recently (2017)'
        $script:status.Awaiting[0].DaysWaiting | Should -BeGreaterOrEqual 2
    }

    # This is what would have surfaced the stuck Hellboy import weeks earlier.
    It "calls out a request that has gone nowhere for 30+ days" {
        @($script:status.Stale).Count | Should -Be 1
        $script:status.Stale[0].Folder | Should -Be 'Sent Long Ago (2018)'
        $script:status.Stale[0].DaysWaiting | Should -BeGreaterOrEqual 44
    }

    It "never reports a movie reviewed as clean" {
        @($script:status.Pending + $script:status.Awaiting + $script:status.Stale |
            Where-Object { $_.Folder -eq 'False Positive (2019)' }).Count | Should -Be 0
    }

    It "honours a custom staleness threshold" {
        $s = Get-HardsubReacquisitionStatus -Path $script:statusRoot -VideoExtensions @('.mkv') -StaleAfterDays 2
        @($s.Stale).Count | Should -Be 2
        @($s.Awaiting).Count | Should -Be 0
    }

    # Records written before self-confirmation are cached against an
    # unchanged file, so they will never be re-audited to pick up the new
    # behaviour. The rule has to apply to them where they sit, or every
    # already-tagged rip in the library stays invisible.
    It "treats a pre-existing TAGGED record with no review as confirmed" {
        $root = Join-Path $TestDrive 'LegacyTaggedLib'
        $folder = New-HardsubFixture -Root $root -Name 'Legacy Tagged (2016)' `
            -VideoName 'Legacy Tagged (2016).mkv' -OriginalFileName 'x.HC.mkv'
        # Verdict TAGGED, Reviewed null: exactly what the old code wrote.
        $rec = [ordered]@{ Verdict = 'TAGGED'; CoveragePct = 100; SamplesTaken = 0; TextFrames = 0
                           Fingerprint = 'fp'; AuditedAt = '2026-09-01 10:00:00'; Reviewed = $null; ReviewedAt = $null }
        $null = Set-HardsubAuditRecord -FolderPath $folder -Record $rec
        $s = Get-HardsubReacquisitionStatus -Path $root -VideoExtensions @('.mkv')
        @($s.Pending).Count | Should -Be 1
        $s.Pending[0].Folder | Should -Be 'Legacy Tagged (2016)'
    }

    # The filename is strong evidence, not a verdict a person can't overrule.
    It "respects an explicit clean review over the release tag" {
        $root = Join-Path $TestDrive 'TaggedButCleanLib'
        $folder = New-HardsubFixture -Root $root -Name 'Tagged But Clean (2016)' `
            -VideoName 'Tagged But Clean (2016).mkv' -OriginalFileName 'x.HC.mkv'
        $rec = [ordered]@{ Verdict = 'TAGGED'; CoveragePct = 100; SamplesTaken = 0; TextFrames = 0
                           Fingerprint = 'fp'; AuditedAt = '2026-09-01 10:00:00'; Reviewed = 'clean'; ReviewedAt = '2026-09-02 10:00:00' }
        $null = Set-HardsubAuditRecord -FolderPath $folder -Record $rec
        $s = Get-HardsubReacquisitionStatus -Path $root -VideoExtensions @('.mkv')
        $s.ConfirmedTotal | Should -Be 0
        @($s.Pending).Count | Should -Be 0
    }

    It "returns empty for a path that does not exist" {
        $s = Get-HardsubReacquisitionStatus -Path (Join-Path $TestDrive 'nope') -VideoExtensions @('.mkv')
        $s.ConfirmedTotal | Should -Be 0
        @($s.Pending).Count | Should -Be 0
    }

    # A request must not be re-asked just because the audit re-ran, but a
    # genuinely different file starts the process over.
    It "keeps the request across a re-audit of the same file" {
        $root = Join-Path $TestDrive 'ReauditLib'
        $folder = New-HardsubFixture -Root $root -Name 'Reaudit Me (2016)' `
            -VideoName 'Reaudit Me (2016).mkv' -OriginalFileName 'r.HC.mkv'
        $null = Invoke-HardsubAudit -Path $root -VideoExtensions @('.mkv') -SampleCount 4
        $null = Set-HardsubReacquisitionRequested -FolderPath $folder
        $stamped = (Get-HardsubAuditRecord -FolderPath $folder).ReacquisitionRequestedAt
        $stamped | Should -Not -BeNullOrEmpty

        $null = Invoke-HardsubAudit -Path $root -VideoExtensions @('.mkv') -SampleCount 4 -Force
        (Get-HardsubAuditRecord -FolderPath $folder).ReacquisitionRequestedAt | Should -Be $stamped
    }

    It "clears the request when the video is replaced" {
        $root = Join-Path $TestDrive 'ReplacedLib'
        $folder = New-HardsubFixture -Root $root -Name 'Replaced Me (2016)' `
            -VideoName 'Replaced Me (2016).mkv' -OriginalFileName 'r.HC.mkv'
        $null = Invoke-HardsubAudit -Path $root -VideoExtensions @('.mkv') -SampleCount 4
        $null = Set-HardsubReacquisitionRequested -FolderPath $folder
        # Different bytes and length: a different release landed.
        Set-Content -LiteralPath (Join-Path $folder 'Replaced Me (2016).mkv') -Value 'a-completely-different-video-file' -NoNewline
        $null = Invoke-HardsubAudit -Path $root -VideoExtensions @('.mkv') -SampleCount 4
        (Get-HardsubAuditRecord -FolderPath $folder).ReacquisitionRequestedAt | Should -BeNullOrEmpty
    }
}
