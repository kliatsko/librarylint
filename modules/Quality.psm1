# Quality.psm1 - Video quality analysis, scoring, and transcoding functions
# Extracted from LibraryLint.ps1 for modularity

# Cache version - bump this when analysis logic changes to invalidate all cached entries
$script:CodecCacheVersion = 3

# Junk-content filters shared by codec analysis and the hardsub audit.
# Single definition so the two walks can't drift apart on what counts as
# a trailer/sample/extra.
$script:JunkNameRegex   = '(?i)(^|[\.\-_\s])(trailer|sample|featurette|behind[\.\-_\s]?the[\.\-_\s]?scenes|extras?|deleted[\.\-_\s]?scenes?)($|[\.\-_\s])'
$script:JunkFolderRegex = '(?i)[/\\](trailers?|samples?|extras?|featurettes?|behind[\.\-_\s]?the[\.\-_\s]?scenes|deleted[\.\-_\s]?scenes?|bonus|extras?[\.\-_\s]?disc)[/\\]'

#region Private Helpers

function Format-QualitySize {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

function Get-CodecCacheKey {
    param([System.IO.FileInfo]$File)
    return "$($File.Name)|$($File.Length)|$($File.LastWriteTimeUtc.ToString('o'))"
}

function Add-SourceScore {
    # Source detection is filename-based in BOTH Get-QualityScore branches
    # (MediaInfo can't detect the release source), so the same chain applies
    # whether media probing succeeded or not. $Quality is a hashtable —
    # mutated in place.
    param(
        [hashtable]$Quality,
        [string]$AnalysisTextLower
    )

    if ($AnalysisTextLower -match 'remux') {
        $Quality.Source = "Remux"
        $Quality.Score += 35
        $Quality.Details += "Remux (+35)"
    }
    elseif ($AnalysisTextLower -match 'bluray|blu-ray|bdrip|brrip') {
        $Quality.Source = "BluRay"
        $Quality.Score += 30
        $Quality.Details += "BluRay (+30)"
    }
    elseif ($AnalysisTextLower -match 'web-dl|webdl') {
        $Quality.Source = "WEB-DL"
        $Quality.Score += 25
        $Quality.Details += "WEB-DL (+25)"
    }
    elseif ($AnalysisTextLower -match 'webrip') {
        $Quality.Source = "WEBRip"
        $Quality.Score += 20
        $Quality.Details += "WEBRip (+20)"
    }
    elseif ($AnalysisTextLower -match 'hdtv') {
        $Quality.Source = "HDTV"
        $Quality.Score += 15
        $Quality.Details += "HDTV (+15)"
    }
    elseif ($AnalysisTextLower -match 'dvdrip') {
        $Quality.Source = "DVDRip"
        $Quality.Score += 10
        $Quality.Details += "DVDRip (+10)"
    }
    elseif ($AnalysisTextLower -match 'hdrip') {
        $Quality.Source = "HDRip"
        $Quality.Score += 8
        $Quality.Details += "HDRip (+8)"
    }
}

#endregion

#region Public Functions

<#
.SYNOPSIS
    Checks whether a movie folder is marked as "best available quality" —
    LibraryLint should skip flagging it as low-quality.
.DESCRIPTION
    Some movies (older animations, films lost to time, indie releases) just
    aren't available in higher quality. The .quality_ok marker file in a
    movie folder signals "this is as good as it gets, stop nagging me."
    Three surfaces consult this marker: Radarr re-acquisition, codec
    analysis's lowest-quality ranking, and codec analysis's quality
    concerns flagging.
.PARAMETER FolderPath
    Path to the movie folder.
.OUTPUTS
    Boolean.
#>
function Test-QualityAccepted {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )
    $marker = Join-Path $FolderPath ".quality_ok"
    return (Test-Path -LiteralPath $marker)
}

<#
.SYNOPSIS
    Marks a movie folder as "best available quality" with an optional reason.
.PARAMETER FolderPath
    Path to the movie folder.
.PARAMETER Reason
    Free-text note explaining why this version is acceptable (e.g., "DVD-
    only release", "no Blu-ray exists", "fan restoration is best available").
    Stored in the marker so future-you knows why this was accepted.
.OUTPUTS
    Boolean (write success).
#>
function Set-QualityAccepted {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,
        [string]$Reason = "best-available"
    )
    $marker = Join-Path $FolderPath ".quality_ok"
    $content = @{
        Marker     = "best-available"
        Reason     = $Reason
        AcceptedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        AcceptedBy = "LibraryLint"
    } | ConvertTo-Json

    try {
        $content | Out-File -LiteralPath $marker -Encoding UTF8 -Force
        return $true
    } catch {
        Write-Host "Failed to create .quality_ok file: $_" -ForegroundColor Yellow
        return $false
    }
}

<#
.SYNOPSIS
    Removes the .quality_ok marker so the folder shows up in low-quality
    flagging again.
#>
function Remove-QualityAccepted {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )
    $marker = Join-Path $FolderPath ".quality_ok"
    if (Test-Path -LiteralPath $marker) {
        Remove-Item -LiteralPath $marker -Force -ErrorAction SilentlyContinue
    }
}

<#
.SYNOPSIS
    Walks a library and returns the quality-acceptance state of every
    movie folder.
.DESCRIPTION
    Movie-shaped folders only — a video-bearing folder at any depth. Each
    row carries the marker's metadata when present so the menu's listing
    UI can show reason + acceptance date.
.PARAMETER Path
    Library root to walk.
.PARAMETER VideoExtensions
    Recognized video extensions used to identify video-bearing folders.
.OUTPUTS
    Array of PSCustomObject: FolderPath, RelativePath, IsAccepted, Reason,
    AcceptedAt, AcceptedBy.
#>
function Get-QualityAcceptedStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [string[]]$VideoExtensions = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov')
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    $videoDirs = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $VideoExtensions -contains $_.Extension.ToLower() } |
        ForEach-Object { [void]$videoDirs.Add($_.DirectoryName) }

    $results = New-Object 'System.Collections.Generic.List[object]'
    foreach ($folder in $videoDirs) {
        $marker = Join-Path $folder ".quality_ok"
        $isAccepted = $false
        $reason = $null
        $acceptedAt = $null
        $acceptedBy = $null

        if (Test-Path -LiteralPath $marker) {
            $isAccepted = $true
            try {
                $data = Get-Content -LiteralPath $marker -Raw -ErrorAction Stop | ConvertFrom-Json
                $reason     = $data.Reason
                $acceptedAt = $data.AcceptedAt
                $acceptedBy = $data.AcceptedBy
            } catch {
                # Marker present but unreadable JSON — still counts as
                # accepted; leave metadata fields null.
            }
        }

        $rel = if ($folder.Length -gt $Path.Length) {
            $folder.Substring($Path.Length).TrimStart('\','/')
        } else { '' }

        $results.Add([PSCustomObject]@{
            FolderPath   = $folder
            RelativePath = $rel
            IsAccepted   = $isAccepted
            Reason       = $reason
            AcceptedAt   = $acceptedAt
            AcceptedBy   = $acceptedBy
        })
    }

    return @($results | Sort-Object RelativePath)
}

function Get-QualityConcerns {
    param(
        [hashtable]$Quality,
        [string]$FileName,
        [string]$ReleaseInfo = $null
    )

    $concerns = @()
    # Combine filename and release info for analysis
    $analysisText = $FileName
    if ($ReleaseInfo) {
        $analysisText = "$FileName $ReleaseInfo"
    }
    $analysisTextLower = $analysisText.ToLower()

    # Calculate bitrate in Mbps
    $bitrateMbps = if ($Quality.Bitrate -gt 0) { [math]::Round($Quality.Bitrate / 1000000, 2) } else { 0 }

    # HEVC/x265 is ~40-50% more efficient than x264 at the same quality, so lower thresholds
    $isHEVC = $Quality.Codec -match 'x265|HEVC|H\.?265' -or $analysisTextLower -match 'x265|hevc|h\.?265'

    # Resolution-based bitrate thresholds (adjusted for codec efficiency)
    if ($bitrateMbps -gt 0) {
        switch ($Quality.Resolution) {
            "2160p" {
                $veryLow = if ($isHEVC) { 5 } else { 8 }
                $low = if ($isHEVC) { 8 } else { 12 }
                $expect = if ($isHEVC) { "10+" } else { "15+" }
                if ($bitrateMbps -lt $veryLow) {
                    $concerns += "Very low bitrate for 4K (${bitrateMbps} Mbps, expect ${expect} Mbps)"
                } elseif ($bitrateMbps -lt $low) {
                    $concerns += "Low bitrate for 4K (${bitrateMbps} Mbps, expect ${expect} Mbps)"
                }
            }
            "1080p" {
                $veryLow = if ($isHEVC) { 1 } else { 2 }
                $low = if ($isHEVC) { 1.5 } else { 3 }
                $expect = if ($isHEVC) { "2.5+" } else { "5+" }
                if ($bitrateMbps -lt $veryLow) {
                    $concerns += "Very low bitrate for 1080p (${bitrateMbps} Mbps, expect ${expect} Mbps)"
                } elseif ($bitrateMbps -lt $low) {
                    $concerns += "Low bitrate for 1080p (${bitrateMbps} Mbps, expect ${expect} Mbps)"
                }
            }
            "720p" {
                $veryLow = if ($isHEVC) { 0.5 } else { 1 }
                if ($bitrateMbps -lt $veryLow) {
                    $concerns += "Very low bitrate for 720p (${bitrateMbps} Mbps, expect $(if ($isHEVC) { '1+' } else { '2+' }) Mbps)"
                }
            }
            "480p" {
                $veryLow = if ($isHEVC) { 0.3 } else { 0.6 }
                if ($bitrateMbps -lt $veryLow) {
                    $concerns += "Very low bitrate for 480p (${bitrateMbps} Mbps, expect $(if ($isHEVC) { '0.6+' } else { '1+' }) Mbps)"
                }
            }
        }
    }

    # Hardcoded-subtitle release tags. KORSUB is the classic case (early
    # rips from Korean retail sources with burned-in Korean subs — the
    # infamous Deadpool 2016 KORSUB rip); HC/HARDCODED/HARDSUB mark the
    # same defect generically. These are full-burn defects, not forced-
    # narrative subs — the whole film carries burned text on (usually) a
    # worse source, so the release is a re-acquisition candidate no
    # matter how good its bitrate looks. Token-bounded to avoid matching
    # inside words ("hchd", titles containing "hc"). Note: SUBBED/SUBS
    # tags are NOT flagged — those usually mean soft subs.
    if ($analysisTextLower -match '(^|[\.\-_\s\[\(])(korsub|hc|hardcoded|hardsub(bed)?)([\.\-_\s\]\)]|$)') {
        $concerns += "Hardcoded subtitles (KORSUB/HC tag) - burned-in subs, re-acquisition candidate"
    }

    # Sub-HD resolution is a re-acquisition candidate in its own right.
    # Without this, a clean-bitrate 480p or 360p file generates zero
    # concerns and never lands in the "Quality Concerns" report — it
    # only ranks low in the score-sorted "Lowest Quality Files" top-20,
    # which can miss titles when the library has many similar-tier files.
    # Skip "Unknown" so weird metadata doesn't flood the report.
    if ($Quality.Resolution -match '^(\d+)p$' -and [int]$Matches[1] -lt 720) {
        $concerns += "Sub-HD resolution ($($Quality.Resolution)) - re-acquisition candidate"
    }

    # File size vs duration check (if we have both)
    if ($Quality.FileSize -gt 0 -and $Quality.Duration -gt 0) {
        $durationMinutes = $Quality.Duration / 60000
        $fileSizeGB = $Quality.FileSize / 1GB

        $minSizePerMovie = if ($isHEVC) { 0.6 } else { 1 }
        $minSize4K = if ($isHEVC) { 3 } else { 5 }

        if ($durationMinutes -ge 80 -and $fileSizeGB -lt $minSizePerMovie) {
            $concerns += "Small file for runtime ($([math]::Round($fileSizeGB, 2)) GB for $([math]::Round($durationMinutes, 0)) min)"
        }

        if ($Quality.Resolution -eq "2160p" -and $durationMinutes -ge 80 -and $fileSizeGB -lt $minSize4K) {
            $concerns += "Small 4K file ($([math]::Round($fileSizeGB, 2)) GB for $([math]::Round($durationMinutes, 0)) min, expect ${minSize4K}+ GB)"
        }
    }

    # Source/bitrate mismatch
    if ($bitrateMbps -gt 0) {
        $isEncode = $analysisTextLower -match 'x264|x265|hevc|h\.?264|h\.?265|avc|web-?dl|web-?rip|hdrip|dvdrip|brrip|yify|yts|rarbg|\d{3,4}p'

        if ($analysisTextLower -match 'remux' -and $bitrateMbps -lt 15) {
            $concerns += "Low bitrate for Remux (${bitrateMbps} Mbps, expect 20+ Mbps)"
        }
        elseif (-not $isEncode -and $analysisTextLower -match 'bluray|blu-ray' -and $bitrateMbps -lt 10) {
            $concerns += "Low bitrate for apparent BluRay rip (${bitrateMbps} Mbps) - if this is an encode, ignore"
        }
    }

    # File size only check
    if ($Quality.FileSize -gt 0 -and $Quality.Duration -eq 0) {
        $fileSizeGB = $Quality.FileSize / 1GB
        $minSize1080 = if ($isHEVC) { 0.4 } else { 0.7 }
        $minSize4K = if ($isHEVC) { 1.2 } else { 2 }

        if ($Quality.Resolution -eq "1080p" -and $fileSizeGB -lt $minSize1080) {
            $concerns += "Very small 1080p file ($([math]::Round($fileSizeGB, 2)) GB)"
        }
        elseif ($Quality.Resolution -eq "2160p" -and $fileSizeGB -lt $minSize4K) {
            $concerns += "Very small 4K file ($([math]::Round($fileSizeGB, 2)) GB)"
        }
    }

    # Known low-quality release groups
    if ($analysisTextLower -match '[\-\.]?(yify|yts)[\.\-\s]') {
        $concerns += "YIFY/YTS release - Known for aggressive compression and low bitrates"
    }

    # Screener/pre-release detection
    if ($analysisTextLower -match '\b(dvdscr|hdscr|bdscr|webscr|screener)\b') {
        $concerns += "SCREENER - Pre-release copy, likely lower quality"
    }
    elseif ($analysisTextLower -match '\b(cam|camrip|hdcam)\b') {
        $concerns += "CAM - Recorded in theater, very poor quality"
    }
    elseif ($analysisTextLower -match '\b(ts|telesync|hdts)\b') {
        $concerns += "TELESYNC - Pre-release, poor audio/video quality"
    }
    elseif ($analysisTextLower -match '\b(tc|telecine)\b') {
        $concerns += "TELECINE - Pre-release, subpar quality"
    }
    elseif ($analysisTextLower -match '\b(r5|r6)\b') {
        $concerns += "R5/R6 - Region 5/6 early release, often lower quality"
    }
    elseif ($analysisTextLower -match '\b(workprint|wp)\b') {
        $concerns += "WORKPRINT - Unfinished version, missing effects/scenes"
    }

    return $concerns
}

<#
.SYNOPSIS
    Calculates a quality score for a video file based on its properties
.PARAMETER FileName
    The filename to analyze
.PARAMETER FilePath
    Path to the video file (used for file size lookup)
.PARAMETER MediaInfo
    Pre-computed MediaInfo hashtable (from Get-MediaInfoDetails in main script)
.PARAMETER ReleaseInfoText
    Pre-computed release info text (from Read-ReleaseInfo in main script)
.PARAMETER HDRFormat
    Pre-computed HDR format string (from Get-MediaInfoHDRFormat in main script)
.OUTPUTS
    Hashtable with Score, Resolution, Codec, Source, and details
#>
function Get-QualityScore {
    param(
        [string]$FileName,
        [string]$FilePath = $null,
        [hashtable]$MediaInfo = $null,
        [string]$ReleaseInfoText = $null,
        [string]$HDRFormat = $null
    )

    $quality = @{
        Score = 0
        Resolution = "Unknown"
        Codec = "Unknown"
        Source = "Unknown"
        Audio = "Unknown"
        HDR = $false
        HDRFormat = $null
        Bitrate = 0
        Width = 0
        Height = 0
        AudioChannels = 0
        Details = @()
        DataSource = "Filename"
        QualityConcerns = @()
        Duration = 0
        FileSize = 0
        ReleaseInfo = $null
        ReleaseGroup = $null
        StreamingService = $null
    }

    # Use pre-computed release info if provided
    $releaseInfoText = $ReleaseInfoText
    $quality.ReleaseInfo = $releaseInfoText

    # Combine filename and release info for analysis
    $analysisText = $FileName
    if ($releaseInfoText) {
        $analysisText = "$FileName $releaseInfoText"
    }
    $analysisTextLower = $analysisText.ToLower()

    # Detect streaming service from release info
    if ($analysisTextLower -match '\b(dsnp|disneyplus|disney\+)\b') {
        $quality.StreamingService = "Disney+"
    } elseif ($analysisTextLower -match '\b(amzn|amazon)\b') {
        $quality.StreamingService = "Amazon"
    } elseif ($analysisTextLower -match '\b(nf|netflix)\b') {
        $quality.StreamingService = "Netflix"
    } elseif ($analysisTextLower -match '\b(hmax|hbomax|max)\b') {
        $quality.StreamingService = "HBO Max"
    } elseif ($analysisTextLower -match '\b(atvp|appletv|atv)\b') {
        $quality.StreamingService = "Apple TV+"
    } elseif ($analysisTextLower -match '\b(pcok|peacock)\b') {
        $quality.StreamingService = "Peacock"
    } elseif ($analysisTextLower -match '\b(pmtp|paramount)\b') {
        $quality.StreamingService = "Paramount+"
    } elseif ($analysisTextLower -match '\b(hulu)\b') {
        $quality.StreamingService = "Hulu"
    }

    # Detect release group (typically at end after hyphen)
    if ($analysisText -match '-([A-Za-z0-9]+)(?:\.[^.]+)?$') {
        $quality.ReleaseGroup = $Matches[1]
    }

    # Use pre-computed MediaInfo if provided
    $mediaInfo = $MediaInfo

    if ($mediaInfo) {
        $quality.DataSource = "MediaInfo"
        $quality.Width = $mediaInfo.Width
        $quality.Height = $mediaInfo.Height
        $quality.Bitrate = $mediaInfo.Bitrate
        $quality.AudioChannels = $mediaInfo.AudioChannels
        $quality.Duration = $mediaInfo.Duration

        # Get file size
        if ($FilePath -and (Test-Path -LiteralPath $FilePath)) {
            $quality.FileSize = (Get-Item -LiteralPath $FilePath).Length
        }

        # Resolution from actual dimensions. Bucket by the LARGER of width
        # or height so wide-aspect releases (2.39:1, 2.76:1) still classify
        # at their mastered tier — e.g. The Creator (1920x696 / 2.76:1) is
        # 1080p even though height alone would put it in the 480p bucket.
        # Symmetric for tall aspects (4:3 1080p = 1440x1080) where height is
        # the deciding signal.
        $w = $mediaInfo.Width
        $h = $mediaInfo.Height
        if ($w -ge 3840 -or $h -ge 2160) {
            $quality.Resolution = "2160p"
            $quality.Score += 100
            $quality.Details += "4K/2160p [MediaInfo: ${w}x${h}] (+100)"
        }
        elseif ($w -ge 1920 -or $h -ge 1080) {
            $quality.Resolution = "1080p"
            $quality.Score += 80
            $quality.Details += "1080p [MediaInfo: ${w}x${h}] (+80)"
        }
        elseif ($w -ge 1280 -or $h -ge 720) {
            $quality.Resolution = "720p"
            $quality.Score += 60
            $quality.Details += "720p [MediaInfo: ${w}x${h}] (+60)"
        }
        elseif ($w -ge 720 -or $h -ge 480) {
            $quality.Resolution = "480p"
            $quality.Score += 40
            $quality.Details += "480p [MediaInfo: ${w}x${h}] (+40)"
        }
        elseif ($h -gt 0) {
            $quality.Resolution = "${h}p"
            $quality.Score += 20
            $quality.Details += "${h}p [MediaInfo: ${w}x${h}] (+20)"
        }

        # Video codec from MediaInfo
        $videoCodec = $mediaInfo.VideoCodec
        if ($videoCodec) {
            switch -Regex ($videoCodec) {
                'HEVC|H\.?265|V_MPEGH' {
                    $quality.Codec = "HEVC/x265"
                    $quality.Score += 20
                    $quality.Details += "HEVC/x265 [MediaInfo: $videoCodec] (+20)"
                }
                'AVC|H\.?264|V_MPEG4/ISO/AVC' {
                    $quality.Codec = "x264"
                    $quality.Score += 15
                    $quality.Details += "x264 [MediaInfo: $videoCodec] (+15)"
                }
                'AV1' {
                    $quality.Codec = "AV1"
                    $quality.Score += 25
                    $quality.Details += "AV1 [MediaInfo] (+25)"
                }
                'VP9' {
                    $quality.Codec = "VP9"
                    $quality.Score += 18
                    $quality.Details += "VP9 [MediaInfo] (+18)"
                }
                'MPEG-4|DivX|XviD' {
                    $quality.Codec = "XviD"
                    $quality.Score += 5
                    $quality.Details += "MPEG-4/XviD [MediaInfo: $videoCodec] (+5)"
                }
                'VC-1|WMV' {
                    $quality.Codec = "VC-1"
                    $quality.Score += 8
                    $quality.Details += "VC-1 [MediaInfo] (+8)"
                }
                default {
                    $quality.Codec = $videoCodec
                    $quality.Score += 10
                    $quality.Details += "$videoCodec [MediaInfo] (+10)"
                }
            }
        }

        # Audio codec from MediaInfo
        $audioCodec = $mediaInfo.AudioCodec
        $channels = $mediaInfo.AudioChannels
        if ($audioCodec) {
            $channelInfo = if ($channels -gt 0) { " ${channels}ch" } else { "" }
            switch -Regex ($audioCodec) {
                'Atmos|E-AC-3.*Atmos|TrueHD.*Atmos' {
                    $quality.Audio = "Atmos"
                    $quality.Score += 15
                    $quality.Details += "Atmos [MediaInfo$channelInfo] (+15)"
                }
                'TrueHD' {
                    $quality.Audio = "TrueHD"
                    $quality.Score += 12
                    $quality.Details += "TrueHD [MediaInfo$channelInfo] (+12)"
                }
                'DTS-HD|DTS.*HD' {
                    $quality.Audio = "DTS-HD"
                    $quality.Score += 10
                    $quality.Details += "DTS-HD [MediaInfo$channelInfo] (+10)"
                }
                'DTS.*X|DTS:X' {
                    $quality.Audio = "DTS:X"
                    $quality.Score += 14
                    $quality.Details += "DTS:X [MediaInfo$channelInfo] (+14)"
                }
                '^DTS$|^DTS\s' {
                    $quality.Audio = "DTS"
                    $quality.Score += 8
                    $quality.Details += "DTS [MediaInfo$channelInfo] (+8)"
                }
                'E-AC-3|EAC3|DD\+|Dolby Digital Plus' {
                    $quality.Audio = "EAC3"
                    $quality.Score += 7
                    $quality.Details += "EAC3/DD+ [MediaInfo$channelInfo] (+7)"
                }
                'AC-3|AC3|Dolby Digital' {
                    $quality.Audio = "AC3"
                    $quality.Score += 5
                    $quality.Details += "AC3 [MediaInfo$channelInfo] (+5)"
                }
                'AAC' {
                    $quality.Audio = "AAC"
                    $quality.Score += 3
                    $quality.Details += "AAC [MediaInfo$channelInfo] (+3)"
                }
                'FLAC' {
                    $quality.Audio = "FLAC"
                    $quality.Score += 6
                    $quality.Details += "FLAC [MediaInfo$channelInfo] (+6)"
                }
                'PCM|LPCM' {
                    $quality.Audio = "PCM"
                    $quality.Score += 4
                    $quality.Details += "PCM [MediaInfo$channelInfo] (+4)"
                }
                'Opus' {
                    $quality.Audio = "Opus"
                    $quality.Score += 4
                    $quality.Details += "Opus [MediaInfo$channelInfo] (+4)"
                }
                'Vorbis' {
                    $quality.Audio = "Vorbis"
                    $quality.Score += 2
                    $quality.Details += "Vorbis [MediaInfo$channelInfo] (+2)"
                }
                'MP3|MPEG Audio' {
                    $quality.Audio = "MP3"
                    $quality.Score += 1
                    $quality.Details += "MP3 [MediaInfo$channelInfo] (+1)"
                }
                default {
                    $quality.Audio = $audioCodec
                    $quality.Score += 2
                    $quality.Details += "$audioCodec [MediaInfo$channelInfo] (+2)"
                }
            }
        }

        # HDR detection from MediaInfo
        if ($mediaInfo.HDR) {
            $quality.HDR = $true
            # Use pre-computed HDR format if available
            if ($HDRFormat) {
                $quality.HDRFormat = $HDRFormat
                switch ($HDRFormat) {
                    "Dolby Vision" {
                        $quality.Score += 18
                        $quality.Details += "Dolby Vision [MediaInfo] (+18)"
                    }
                    "HDR10+" {
                        $quality.Score += 16
                        $quality.Details += "HDR10+ [MediaInfo] (+16)"
                    }
                    "HDR10" {
                        $quality.Score += 12
                        $quality.Details += "HDR10 [MediaInfo] (+12)"
                    }
                    "HLG" {
                        $quality.Score += 10
                        $quality.Details += "HLG [MediaInfo] (+10)"
                    }
                    default {
                        $quality.Score += 10
                        $quality.Details += "HDR [MediaInfo] (+10)"
                    }
                }
            }
            else {
                $quality.Score += 10
                $quality.Details += "HDR [MediaInfo] (+10)"
            }
        }

        # Source detection from filename and release info (MediaInfo can't detect source)
        Add-SourceScore -Quality $quality -AnalysisTextLower $analysisTextLower

        # Streaming service bonus (known high-quality sources)
        if ($quality.StreamingService) {
            $quality.Score += 5
            $quality.Details += "$($quality.StreamingService) source (+5)"
        }

        # Known low-quality release groups (penalty)
        if ($quality.ReleaseGroup -and $quality.ReleaseGroup -match '^(YIFY|YTS|RARBG|EVO|FGT)$') {
            $quality.Score -= 15
            $quality.Details += "Known low-bitrate group [$($quality.ReleaseGroup)] (-15)"
        }

        # Bitrate bonus (higher bitrate = better quality)
        if ($quality.Bitrate -gt 0) {
            $bitrateMbps = [math]::Round($quality.Bitrate / 1000000, 1)
            if ($bitrateMbps -ge 40) {
                $quality.Score += 20
                $quality.Details += "High Bitrate [${bitrateMbps} Mbps] (+20)"
            }
            elseif ($bitrateMbps -ge 20) {
                $quality.Score += 15
                $quality.Details += "Good Bitrate [${bitrateMbps} Mbps] (+15)"
            }
            elseif ($bitrateMbps -ge 10) {
                $quality.Score += 10
                $quality.Details += "Moderate Bitrate [${bitrateMbps} Mbps] (+10)"
            }
            elseif ($bitrateMbps -ge 5) {
                $quality.Score += 5
                $quality.Details += "Low Bitrate [${bitrateMbps} Mbps] (+5)"
            }
        }

        # Quality Concerns detection (MediaInfo path)
        $quality.QualityConcerns = @(Get-QualityConcerns -Quality $quality -FileName $FileName -ReleaseInfo $releaseInfoText)

        return $quality
    }

    # Fallback to filename parsing if MediaInfo not available
    $quality.DataSource = "Filename"

    # Resolution scoring (use combined analysis text)
    if ($analysisTextLower -match '2160p|4k|uhd') {
        $quality.Resolution = "2160p"
        $quality.Score += 100
        $quality.Details += "4K/2160p (+100)"
    }
    elseif ($analysisTextLower -match '1080p') {
        $quality.Resolution = "1080p"
        $quality.Score += 80
        $quality.Details += "1080p (+80)"
    }
    elseif ($analysisTextLower -match '720p') {
        $quality.Resolution = "720p"
        $quality.Score += 60
        $quality.Details += "720p (+60)"
    }
    elseif ($analysisTextLower -match '480p|dvd') {
        $quality.Resolution = "480p"
        $quality.Score += 40
        $quality.Details += "480p (+40)"
    }

    # Source scoring
    Add-SourceScore -Quality $quality -AnalysisTextLower $analysisTextLower

    # Codec scoring
    if ($analysisTextLower -match 'av1') {
        $quality.Codec = "AV1"
        $quality.Score += 25
        $quality.Details += "AV1 (+25)"
    }
    elseif ($analysisTextLower -match 'x265|h\.?265|hevc') {
        $quality.Codec = "HEVC/x265"
        $quality.Score += 20
        $quality.Details += "HEVC/x265 (+20)"
    }
    elseif ($analysisTextLower -match 'vp9') {
        $quality.Codec = "VP9"
        $quality.Score += 18
        $quality.Details += "VP9 (+18)"
    }
    elseif ($analysisTextLower -match 'x264|h\.?264|avc') {
        $quality.Codec = "x264"
        $quality.Score += 15
        $quality.Details += "x264 (+15)"
    }
    elseif ($analysisTextLower -match 'xvid|divx') {
        $quality.Codec = "XviD"
        $quality.Score += 5
        $quality.Details += "XviD (+5)"
    }

    # Audio scoring
    if ($analysisTextLower -match 'atmos') {
        $quality.Audio = "Atmos"
        $quality.Score += 15
        $quality.Details += "Atmos (+15)"
    }
    elseif ($analysisTextLower -match 'dts[\s\.\-]?x|dtsx') {
        $quality.Audio = "DTS:X"
        $quality.Score += 14
        $quality.Details += "DTS:X (+14)"
    }
    elseif ($analysisTextLower -match 'truehd') {
        $quality.Audio = "TrueHD"
        $quality.Score += 12
        $quality.Details += "TrueHD (+12)"
    }
    elseif ($analysisTextLower -match 'dts-hd|dtshd|dts[\s\.\-]?hd[\s\.\-]?ma') {
        $quality.Audio = "DTS-HD"
        $quality.Score += 10
        $quality.Details += "DTS-HD (+10)"
    }
    elseif ($analysisTextLower -match 'dts') {
        $quality.Audio = "DTS"
        $quality.Score += 8
        $quality.Details += "DTS (+8)"
    }
    elseif ($analysisTextLower -match 'eac3|ddp|dd\+|dolby\s*digital\s*plus') {
        $quality.Audio = "EAC3"
        $quality.Score += 7
        $quality.Details += "EAC3/DD+ (+7)"
    }
    elseif ($analysisTextLower -match 'ac3|dd5\.?1') {
        $quality.Audio = "AC3"
        $quality.Score += 5
        $quality.Details += "AC3 (+5)"
    }
    elseif ($analysisTextLower -match 'flac') {
        $quality.Audio = "FLAC"
        $quality.Score += 6
        $quality.Details += "FLAC (+6)"
    }
    elseif ($analysisTextLower -match 'aac') {
        $quality.Audio = "AAC"
        $quality.Score += 3
        $quality.Details += "AAC (+3)"
    }
    elseif ($analysisTextLower -match 'opus') {
        $quality.Audio = "Opus"
        $quality.Score += 4
        $quality.Details += "Opus (+4)"
    }

    # HDR scoring
    if ($analysisTextLower -match 'dolby[\s\.\-]?vision|dovi|dv[\s\.\-]hdr|\.dv\.') {
        $quality.HDR = $true
        $quality.HDRFormat = "Dolby Vision"
        $quality.Score += 18
        $quality.Details += "Dolby Vision (+18)"
    }
    elseif ($analysisTextLower -match 'hdr10\+|hdr10plus') {
        $quality.HDR = $true
        $quality.HDRFormat = "HDR10+"
        $quality.Score += 16
        $quality.Details += "HDR10+ (+16)"
    }
    elseif ($analysisTextLower -match 'hdr10') {
        $quality.HDR = $true
        $quality.HDRFormat = "HDR10"
        $quality.Score += 12
        $quality.Details += "HDR10 (+12)"
    }
    elseif ($analysisTextLower -match 'hlg') {
        $quality.HDR = $true
        $quality.HDRFormat = "HLG"
        $quality.Score += 10
        $quality.Details += "HLG (+10)"
    }
    elseif ($analysisTextLower -match 'hdr') {
        $quality.HDR = $true
        $quality.HDRFormat = "HDR"
        $quality.Score += 10
        $quality.Details += "HDR (+10)"
    }

    # Streaming service bonus (known high-quality sources)
    if ($quality.StreamingService) {
        $quality.Score += 5
        $quality.Details += "$($quality.StreamingService) source (+5)"
    }

    # Known low-quality release groups (penalty)
    if ($quality.ReleaseGroup -and $quality.ReleaseGroup -match '^(YIFY|YTS|RARBG|EVO|FGT)$') {
        $quality.Score -= 15
        $quality.Details += "Known low-bitrate group [$($quality.ReleaseGroup)] (-15)"
    }

    # Get file size for filename-based analysis
    if ($FilePath -and (Test-Path -LiteralPath $FilePath)) {
        $quality.FileSize = (Get-Item -LiteralPath $FilePath).Length
    }

    # Quality Concerns detection (Filename path - limited without MediaInfo)
    $quality.QualityConcerns = @(Get-QualityConcerns -Quality $quality -FileName $FileName -ReleaseInfo $releaseInfoText)

    return $quality
}

function Get-VideoCodecInfo {
    param(
        [string]$FilePath,
        [scriptblock]$QualityScorer = $null
    )

    $info = @{
        FileName = [System.IO.Path]::GetFileName($FilePath)
        FileSize = 0
        VideoCodec = "Unknown"
        AudioCodec = "Unknown"
        Resolution = "Unknown"
        HDR = $false
        Container = "Unknown"
        NeedsTranscode = $false
        TranscodeReason = @()
    }

    try {
        $file = Get-Item $FilePath -ErrorAction Stop
        $info.FileSize = $file.Length
        $info.Container = $file.Extension.TrimStart('.').ToUpper()

        # Get quality info using MediaInfo when available
        $quality = if ($QualityScorer) { & $QualityScorer $file.Name $file.FullName } else { Get-QualityScore -FileName $file.Name -FilePath $file.FullName }
        $info.Resolution = $quality.Resolution
        $info.VideoCodec = $quality.Codec
        $info.AudioCodec = $quality.Audio
        $info.HDR = $quality.HDR

        # Determine processing needed based on codec and container
        $info.TranscodeMode = "none"

        if ($info.VideoCodec -eq "XviD" -or $info.VideoCodec -eq "DivX" -or $info.VideoCodec -eq "MPEG-4") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "transcode"
            $info.TranscodeReason += "XviD/DivX/MPEG-4 is a legacy codec - will transcode to H.264"
        }
        elseif ($info.VideoCodec -eq "VC-1" -or $info.VideoCodec -eq "WMV") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "transcode"
            $info.TranscodeReason += "VC-1/WMV is a legacy codec - will transcode to H.265"
        }
        elseif (($info.VideoCodec -eq "H264" -or $info.VideoCodec -eq "AVC" -or $info.VideoCodec -eq "H.264") -and $info.Container -eq "AVI") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "remux"
            $info.TranscodeReason += "H.264 in AVI container - will remux to MKV (no re-encoding)"
        }
        elseif ($info.Container -eq "AVI") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "transcode"
            $info.TranscodeReason += "AVI with legacy codec - will transcode to H.264"
        }

        return $info
    }
    catch {
        Write-Host "  Error analyzing file $FilePath : $_" -ForegroundColor Red
        return $info
    }
}

function Invoke-CodecAnalysis {
    param(
        [string]$Path,
        [string]$ExportPath = $null,
        [string[]]$VideoExtensions,
        [string]$ReportsFolder,
        [scriptblock]$QualityScorer = $null,
        [string]$CachePath = $null,
        [switch]$ForceRescan
    )

    Write-Host "`n=== CODEC ANALYSIS ===" -ForegroundColor Cyan

    Write-Host "  Starting codec analysis for: $Path" -ForegroundColor Gray

    try {
        # Find all video files. Exclude trailers, samples, and extras — they
        # have low bitrates/resolutions by design and pollute the lowest-quality
        # report if surfaced as "movie files". The old `-trailer\.` regex only
        # caught dash-separated names like `Movie-trailer.mkv`; this widened
        # pattern also catches dot-separated (`.trailer.`), underscore, or bare
        # `trailer.mkv`, plus samples/featurettes/extras/behind-the-scenes.
        # Also drop anything inside a `Trailers/`, `Sample/`, `Extras/`,
        # `Featurettes/`, or `Behind The Scenes/` subfolder — those exist
        # specifically to segregate non-feature content.
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and
                           $_.Name -notmatch $script:JunkNameRegex -and
                           $_.FullName -notmatch $script:JunkFolderRegex }

        if ($videoFiles.Count -eq 0) {
            Write-Host "No video files found" -ForegroundColor Cyan
            return
        }

        Write-Host "`nAnalyzing $($videoFiles.Count) video file(s)..." -ForegroundColor Yellow

        $analysis = @{
            TotalFiles = $videoFiles.Count
            TotalSize = 0
            ByResolution = @{}
            ByCodec = @{}
            ByContainer = @{}
            NeedTranscode = @()
            AllFiles = @()
            FilesWithConcerns = 0
        }

        $current = 1
        $cacheHits = 0

        # Load central codec cache
        $codecCache = @{}
        $cacheLoaded = $false
        if ($CachePath -and -not $ForceRescan -and (Test-Path -LiteralPath $CachePath)) {
            try {
                $rawCache = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                if ($rawCache.Version -eq $script:CodecCacheVersion -and $rawCache.Entries) {
                    foreach ($prop in $rawCache.Entries.PSObject.Properties) {
                        $codecCache[$prop.Name] = $prop.Value
                    }
                    $cacheLoaded = $true
                } else {
                    Write-Host "  Cache version mismatch - rebuilding cache" -ForegroundColor DarkGray
                }
            } catch {
                Write-Host "  Cache file corrupt - rebuilding cache" -ForegroundColor DarkGray
            }
        }
        $cacheDirty = $false

        foreach ($file in $videoFiles) {
            $percentComplete = ($current / $videoFiles.Count) * 100
            Write-Progress -Activity "Analyzing codecs" -Status "[$current/$($videoFiles.Count)] $($file.Name)" -PercentComplete $percentComplete

            # Check central codec cache
            $cacheKey = Get-CodecCacheKey -File $file
            $cachedInfo = $null
            if ($cacheLoaded -and $codecCache.ContainsKey($cacheKey)) {
                $cachedInfo = $codecCache[$cacheKey]
            }

            if ($cachedInfo) {
                # Use cached data
                $fileInfo = @{
                    Path = $file.FullName
                    FileName = $cachedInfo.FileName
                    FolderName = $file.Directory.Name
                    Size = $cachedInfo.Size
                    Resolution = $cachedInfo.Resolution
                    Codec = $cachedInfo.Codec
                    AudioCodec = $cachedInfo.AudioCodec
                    Container = $cachedInfo.Container
                    HDR = [bool]$cachedInfo.HDR
                    QualityScore = $cachedInfo.QualityScore
                    Bitrate = $cachedInfo.Bitrate
                    QualityConcerns = @($cachedInfo.QualityConcerns)
                    HasConcerns = ($cachedInfo.QualityConcerns.Count -gt 0)
                    NeedsTranscode = [bool]$cachedInfo.NeedsTranscode
                    TranscodeMode = $cachedInfo.TranscodeMode
                    TranscodeReason = $cachedInfo.TranscodeReason
                }
                $cacheHits++
            } else {
                # Run full MediaInfo analysis
                $info = Get-VideoCodecInfo -FilePath $file.FullName -QualityScorer $QualityScorer
                $quality = if ($QualityScorer) { & $QualityScorer $file.Name $file.FullName } else { Get-QualityScore -FileName $file.Name -FilePath $file.FullName }

                $fileInfo = @{
                    Path = $file.FullName
                    FileName = $info.FileName
                    FolderName = $file.Directory.Name
                    Size = $info.FileSize
                    Resolution = $info.Resolution
                    Codec = $info.VideoCodec
                    AudioCodec = $info.AudioCodec
                    Container = $info.Container
                    HDR = $info.HDR
                    QualityScore = $quality.Score
                    Bitrate = $quality.Bitrate
                    QualityConcerns = $quality.QualityConcerns
                    HasConcerns = ($quality.QualityConcerns.Count -gt 0)
                    NeedsTranscode = $info.NeedsTranscode
                    TranscodeMode = if ($info.TranscodeMode) { $info.TranscodeMode } else { "none" }
                    TranscodeReason = $info.TranscodeReason -join "; "
                }

                # Add to central cache
                if ($CachePath) {
                    $codecCache[$cacheKey] = @{
                        FileName = $fileInfo.FileName
                        Size = $fileInfo.Size
                        LastModified = $file.LastWriteTimeUtc.ToString('o')
                        Resolution = $fileInfo.Resolution
                        Codec = $fileInfo.Codec
                        AudioCodec = $fileInfo.AudioCodec
                        Container = $fileInfo.Container
                        HDR = $fileInfo.HDR
                        QualityScore = $fileInfo.QualityScore
                        Bitrate = $fileInfo.Bitrate
                        QualityConcerns = @($fileInfo.QualityConcerns)
                        NeedsTranscode = $fileInfo.NeedsTranscode
                        TranscodeMode = $fileInfo.TranscodeMode
                        TranscodeReason = $fileInfo.TranscodeReason
                        CachedAt = (Get-Date).ToString("o")
                    }
                    $cacheDirty = $true
                }
            }

            $analysis.TotalSize += $fileInfo.Size

            # Count by resolution
            if (-not $analysis.ByResolution.ContainsKey($fileInfo.Resolution)) {
                $analysis.ByResolution[$fileInfo.Resolution] = 0
            }
            $analysis.ByResolution[$fileInfo.Resolution]++

            # Count by codec
            if (-not $analysis.ByCodec.ContainsKey($fileInfo.Codec)) {
                $analysis.ByCodec[$fileInfo.Codec] = 0
            }
            $analysis.ByCodec[$fileInfo.Codec]++

            # Count by container
            if (-not $analysis.ByContainer.ContainsKey($fileInfo.Container)) {
                $analysis.ByContainer[$fileInfo.Container] = 0
            }
            $analysis.ByContainer[$fileInfo.Container]++

            $analysis.AllFiles += $fileInfo

            # Track files with quality concerns
            if ($fileInfo.QualityConcerns.Count -gt 0) {
                $analysis.FilesWithConcerns++
            }

            # Add to transcode queue if needed
            if ($fileInfo.NeedsTranscode) {
                $analysis.NeedTranscode += $fileInfo
            }

            $current++
        }
        Write-Progress -Activity "Analyzing codecs" -Completed

        if ($cacheHits -gt 0) {
            Write-Host "$cacheHits/$($videoFiles.Count) file(s) loaded from cache" -ForegroundColor DarkGray
        }

        # Save central codec cache (single write at end)
        if ($CachePath -and $cacheDirty) {
            try {
                $cacheContainer = @{
                    Version = $script:CodecCacheVersion
                    UpdatedAt = (Get-Date).ToString("o")
                    EntryCount = $codecCache.Count
                    Entries = $codecCache
                }
                $cacheJson = $cacheContainer | ConvertTo-Json -Depth 4
                $cacheDir = Split-Path $CachePath -Parent
                if (-not (Test-Path $cacheDir)) {
                    New-Item -Path $cacheDir -ItemType Directory -Force | Out-Null
                }
                [System.IO.File]::WriteAllText($CachePath, $cacheJson, [System.Text.Encoding]::UTF8)
                Write-Host "  Codec cache saved ($($codecCache.Count) entries)" -ForegroundColor DarkGray
            } catch {
                Write-Host "  Warning: Could not save codec cache: $_" -ForegroundColor DarkYellow
            }
        }

        # Display results
        Write-Host "`n=== Library Statistics ===" -ForegroundColor Cyan
        Write-Host "Total Files: $($analysis.TotalFiles)" -ForegroundColor White
        Write-Host "Total Size: $(Format-QualitySize $analysis.TotalSize)" -ForegroundColor White

        Write-Host "`n=== By Resolution ===" -ForegroundColor Cyan
        $analysis.ByResolution.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            $pct = [math]::Round(($_.Value / $analysis.TotalFiles) * 100, 1)
            Write-Host "  $($_.Key): $($_.Value) ($pct%)" -ForegroundColor White
        }

        Write-Host "`n=== By Video Codec ===" -ForegroundColor Cyan
        $analysis.ByCodec.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            $pct = [math]::Round(($_.Value / $analysis.TotalFiles) * 100, 1)
            Write-Host "  $($_.Key): $($_.Value) ($pct%)" -ForegroundColor White
        }

        Write-Host "`n=== By Container ===" -ForegroundColor Cyan
        $analysis.ByContainer.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            $pct = [math]::Round(($_.Value / $analysis.TotalFiles) * 100, 1)
            Write-Host "  $($_.Key): $($_.Value) ($pct%)" -ForegroundColor White
        }

        # Quality score range
        $scores = $analysis.AllFiles | ForEach-Object { $_.QualityScore }
        $minScore = ($scores | Measure-Object -Minimum).Minimum
        $maxScore = ($scores | Measure-Object -Maximum).Maximum
        $avgScore = [math]::Round(($scores | Measure-Object -Average).Average, 1)
        Write-Host "`n=== Quality Scores ===" -ForegroundColor Cyan
        Write-Host "  Range: $minScore - $maxScore (Average: $avgScore)" -ForegroundColor White

        # Show quality concerns summary
        if ($analysis.FilesWithConcerns -gt 0) {
            Write-Host "`n=== Quality Concerns ===" -ForegroundColor Yellow
            Write-Host "  Files with concerns: $($analysis.FilesWithConcerns)" -ForegroundColor $(if ($analysis.FilesWithConcerns -gt 10) { 'Red' } else { 'Yellow' })
        }

        # The interactive report browser (7-option loop) was retired: its jobs
        # moved to flows that act instead of list — Radarr Re-acquisition scans
        # the library directly (below-resolution / by-concerns modes), the
        # hardsub audit exports its own re-acquisition CSV, and the transcode
        # queue below still offers the legacy-codec action. What remains here
        # is the summary above plus an optional full CSV for offline digging.
        if ($analysis.FilesWithConcerns -gt 0) {
            Write-Host "  Act on concerns via Utilities > Radarr Re-acquisition (scan mode 4: by quality concerns)." -ForegroundColor DarkGray
        }

        $exportAns = Read-Host "`nExport full quality report to CSV? (Y/N) [N]"
        if ($exportAns -match '^[Yy]') {
            $qualityExportPath = Join-Path $ReportsFolder "QualityReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $analysis.AllFiles | Sort-Object QualityScore |
                Select-Object FolderName, FileName, QualityScore, Resolution, Codec, AudioCodec, Container, HDR, @{N='SizeMB';E={[math]::Round($_.Size/1MB,2)}}, @{N='BitrateMbps';E={if($_.Bitrate -gt 0){[math]::Round($_.Bitrate/1000000,1)}else{'N/A'}}}, @{N='QualityConcerns';E={$_.QualityConcerns -join '; '}}, NeedsTranscode, TranscodeReason, Path |
                Export-Csv -Path $qualityExportPath -NoTypeInformation -Encoding UTF8
            Write-Host "Quality report exported to: $qualityExportPath" -ForegroundColor Green
        }

        # Transcode queue
        if ($analysis.NeedTranscode.Count -gt 0) {
            Write-Host "`n=== Potential Transcoding Queue ===" -ForegroundColor Yellow
            Write-Host "Found $($analysis.NeedTranscode.Count) file(s) that may benefit from transcoding:" -ForegroundColor White

            $analysis.NeedTranscode | Sort-Object QualityScore | Select-Object -First 10 | ForEach-Object {
                Write-Host "`n  $($_.FolderName)" -ForegroundColor White
                Write-Host "    Score: $($_.QualityScore) | $(Format-QualitySize $_.Size) | $($_.Resolution) | $($_.Codec)" -ForegroundColor Gray
                Write-Host "    Reason: $($_.TranscodeReason)" -ForegroundColor Yellow
            }

            if ($analysis.NeedTranscode.Count -gt 10) {
                Write-Host "`n  ... and $($analysis.NeedTranscode.Count - 10) more files" -ForegroundColor Gray
            }

            # Export option
            if (-not $ExportPath) {
                $exportInput = Read-Host "`nExport transcoding queue to CSV? (Y/N) [N]"
                if ($exportInput -eq 'Y' -or $exportInput -eq 'y') {
                    $ExportPath = Join-Path $ReportsFolder "TranscodeQueue_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                }
            }

            if ($ExportPath) {
                $analysis.NeedTranscode | Sort-Object QualityScore |
                    Select-Object FolderName, FileName, QualityScore, @{N='SizeMB';E={[math]::Round($_.Size/1MB,2)}}, Resolution, Codec, Container, TranscodeReason, Path |
                    Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
                Write-Host "`nTranscoding queue exported to: $ExportPath" -ForegroundColor Green
                Write-Host "  Transcoding queue exported to: $ExportPath" -ForegroundColor Gray
            }
        } else {
            Write-Host "`nNo files require transcoding for compatibility." -ForegroundColor Green
        }

        Write-Host "  Codec analysis completed: $($analysis.TotalFiles) files analyzed" -ForegroundColor Gray
        return $analysis
    }
    catch {
        Write-Host "Error during codec analysis: $_" -ForegroundColor Red
        return $null
    }
}

function Invoke-Transcode {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [array]$TranscodeQueue,
        [string]$TargetCodec = "libx264",
        [string]$FFmpegPath = "ffmpeg"
    )

    if ($TranscodeQueue.Count -eq 0) {
        Write-Host "No files to transcode" -ForegroundColor Cyan
        return
    }

    # Check FFmpeg
    $ffmpegCmd = Get-Command $FFmpegPath -ErrorAction SilentlyContinue
    if (-not $ffmpegCmd) {
        Write-Host "`nFFmpeg not found!" -ForegroundColor Red
        Write-Host "Please install FFmpeg to use transcoding features." -ForegroundColor Yellow
        Write-Host "Download from: https://ffmpeg.org/download.html" -ForegroundColor Cyan
        Write-Host "Or install via: winget install ffmpeg" -ForegroundColor Cyan
        return
    }

    $total = $TranscodeQueue.Count
    $current = 0
    $remuxed = 0
    $transcoded = 0
    $failed = 0

    Write-Host "`n=== Starting Transcode ===" -ForegroundColor Cyan
    Write-Host "Files to process: $total" -ForegroundColor White
    Write-Host "Press Ctrl+C to abort at any time`n" -ForegroundColor Gray

    foreach ($item in $TranscodeQueue) {
        $current++
        $inputPath = $item.Path
        $inputDir = [System.IO.Path]::GetDirectoryName($inputPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.FileName)
        $outputPath = Join-Path $inputDir ($baseName + ".mkv")
        $mode = if ($item.TranscodeMode) { $item.TranscodeMode } else { "transcode" }

        # Skip remux if input and output are the same (already .mkv)
        if ($mode -eq "remux" -and $inputPath -eq $outputPath) {
            Write-Host "[$current/$total] Skipping remux (already .mkv): $($item.FileName)" -ForegroundColor Gray
            continue
        }

        # Use temp file during processing
        $tempOutput = $outputPath + ".tmp.mkv"

        # -hide_banner   : drop ffmpeg's version/config banner at startup
        # -loglevel error: suppress the per-frame warning flood from older
        #                  codecs (the "Discarding excessive bitstream in
        #                  packed xvid" spam on XviD/MPEG-4 sources). Real
        #                  failures still print.
        # -stats         : keep the live single-line progress counter so the
        #                  user can see speed/ETA without the noise.
        $ffmpegQuiet = "-hide_banner -loglevel error -stats"

        if ($mode -eq "remux") {
            Write-Host "[$current/$total] Remuxing: $($item.FileName)" -ForegroundColor Cyan
            $ffmpegArgs = "$ffmpegQuiet -i `"$inputPath`" -c:v copy -c:a copy -c:s copy `"$tempOutput`" -y"
        } else {
            Write-Host "[$current/$total] Transcoding: $($item.FileName)" -ForegroundColor Yellow
            Write-Host "  Reason: $($item.TranscodeReason -join '; ')" -ForegroundColor Gray
            # Use H.265 for VC-1 (HD content benefits from HEVC), H.264 for everything else
            $codec = if ($item.Codec -match 'VC-1|WMV') { "libx265" } else { $TargetCodec }
            $crf = if ($codec -eq "libx265") { "28" } else { "23" }
            $ffmpegArgs = "$ffmpegQuiet -i `"$inputPath`" -map 0:v -map 0:a -map 0:s? -c:v $codec -crf $crf -preset medium -c:a aac -b:a 192k -c:s copy `"$tempOutput`" -y"
        }

        # Run FFmpeg
        $process = Start-Process -FilePath $FFmpegPath -ArgumentList $ffmpegArgs -NoNewWindow -Wait -PassThru

        if ($process.ExitCode -eq 0 -and (Test-Path $tempOutput)) {
            $outSize = (Get-Item $tempOutput).Length
            if ($outSize -gt 0) {
                if ($PSCmdlet.ShouldProcess($inputPath, "Transcode and replace")) {
                    # Rename associated files (NFO, images) if extension changed
                    $inputExt = [System.IO.Path]::GetExtension($inputPath)
                    $outputExt = [System.IO.Path]::GetExtension($outputPath)

                    if ($inputExt -ne $outputExt) {
                        # Find and rename associated files that match the original base name
                        $oldBaseName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
                        $newBaseName = [System.IO.Path]::GetFileNameWithoutExtension($outputPath)

                        # Rename NFO file if it exists
                        $oldNfo = Join-Path $inputDir "$oldBaseName.nfo"
                        $newNfo = Join-Path $inputDir "$newBaseName.nfo"
                        if ((Test-Path $oldNfo) -and $oldNfo -ne $newNfo) {
                            Rename-Item -LiteralPath $oldNfo -NewName "$newBaseName.nfo" -ErrorAction SilentlyContinue
                            Write-Host "  Renamed NFO: $oldBaseName.nfo -> $newBaseName.nfo" -ForegroundColor Gray
                        }

                        # Rename associated image files (poster, fanart, etc.)
                        $imagePatterns = @("$oldBaseName-poster.*", "$oldBaseName-fanart.*", "$oldBaseName-thumb.*", "$oldBaseName-banner.*", "$oldBaseName-landscape.*")
                        foreach ($pattern in $imagePatterns) {
                            $oldImages = Get-ChildItem -LiteralPath $inputDir -Filter $pattern -ErrorAction SilentlyContinue
                            foreach ($oldImage in $oldImages) {
                                $newImageName = $oldImage.Name -replace [regex]::Escape($oldBaseName), $newBaseName
                                if ($oldImage.Name -ne $newImageName) {
                                    Rename-Item -LiteralPath $oldImage.FullName -NewName $newImageName -ErrorAction SilentlyContinue
                                    Write-Host "  Renamed image: $($oldImage.Name) -> $newImageName" -ForegroundColor Gray
                                }
                            }
                        }
                    }

                    # Safe swap: keep the original as a .bak until the new file is
                    # confirmed in place, so a failure mid-swap never leaves the
                    # folder without a playable file.
                    $backupPath = "$inputPath.replaced.bak"
                    $swapOk = $false

                    try {
                        Rename-Item -LiteralPath $inputPath -NewName ([System.IO.Path]::GetFileName($backupPath)) -Confirm:$false -ErrorAction Stop
                    } catch {
                        Write-Host "  Failed to stage original for replacement: $($_.Exception.Message)" -ForegroundColor Red
                        Remove-Item -LiteralPath $tempOutput -Force -Confirm:$false -ErrorAction SilentlyContinue
                        $failed++
                        continue
                    }

                    try {
                        Rename-Item -LiteralPath $tempOutput -NewName ([System.IO.Path]::GetFileName($outputPath)) -Confirm:$false -ErrorAction Stop
                        $swapOk = $true
                    } catch {
                        Write-Host "  Failed to move new file into place: $($_.Exception.Message)" -ForegroundColor Red
                        # Put the original back under its real name before cleaning up.
                        try {
                            Rename-Item -LiteralPath $backupPath -NewName ([System.IO.Path]::GetFileName($inputPath)) -Confirm:$false -ErrorAction Stop
                        } catch {
                            Write-Host "  CRITICAL: could not restore original from backup: $backupPath ($($_.Exception.Message))" -ForegroundColor Red
                        }
                        Remove-Item -LiteralPath $tempOutput -Force -Confirm:$false -ErrorAction SilentlyContinue
                        $failed++
                    }

                    if ($swapOk) {
                        # New file confirmed in place - the backup can go.
                        try {
                            Remove-Item -LiteralPath $backupPath -Force -Confirm:$false -ErrorAction Stop
                        } catch {
                            Write-Host "  Warning: could not delete backup: $backupPath ($($_.Exception.Message))" -ForegroundColor Yellow
                        }

                        if ($mode -eq "remux") {
                            Write-Host "  Remuxed successfully" -ForegroundColor Green
                            $remuxed++
                        } else {
                            Write-Host "  Transcoded successfully" -ForegroundColor Green
                            $transcoded++
                        }
                    }
                } else {
                    # -WhatIf: ffmpeg already wrote the temp file; remove it so
                    # the library folder is left exactly as it was found.
                    Remove-Item -LiteralPath $tempOutput -Force -WhatIf:$false -Confirm:$false -ErrorAction SilentlyContinue
                }
            } else {
                Write-Host "  Failed (output empty)" -ForegroundColor Red
                Remove-Item -LiteralPath $tempOutput -Force -ErrorAction SilentlyContinue
                $failed++
            }
        } else {
            Write-Host "  Failed" -ForegroundColor Red
            Remove-Item -LiteralPath $tempOutput -Force -ErrorAction SilentlyContinue
            $failed++
        }
    }

    Write-Host "`n=== Transcode Complete ===" -ForegroundColor Cyan
    Write-Host "  Remuxed: $remuxed" -ForegroundColor Cyan
    Write-Host "  Transcoded: $transcoded" -ForegroundColor Yellow
    Write-Host "  Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
    Write-Host "  Transcode complete - Remuxed: $remuxed, Transcoded: $transcoded, Failed: $failed" -ForegroundColor Gray
}

function New-TranscodeScript {
    param(
        [array]$TranscodeQueue,
        [string]$TargetCodec = "libx264",
        [string]$TargetResolution = $null,
        [string]$FFmpegPath = "ffmpeg",
        [string]$OutputFolder
    )

    if ($TranscodeQueue.Count -eq 0) {
        Write-Host "No files in transcode queue" -ForegroundColor Cyan
        return
    }

    # Check FFmpeg
    $ffmpegCmd = Get-Command $FFmpegPath -ErrorAction SilentlyContinue
    if (-not $ffmpegCmd) {
        Write-Host "`nFFmpeg not found!" -ForegroundColor Red
        Write-Host "Please install FFmpeg to use transcoding features." -ForegroundColor Yellow
        Write-Host "Download from: https://ffmpeg.org/download.html" -ForegroundColor Cyan
        Write-Host "Or install via: winget install ffmpeg" -ForegroundColor Cyan
        return $null
    }

    # Count files by mode
    $remuxCount = ($TranscodeQueue | Where-Object { $_.TranscodeMode -eq "remux" }).Count
    $transcodeCount = ($TranscodeQueue | Where-Object { $_.TranscodeMode -eq "transcode" -or $_.TranscodeMode -eq $null }).Count

    $scriptPath = Join-Path $OutputFolder "transcode_$(Get-Date -Format 'yyyyMMdd_HHmmss').ps1"

    $scriptContent = @"
# LibraryLint Transcode Script
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# Target Codec: $TargetCodec
# Files to process: $($TranscodeQueue.Count)
#   - Remux only (no re-encoding): $remuxCount
#   - Full transcode: $transcodeCount
#
# Output files are saved in the same folder as the original.
# Original files are deleted after successful processing.

`$ffmpegPath = "$FFmpegPath"  # Update this path if ffmpeg is not in PATH

`$files = @(
"@

    $itemCount = $TranscodeQueue.Count
    $currentIndex = 0
    foreach ($item in $TranscodeQueue) {
        $currentIndex++
        # Output goes to same folder as input, with .mkv extension
        $inputDir = [System.IO.Path]::GetDirectoryName($item.Path)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.FileName)
        $mode = if ($item.TranscodeMode) { $item.TranscodeMode } else { "transcode" }

        # Output is always .mkv (same name as input, different extension if needed)
        # If input is already .mkv, the script uses a temp file during processing
        $outputFile = Join-Path $inputDir ($baseName + ".mkv")
        # Escape single quotes by doubling them for PowerShell string literals
        $escapedInput = $item.Path -replace "'", "''"
        $escapedOutput = $outputFile -replace "'", "''"
        # Only add comma if not the last item
        $comma = if ($currentIndex -lt $itemCount) { "," } else { "" }
        $scriptContent += "`n    @{ Input = '$escapedInput'; Output = '$escapedOutput'; Mode = '$mode' }$comma"
    }

    $resolutionParam = if ($TargetResolution) { "-vf scale=$TargetResolution" } else { "" }

    $scriptContent += @"

)

`$total = `$files.Count
`$current = 1
`$remuxed = 0
`$transcoded = 0
`$failed = 0
`$deleted = 0

foreach (`$file in `$files) {
    # Skip remux if input and output are the same file (already .mkv container)
    if (`$file.Mode -eq "remux" -and `$file.Input -eq `$file.Output) {
        Write-Host "[`$current/`$total] Skipping remux (already .mkv): `$(`$file.Input)" -ForegroundColor Gray
        `$current++
        continue
    }

    # Use temp file to avoid issues when input/output have same name
    `$tempOutput = `$file.Output + ".tmp.mkv"

    if (`$file.Mode -eq "remux") {
        # REMUX: Copy streams without re-encoding (fast, no quality loss)
        Write-Host "[`$current/`$total] Remuxing (no re-encode): `$(`$file.Input)" -ForegroundColor Cyan
        & `$ffmpegPath -hide_banner -loglevel error -stats -i "`$(`$file.Input)" -c:v copy -c:a copy -c:s copy "`$tempOutput" -y

        if (`$LASTEXITCODE -eq 0 -and (Test-Path `$tempOutput)) {
            # Verify output file is valid (has size > 0)
            `$outSize = (Get-Item `$tempOutput).Length
            if (`$outSize -gt 0) {
                # Rename associated files (NFO, images) if extension changed
                `$inputDir = [System.IO.Path]::GetDirectoryName(`$file.Input)
                `$inputExt = [System.IO.Path]::GetExtension(`$file.Input)
                `$outputExt = [System.IO.Path]::GetExtension(`$file.Output)
                if (`$inputExt -ne `$outputExt) {
                    `$oldBaseName = [System.IO.Path]::GetFileNameWithoutExtension(`$file.Input)
                    `$newBaseName = [System.IO.Path]::GetFileNameWithoutExtension(`$file.Output)
                    # Rename NFO
                    `$oldNfo = Join-Path `$inputDir "`$oldBaseName.nfo"
                    if (Test-Path `$oldNfo) {
                        Rename-Item -LiteralPath `$oldNfo -NewName "`$newBaseName.nfo" -ErrorAction SilentlyContinue
                        Write-Host "  Renamed NFO: `$oldBaseName.nfo -> `$newBaseName.nfo" -ForegroundColor Gray
                    }
                    # Rename images
                    @("poster", "fanart", "thumb", "banner", "landscape") | ForEach-Object {
                        Get-ChildItem -LiteralPath `$inputDir -Filter "`$oldBaseName-`$_.*" -ErrorAction SilentlyContinue | ForEach-Object {
                            `$newName = `$_.Name -replace [regex]::Escape(`$oldBaseName), `$newBaseName
                            Rename-Item -LiteralPath `$_.FullName -NewName `$newName -ErrorAction SilentlyContinue
                            Write-Host "  Renamed: `$(`$_.Name) -> `$newName" -ForegroundColor Gray
                        }
                    }
                }
                # Delete original and rename temp to final
                Remove-Item -LiteralPath `$file.Input -Force
                Rename-Item -LiteralPath `$tempOutput -NewName ([System.IO.Path]::GetFileName(`$file.Output))
                Write-Host "  -> Remuxed & replaced: `$(`$file.Output)" -ForegroundColor Green
                `$remuxed++
                `$deleted++
            } else {
                Write-Host "  -> Failed (output empty): `$(`$file.Input)" -ForegroundColor Red
                Remove-Item -LiteralPath `$tempOutput -Force -ErrorAction SilentlyContinue
                `$failed++
            }
        } else {
            Write-Host "  -> Failed: `$(`$file.Input)" -ForegroundColor Red
            Remove-Item -LiteralPath `$tempOutput -Force -ErrorAction SilentlyContinue
            `$failed++
        }
    } else {
        # TRANSCODE: Re-encode video to H.264
        Write-Host "[`$current/`$total] Transcoding: `$(`$file.Input)" -ForegroundColor Yellow
        & `$ffmpegPath -hide_banner -loglevel error -stats -i "`$(`$file.Input)" -map 0:v -map 0:a -map 0:s? -c:v $TargetCodec -crf 23 -preset medium $resolutionParam -c:a aac -b:a 192k -c:s copy "`$tempOutput" -y

        if (`$LASTEXITCODE -eq 0 -and (Test-Path `$tempOutput)) {
            # Verify output file is valid (has size > 0)
            `$outSize = (Get-Item `$tempOutput).Length
            if (`$outSize -gt 0) {
                # Rename associated files (NFO, images) if extension changed
                `$inputDir = [System.IO.Path]::GetDirectoryName(`$file.Input)
                `$inputExt = [System.IO.Path]::GetExtension(`$file.Input)
                `$outputExt = [System.IO.Path]::GetExtension(`$file.Output)
                if (`$inputExt -ne `$outputExt) {
                    `$oldBaseName = [System.IO.Path]::GetFileNameWithoutExtension(`$file.Input)
                    `$newBaseName = [System.IO.Path]::GetFileNameWithoutExtension(`$file.Output)
                    # Rename NFO
                    `$oldNfo = Join-Path `$inputDir "`$oldBaseName.nfo"
                    if (Test-Path `$oldNfo) {
                        Rename-Item -LiteralPath `$oldNfo -NewName "`$newBaseName.nfo" -ErrorAction SilentlyContinue
                        Write-Host "  Renamed NFO: `$oldBaseName.nfo -> `$newBaseName.nfo" -ForegroundColor Gray
                    }
                    # Rename images
                    @("poster", "fanart", "thumb", "banner", "landscape") | ForEach-Object {
                        Get-ChildItem -LiteralPath `$inputDir -Filter "`$oldBaseName-`$_.*" -ErrorAction SilentlyContinue | ForEach-Object {
                            `$newName = `$_.Name -replace [regex]::Escape(`$oldBaseName), `$newBaseName
                            Rename-Item -LiteralPath `$_.FullName -NewName `$newName -ErrorAction SilentlyContinue
                            Write-Host "  Renamed: `$(`$_.Name) -> `$newName" -ForegroundColor Gray
                        }
                    }
                }
                # Delete original and rename temp to final
                Remove-Item -LiteralPath `$file.Input -Force
                Rename-Item -LiteralPath `$tempOutput -NewName ([System.IO.Path]::GetFileName(`$file.Output))
                Write-Host "  -> Transcoded & replaced: `$(`$file.Output)" -ForegroundColor Green
                `$transcoded++
                `$deleted++

                # 5-second pause to allow user to stop
                Write-Host "  Press any key within 5 seconds to stop, or wait to continue..." -ForegroundColor Magenta
                `$timeout = 5
                `$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                while (`$stopwatch.Elapsed.TotalSeconds -lt `$timeout) {
                    if ([Console]::KeyAvailable) {
                        `$null = [Console]::ReadKey(`$true)
                        Write-Host "``n  User requested stop. Exiting..." -ForegroundColor Yellow
                        Write-Host "``n========== Summary (Stopped Early) ==========" -ForegroundColor Cyan
                        Write-Host "Remuxed (fast, no quality loss): `$remuxed" -ForegroundColor Cyan
                        Write-Host "Transcoded (re-encoded): `$transcoded" -ForegroundColor Yellow
                        Write-Host "Original files deleted: `$deleted" -ForegroundColor Green
                        Write-Host "Failed: `$failed" -ForegroundColor Red
                        Write-Host "==============================================" -ForegroundColor Cyan
                        exit
                    }
                    Start-Sleep -Milliseconds 100
                }
            } else {
                Write-Host "  -> Failed (output empty): `$(`$file.Input)" -ForegroundColor Red
                Remove-Item -LiteralPath `$tempOutput -Force -ErrorAction SilentlyContinue
                `$failed++
            }
        } else {
            Write-Host "  -> Failed: `$(`$file.Input)" -ForegroundColor Red
            Remove-Item -LiteralPath `$tempOutput -Force -ErrorAction SilentlyContinue
            `$failed++
        }
    }

    `$current++
}

Write-Host "`n========== Summary ==========" -ForegroundColor Cyan
Write-Host "Remuxed (fast, no quality loss): `$remuxed" -ForegroundColor Cyan
Write-Host "Transcoded (re-encoded): `$transcoded" -ForegroundColor Yellow
Write-Host "Original files deleted: `$deleted" -ForegroundColor Green
Write-Host "Failed: `$failed" -ForegroundColor $(if (`$failed -gt 0) { 'Red' } else { 'Green' })
Write-Host "==============================" -ForegroundColor Cyan
"@

    $scriptContent | Out-File -FilePath $scriptPath -Encoding UTF8
    Write-Host "`nTranscode script created: $scriptPath" -ForegroundColor Green
    Write-Host "  - Remux only (fast, no quality loss): $remuxCount files" -ForegroundColor Cyan
    Write-Host "  - Full transcode (re-encode): $transcodeCount files" -ForegroundColor Yellow
    Write-Host "`nRun the script to start processing, or edit it to customize parameters." -ForegroundColor Cyan
    Write-Host "  Transcode script created: $scriptPath - Remux: $remuxCount, Transcode: $transcodeCount" -ForegroundColor Gray

    return $scriptPath
}

function Remove-CodecSidecarFiles {
    [CmdletBinding(SupportsShouldProcess)]
    param([string]$Path)
    $removed = 0
    Get-ChildItem -LiteralPath $Path -Recurse -Filter "codec-info.json" -File -ErrorAction SilentlyContinue | ForEach-Object {
        if ($PSCmdlet.ShouldProcess($_.FullName, "Remove deprecated codec sidecar")) {
            try {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                $removed++
            } catch { }
        } else {
            $removed++
        }
    }
    return $removed
}

<#
.SYNOPSIS
    Locates the tesseract OCR executable.
.DESCRIPTION
    Checks PATH first, then the standard UB-Mannheim Windows install
    location (what `winget install UB-Mannheim.TesseractOCR` produces).
.OUTPUTS
    Full path to tesseract executable, or $null when not found.
#>
function Test-TesseractInstallation {
    [CmdletBinding()]
    param()

    $cmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $candidates = @(
        "$env:ProgramFiles\Tesseract-OCR\tesseract.exe",
        "${env:ProgramFiles(x86)}\Tesseract-OCR\tesseract.exe",
        "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return $c }
    }
    return $null
}

<#
.SYNOPSIS
    Detects burned-in (hardcoded) subtitles by OCR-sampling video frames.
.DESCRIPTION
    Release tags (KORSUB/HC/HARDSUB) only catch rips that admit to the
    defect in their name. This audit looks at the actual pixels: for each
    movie it extracts N frames spread across the runtime, crops the lower
    third (where subtitles live), runs tesseract OCR on each, and reports
    what fraction of sampled frames contain text.

    Classification by coverage:
      >= FullThresholdPct    : FULL hardsub — text on most frames means the
                               whole film is burned. Re-acquisition candidate.
      >= SuspectThresholdPct : suspicious — could be a partial burn, heavy
                               credits, or on-screen-text-heavy film. Review.
      below                  : clean. Sparse hits are normal (title cards,
                               credits, forced-narrative subs like Deadpool's
                               foreign-dialogue lines — NOT defects).

    Caveats:
      - OCR language defaults to English glyph detection. Korean/CJK burns
        (classic KORSUB) may under-detect unless the matching tesseract
        language pack is installed and passed via -OcrLanguages 'eng+kor'.
      - Text-heavy films (documentaries with lower-third captions) can land
        in "suspicious" legitimately. That's why nothing is auto-actioned.
.PARAMETER Path
    Movie library root. Each subfolder is treated as one movie.
.PARAMETER VideoExtensions
    Video file extensions to consider (e.g. @('.mkv','.mp4')).
.PARAMETER FFmpegPath
    Path to ffmpeg. Defaults to 'ffmpeg' on PATH.
.PARAMETER SampleCount
    Frames sampled per movie, spread across 10%%-90%% of the runtime.
    More samples = better estimate, slower audit.
.PARAMETER Limit
    Cap the number of movies audited (0 = all).
.PARAMETER SkipQualityAccepted
    Skip folders carrying the .quality_ok marker.
.PARAMETER FullThresholdPct / SuspectThresholdPct
    Classification cut lines (percent of sampled frames with text).
.PARAMETER OcrLanguages
    Tesseract language spec (default 'eng'; e.g. 'eng+kor' for KORSUB).
.OUTPUTS
    Hashtable: Scanned, Failed, Results (per-movie objects with
    Folder, VideoPath, SamplesTaken, TextFrames, CoveragePct, Classification).
#>
function Invoke-HardsubAudit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [string[]]$VideoExtensions = @('.mkv', '.mp4', '.avi', '.m4v', '.mpg', '.mpeg', '.wmv', '.ts'),
        [string]$FFmpegPath = 'ffmpeg',
        [int]$SampleCount = 48,
        [int]$Limit = 0,
        [switch]$SkipQualityAccepted,
        [int]$FullThresholdPct = 50,
        [int]$SuspectThresholdPct = 20,
        [string]$OcrLanguages = 'eng',
        # When set, a self-contained HTML report (verdicts + the actual
        # flagged frame pairs as embedded images) is written here.
        [string]$HtmlReportPath
    )

    $tesseract = Test-TesseractInstallation
    if (-not $tesseract) {
        Write-Host "  tesseract not found — install via: winget install UB-Mannheim.TesseractOCR" -ForegroundColor Red
        return $null
    }
    $ffmpegCmd = Get-Command $FFmpegPath -ErrorAction SilentlyContinue
    if (-not $ffmpegCmd) {
        Write-Host "  ffmpeg not found — install via Settings > Install/Update Dependencies." -ForegroundColor Red
        return $null
    }

    # Reuse the codec-analysis junk filter so trailers/samples/extras never
    # get picked as the "primary video".
    $junkNameRegex = $script:JunkNameRegex

    # Auto-extend OCR languages with any installed non-Latin packs. KORSUB-
    # style burns (Korean subs on an English film) OCR as nothing under
    # eng-only tessdata — the audit's canonical target would be invisible.
    try {
        $installedLangs = @(& $tesseract --list-langs 2>$null)
        foreach ($extra in @('kor', 'jpn', 'chi_sim', 'chi_tra', 'rus', 'tha', 'vie', 'ara')) {
            if ($installedLangs -contains $extra -and $OcrLanguages -notmatch [regex]::Escape($extra)) {
                $OcrLanguages = "$OcrLanguages+$extra"
            }
        }
    } catch {}

    # OCR a frame and return only CONFIDENT words: tesseract TSV rows with
    # confidence >= 60 and >= 2 letters. The old check ("any 5 letters in
    # the raw dump") counted grain/artifact hallucinations as text and
    # ranked noisy transfers above actual hardsubs. --psm 11 = sparse text,
    # the right mode for an isolated subtitle line (psm 6 assumes the frame
    # IS a text block, which actively encourages hallucination on noise).
    $getConfidentWords = {
        param($ImagePath)
        $words = @()
        $tsv = & $tesseract $ImagePath stdout -l $OcrLanguages --psm 11 tsv 2>$null
        foreach ($line in $tsv) {
            $cols = $line -split "`t"
            if ($cols.Count -ge 12 -and $cols[10] -match '^\d+(\.\d+)?$') {
                $conf = [double]$cols[10]
                $word = $cols[11].Trim()
                $letters = ($word -replace '[^\p{L}]', '')
                if ($conf -ge 60 -and $letters.Length -ge 2) { $words += $word.ToLower() }
            }
        }
        return $words
    }

    $folders = @(Get-ChildItem -LiteralPath $Path -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '_*' })

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "LibraryLint_HardsubAudit_$([guid]::NewGuid().ToString('N').Substring(0,8))"
    New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null

    $results = @()
    $scanned = 0
    $failed  = 0
    $skippedAccepted = 0
    $reportWritten = $null
    $total   = $folders.Count

    try {
        foreach ($folder in $folders) {
            if ($Limit -gt 0 -and $scanned -ge $Limit) { break }

            if ($SkipQualityAccepted -and (Test-QualityAccepted -FolderPath $folder.FullName)) {
                $skippedAccepted++
                continue
            }

            $video = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and $_.Name -notmatch $junkNameRegex } |
                Sort-Object Length -Descending | Select-Object -First 1
            if (-not $video) { continue }

            $scanned++
            # WindowWidth throws in redirected consoles — fall back rather
            # than let a progress cosmetic abort the movie's whole iteration.
            $consoleWidth = try { [Console]::WindowWidth } catch { 120 }
            Write-Host "`r  [$scanned/$total] $($folder.Name)".PadRight([Math]::Max(40, $consoleWidth - 1)) -NoNewline -ForegroundColor Gray

            # Duration from ffmpeg's stderr banner ("Duration: 01:48:20.06").
            # Avoids assuming ffprobe sits next to ffmpeg.
            $durationSec = 0
            try {
                $ffInfo = & $FFmpegPath -hide_banner -i $video.FullName 2>&1 | Out-String
                if ($ffInfo -match 'Duration:\s*(\d+):(\d+):(\d+(\.\d+)?)') {
                    $durationSec = [int]$Matches[1] * 3600 + [int]$Matches[2] * 60 + [double]$Matches[3]
                }
            } catch {}
            if ($durationSec -lt 120) {
                # Unparseable or absurdly short — skip rather than misreport.
                $failed++
                continue
            }

            # Sample between 10% and 90% of runtime — skips studio logos at
            # the head and the credit roll at the tail (credits would OCR as
            # text on every frame and false-positive the whole film).
            $textFrames = 0
            $samplesTaken = 0
            $sampleSnippets = @()
            $evidencePairs = @()
            $framePng  = Join-Path $tempRoot "frame.png"
            $framePng2 = Join-Path $tempRoot "frame2.png"
            for ($i = 0; $i -lt $SampleCount; $i++) {
                $t = $durationSec * (0.10 + 0.80 * ($i / [Math]::Max(1, $SampleCount - 1)))
                $ts = [TimeSpan]::FromSeconds($t).ToString('hh\:mm\:ss')

                # -ss before -i = fast keyframe seek. Crop to the lower third
                # (where subs render), grayscale for cleaner OCR.
                Remove-Item -LiteralPath $framePng -Force -ErrorAction SilentlyContinue
                & $FFmpegPath -hide_banner -loglevel error -ss $ts -i $video.FullName `
                    -frames:v 1 -vf "crop=iw:ih/3:0:2*ih/3,format=gray" -y $framePng 2>&1 | Out-Null
                if (-not (Test-Path -LiteralPath $framePng)) { continue }
                $samplesTaken++

                $wordsA = @(& $getConfidentWords $framePng)
                if ($wordsA.Count -lt 2) { continue }   # subtitle lines are multi-word

                # Temporal check — the discriminator OCR alone can't provide.
                # A subtitle cue changes or vanishes ~2.5s later; burned-in
                # SCENE text (signs, storefronts, location cards) persists.
                # Only candidate frames pay for the second extraction.
                $t2 = [Math]::Min($t + 2.5, $durationSec * 0.95)
                $ts2 = [TimeSpan]::FromSeconds($t2).ToString('hh\:mm\:ss')
                Remove-Item -LiteralPath $framePng2 -Force -ErrorAction SilentlyContinue
                & $FFmpegPath -hide_banner -loglevel error -ss $ts2 -i $video.FullName `
                    -frames:v 1 -vf "crop=iw:ih/3:0:2*ih/3,format=gray" -y $framePng2 2>&1 | Out-Null
                $wordsB = if (Test-Path -LiteralPath $framePng2) { @(& $getConfidentWords $framePng2) } else { @() }

                $setA = New-Object 'System.Collections.Generic.HashSet[string]' ([string[]]$wordsA), ([System.StringComparer]::OrdinalIgnoreCase)
                $overlap = @($wordsB | Where-Object { $setA.Contains($_) }).Count
                $isStaticText = ($wordsB.Count -ge 2) -and
                    ($overlap -ge [Math]::Ceiling(0.7 * [Math]::Min($wordsA.Count, $wordsB.Count)))

                if (-not $isStaticText) {
                    $textFrames++
                    if ($sampleSnippets.Count -lt 3) {
                        $sampleSnippets += (($wordsA | Select-Object -First 6) -join ' ')
                    }
                    # Keep small JPEG copies of the first few flagged pairs
                    # as report evidence — re-encoded from the already-
                    # extracted PNGs, so no extra video seeks. Seeing the
                    # detected frame next to its +2.5s partner is what lets
                    # a human confirm sub-vs-scene-text at a glance.
                    if ($HtmlReportPath -and $evidencePairs.Count -lt 4) {
                        $evA = Join-Path $tempRoot ("ev_{0}_{1}_A.jpg" -f $scanned, $evidencePairs.Count)
                        $evB = Join-Path $tempRoot ("ev_{0}_{1}_B.jpg" -f $scanned, $evidencePairs.Count)
                        & $FFmpegPath -hide_banner -loglevel error -i $framePng -vf "scale=640:-1" -q:v 7 -y $evA 2>&1 | Out-Null
                        if (Test-Path -LiteralPath $framePng2) {
                            & $FFmpegPath -hide_banner -loglevel error -i $framePng2 -vf "scale=640:-1" -q:v 7 -y $evB 2>&1 | Out-Null
                        }
                        $evidencePairs += [PSCustomObject]@{
                            TimeA  = $ts
                            TimeB  = $ts2
                            ImageA = if (Test-Path -LiteralPath $evA) { $evA } else { $null }
                            ImageB = if (Test-Path -LiteralPath $evB) { $evB } else { $null }
                            Words  = (($wordsA | Select-Object -First 8) -join ' ')
                        }
                    }
                }
            }

            if ($samplesTaken -eq 0) {
                $failed++
                continue
            }

            $coverage = [math]::Round(100 * $textFrames / $samplesTaken, 0)
            $classification = if ($coverage -ge $FullThresholdPct) { 'FULL' }
                              elseif ($coverage -ge $SuspectThresholdPct) { 'SUSPECT' }
                              else { 'CLEAN' }

            $results += [PSCustomObject]@{
                Folder         = $folder.Name
                VideoPath      = $video.FullName
                SamplesTaken   = $samplesTaken
                TextFrames     = $textFrames
                CoveragePct    = $coverage
                Classification = $classification
                # What OCR actually read on flagged frames — makes every
                # FULL/SUSPECT verdict auditable instead of a bare number.
                SampleText     = ($sampleSnippets -join ' / ')
                Evidence       = $evidencePairs
            }
        }

        # Report must render inside the try — the evidence JPEGs live in
        # $tempRoot, which the finally block deletes.
        if ($HtmlReportPath -and $results.Count -gt 0) {
            try {
                New-HardsubAuditHtmlReport -Results $results -OutPath $HtmlReportPath `
                    -Scanned $scanned -Failed $failed -SampleCount $SampleCount `
                    -FullThresholdPct $FullThresholdPct -SuspectThresholdPct $SuspectThresholdPct
                $reportWritten = $HtmlReportPath
            } catch {
                Write-Host "  HTML report generation failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    } finally {
        # Clear the progress line, then the temp frames.
        $clearWidth = try { [Math]::Max(0, [Console]::WindowWidth - 1) } catch { 0 }
        if ($clearWidth -gt 0) { Write-Host "`r$(' ' * $clearWidth)`r" -NoNewline }
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    return @{
        Scanned         = $scanned
        Failed          = $failed
        SkippedAccepted = $skippedAccepted
        Results         = @($results | Sort-Object CoveragePct -Descending)
        ReportPath      = $reportWritten
    }
}

<#
.SYNOPSIS
    Writes the hardsub audit's self-contained HTML report.
.DESCRIPTION
    One file, images embedded as base64 — portable, no sidecar folder to
    keep in sync. Each FULL/SUSPECT movie shows its flagged frame pairs:
    the frame where OCR found text next to the +2.5s comparison frame.
    Text that changed or vanished = subtitle; text that persisted would
    have been rejected as scene text before reaching this report. That
    lets a human confirm or veto every verdict without opening a player.
#>
function New-HardsubAuditHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object[]]$Results,
        [Parameter(Mandatory)] [string]$OutPath,
        [int]$Scanned = 0,
        [int]$Failed = 0,
        [int]$SampleCount = 0,
        [int]$FullThresholdPct = 50,
        [int]$SuspectThresholdPct = 20
    )

    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }
    $img64 = {
        param($p)
        if ($p -and (Test-Path -LiteralPath $p)) {
            'data:image/jpeg;base64,' + [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($p))
        } else { $null }
    }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine(@'
<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Hardsub Audit</title>
<style>
  body { background:#16181d; color:#d8dce2; font:14px/1.5 "Segoe UI",system-ui,sans-serif; margin:0; padding:24px; }
  h1 { font-size:20px; margin:0 0 4px; }
  h2 { font-size:16px; margin:28px 0 10px; padding-bottom:4px; border-bottom:1px solid #2c3038; }
  h2.full { color:#ff6b6b; } h2.suspect { color:#f0b04c; } h2.clean { color:#5fbf77; }
  .meta { color:#8a919c; margin-bottom:8px; }
  .movie { background:#1d2026; border:1px solid #2c3038; border-radius:8px; padding:14px 16px; margin:12px 0; }
  .movie h3 { margin:0 0 8px; font-size:15px; }
  .cov { font-weight:normal; color:#8a919c; font-size:13px; margin-left:8px; }
  .pair { display:flex; gap:10px; flex-wrap:wrap; margin:10px 0; }
  figure { margin:0; flex:1 1 300px; max-width:640px; }
  figure img { width:100%; border-radius:4px; display:block; }
  figcaption { color:#8a919c; font-size:12px; margin-top:3px; }
  .ocr { color:#b6bdc7; font-size:13px; font-style:italic; }
  table { border-collapse:collapse; margin-top:8px; }
  td, th { padding:3px 14px 3px 0; text-align:left; color:#8a919c; font-size:13px; }
  .verdict-tip { color:#8a919c; font-size:13px; margin:4px 0 0; }
</style></head><body>
'@)

    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm'
    [void]$sb.AppendLine("<h1>Hardsub Audit</h1>")
    [void]$sb.AppendLine("<div class='meta'>$ts &middot; $Scanned scanned, $Failed failed &middot; $SampleCount frames sampled per movie &middot; FULL &ge; $FullThresholdPct% &middot; SUSPECT &ge; $SuspectThresholdPct%</div>")
    [void]$sb.AppendLine("<div class='verdict-tip'>Each pair shows the frame where text was detected beside the same shot ~2.5s later. Changed or vanished text = subtitle. Persistent scene text was already rejected before reaching this page.</div>")

    foreach ($class in 'FULL', 'SUSPECT') {
        $rows = @($Results | Where-Object { $_.Classification -eq $class } | Sort-Object CoveragePct -Descending)
        if ($rows.Count -eq 0) { continue }
        $label = if ($class -eq 'FULL') { "FULL — burned across the film ($($rows.Count))" } else { "SUSPECT — partial burn or text-heavy film ($($rows.Count))" }
        [void]$sb.AppendLine("<h2 class='$($class.ToLower())'>$(& $enc $label)</h2>")
        foreach ($r in $rows) {
            [void]$sb.AppendLine("<div class='movie'>")
            [void]$sb.AppendLine("<h3>$(& $enc $r.Folder)<span class='cov'>$($r.CoveragePct)% coverage ($($r.TextFrames)/$($r.SamplesTaken) frames)</span></h3>")
            foreach ($ev in @($r.Evidence)) {
                if (-not $ev) { continue }
                [void]$sb.AppendLine("<div class='pair'>")
                $srcA = & $img64 $ev.ImageA
                if ($srcA) {
                    [void]$sb.AppendLine("<figure><img src='$srcA' alt='detected frame'><figcaption>$(& $enc $ev.TimeA) — detected: <span class='ocr'>&quot;$(& $enc $ev.Words)&quot;</span></figcaption></figure>")
                }
                $srcB = & $img64 $ev.ImageB
                if ($srcB) {
                    [void]$sb.AppendLine("<figure><img src='$srcB' alt='comparison frame 2.5s later'><figcaption>$(& $enc $ev.TimeB) — same shot ~2.5s later</figcaption></figure>")
                }
                [void]$sb.AppendLine("</div>")
            }
            if (-not @($r.Evidence)) {
                [void]$sb.AppendLine("<div class='ocr'>No evidence frames retained. OCR read: $(& $enc $r.SampleText)</div>")
            }
            [void]$sb.AppendLine("</div>")
        }
    }

    $cleanRows = @($Results | Where-Object { $_.Classification -eq 'CLEAN' } | Sort-Object CoveragePct -Descending)
    if ($cleanRows.Count -gt 0) {
        [void]$sb.AppendLine("<h2 class='clean'>CLEAN ($($cleanRows.Count))</h2>")
        [void]$sb.AppendLine("<table><tr><th>Coverage</th><th>Movie</th></tr>")
        foreach ($r in $cleanRows) {
            [void]$sb.AppendLine("<tr><td>$($r.CoveragePct)%</td><td>$(& $enc $r.Folder)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }

    [void]$sb.AppendLine("</body></html>")

    $outDir = Split-Path $OutPath -Parent
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -Path $outDir -ItemType Directory -Force | Out-Null
    }
    Set-Content -LiteralPath $OutPath -Value $sb.ToString() -Encoding UTF8
}

#endregion

Export-ModuleMember -Function Get-QualityConcerns, Get-QualityScore, Get-VideoCodecInfo, Invoke-CodecAnalysis, Invoke-Transcode, New-TranscodeScript, Remove-CodecSidecarFiles,
    Test-QualityAccepted, Set-QualityAccepted, Remove-QualityAccepted, Get-QualityAcceptedStatus, Test-TesseractInstallation, Invoke-HardsubAudit, New-HardsubAuditHtmlReport
