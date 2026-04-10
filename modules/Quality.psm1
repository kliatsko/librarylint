# Quality.psm1 - Video quality analysis, scoring, and transcoding functions
# Extracted from LibraryLint.ps1 for modularity

# Cache version - bump this when analysis logic changes to invalidate all cached entries
$script:CodecCacheVersion = 1

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

#endregion

#region Public Functions

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
        }
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
        if ($FilePath -and (Test-Path $FilePath)) {
            $quality.FileSize = (Get-Item $FilePath).Length
        }

        # Resolution from actual dimensions
        if ($mediaInfo.Height -ge 2160) {
            $quality.Resolution = "2160p"
            $quality.Score += 100
            $quality.Details += "4K/2160p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+100)"
        }
        elseif ($mediaInfo.Height -ge 1080) {
            $quality.Resolution = "1080p"
            $quality.Score += 80
            $quality.Details += "1080p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+80)"
        }
        elseif ($mediaInfo.Height -ge 720) {
            $quality.Resolution = "720p"
            $quality.Score += 60
            $quality.Details += "720p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+60)"
        }
        elseif ($mediaInfo.Height -ge 480) {
            $quality.Resolution = "480p"
            $quality.Score += 40
            $quality.Details += "480p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+40)"
        }
        elseif ($mediaInfo.Height -gt 0) {
            $quality.Resolution = "$($mediaInfo.Height)p"
            $quality.Score += 20
            $quality.Details += "$($mediaInfo.Height)p [MediaInfo] (+20)"
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
        if ($analysisTextLower -match 'remux') {
            $quality.Source = "Remux"
            $quality.Score += 35
            $quality.Details += "Remux (+35)"
        }
        elseif ($analysisTextLower -match 'bluray|blu-ray|bdrip|brrip') {
            $quality.Source = "BluRay"
            $quality.Score += 30
            $quality.Details += "BluRay (+30)"
        }
        elseif ($analysisTextLower -match 'web-dl|webdl') {
            $quality.Source = "WEB-DL"
            $quality.Score += 25
            $quality.Details += "WEB-DL (+25)"
        }
        elseif ($analysisTextLower -match 'webrip') {
            $quality.Source = "WEBRip"
            $quality.Score += 20
            $quality.Details += "WEBRip (+20)"
        }
        elseif ($analysisTextLower -match 'hdtv') {
            $quality.Source = "HDTV"
            $quality.Score += 15
            $quality.Details += "HDTV (+15)"
        }
        elseif ($analysisTextLower -match 'dvdrip') {
            $quality.Source = "DVDRip"
            $quality.Score += 10
            $quality.Details += "DVDRip (+10)"
        }
        elseif ($analysisTextLower -match 'hdrip') {
            $quality.Source = "HDRip"
            $quality.Score += 8
            $quality.Details += "HDRip (+8)"
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
    if ($analysisTextLower -match 'remux') {
        $quality.Source = "Remux"
        $quality.Score += 35
        $quality.Details += "Remux (+35)"
    }
    elseif ($analysisTextLower -match 'bluray|blu-ray|bdrip|brrip') {
        $quality.Source = "BluRay"
        $quality.Score += 30
        $quality.Details += "BluRay (+30)"
    }
    elseif ($analysisTextLower -match 'web-dl|webdl') {
        $quality.Source = "WEB-DL"
        $quality.Score += 25
        $quality.Details += "WEB-DL (+25)"
    }
    elseif ($analysisTextLower -match 'webrip') {
        $quality.Source = "WEBRip"
        $quality.Score += 20
        $quality.Details += "WEBRip (+20)"
    }
    elseif ($analysisTextLower -match 'hdtv') {
        $quality.Source = "HDTV"
        $quality.Score += 15
        $quality.Details += "HDTV (+15)"
    }
    elseif ($analysisTextLower -match 'dvdrip') {
        $quality.Source = "DVDRip"
        $quality.Score += 10
        $quality.Details += "DVDRip (+10)"
    }
    elseif ($analysisTextLower -match 'hdrip') {
        $quality.Source = "HDRip"
        $quality.Score += 8
        $quality.Details += "HDRip (+8)"
    }

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
    if ($FilePath -and (Test-Path $FilePath)) {
        $quality.FileSize = (Get-Item $FilePath).Length
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
        # Find all video files (exclude trailers - they have low bitrates by design)
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and
                           $_.Name -notmatch '-trailer\.' }

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

        # Interactive quality report menu - loop until user exits
        $continueReports = $true
        while ($continueReports) {
            Write-Host "`n=== Quality Reports ===" -ForegroundColor Yellow
            Write-Host "1. Show lowest quality files (re-acquisition candidates)"
            Write-Host "2. Show highest quality files"
            Write-Host "3. Show files by resolution (ascending)"
            Write-Host "4. Show legacy codec files (XviD, DivX, etc.)"
            Write-Host "5. Show files with quality concerns"
            Write-Host "6. Export full quality report to CSV"
            Write-Host "7. Done with reports"

            $qualityChoice = Read-Host "`nSelect option [7]"

            switch ($qualityChoice) {
            "1" {
                # Lowest quality files
                Write-Host "`n=== Lowest Quality Files (Re-acquisition Candidates) ===" -ForegroundColor Red
                $lowestQuality = $analysis.AllFiles | Sort-Object QualityScore | Select-Object -First 20
                $rank = 1
                foreach ($file in $lowestQuality) {
                    $hdrTag = if ($file.HDR) { " [HDR]" } else { "" }
                    $color = if ($file.QualityScore -lt 50) { "Red" } elseif ($file.QualityScore -lt 80) { "Yellow" } else { "White" }
                    Write-Host "`n  $rank. $($file.FolderName)" -ForegroundColor $color
                    Write-Host "     Score: $($file.QualityScore) | $($file.Resolution) | $($file.Codec)$hdrTag | $(Format-QualitySize $file.Size)" -ForegroundColor Gray
                    $rank++
                }
            }
            "2" {
                # Highest quality files
                Write-Host "`n=== Highest Quality Files ===" -ForegroundColor Green
                $highestQuality = $analysis.AllFiles | Sort-Object QualityScore -Descending | Select-Object -First 20
                $rank = 1
                foreach ($file in $highestQuality) {
                    $hdrTag = if ($file.HDR) { " [HDR]" } else { "" }
                    Write-Host "`n  $rank. $($file.FolderName)" -ForegroundColor Green
                    Write-Host "     Score: $($file.QualityScore) | $($file.Resolution) | $($file.Codec)$hdrTag | $(Format-QualitySize $file.Size)" -ForegroundColor Gray
                    $rank++
                }
            }
            "3" {
                # By resolution ascending (worst first)
                Write-Host "`n=== Files by Resolution (Lowest First) ===" -ForegroundColor Yellow
                $resolutionOrder = @{
                    "Unknown" = 0; "360p" = 1; "480p" = 2; "576p" = 3; "720p" = 4; "1080p" = 5; "2160p" = 6
                }
                $byResolution = $analysis.AllFiles | Sort-Object {
                    $res = $_.Resolution
                    if ($resolutionOrder.ContainsKey($res)) { $resolutionOrder[$res] }
                    else {
                        # Extract numeric value for non-standard resolutions
                        if ($res -match '(\d+)p') { [int]$matches[1] / 1000 }
                        else { 0 }
                    }
                }, QualityScore | Select-Object -First 30

                $currentRes = ""
                foreach ($file in $byResolution) {
                    if ($file.Resolution -ne $currentRes) {
                        $currentRes = $file.Resolution
                        Write-Host "`n  --- $currentRes ---" -ForegroundColor Cyan
                    }
                    $hdrTag = if ($file.HDR) { " [HDR]" } else { "" }
                    Write-Host "    $($file.FolderName) | $($file.Codec)$hdrTag | Score: $($file.QualityScore)" -ForegroundColor Gray
                }
            }
            "4" {
                # Legacy codec files - use regex matching for flexibility
                Write-Host "`n=== Legacy Codec Files ===" -ForegroundColor Yellow
                $legacyFiles = $analysis.AllFiles | Where-Object {
                    $_.Codec -match 'XviD|DivX|MPEG-4|VC-1|WMV|MPEG4' -or
                    $_.Container -eq "AVI"
                } | Sort-Object QualityScore

                if ($legacyFiles.Count -eq 0) {
                    Write-Host "  No legacy codec files found!" -ForegroundColor Green
                } else {
                    Write-Host "  Found $($legacyFiles.Count) file(s) with legacy codecs:" -ForegroundColor White
                    foreach ($file in $legacyFiles) {
                        Write-Host "`n    $($file.FolderName)" -ForegroundColor Yellow
                        Write-Host "      $($file.Resolution) | $($file.Codec) | $($file.Container) | Score: $($file.QualityScore)" -ForegroundColor Gray
                    }

                    # Offer to transcode
                    Write-Host ""
                    $transcodeChoice = Read-Host "Would you like to transcode these files to H.264? (Y/N) [N]"
                    if ($transcodeChoice -match '^[Yy]') {
                        # Ensure TranscodeMode is set for files that need it
                        $transcodeQueue = $legacyFiles | ForEach-Object {
                            $file = $_
                            if (-not $file.TranscodeMode -or $file.TranscodeMode -eq "none") {
                                $file.TranscodeMode = "transcode"
                            }
                            $file
                        }

                        $totalSize = ($transcodeQueue | Measure-Object -Property Size -Sum).Sum
                        Write-Host "`nThis will transcode $($transcodeQueue.Count) file(s) ($(Format-QualitySize $totalSize))." -ForegroundColor Yellow
                        Write-Host "Original files will be replaced after successful conversion." -ForegroundColor Yellow
                        $confirm = Read-Host "Are you sure you want to proceed? (Y/N) [N]"
                        if ($confirm -match '^[Yy]') {
                            Invoke-Transcode -TranscodeQueue $transcodeQueue
                        } else {
                            Write-Host "Transcode cancelled" -ForegroundColor Gray
                        }
                    }
                }
            }
            "5" {
                # Files with quality concerns
                Write-Host "`n=== Files with Quality Concerns ===" -ForegroundColor Yellow
                $concernFiles = $analysis.AllFiles | Where-Object { $_.HasConcerns } | Sort-Object QualityScore

                if ($concernFiles.Count -eq 0) {
                    Write-Host "  No quality concerns found!" -ForegroundColor Green
                } else {
                    Write-Host "  Found $($concernFiles.Count) file(s) with potential quality issues:" -ForegroundColor White
                    foreach ($file in $concernFiles) {
                        $bitrateMbps = if ($file.Bitrate -gt 0) { "$([math]::Round($file.Bitrate / 1000000, 1)) Mbps" } else { "N/A" }
                        Write-Host "`n    $($file.FolderName)" -ForegroundColor Yellow
                        Write-Host "      $($file.Resolution) | $($file.Codec) | $bitrateMbps | $(Format-QualitySize $file.Size)" -ForegroundColor Gray
                        foreach ($concern in $file.QualityConcerns) {
                            Write-Host "      ! $concern" -ForegroundColor Red
                        }
                    }
                }
            }
            "6" {
                # Export full quality report
                $qualityExportPath = Join-Path $ReportsFolder "QualityReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                $analysis.AllFiles | Sort-Object QualityScore |
                    Select-Object FolderName, FileName, QualityScore, Resolution, Codec, AudioCodec, Container, HDR, @{N='SizeMB';E={[math]::Round($_.Size/1MB,2)}}, @{N='BitrateMbps';E={if($_.Bitrate -gt 0){[math]::Round($_.Bitrate/1000000,1)}else{'N/A'}}}, @{N='QualityConcerns';E={$_.QualityConcerns -join '; '}}, NeedsTranscode, TranscodeReason, Path |
                    Export-Csv -Path $qualityExportPath -NoTypeInformation -Encoding UTF8
                Write-Host "`nQuality report exported to: $qualityExportPath" -ForegroundColor Green
                Write-Host "  Quality report exported to: $qualityExportPath" -ForegroundColor Gray
            }
            default {
                $continueReports = $false
            }
        }
        } # End while loop

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

        if ($mode -eq "remux") {
            Write-Host "[$current/$total] Remuxing: $($item.FileName)" -ForegroundColor Cyan
            $ffmpegArgs = "-i `"$inputPath`" -c:v copy -c:a copy -c:s copy `"$tempOutput`" -y"
        } else {
            Write-Host "[$current/$total] Transcoding: $($item.FileName)" -ForegroundColor Yellow
            Write-Host "  Reason: $($item.TranscodeReason -join '; ')" -ForegroundColor Gray
            # Use H.265 for VC-1 (HD content benefits from HEVC), H.264 for everything else
            $codec = if ($item.Codec -match 'VC-1|WMV') { "libx265" } else { $TargetCodec }
            $crf = if ($codec -eq "libx265") { "28" } else { "23" }
            $ffmpegArgs = "-i `"$inputPath`" -map 0:v -map 0:a -map 0:s? -c:v $codec -crf $crf -preset medium -c:a aac -b:a 192k -c:s copy `"$tempOutput`" -y"
        }

        # Run FFmpeg
        $process = Start-Process -FilePath $FFmpegPath -ArgumentList $ffmpegArgs -NoNewWindow -Wait -PassThru

        if ($process.ExitCode -eq 0 -and (Test-Path $tempOutput)) {
            $outSize = (Get-Item $tempOutput).Length
            if ($outSize -gt 0) {
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

                # Delete original and rename temp to final
                Remove-Item -LiteralPath $inputPath -Force
                Rename-Item -LiteralPath $tempOutput -NewName ([System.IO.Path]::GetFileName($outputPath))

                if ($mode -eq "remux") {
                    Write-Host "  Remuxed successfully" -ForegroundColor Green
                    $remuxed++
                } else {
                    Write-Host "  Transcoded successfully" -ForegroundColor Green
                    $transcoded++
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
        & `$ffmpegPath -i "`$(`$file.Input)" -c:v copy -c:a copy -c:s copy "`$tempOutput" -y

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
        & `$ffmpegPath -i "`$(`$file.Input)" -map 0:v -map 0:a -map 0:s? -c:v $TargetCodec -crf 23 -preset medium $resolutionParam -c:a aac -b:a 192k -c:s copy "`$tempOutput" -y

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
    param([string]$Path)
    $removed = 0
    Get-ChildItem -Path $Path -Recurse -Filter "codec-info.json" -File -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
            $removed++
        } catch { }
    }
    return $removed
}

#endregion

Export-ModuleMember -Function Get-QualityConcerns, Get-QualityScore, Get-VideoCodecInfo, Invoke-CodecAnalysis, Invoke-Transcode, New-TranscodeScript, Remove-CodecSidecarFiles
