# Subtitles.psm1
# Subtitle detection, download, sync, and repair functions
# Handles external/embedded subtitle checking, SubDL API, ffsubsync, and orphan repair
# Part of LibraryLint suite

#region Public Functions

function Test-SubtitlesExist {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,
        [string]$VideoPath,
        [string]$Language = "en",
        [string]$MediaInfoPath
    )

    $result = @{
        HasSubtitles = $false
        ExternalSubs = @()
        EmbeddedSubs = @()
        EmbeddedEnglish = $false
        Details = ""
    }

    # Check for external subtitle files
    $subtitleExtensions = @('.srt', '.sub', '.ass', '.ssa', '.vtt', '.idx')
    $externalSubs = Get-ChildItem -LiteralPath $FolderPath -File -ErrorAction SilentlyContinue |
        Where-Object { $subtitleExtensions -contains $_.Extension.ToLower() }

    if ($externalSubs) {
        # Filter by language if specified
        if ($Language -ne "*") {
            $langPattern = "\.$Language\.|[\._-]$Language[\._-]|^$Language[\._-]"
            $matchedSubs = $externalSubs | Where-Object { $_.Name -match $langPattern -or $_.Name -notmatch '\.(en|es|fr|de|it|pt|ru|ja|ko|zh|nl|pl|sv|da|no|fi)\.' }
        } else {
            $matchedSubs = $externalSubs
        }

        if ($matchedSubs) {
            $result.ExternalSubs = @($matchedSubs)
            $result.HasSubtitles = $true
            $result.Details = "External: $($matchedSubs.Count) file(s)"
        }
    }

    # Check for embedded subtitles in video file
    if (-not $VideoPath) {
        $VideoPath = Get-ChildItem -LiteralPath $FolderPath -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '\.(mkv|mp4|avi|mov|wmv|m4v)$' } |
            Sort-Object Length -Descending |
            Select-Object -First 1 -ExpandProperty FullName
    }

    if ($VideoPath -and (Test-MediaInfoInstallation)) {
        try {
            $jsonOutput = & $MediaInfoPath --Output=JSON "$VideoPath" 2>$null
            if ($jsonOutput) {
                $mediaData = $jsonOutput | ConvertFrom-Json
                if ($mediaData.media -and $mediaData.media.track) {
                    $textTracks = @($mediaData.media.track | Where-Object { $_.'@type' -eq 'Text' })

                    if ($textTracks.Count -gt 0) {
                        $result.EmbeddedSubs = $textTracks
                        $result.HasSubtitles = $true

                        # Check for English specifically
                        $englishTracks = $textTracks | Where-Object {
                            $_.Language -match '^en' -or
                            $_.Title -match 'english' -or
                            (-not $_.Language -and -not $_.Title)  # Assume unlabeled is English
                        }
                        $result.EmbeddedEnglish = ($englishTracks.Count -gt 0)

                        if ($result.Details) {
                            $result.Details += ", Embedded: $($textTracks.Count) track(s)"
                        } else {
                            $result.Details = "Embedded: $($textTracks.Count) track(s)"
                        }
                    }
                }
            }
        }
        catch {
            # Debug: Error checking embedded subtitles
        }
    }

    return $result
}

<#
.SYNOPSIS
    Checks if a movie folder has verified/confirmed good subtitles
.PARAMETER FolderPath
    Path to the movie folder
.OUTPUTS
    Boolean indicating if subtitles are verified
#>
function Test-SubtitlesVerified {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )

    $verifiedFile = Join-Path $FolderPath ".subs_ok"
    return (Test-Path -LiteralPath $verifiedFile)
}

<#
.SYNOPSIS
    Marks a movie folder as having verified good subtitles
.PARAMETER FolderPath
    Path to the movie folder
.PARAMETER Source
    Source of the subtitles (e.g., "embedded", "included", "subdl", "manual")
#>
function Set-SubtitlesVerified {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,
        [string]$Source = "unknown"
    )

    $verifiedFile = Join-Path $FolderPath ".subs_ok"
    $content = @{
        VerifiedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Source = $Source
        VerifiedBy = "LibraryLint"
    } | ConvertTo-Json

    try {
        $content | Out-File -LiteralPath $verifiedFile -Encoding UTF8 -Force
        return $true
    }
    catch {
        Write-Host "Failed to create .subs_ok file: $_" -ForegroundColor Yellow
        return $false
    }
}

<#
.SYNOPSIS
    Removes the verified subtitle marker from a folder
.PARAMETER FolderPath
    Path to the movie folder
#>
function Remove-SubtitlesVerified {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )

    $verifiedFile = Join-Path $FolderPath ".subs_ok"
    if (Test-Path -LiteralPath $verifiedFile) {
        Remove-Item -LiteralPath $verifiedFile -Force -ErrorAction SilentlyContinue
    }
}

<#
.SYNOPSIS
    Checks if ffsubsync is installed
.OUTPUTS
    Boolean indicating if ffsubsync is available
#>
function Test-FFSubSyncInstallation {
    $ffsubsync = Get-Command "ffsubsync" -ErrorAction SilentlyContinue
    if ($ffsubsync) {
        return $true
    }

    # Also check common pip install locations
    $pythonScripts = @(
        "$env:LOCALAPPDATA\Programs\Python\Python*\Scripts\ffsubsync.exe",
        "$env:APPDATA\Python\Python*\Scripts\ffsubsync.exe",
        "$env:USERPROFILE\.local\bin\ffsubsync"
    )

    foreach ($pattern in $pythonScripts) {
        $found = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) {
            return $true
        }
    }

    return $false
}

<#
.SYNOPSIS
    Syncs a subtitle file to match the audio of a video using ffsubsync
.PARAMETER VideoPath
    Path to the video file
.PARAMETER SubtitlePath
    Path to the subtitle file to sync
.PARAMETER OutputPath
    Optional: Output path for synced subtitle. If not specified, overwrites original.
.OUTPUTS
    Boolean indicating success
#>
function Invoke-FFSubSync {
    param(
        [Parameter(Mandatory=$true)]
        [string]$VideoPath,
        [Parameter(Mandatory=$true)]
        [string]$SubtitlePath,
        [string]$OutputPath,
        [switch]$Backup
    )

    if (-not (Test-FFSubSyncInstallation)) {
        return $false
    }

    if (-not $OutputPath) {
        $OutputPath = $SubtitlePath
    }

    try {
        Write-Host "    Syncing subtitle timing..." -ForegroundColor Gray

        # Create backup if requested and we're overwriting
        if ($Backup -and $OutputPath -eq $SubtitlePath) {
            $backupPath = $SubtitlePath + ".bak"
            Copy-Item -LiteralPath $SubtitlePath -Destination $backupPath -Force
        }

        # Create temp output if overwriting
        $tempOutput = $null
        if ($OutputPath -eq $SubtitlePath) {
            $tempOutput = Join-Path $env:TEMP "ffsubsync_$(Get-Random).srt"
            $actualOutput = $tempOutput
        } else {
            $actualOutput = $OutputPath
        }

        # Run ffsubsync
        $process = Start-Process -FilePath "ffsubsync" -ArgumentList "`"$VideoPath`"", "-i", "`"$SubtitlePath`"", "-o", "`"$actualOutput`"" -Wait -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\ffsubsync_out.txt" -RedirectStandardError "$env:TEMP\ffsubsync_err.txt"

        if ($process.ExitCode -eq 0 -and (Test-Path $actualOutput)) {
            # If we used a temp file, move it to the final location
            if ($tempOutput) {
                Move-Item -Path $tempOutput -Destination $OutputPath -Force
            }

            Write-Host "    Subtitle synced successfully" -ForegroundColor Green
            Write-Host "  Synced subtitle: $SubtitlePath" -ForegroundColor Gray
            return $true
        } else {
            $errorOutput = Get-Content "$env:TEMP\ffsubsync_err.txt" -Raw -ErrorAction SilentlyContinue

            # Parse error for common failure reasons
            $failureReason = "Unknown error"
            if ($errorOutput -match "failed to sync") {
                $failureReason = "Speech detection failed - audio may be incompatible"
            } elseif ($errorOutput -match "codec" -or $errorOutput -match "audio") {
                $failureReason = "Audio codec not supported"
            } elseif ($errorOutput -match "encoding" -or $errorOutput -match "decode") {
                $failureReason = "Subtitle encoding issue"
            } elseif ($errorOutput -match "memory" -or $errorOutput -match "MemoryError") {
                $failureReason = "Out of memory - file may be too large"
            } elseif ($errorOutput -match "permission" -or $errorOutput -match "access") {
                $failureReason = "Permission denied"
            } elseif (-not [string]::IsNullOrWhiteSpace($errorOutput)) {
                # Extract last meaningful error line
                $errorLines = $errorOutput -split "`n" | Where-Object { $_ -match "ERROR|error|failed" } | Select-Object -Last 1
                if ($errorLines) {
                    $failureReason = $errorLines.Trim()
                }
            }

            Write-Host "    Subtitle sync failed: $failureReason" -ForegroundColor Yellow
            Write-Host "ffsubsync failed for $SubtitlePath" -ForegroundColor Yellow
            Write-Host "  Reason: $failureReason" -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        Write-Host "Error running ffsubsync: $_" -ForegroundColor Yellow
        return $false
    }
    finally {
        # Cleanup temp files
        Remove-Item "$env:TEMP\ffsubsync_out.txt" -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:TEMP\ffsubsync_err.txt" -Force -ErrorAction SilentlyContinue
        if ($tempOutput -and (Test-Path $tempOutput)) {
            Remove-Item $tempOutput -Force -ErrorAction SilentlyContinue
        }
    }
}

<#
.SYNOPSIS
    Searches Subdl.com for subtitles by IMDB ID or movie title
.PARAMETER IMDBID
    The IMDB ID of the movie (e.g., tt0107290)
.PARAMETER Title
    The movie title (fallback if no IMDB ID)
.PARAMETER Year
    The movie year (used with title search)
.PARAMETER Language
    Language code (e.g., en, es, fr, de)
.OUTPUTS
    Array of subtitle results or $null
#>
function Search-SubdlSubtitle {
    param(
        [string]$IMDBID,
        [string]$Title,
        [string]$Year,
        [string]$Language = "en",
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        Write-Host "Subdl API key not provided" -ForegroundColor Yellow
        return $null
    }

    try {
        $baseUrl = "https://api.subdl.com/api/v1/subtitles"
        $results = $null

        # Try IMDB ID first (most accurate)
        if ($IMDBID) {
            $url = "$baseUrl`?api_key=$ApiKey&imdb_id=$IMDBID&languages=$Language&subs_per_page=10"

            $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

            if ($response.status -and $response.subtitles -and $response.subtitles.Count -gt 0) {
                $results = $response.subtitles
            }
        }

        # Fallback to title search
        if (-not $results -and $Title) {
            $encodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)
            $url = "$baseUrl`?api_key=$ApiKey&film_name=$encodedTitle&languages=$Language&subs_per_page=10"
            if ($Year) {
                $url += "&year=$Year"
            }

            $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

            if ($response.status -and $response.subtitles -and $response.subtitles.Count -gt 0) {
                $results = $response.subtitles
            }
        }

        if ($results) {
            return $results
        }

        return $null
    }
    catch {
        Write-Host "Error searching Subdl: $_" -ForegroundColor Yellow
        return $null
    }
}

<#
.SYNOPSIS
    Downloads and saves a subtitle from Subdl.com
.PARAMETER SubtitleUrl
    The URL to download the subtitle from
.PARAMETER MovieFolder
    The folder to save the subtitle to
.PARAMETER MovieTitle
    The movie title (used for naming the file)
.PARAMETER Language
    Language code for the subtitle filename
.OUTPUTS
    Boolean indicating success
#>
function Save-MovieSubtitle {
    param(
        [string]$IMDBID,
        [string]$Title,
        [string]$Year,
        [string]$MovieFolder,
        [string]$MovieTitle,
        [string]$Language = "en",
        [string]$ApiKey,
        [switch]$AutoSyncSubtitles
    )

    # Check if API key is available
    if (-not $ApiKey) {
        Write-Host "Subdl API key not configured" -ForegroundColor Yellow
        return $false
    }

    # Check if subtitles already verified
    if (Test-SubtitlesVerified -FolderPath $MovieFolder) {
        return $true
    }

    # Check for existing subtitles (external or embedded)
    $subCheck = Test-SubtitlesExist -FolderPath $MovieFolder -Language $Language

    if ($subCheck.HasSubtitles) {
        # Has subtitles - mark as verified (came with the release)
        if ($subCheck.ExternalSubs.Count -gt 0) {
            Set-SubtitlesVerified -FolderPath $MovieFolder -Source "included"
        } elseif ($subCheck.EmbeddedEnglish -or $subCheck.EmbeddedSubs.Count -gt 0) {
            Set-SubtitlesVerified -FolderPath $MovieFolder -Source "embedded"
        }
        return $true
    }

    # Rate limiting: 1 request per second for Subdl
    Start-Sleep -Seconds 1

    # Search for subtitles
    $subtitles = Search-SubdlSubtitle -IMDBID $IMDBID -Title $Title -Year $Year -Language $Language -ApiKey $ApiKey

    if (-not $subtitles) {
        return $false
    }

    # Get the best subtitle (first result, sorted by Subdl's ranking)
    $subtitle = $subtitles | Select-Object -First 1

    if (-not $subtitle.url) {
        Write-Host "Subtitle has no download URL" -ForegroundColor Yellow
        return $false
    }

    try {
        Write-Host "    Downloading subtitle ($Language)..." -ForegroundColor Gray

        # Create temp folder for extraction
        $tempFolder = Join-Path $env:TEMP "LibraryLint_Sub_$(Get-Random)"
        New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null

        # Download the subtitle zip
        $zipPath = Join-Path $tempFolder "subtitle.zip"
        $downloadUrl = $subtitle.url

        # Subdl returns a path, construct full URL
        if (-not $downloadUrl.StartsWith("http")) {
            $downloadUrl = "https://dl.subdl.com$downloadUrl"
        }

        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -ErrorAction Stop

        # Extract the zip
        Expand-Archive -Path $zipPath -DestinationPath $tempFolder -Force

        # Find SRT file in extracted content
        $srtFiles = Get-ChildItem -Path $tempFolder -Filter "*.srt" -Recurse -ErrorAction SilentlyContinue

        if ($srtFiles) {
            # Use the first/largest SRT file
            $srtFile = $srtFiles | Sort-Object Length -Descending | Select-Object -First 1

            # Clean movie title for filename
            $cleanTitle = $MovieTitle -replace '[\\/:*?"<>|]', ''
            $destPath = Join-Path $MovieFolder "$cleanTitle.$Language.srt"

            Copy-Item -LiteralPath $srtFile.FullName -Destination $destPath -Force

            Write-Host "    Subtitle downloaded ($Language)" -ForegroundColor Green
            Write-Host "  Downloaded subtitle for $MovieTitle from SubDL" -ForegroundColor Gray

            # Run ffsubsync to fix timing (SubDL subs often have sync issues)
            if ($AutoSyncSubtitles) {
                $videoFile = Get-ChildItem -LiteralPath $MovieFolder -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.Extension -match '\.(mkv|mp4|avi|mov|wmv|m4v)$' } |
                    Sort-Object Length -Descending |
                    Select-Object -First 1

                if ($videoFile -and (Test-FFSubSyncInstallation)) {
                    Write-Host "    Syncing subtitle timing..." -ForegroundColor Gray
                    $syncResult = Invoke-FFSubSync -VideoPath $videoFile.FullName -SubtitlePath $destPath
                    if ($syncResult) {
                        # Mark as verified after successful sync
                        Set-SubtitlesVerified -FolderPath $MovieFolder -Source "subdl-synced"
                    } else {
                        Write-Host "Subtitle downloaded but sync failed - timing may be off" -ForegroundColor Yellow
                    }
                } elseif (-not (Test-FFSubSyncInstallation)) {
                    Write-Host "    ffsubsync not installed - timing may be off" -ForegroundColor Yellow
                    Write-Host "    Install with: pip install ffsubsync" -ForegroundColor DarkGray
                }
            }

            # Cleanup
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue

            return $true
        } else {
            Write-Host "No SRT file found in downloaded archive" -ForegroundColor Yellow
        }

        # Cleanup
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        return $false
    }
    catch {
        Write-Host "Failed to download subtitle for $MovieTitle : $_" -ForegroundColor Yellow
        # Cleanup on error
        if (Test-Path $tempFolder) {
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        return $false
    }
}

function Repair-OrphanedSubtitles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [switch]$WhatIf,
        [string[]]$VideoExtensions = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.m4v'),
        [string[]]$SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt')
    )

    Write-Host "`n--- Fix Orphaned Subtitles ---" -ForegroundColor Yellow
    if ($WhatIf) {
        Write-Host "(Dry run - no changes will be made)" -ForegroundColor Cyan
    }
    Write-Host "  Starting orphaned subtitle repair in: $Path (WhatIf: $WhatIf)" -ForegroundColor Gray

    $stats = @{
        Total = 0
        Renamed = 0
        Deleted = 0
        NoVideo = 0
        MultipleVideos = 0
        AlreadyHasMatch = 0
        Errors = 0
    }

    # Get all subtitle files
    $subtitleFiles = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $SubtitleExtensions -contains $_.Extension.ToLower() }

    Write-Host "Scanning $($subtitleFiles.Count) subtitle files..." -ForegroundColor Cyan

    foreach ($sub in $subtitleFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sub.Name)
        # Extract language suffix if present
        $langSuffix = ""
        if ($baseName -match '\.(eng|en|english|spa|es|spanish|fre|fr|french|ger|de|german|por|pt|portuguese|ita|it|italian|rus|ru|russian|jpn|ja|japanese|chi|zh|chinese|kor|ko|korean|ara|ar|arabic|hin|hi|hindi|dut|nl|dutch|swe|sv|swedish|nor|no|norwegian|dan|da|danish|fin|fi|finnish|pol|pl|polish|tur|tr|turkish|heb|he|hebrew|tha|th|thai|vie|vi|vietnamese|ind|id|indonesian|msa|ms|malay|forced)$') {
            $langSuffix = $matches[0]
            $baseName = $baseName -replace '\.(eng|en|english|spa|es|spanish|fre|fr|french|ger|de|german|por|pt|portuguese|ita|it|italian|rus|ru|russian|jpn|ja|japanese|chi|zh|chinese|kor|ko|korean|ara|ar|arabic|hin|hi|hindi|dut|nl|dutch|swe|sv|swedish|nor|no|norwegian|dan|da|danish|fin|fi|finnish|pol|pl|polish|tur|tr|turkish|heb|he|hebrew|tha|th|thai|vie|vi|vietnamese|ind|id|indonesian|msa|ms|malay|forced)$', ''
        }

        # Check if this subtitle matches a video file
        $hasMatchingVideo = Get-ChildItem -Path $sub.DirectoryName -File -ErrorAction SilentlyContinue |
            Where-Object {
                $VideoExtensions -contains $_.Extension.ToLower() -and
                [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -eq $baseName
            } | Select-Object -First 1

        if ($hasMatchingVideo) {
            # Not orphaned, skip
            continue
        }

        $stats.Total++

        # Find all video files in this folder
        $videoFiles = @(Get-ChildItem -Path $sub.DirectoryName -File -ErrorAction SilentlyContinue |
            Where-Object { $VideoExtensions -contains $_.Extension.ToLower() } |
            Where-Object { $_.Name -notmatch 'sam?ple|sampe|smaple|preview|trailer|teaser' })

        if ($videoFiles.Count -eq 0) {
            # No video file - delete the orphaned subtitle
            $action = if ($WhatIf) { '[would delete]' } else { '[deleted] ' }
            Write-Host "  $action $($sub.Name) — no matching video" -ForegroundColor Gray
            if (-not $WhatIf) {
                try {
                    Remove-Item -LiteralPath $sub.FullName -Force
                } catch {
                    Write-Host "    error: $_" -ForegroundColor Red
                    $stats.Errors++
                }
            }
            $stats.Deleted++
            continue
        }

        if ($videoFiles.Count -gt 1) {
            # Multiple videos - can't determine which one to match
            Write-Host "  [skip]         $($sub.Name) — multiple videos in folder" -ForegroundColor Yellow
            $stats.MultipleVideos++
            continue
        }

        # Exactly one video - rename subtitle to match
        $video = $videoFiles[0]
        $videoBaseName = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)

        # Check if a matching subtitle already exists for this video
        $existingMatch = Get-ChildItem -Path $sub.DirectoryName -File -ErrorAction SilentlyContinue |
            Where-Object {
                $SubtitleExtensions -contains $_.Extension.ToLower() -and
                $_.FullName -ne $sub.FullName -and
                [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -match "^$([regex]::Escape($videoBaseName))"
            } | Select-Object -First 1

        if ($existingMatch) {
            # Video already has a matching subtitle - delete this orphaned one
            $action = if ($WhatIf) { '[would delete]' } else { '[deleted] ' }
            Write-Host "  $action $($sub.Name) — video already has $($existingMatch.Name)" -ForegroundColor Gray
            if (-not $WhatIf) {
                try {
                    Remove-Item -LiteralPath $sub.FullName -Force
                } catch {
                    Write-Host "    error: $_" -ForegroundColor Red
                    $stats.Errors++
                }
            }
            $stats.AlreadyHasMatch++
            continue
        }

        # Build new filename
        $newName = $videoBaseName + $langSuffix + $sub.Extension
        $action = if ($WhatIf) { '[would rename]' } else { '[renamed] ' }
        Write-Host "  $action $($sub.Name) -> $newName" -ForegroundColor Green

        if (-not $WhatIf) {
            try {
                Rename-Item -LiteralPath $sub.FullName -NewName $newName -Force
                $stats.Renamed++
            } catch {
                Write-Host "    error: $_" -ForegroundColor Red
                $stats.Errors++
            }
        } else {
            $stats.Renamed++
        }
    }

    # One-line summary for standalone callers. Repair-SubtitlePlacement also
    # prints a fuller consolidated block; one duplicated count line is the
    # accepted trade-off for keeping this function self-sufficient.
    $verb = if ($WhatIf) { 'would ' } else { '' }
    $deleted = $stats.Deleted + $stats.AlreadyHasMatch
    Write-Host ("  Orphaned: {0} found, {1}rename {2}, {1}delete {3}, skip {4}, errors {5}" -f `
        $stats.Total, $verb, $stats.Renamed, $deleted, $stats.MultipleVideos, $stats.Errors) -ForegroundColor Gray

    return $stats
}

<#
.SYNOPSIS
    Finds and fixes subtitle files that are buried in subfolders or misnamed
.DESCRIPTION
    Scans a library for subtitle files in Subs/Sub/Subtitles subfolders,
    moves them alongside the video file with proper naming for player auto-detection,
    then runs orphaned subtitle repair for any remaining mismatches.
.PARAMETER Path
    The root path of the media library
.PARAMETER WhatIf
    If set, shows what would be done without making changes
#>
function Repair-SubtitlePlacement {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [switch]$WhatIf,
        [string[]]$VideoExtensions = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.m4v'),
        [string[]]$SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt'),
        [string[]]$PreferredSubtitleLanguages = @('eng', 'en', 'english'),
        [bool]$KeepSubtitles = $true
    )

    Write-Host "`n=== Repair Subtitle Placement ===" -ForegroundColor Cyan
    if ($WhatIf) {
        Write-Host "(Dry run - no changes will be made)" -ForegroundColor Yellow
    }
    Write-Host "  Starting subtitle placement repair in: $Path (WhatIf: $WhatIf)" -ForegroundColor Gray

    $stats = @{
        SubfoldersFound = 0
        SubsMoved = 0
        SubsSkipped = 0
        SubsDeleted = 0
        FoldersRemoved = 0
        OrphanedFixed = 0
        OrphansDeletedNoVideo = 0
        OrphansDeletedHadMatch = 0
        OrphansMultipleVideos = 0
        Errors = 0
    }

    # Language detection helper (maps subtitle name patterns to ISO 639-1 codes)
    $langMap = [ordered]@{
        '[\._](eng|en|english)'    = '.en'
        '[\._](spa|es|spanish)'    = '.es'
        '[\._](fre|fr|french)'     = '.fr'
        '[\._](ger|de|german)'     = '.de'
        '[\._](ita|it|italian)'    = '.it'
        '[\._](por|pt|portuguese)' = '.pt'
        '[\._](rus|ru|russian)'    = '.ru'
        '[\._](jpn|ja|japanese)'   = '.ja'
        '[\._](chi|zh|chinese)'    = '.zh'
        '[\._](kor|ko|korean)'     = '.ko'
        '[\._](ara|ar|arabic)'     = '.ar'
        '[\._](hin|hi|hindi)'      = '.hi'
        '[\._](dut|nl|dutch)'      = '.nl'
        '[\._](pol|pl|polish)'     = '.pl'
        '[\._](swe|sv|swedish)'    = '.sv'
        '[\._](nor|no|norwegian)'  = '.no'
        '[\._](dan|da|danish)'     = '.da'
        '[\._](fin|fi|finnish)'    = '.fi'
        '[\._](tur|tr|turkish)'    = '.tr'
        '[\._](heb|he|hebrew)'     = '.he'
        '[\._](tha|th|thai)'       = '.th'
        '[\._](vie|vi|vietnamese)' = '.vi'
        '[\._](ind|id|indonesian)' = '.id'
        '[\._](msa|ms|malay)'     = '.ms'
    }

    # Step 1: Find subtitle subfolders and move subs up to video level
    Write-Host "`n--- Step 1: Fix subtitles in subfolders ---" -ForegroundColor Yellow

    $subFolderNames = @('Subs', 'Sub', 'Subtitles', 'Subtitle', 'SRT', 'subs', 'sub', 'subtitles')
    $subFolders = Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -in $subFolderNames }

    if ($subFolders.Count -eq 0) {
        Write-Host "  No subtitle subfolders found" -ForegroundColor Gray
    } else {
        Write-Host "  Found $($subFolders.Count) subtitle subfolder(s)" -ForegroundColor Cyan

        foreach ($subFolder in $subFolders) {
            $stats.SubfoldersFound++
            $movieFolder = $subFolder.Parent.FullName

            # Find the main video file in the parent folder
            $videoFile = Get-ChildItem -LiteralPath $movieFolder -File -ErrorAction SilentlyContinue |
                Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and $_.Name -notmatch '-trailer\.' } |
                Sort-Object Length -Descending | Select-Object -First 1

            if (-not $videoFile) {
                Write-Host "  [skip]         $($subFolder.Parent.Name)\$($subFolder.Name) — parent folder has no video" -ForegroundColor DarkGray
                continue
            }

            $videoBaseName = [System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)

            # Get all subtitle files in this subfolder (and nested subfolders)
            $subtitleFiles = Get-ChildItem -LiteralPath $subFolder.FullName -File -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $SubtitleExtensions -contains $_.Extension.ToLower() }

            foreach ($sub in $subtitleFiles) {
                $subNameLower = $sub.BaseName.ToLower()

                # Detect language
                $detectedLang = ""
                foreach ($pattern in $langMap.Keys) {
                    if ($subNameLower -match $pattern) {
                        $detectedLang = $langMap[$pattern]
                        break
                    }
                }

                # Check if it's a preferred language
                $isPreferred = $false
                foreach ($lang in $PreferredSubtitleLanguages) {
                    if ($subNameLower -match "[\._]$lang") {
                        $isPreferred = $true
                        break
                    }
                }
                # No language tag = assume preferred
                $hasLangTag = $subNameLower -match '[\._](eng|en|english|spa|es|spanish|fre|fr|french|ger|de|german|ita|it|italian|por|pt|portuguese|rus|ru|russian|jpn|ja|japanese|chi|zh|chinese|kor|ko|korean|ara|ar|arabic|hin|hi|hindi|dut|nl|dutch|pol|pl|polish|swe|sv|swedish|nor|no|norwegian|dan|da|danish|fin|fi|finnish|tur|tr|turkish|heb|he|hebrew|tha|th|thai|vie|vi|vietnamese|ind|id|indonesian|msa|ms|malay|forced)'
                if (-not $hasLangTag) { $isPreferred = $true }

                if (-not $isPreferred -and -not $KeepSubtitles) {
                    $action = if ($WhatIf) { '[would delete]' } else { '[deleted] ' }
                    Write-Host "  $action $($sub.Name) — non-preferred language" -ForegroundColor Gray
                    if (-not $WhatIf) {
                        Remove-Item -LiteralPath $sub.FullName -Force -ErrorAction SilentlyContinue
                    }
                    $stats.SubsDeleted++
                    continue
                }

                $newSubName = "$videoBaseName$detectedLang$($sub.Extension)"
                $newSubPath = Join-Path $movieFolder $newSubName

                if (Test-Path -LiteralPath $newSubPath) {
                    Write-Host "  [skip]         $($sub.Name) — destination $newSubName already exists" -ForegroundColor DarkGray
                    $stats.SubsSkipped++
                    continue
                }

                $action = if ($WhatIf) { '[would move]  ' } else { '[moved]   ' }
                Write-Host "  $action $($subFolder.Name)\$($sub.Name) -> $newSubName" -ForegroundColor Green
                if (-not $WhatIf) {
                    try {
                        Move-Item -LiteralPath $sub.FullName -Destination $newSubPath -Force
                    }
                    catch {
                        Write-Host "    error: $_" -ForegroundColor Red
                        $stats.Errors++
                        continue
                    }
                }
                $stats.SubsMoved++
            }

            # Clean up empty subtitle subfolder
            if (-not $WhatIf) {
                $remaining = Get-ChildItem -LiteralPath $subFolder.FullName -Recurse -File -ErrorAction SilentlyContinue
                if (-not $remaining -or $remaining.Count -eq 0) {
                    Remove-Item -LiteralPath $subFolder.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    $stats.FoldersRemoved++
                }
            }
        }
    }

    # Step 2: Fix orphaned/misnamed subtitles already at video level
    Write-Host "`n--- Step 2: Fix misnamed subtitles ---" -ForegroundColor Yellow
    $orphanStats = Repair-OrphanedSubtitles -Path $Path -WhatIf:$WhatIf -VideoExtensions $VideoExtensions -SubtitleExtensions $SubtitleExtensions
    if ($orphanStats) {
        $stats.OrphanedFixed = $orphanStats.Renamed
        $stats.OrphansDeletedNoVideo = $orphanStats.Deleted
        $stats.OrphansDeletedHadMatch = $orphanStats.AlreadyHasMatch
        $stats.OrphansMultipleVideos = $orphanStats.MultipleVideos
        $stats.Errors += $orphanStats.Errors
    }

    # Single consolidated summary. Use 'would' phrasing in dry-run so the
    # numbers visibly describe a plan rather than a fait accompli.
    $verb = if ($WhatIf) { 'would ' } else { '' }
    Write-Host "`n=== Subtitle Repair Summary ===" -ForegroundColor Cyan
    Write-Host ("  Subfolders found:        {0}" -f $stats.SubfoldersFound) -ForegroundColor White
    Write-Host ("  Subs ${verb}moved to video:  {0}" -f $stats.SubsMoved) -ForegroundColor Green
    Write-Host ("  Subs ${verb}renamed:         {0}" -f $stats.OrphanedFixed) -ForegroundColor Green
    Write-Host ("  Subs ${verb}deleted (no video):     {0}" -f $stats.OrphansDeletedNoVideo) -ForegroundColor Gray
    Write-Host ("  Subs ${verb}deleted (had match):    {0}" -f $stats.OrphansDeletedHadMatch) -ForegroundColor Gray
    Write-Host ("  Subs ${verb}deleted (non-preferred):{0}" -f $stats.SubsDeleted) -ForegroundColor Gray
    Write-Host ("  Skipped (already correct):     {0}" -f $stats.SubsSkipped) -ForegroundColor Gray
    Write-Host ("  Skipped (multiple videos):     {0}" -f $stats.OrphansMultipleVideos) -ForegroundColor Gray
    Write-Host ("  Empty folders ${verb}removed:    {0}" -f $stats.FoldersRemoved) -ForegroundColor Gray
    Write-Host ("  Errors:                  {0}" -f $stats.Errors) -ForegroundColor $(if ($stats.Errors -gt 0) { 'Red' } else { 'Gray' })

    return $stats
}

<#
.SYNOPSIS
    Syncs subtitle files using ffsubsync to fix timing issues
.DESCRIPTION
    Scans for external subtitle files and uses ffsubsync to automatically
    adjust their timing to match the video's audio track.
.PARAMETER Path
    The root path of the movie library
.PARAMETER Force
    If specified, syncs all subtitles even if already verified
.PARAMETER Backup
    If specified, creates .bak backup of original subtitles before syncing
.PARAMETER WhatIf
    If specified, shows what would be synced without making changes
#>
function Invoke-SubtitleSync {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [switch]$Force,
        [switch]$Backup,
        [switch]$WhatIf
    )

    Write-Host "`n--- Subtitle Sync (ffsubsync) ---" -ForegroundColor Yellow

    if (-not (Test-FFSubSyncInstallation)) {
        Write-Host "ffsubsync is not installed." -ForegroundColor Red
        Write-Host "Install it with: pip install ffsubsync" -ForegroundColor Gray
        Write-Host "Requires Python to be installed first." -ForegroundColor Gray
        Write-Host "ffsubsync not found - cannot sync subtitles" -ForegroundColor Yellow
        return
    }

    if ($WhatIf) {
        Write-Host "(Dry run - no changes will be made)" -ForegroundColor Cyan
    }
    if ($Backup) {
        Write-Host "(Backup mode - original subtitles saved as .bak)" -ForegroundColor Cyan
    }
    Write-Host "  Starting subtitle sync in: $Path (Force: $Force, Backup: $Backup, WhatIf: $WhatIf)" -ForegroundColor Gray

    $stats = @{
        Total = 0
        Synced = 0
        Skipped = 0
        AlreadyVerified = 0
        NoVideo = 0
        Failed = 0
        BackupsCreated = 0
    }

    # Get all movie folders
    $movieFolders = Get-ChildItem -LiteralPath $Path -Directory -ErrorAction SilentlyContinue

    $totalFolders = $movieFolders.Count
    $processed = 0

    Write-Host "Scanning $totalFolders movie folders..." -ForegroundColor Cyan

    foreach ($folder in $movieFolders) {
        $processed++
        $percentComplete = [math]::Round(($processed / $totalFolders) * 100)
        Write-Progress -Activity "Syncing Subtitles" -Status "$processed of $totalFolders - $($folder.Name)" -PercentComplete $percentComplete

        # Skip if already verified (unless Force)
        if (-not $Force -and (Test-SubtitlesVerified -FolderPath $folder.FullName)) {
            $stats.AlreadyVerified++
            continue
        }

        # Find external subtitle files
        $subtitleFiles = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '\.(srt|sub|ass|ssa)$' }

        if (-not $subtitleFiles) {
            continue
        }

        # Find the main video file
        $videoFile = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '\.(mkv|mp4|avi|mov|wmv|m4v)$' } |
            Sort-Object Length -Descending |
            Select-Object -First 1

        if (-not $videoFile) {
            $stats.NoVideo++
            continue
        }

        foreach ($sub in $subtitleFiles) {
            $stats.Total++

            Write-Host "  [$processed/$totalFolders] $($folder.Name) - $($sub.Name)" -ForegroundColor Gray -NoNewline

            if ($WhatIf) {
                Write-Host " [would sync]" -ForegroundColor Cyan
                $stats.Synced++
                continue
            }

            if ($Backup) {
                $result = Invoke-FFSubSync -VideoPath $videoFile.FullName -SubtitlePath $sub.FullName -Backup
                $stats.BackupsCreated++
            } else {
                $result = Invoke-FFSubSync -VideoPath $videoFile.FullName -SubtitlePath $sub.FullName
            }

            if ($result) {
                Write-Host " [synced]" -ForegroundColor Green
                $stats.Synced++
            } else {
                Write-Host " [failed]" -ForegroundColor Red
                $stats.Failed++
            }
        }

        # Mark as verified if all subs synced successfully
        if (-not $WhatIf -and $stats.Failed -eq 0) {
            Set-SubtitlesVerified -FolderPath $folder.FullName -Source "ffsubsync"
        }
    }

    Write-Progress -Activity "Syncing Subtitles" -Completed

    # Summary
    Write-Host "`n--- Subtitle Sync Summary ---" -ForegroundColor Cyan
    Write-Host "Total subtitles processed: $($stats.Total)" -ForegroundColor White
    Write-Host "Successfully synced:       $($stats.Synced)" -ForegroundColor Green
    if ($Backup) {
        Write-Host "Backups created:           $($stats.BackupsCreated)" -ForegroundColor Cyan
    }
    Write-Host "Already verified:          $($stats.AlreadyVerified)" -ForegroundColor Gray
    Write-Host "Failed:                    $($stats.Failed)" -ForegroundColor $(if ($stats.Failed -gt 0) { 'Red' } else { 'Gray' })
    Write-Host "No video file:             $($stats.NoVideo)" -ForegroundColor Yellow

    Write-Host "  Subtitle sync complete - Synced: $($stats.Synced), Backups: $($stats.BackupsCreated), Failed: $($stats.Failed), Verified: $($stats.AlreadyVerified)" -ForegroundColor Gray

    return $stats
}

<#
.SYNOPSIS
    Restores subtitle files from .bak backups
.DESCRIPTION
    Two backup-naming conventions are handled:

      1. Chained extension (e.g. "Movie.en.srt.bak"). This is what
         Invoke-FFSubSync's -Backup mode produces. Target = strip the trailing
         ".bak" — straightforward.

      2. Lone .bak (e.g. "Movie.bak"). Common for release-included sub
         backups that predate ffsubsync. Target is resolved by sniffing the
         .bak content (must be SRT-formatted) and finding a single subtitle
         sibling in the same folder. If zero or multiple sibling subs exist,
         the .bak is skipped with a reason rather than guessing.

    On restore, the .subs_ok marker is removed (since we're rolling back to
    a state that wasn't verified by ffsubsync). Pass -KeepBackup to preserve
    the .bak file after restore as a permanent reference; default deletes it.
.PARAMETER Path
    The root path of the movie library
.PARAMETER WhatIf
    If specified, shows what would be restored without making changes
.PARAMETER KeepBackup
    If specified, preserves the .bak file after restoring (default removes it)
#>
function Restore-SubtitleBackups {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [switch]$WhatIf,
        [switch]$KeepBackup
    )

    Write-Host "`n--- Restore Subtitle Backups ---" -ForegroundColor Yellow

    if ($WhatIf) {
        Write-Host "(Dry run - no changes will be made)" -ForegroundColor Cyan
    }
    if ($KeepBackup) {
        Write-Host "(Keep-backup mode - .bak files will not be deleted after restore)" -ForegroundColor Cyan
    }
    Write-Host "  Starting subtitle backup restore in: $Path (WhatIf: $WhatIf, KeepBackup: $KeepBackup)" -ForegroundColor Gray

    $stats = @{
        Found = 0
        Restored = 0
        Skipped = 0
        Failed = 0
    }

    # Find every .bak file in the tree, then classify each.
    $backupFiles = Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -eq '.bak' }

    $stats.Found = $backupFiles.Count
    if ($stats.Found -eq 0) {
        Write-Host "No .bak files found." -ForegroundColor Gray
        return $stats
    }

    $subExtensions = @('.srt', '.sub', '.ass', '.ssa', '.vtt')

    foreach ($backup in $backupFiles) {
        $folderPath = Split-Path $backup.FullName -Parent
        $folderName = Split-Path $folderPath -Leaf
        $targetPath = $null
        $skipReason = $null

        # --- Pattern 1: chained extension like "Movie.en.srt.bak"
        if ($backup.Name -match '\.(srt|sub|ass|ssa|vtt)\.bak$') {
            $targetPath = $backup.FullName -replace '\.bak$', ''
        }
        # --- Pattern 2: lone .bak — sniff content + find sibling
        else {
            # SRT format check: starts with a numeric cue index then a timestamp
            # line. Allows a leading BOM and either CRLF or LF endings. Without
            # this guard we'd risk treating an unrelated .bak (NFO backup,
            # config snapshot, anything) as a sub backup and overwriting good
            # data.
            $head = $null
            try {
                $head = (Get-Content -LiteralPath $backup.FullName -TotalCount 4 -ErrorAction Stop) -join "`n"
            } catch {
                $skipReason = "could not read .bak: $_"
            }

            if (-not $skipReason) {
                $isSrt = $head -match '^﻿?\s*\d+\s*[\r\n]+\d{2}:\d{2}:\d{2}[,\.]\d{3}\s*-->'
                if (-not $isSrt) {
                    $skipReason = "content does not look like an SRT subtitle"
                } else {
                    # Find subtitle siblings (excluding any other .bak files)
                    $siblings = @(Get-ChildItem -LiteralPath $folderPath -File -ErrorAction SilentlyContinue |
                        Where-Object { $subExtensions -contains $_.Extension.ToLower() })
                    if ($siblings.Count -eq 0) {
                        # No sibling — promote the .bak to a .srt next to it.
                        $targetPath = [System.IO.Path]::ChangeExtension($backup.FullName, '.srt')
                    } elseif ($siblings.Count -eq 1) {
                        $targetPath = $siblings[0].FullName
                    } else {
                        $skipReason = "$($siblings.Count) subtitle siblings present — ambiguous target"
                    }
                }
            }
        }

        Write-Host "  $folderName / $($backup.Name)" -ForegroundColor Gray -NoNewline

        if ($skipReason) {
            Write-Host " [skip: $skipReason]" -ForegroundColor Yellow
            $stats.Skipped++
            continue
        }

        $targetLeaf = Split-Path $targetPath -Leaf
        if ($WhatIf) {
            Write-Host " [would restore -> $targetLeaf]" -ForegroundColor Cyan
            $stats.Restored++
            continue
        }

        try {
            Copy-Item -LiteralPath $backup.FullName -Destination $targetPath -Force
            Write-Host " [restored -> $targetLeaf]" -ForegroundColor Green
            $stats.Restored++

            if (-not $KeepBackup) {
                Remove-Item -LiteralPath $backup.FullName -Force
            }

            # Drop the .subs_ok marker — we're rolling back to a state ffsubsync
            # didn't verify, so the existing marker no longer reflects reality.
            $verifiedFile = Join-Path $folderPath ".subs_ok"
            if (Test-Path -LiteralPath $verifiedFile) {
                Remove-Item -LiteralPath $verifiedFile -Force
            }
        }
        catch {
            Write-Host " [failed: $_]" -ForegroundColor Red
            $stats.Failed++
        }
    }

    # Summary
    Write-Host ""
    Write-Host ("  .bak files found:  {0}" -f $stats.Found) -ForegroundColor White
    $verb = if ($WhatIf) { 'would restore:    ' } else { 'restored:         ' }
    Write-Host ("  $verb {0}" -f $stats.Restored) -ForegroundColor Green
    Write-Host ("  skipped:           {0}" -f $stats.Skipped) -ForegroundColor Gray
    Write-Host ("  failed:            {0}" -f $stats.Failed) -ForegroundColor $(if ($stats.Failed -gt 0) { 'Red' } else { 'Gray' })

    return $stats
}

#endregion Public Functions

# Export public functions
Export-ModuleMember -Function Test-SubtitlesExist, Test-SubtitlesVerified, Set-SubtitlesVerified,
    Remove-SubtitlesVerified, Test-FFSubSyncInstallation, Invoke-FFSubSync,
    Search-SubdlSubtitle, Save-MovieSubtitle,
    Repair-OrphanedSubtitles, Repair-SubtitlePlacement, Invoke-SubtitleSync, Restore-SubtitleBackups
