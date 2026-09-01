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
.DESCRIPTION
    Writes the .subs_ok marker as JSON. Provenance-aware: alongside the
    free-text Source it records a machine-readable Provider, and appends
    each event to a History array instead of clobbering earlier records —
    a sub downloaded by one tool and later corrected by another keeps
    both events on file for later parsing.
.PARAMETER FolderPath
    Path to the movie folder
.PARAMETER Source
    Human-readable description of the event
    (e.g. "sync-verified (offset 0.2s)", "opensubtitles-hashmatch (...)")
.PARAMETER Provider
    Machine-readable origin of the subtitle file itself. When omitted, a
    .sub_pending note (written at download time) is consumed if present;
    otherwise the sub predates provenance tracking and is recorded as
    'release' (shipped alongside the video or added before tracking).
#>
function Set-SubtitlesVerified {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,
        [string]$Source = "unknown",
        [ValidateSet('opensubtitles', 'subdl', 'whisper', 'embedded', 'release', 'manual', 'unknown')]
        [string]$Provider
    )

    $verifiedFile = Join-Path $FolderPath ".subs_ok"
    $pendingFile  = Join-Path $FolderPath ".sub_pending"

    # Download-time provenance note. Always consumed here — once the
    # folder is verified, pending state is moot either way.
    $pending = $null
    if (Test-Path -LiteralPath $pendingFile) {
        try { $pending = Get-Content -LiteralPath $pendingFile -Raw | ConvertFrom-Json } catch {}
        Remove-Item -LiteralPath $pendingFile -Force -ErrorAction SilentlyContinue
    }

    # Carry earlier records forward. Pre-provenance markers (no History)
    # become the first history entry, with Provider inferred from their
    # free-text Source so old records stay parseable too.
    $existing = $null
    $history = @()
    if (Test-Path -LiteralPath $verifiedFile) {
        try {
            $existing = Get-Content -LiteralPath $verifiedFile -Raw | ConvertFrom-Json
            if ($existing.History) {
                $history = @($existing.History)
            } elseif ($existing.Source) {
                $history = @([PSCustomObject]@{
                    Date     = $existing.VerifiedDate
                    Source   = $existing.Source
                    Provider = Resolve-SubtitleProviderFromSource -Source $existing.Source
                })
            }
        } catch {}
    }

    # Provider resolution: explicit caller value > download-time pending
    # note > what the existing marker already knows (a re-verification
    # doesn't change where the file came from) > 'release' for subs that
    # predate provenance tracking entirely.
    if (-not $Provider) {
        if ($pending -and $pending.Provider) {
            $Provider = [string]$pending.Provider
        } elseif ($existing -and $existing.Provider) {
            $Provider = [string]$existing.Provider
        } elseif ($existing -and $existing.Source) {
            $inferred = Resolve-SubtitleProviderFromSource -Source $existing.Source
            $Provider = if ($inferred -ne 'unknown') { $inferred } else { 'release' }
        } else {
            $Provider = 'release'
        }
    }
    $history += [PSCustomObject]@{
        Date     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Source   = $Source
        Provider = $Provider
    }

    $content = @{
        VerifiedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Source       = $Source
        Provider     = $Provider
        VerifiedBy   = "LibraryLint"
        History      = $history
    } | ConvertTo-Json -Depth 4

    try {
        $content | Out-File -LiteralPath $verifiedFile -Encoding UTF8 -Force
        return $true
    }
    catch {
        Write-Host "Failed to create .subs_ok file: $_" -ForegroundColor Yellow
        return $false
    }
}

# Maps a pre-provenance free-text Source string to the Provider enum so
# legacy markers migrate into History with a usable value.
function Resolve-SubtitleProviderFromSource {
    param([string]$Source)
    switch -Regex ($Source) {
        '^opensubtitles' { return 'opensubtitles' }
        '^subdl'         { return 'subdl' }
        '^whisper'       { return 'whisper' }
        '^embedded'      { return 'embedded' }
        '^included'      { return 'release' }
        default          { return 'unknown' }
    }
}

<#
.SYNOPSIS
    Records where a just-downloaded (not yet verified) subtitle came from.
.DESCRIPTION
    Name-matched downloads wait for the verification gate before the
    folder gets its .subs_ok marker — by then the download event is long
    past and its origin would be lost. This drops a .sub_pending note at
    download time; the next Set-SubtitlesVerified call on the folder
    consumes it, so Provider survives the gap between download and
    verification (and unverified downloads are parseable meanwhile).
#>
function Set-SubtitlePendingProvenance {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,
        [Parameter(Mandatory=$true)]
        [ValidateSet('opensubtitles', 'subdl', 'whisper', 'embedded', 'release', 'manual')]
        [string]$Provider,
        [string]$Detail = ''
    )
    $pendingFile = Join-Path $FolderPath ".sub_pending"
    try {
        @{
            Provider = $Provider
            Detail   = $Detail
            Date     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        } | ConvertTo-Json | Out-File -LiteralPath $pendingFile -Encoding UTF8 -Force
    } catch {
        Write-Host "Failed to write .sub_pending note: $_" -ForegroundColor Yellow
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
    Scans a library and returns counts + samples of every subtitle-cleanup
    condition the existing fix functions know how to address.
.DESCRIPTION
    Read-only health check designed as the front door to a "Subtitle Cleanup"
    workflow: scan first, see the shape of the problem, then pick which of
    the existing fix functions to run. Reports:

      - Verification state: folders with .subs_ok markers vs unverified
        folders that have external subs (the "needs review" pool).
      - Misplaced subs: subtitle files still living inside Subs/ /
        Subtitles/ subfolders (input for Repair-SubtitlePlacement).
      - Orphan subs: subtitle files whose basename (with language suffix
        stripped) doesn't match any video in the same folder (input for
        Repair-OrphanedSubtitles).
      - Non-preferred language subs: subtitle files whose detected language
        code isn't in PreferredLanguages (input for Invoke-SubtitleLanguagePrune).

    Each category returns a small sample (default 5 entries) so the report
    can show concrete examples without dumping thousands of paths.
.PARAMETER Path
    Library root to scan.
.PARAMETER PreferredLanguages
    Lowercase language codes / names treated as keep-worthy.
.PARAMETER VideoExtensions
    Extensions used to identify video-bearing folders + match-by-basename.
.PARAMETER SubtitleExtensions
    Extensions counted as subtitle files.
.PARAMETER SampleSize
    Number of example entries to include per category.
.OUTPUTS
    Hashtable with count fields and Sample* arrays.
#>
function Get-SubtitleHealthSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [string[]]$PreferredLanguages = @('eng', 'en', 'english'),
        [string[]]$VideoExtensions    = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov'),
        [string[]]$SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt'),
        [int]$SampleSize = 5
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Host "Path does not exist: $Path" -ForegroundColor Yellow
        return $null
    }

    # 1) Verification state via the existing walker. Inherited markers are
    # honored (a show-level .subs_ok covers all its seasons).
    $verifyStatus = Get-VerifiedSubtitleStatus -Path $Path -VideoExtensions $VideoExtensions -SubtitleExtensions $SubtitleExtensions
    $verified = @($verifyStatus | Where-Object { $_.IsVerified })
    $unverifiedWithSubs = @($verifyStatus | Where-Object { -not $_.IsVerified -and $_.ExternalSubCount -gt 0 })

    # 2) Misplaced subs — anything inside a Subs/Sub/Subtitles/Subtitle/SRT
    # subfolder. Repair-SubtitlePlacement moves these out.
    $subFolderNames = @('subs', 'sub', 'subtitles', 'subtitle', 'srt')
    $misplacedSubs = @()
    $misplacedFolders = @(Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $subFolderNames -contains $_.Name.ToLower() })
    foreach ($mf in $misplacedFolders) {
        $misplacedSubs += @(Get-ChildItem -LiteralPath $mf.FullName -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $SubtitleExtensions -contains $_.Extension.ToLower() })
    }

    # 3) Orphan subs + 4) Non-preferred-language subs — single recursive
    # pass over all subs OUTSIDE Subs/ folders (those are the misplaced set
    # above; double-counting them as orphans is misleading).
    $orphanSubs = New-Object 'System.Collections.Generic.List[object]'
    $nonPreferredSubs = New-Object 'System.Collections.Generic.List[object]'
    $prefSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $PreferredLanguages) { if ($p) { [void]$prefSet.Add($p.ToLower().Trim()) } }

    $allRelevantSubs = @(Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object {
            $SubtitleExtensions -contains $_.Extension.ToLower() -and
            ($subFolderNames -notcontains (Split-Path (Split-Path $_.FullName -Parent) -Leaf).ToLower())
        })

    # Cache per-folder video basenames so we don't enumerate the same dir
    # repeatedly when a folder has many subtitle files.
    $videoBasesByDir = @{}
    foreach ($sub in $allRelevantSubs) {
        $dir = $sub.DirectoryName

        # Orphan detection: strip language + modifier suffix, look for a
        # video sharing the resulting basename in the same folder.
        $base = [System.IO.Path]::GetFileNameWithoutExtension($sub.Name)
        $base = $base -replace '(?i)\.(forced|sdh|cc|hi|hearing[\.\-_]?impaired)$', ''
        $lang = Get-SubtitleLanguageCode -FileName $sub.Name
        if ($lang) {
            $base = $base -replace "(?i)\.$([regex]::Escape($lang))$", ''
        }
        if (-not $videoBasesByDir.ContainsKey($dir)) {
            $videoBasesByDir[$dir] = @(Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
                Where-Object { $VideoExtensions -contains $_.Extension.ToLower() } |
                ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name).ToLower() })
        }
        if (-not ($videoBasesByDir[$dir] -contains $base.ToLower())) {
            $orphanSubs.Add($sub)
        }

        # Non-preferred detection.
        if ($lang -and -not $prefSet.Contains($lang)) {
            $nonPreferredSubs.Add([PSCustomObject]@{
                Path     = $sub.FullName
                Language = $lang
                Size     = $sub.Length
            })
        }
    }

    return @{
        TotalVideoFolders     = $verifyStatus.Count
        Verified              = $verified.Count
        UnverifiedWithSubs    = $unverifiedWithSubs.Count
        TotalSubs             = $allRelevantSubs.Count + $misplacedSubs.Count
        MisplacedSubsCount    = $misplacedSubs.Count
        MisplacedFoldersCount = $misplacedFolders.Count
        OrphanSubsCount       = $orphanSubs.Count
        NonPreferredCount     = $nonPreferredSubs.Count
        SampleMisplaced       = @($misplacedSubs | Select-Object -First $SampleSize | ForEach-Object { $_.FullName })
        SampleOrphans         = @($orphanSubs | Select-Object -First $SampleSize | ForEach-Object { $_.FullName })
        SampleNonPreferred    = @($nonPreferredSubs | Select-Object -First $SampleSize)
    }
}

<#
.SYNOPSIS
    Probes a video file's embedded subtitle tracks via MediaInfo.
.DESCRIPTION
    Embedded text subs (SRT/ASS inside the MKV) are soft subs that are
    synced BY CONSTRUCTION — they shipped with the release. For the
    "every movie has English soft subs" goal they beat any download,
    so the census checks them before flagging a movie as missing subs.
.OUTPUTS
    Hashtable: HasEnglishText, HasEnglishBitmap, HasUntaggedText,
    TextLanguages (string[]), Error (string on probe failure).
#>
function Get-EmbeddedSubtitleInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$VideoPath,
        [string]$MediaInfoPath = 'mediainfo'
    )

    $result = @{
        HasEnglishText    = $false
        HasEnglishBitmap  = $false
        HasUntaggedText   = $false
        TextLanguages     = @()
        # Per-track detail for the extractor: StreamOrder is the overall
        # stream index MediaInfo reports, which maps 1:1 to ffmpeg's 0:N.
        EnglishTextTracks = @()
        Error             = $null
    }

    # Text = extractable/renderable character subs; everything else in a
    # Text-type track (PGS/VobSub/DVB) is bitmap — still soft and playable,
    # but not ffsubsync-able and not extractable to .srt without OCR.
    $textFormats = @('UTF-8', 'SubRip', 'SRT', 'ASS', 'SSA', 'WebVTT', 'Timed Text', 'mov_text', 'Text')

    try {
        $json = & $MediaInfoPath --Output=JSON $VideoPath 2>$null | ConvertFrom-Json
        $tracks = @($json.media.track | Where-Object { $_.'@type' -eq 'Text' })
        foreach ($t in $tracks) {
            $lang = if ($t.Language) { $t.Language.ToString().ToLower() } else { '' }
            $isEnglish = $lang -match '^en($|[-_])' -or $lang -eq 'eng' -or $lang -eq 'english'
            $isText = $textFormats -contains $t.Format

            if ($lang) { $result.TextLanguages += $lang }
            if ($isText) {
                if ($isEnglish) {
                    $result.HasEnglishText = $true
                    $result.EnglishTextTracks += [PSCustomObject]@{
                        StreamOrder = if ($t.StreamOrder -match '^\d+$') { [int]$t.StreamOrder } else { $null }
                        Format      = $t.Format
                        Forced      = ($t.Forced -eq 'Yes')
                        Default     = ($t.Default -eq 'Yes')
                        Title       = $t.Title
                        ElementCount = if ($t.ElementCount -match '^\d+$') { [int]$t.ElementCount } else { 0 }
                    }
                }
                elseif (-not $lang)  { $result.HasUntaggedText = $true }
            } elseif ($isEnglish) {
                $result.HasEnglishBitmap = $true
            }
        }
    } catch {
        $result.Error = $_.Exception.Message
    }
    return $result
}

# --- Embedded-subtitle probe cache -------------------------------------
# MediaInfo costs ~100-300ms per file; a full-library census over 1000+
# movies is minutes of probing. Cache keyed by name|size|mtime (same
# convention as the codec cache) makes re-runs near-instant.
function Get-EmbeddedSubCachePath {
    return "$env:LOCALAPPDATA\LibraryLint\embedded_subs_cache.json"
}

function Read-EmbeddedSubCache {
    $path = Get-EmbeddedSubCachePath
    if (-not (Test-Path -LiteralPath $path)) { return @{ version = 1; entries = @{} } }
    try {
        $loaded = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -AsHashtable
        if (-not $loaded -or -not $loaded.entries) { return @{ version = 1; entries = @{} } }
        return $loaded
    } catch {
        return @{ version = 1; entries = @{} }
    }
}

function Save-EmbeddedSubCache {
    param([hashtable]$Cache)
    $dir = Split-Path (Get-EmbeddedSubCachePath) -Parent
    if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    $Cache | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Get-EmbeddedSubCachePath) -Encoding UTF8
}

function Get-SubtitleCensusSummaryPath {
    return "$env:LOCALAPPDATA\LibraryLint\subtitle_census.json"
}

<#
.SYNOPSIS
    Full-library English-subtitle census: buckets every movie folder by
    how (or whether) the soft-English-subs goal is met.
.DESCRIPTION
    Buckets, in priority order:
      Verified          — .subs_ok marker (user- or tool-verified)
      EmbeddedText      — English text track inside the video (synced by
                          construction; also the Phase-2 extraction pool)
      ExternalPreferred — external sub tagged with a preferred language,
                          not yet verified (needs a sync check)
      ExternalUntagged  — external sub with no language tag (probably the
                          right language, needs tagging + verification)
      EmbeddedBitmapOnly— only English PGS/VobSub inside the video
                          (playable soft subs, but not syncable/extractable
                          without OCR)
      Missing           — no English subtitle in any form
    Persists a summary JSON for the Status dashboard's maintenance line.
.OUTPUTS
    Hashtable with per-bucket counts, per-bucket folder lists, probe-error
    count, and ExtractionCandidates (EmbeddedText folders lacking external
    subs — the Phase-2 harvest list).
#>
function Get-SubtitleCensus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [string[]]$PreferredLanguages = @('eng', 'en', 'english'),
        [string[]]$VideoExtensions    = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov'),
        [string[]]$SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt'),
        [string]$MediaInfoPath = 'mediainfo'
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Host "Path does not exist: $Path" -ForegroundColor Yellow
        return $null
    }

    $prefSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $PreferredLanguages) { if ($p) { [void]$prefSet.Add($p.ToLower().Trim()) } }

    $verifyStatus = Get-VerifiedSubtitleStatus -Path $Path -VideoExtensions $VideoExtensions -SubtitleExtensions $SubtitleExtensions
    $cache = Read-EmbeddedSubCache
    $cacheDirty = $false

    $buckets = @{
        Verified           = New-Object 'System.Collections.Generic.List[object]'
        EmbeddedText       = New-Object 'System.Collections.Generic.List[object]'
        ExternalPreferred  = New-Object 'System.Collections.Generic.List[object]'
        ExternalUntagged   = New-Object 'System.Collections.Generic.List[object]'
        EmbeddedBitmapOnly = New-Object 'System.Collections.Generic.List[object]'
        Missing            = New-Object 'System.Collections.Generic.List[object]'
    }
    $extractionCandidates = New-Object 'System.Collections.Generic.List[object]'
    $probeErrors = 0
    $scanned = 0
    $junkNameRegex = '(?i)(^|[\.\-_\s])(trailer|sample|featurette|extras?)($|[\.\-_\s])'

    foreach ($folder in $verifyStatus) {
        $scanned++
        if ($scanned % 25 -eq 0) {
            Write-Host "`r  Census: $scanned / $($verifyStatus.Count) folders..." -NoNewline -ForegroundColor DarkGray
        }

        # External sub language classification for this folder.
        $externalPref = $false
        $externalUntagged = $false
        foreach ($sub in @(Get-ChildItem -LiteralPath $folder.FolderPath -File -ErrorAction SilentlyContinue |
                Where-Object { $SubtitleExtensions -contains $_.Extension.ToLower() })) {
            $lang = Get-SubtitleLanguageCode -FileName $sub.Name
            if ($lang -and $prefSet.Contains($lang)) { $externalPref = $true }
            elseif (-not $lang) { $externalUntagged = $true }
        }

        # Embedded probe on the primary (largest non-junk) video, cached.
        $primary = @(Get-ChildItem -LiteralPath $folder.FolderPath -File -ErrorAction SilentlyContinue |
            Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and $_.Name -notmatch $junkNameRegex } |
            Sort-Object Length -Descending) | Select-Object -First 1
        $embedded = $null
        if ($primary) {
            $cacheKey = "$($primary.Name)|$($primary.Length)|$($primary.LastWriteTimeUtc.Ticks)"
            if ($cache.entries.ContainsKey($cacheKey)) {
                $embedded = $cache.entries[$cacheKey]
            } else {
                $embedded = Get-EmbeddedSubtitleInfo -VideoPath $primary.FullName -MediaInfoPath $MediaInfoPath
                if ($embedded.Error) {
                    $probeErrors++
                } else {
                    $cache.entries[$cacheKey] = @{
                        HasEnglishText   = $embedded.HasEnglishText
                        HasEnglishBitmap = $embedded.HasEnglishBitmap
                        HasUntaggedText  = $embedded.HasUntaggedText
                    }
                    $cacheDirty = $true
                }
            }
        }

        $entry = [PSCustomObject]@{
            Folder       = $folder.FolderPath
            RelativePath = $folder.RelativePath
        }

        # Bucket assignment, priority order.
        if ($folder.IsVerified) {
            $buckets.Verified.Add($entry)
        } elseif ($embedded -and $embedded.HasEnglishText) {
            $buckets.EmbeddedText.Add($entry)
        } elseif ($externalPref) {
            $buckets.ExternalPreferred.Add($entry)
        } elseif ($externalUntagged) {
            $buckets.ExternalUntagged.Add($entry)
        } elseif ($embedded -and $embedded.HasEnglishBitmap) {
            $buckets.EmbeddedBitmapOnly.Add($entry)
        } else {
            $buckets.Missing.Add($entry)
        }

        # Phase-2 harvest pool: embedded English text with no external sub.
        if ($embedded -and $embedded.HasEnglishText -and -not $externalPref -and -not $externalUntagged) {
            $extractionCandidates.Add($entry)
        }
    }
    Write-Host "`r$(' ' * 60)`r" -NoNewline

    if ($cacheDirty) { Save-EmbeddedSubCache -Cache $cache }

    # .ToArray() rather than @(...): wrapping a Generic.List[object] in an
    # array subexpression trips a PS7 dynamic-binder bug ("Argument types
    # do not match" via PSToObjectArrayBinder) — same failure previously
    # hit in Sync.psm1's extraction filter.
    $missingArr = $buckets.Missing.ToArray()
    $candidatesArr = $extractionCandidates.ToArray()

    $result = @{
        Path                 = $Path
        TotalFolders         = $verifyStatus.Count
        Verified             = $buckets.Verified.Count
        EmbeddedText         = $buckets.EmbeddedText.Count
        ExternalPreferred    = $buckets.ExternalPreferred.Count
        ExternalUntagged     = $buckets.ExternalUntagged.Count
        EmbeddedBitmapOnly   = $buckets.EmbeddedBitmapOnly.Count
        Missing              = $buckets.Missing.Count
        MissingList          = $missingArr
        ExtractionCandidates = $candidatesArr
        ProbeErrors          = $probeErrors
    }

    # Persist the summary so the Status dashboard's maintenance section can
    # report subtitle debt without re-running the (cached but non-trivial)
    # census on every S keystroke.
    $summary = @{
        RanAt              = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Path               = $Path
        TotalFolders       = $result.TotalFolders
        Verified           = $result.Verified
        EmbeddedText       = $result.EmbeddedText
        ExternalPreferred  = $result.ExternalPreferred
        ExternalUntagged   = $result.ExternalUntagged
        EmbeddedBitmapOnly = $result.EmbeddedBitmapOnly
        Missing            = $result.Missing
        ExtractionCandidateCount = $extractionCandidates.Count
        MissingSample      = @($missingArr | Select-Object -First 50 | ForEach-Object { $_.RelativePath })
    }
    $dir = Split-Path (Get-SubtitleCensusSummaryPath) -Parent
    if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    $summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Get-SubtitleCensusSummaryPath) -Encoding UTF8

    # Rebuild the acquisition queue from the fresh missing list — the
    # Status flow trickles the daily OpenSubtitles quota against it.
    $null = Update-SubtitleAcquireQueue -MissingList $missingArr

    return $result
}

<#
.SYNOPSIS
    Phase-2 harvest: extracts embedded English text subtitles to external
    .en.srt files for the census's extraction candidates.
.DESCRIPTION
    Embedded text subs shipped with the release, so the extracted .srt is
    synced by construction — zero sync risk, no downloads. Per candidate:
      1. Re-probe the primary video's English text tracks (live, not the
         boolean cache — the extractor needs stream indexes).
      2. Pick the best track: non-forced full subs preferred (SDH accepted),
         most cues wins on a tie. A forced-only track still extracts, but
         to .en.forced.srt and does NOT satisfy the full-subs goal.
      3. ffmpeg -map 0:<StreamOrder> -c:s srt (converts ASS/WebVTT too).
      4. Validate the output actually contains cues before keeping it.
      5. Optionally mark the folder .subs_ok (Source: embedded-extraction).
    Supports -WhatIf (lists what would be extracted, writes nothing).
.PARAMETER Candidates
    Entries from Get-SubtitleCensus's ExtractionCandidates (objects with a
    Folder property). Pass the census output — this function doesn't
    re-derive the pool.
.OUTPUTS
    Hashtable: Extracted, ForcedOnly, Failed, Skipped, Marked.
#>
function Invoke-EmbeddedSubtitleExtraction {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [object[]]$Candidates,
        [string[]]$VideoExtensions    = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov'),
        [string[]]$SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt'),
        [string]$MediaInfoPath = 'mediainfo',
        [string]$FFmpegPath = 'ffmpeg',
        [int]$Limit = 0,
        [switch]$MarkVerified
    )

    $stats = @{ Extracted = 0; ForcedOnly = 0; Failed = 0; Skipped = 0; Marked = 0 }
    $junkNameRegex = '(?i)(^|[\.\-_\s])(trailer|sample|featurette|extras?)($|[\.\-_\s])'
    $processed = 0
    $total = if ($Limit -gt 0) { [Math]::Min($Limit, $Candidates.Count) } else { $Candidates.Count }

    foreach ($cand in $Candidates) {
        if ($Limit -gt 0 -and $processed -ge $Limit) { break }
        $processed++

        $video = Get-ChildItem -LiteralPath $cand.Folder -File -ErrorAction SilentlyContinue |
            Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and $_.Name -notmatch $junkNameRegex } |
            Sort-Object Length -Descending | Select-Object -First 1
        if (-not $video) { $stats.Skipped++; continue }

        # Safety re-check: skip if an external sub appeared since the census.
        $existingSub = Get-ChildItem -LiteralPath $cand.Folder -File -ErrorAction SilentlyContinue |
            Where-Object { $SubtitleExtensions -contains $_.Extension.ToLower() } |
            Select-Object -First 1
        if ($existingSub) { $stats.Skipped++; continue }

        $probe = Get-EmbeddedSubtitleInfo -VideoPath $video.FullName -MediaInfoPath $MediaInfoPath
        $tracks = @($probe.EnglishTextTracks | Where-Object { $null -ne $_.StreamOrder })
        if ($tracks.Count -eq 0) { $stats.Skipped++; continue }

        # Best track: full subs (non-forced) first, then most cues.
        $fullTracks = @($tracks | Where-Object { -not $_.Forced } | Sort-Object ElementCount -Descending)
        $isForcedOnly = ($fullTracks.Count -eq 0)
        $track = if ($isForcedOnly) { ($tracks | Sort-Object ElementCount -Descending)[0] } else { $fullTracks[0] }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)
        $suffix = if ($isForcedOnly) { '.en.forced.srt' } else { '.en.srt' }
        $outPath = Join-Path $cand.Folder "$baseName$suffix"

        $processedLabel = "[$processed/$total] $(Split-Path $cand.Folder -Leaf)"
        if (-not $PSCmdlet.ShouldProcess($video.FullName, "Extract embedded EN sub (stream $($track.StreamOrder), $($track.Format)) -> $(Split-Path $outPath -Leaf)")) {
            Write-Host "  [DRY-RUN] $processedLabel -> $(Split-Path $outPath -Leaf) (stream $($track.StreamOrder), $($track.Format))" -ForegroundColor Yellow
            continue
        }

        & $FFmpegPath -hide_banner -loglevel error -i $video.FullName `
            -map "0:$($track.StreamOrder)" -c:s srt -y $outPath 2>&1 | Out-Null

        # Validate: file exists and actually contains SRT cues. ffmpeg can
        # exit 0 with an empty/cueless file on odd containers.
        $valid = $false
        if (Test-Path -LiteralPath $outPath) {
            $outItem = Get-Item -LiteralPath $outPath
            if ($outItem.Length -gt 200) {
                $head = Get-Content -LiteralPath $outPath -TotalCount 20 -ErrorAction SilentlyContinue | Out-String
                $valid = $head -match '-->'
            }
        }

        if (-not $valid) {
            Remove-Item -LiteralPath $outPath -Force -ErrorAction SilentlyContinue
            Write-Host "  X $processedLabel — extraction produced no usable cues" -ForegroundColor Red
            $stats.Failed++
            continue
        }

        if ($isForcedOnly) {
            Write-Host "  ~ $processedLabel -> $(Split-Path $outPath -Leaf) (forced-only track; full subs still needed)" -ForegroundColor DarkYellow
            $stats.ForcedOnly++
        } else {
            Write-Host "  + $processedLabel -> $(Split-Path $outPath -Leaf)" -ForegroundColor Green
            $stats.Extracted++
            # Embedded subs shipped with this exact release — synced by
            # construction, so verification is warranted without a check.
            if ($MarkVerified) {
                if (Set-SubtitlesVerified -FolderPath $cand.Folder -Source 'embedded-extraction' -Provider 'embedded') {
                    $stats.Marked++
                }
            }
        }
    }

    return $stats
}

<#
.SYNOPSIS
    Parses a language code out of a subtitle filename.
.DESCRIPTION
    Common patterns:
      Movie.srt              -> '' (no code)
      Movie.eng.srt          -> 'eng'
      Movie.en.srt           -> 'en'
      Movie.eng.forced.srt   -> 'eng' (modifier stripped)
      Movie.en.sdh.srt       -> 'en'
      Movie.2020.eng.srt     -> 'eng'
      Movie.english.srt      -> 'english' (matched against the known-name list)

    Strategy: strip recognized modifier suffixes (forced / sdh / cc / hi),
    then look at the trailing dot segment. Accept it as a language code
    when it's 2 or 3 letters, OR when it matches a known full-name alias.
    Anything else (numbers, longer words like "Name") returns empty,
    signaling "unknown" — the caller can decide whether to keep or prune.

    Match-by-content (detecting English vs French from the actual text) is
    intentionally out of scope; release-naming conventions are reliable
    enough that the filename-only approach catches >99% of real cases.
.PARAMETER FileName
    Subtitle file's name (with or without path).
.OUTPUTS
    String — the detected language code in lowercase, or '' if unknown.
#>
function Get-SubtitleLanguageCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$FileName
    )

    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    if (-not $base) { return '' }

    # Strip trailing modifier(s). Most releases use one, but some chain them
    # ("Movie.eng.forced.sdh.srt" — uncommon but seen). Apply iteratively
    # until no more modifiers match so we don't have to enumerate orderings.
    $modifierPattern = '(?i)\.(forced|sdh|cc|hi|hearing[\.\-_]?impaired)$'
    while ($base -match $modifierPattern) {
        $base = $base -replace $modifierPattern, ''
    }

    if ($base -notmatch '\.([A-Za-z]+)$') { return '' }
    $candidate = $Matches[1].ToLower()

    # 2 or 3 letters: high-confidence ISO 639-1 / 639-2 code.
    if ($candidate.Length -ge 2 -and $candidate.Length -le 3) {
        return $candidate
    }

    # Longer: only accept if it's a known full-name alias. Otherwise it's
    # something like "Movie Name" with a dotted title — definitely not a
    # language code.
    $fullNames = @(
        'english','spanish','french','german','italian','portuguese',
        'japanese','korean','chinese','dutch','swedish','norwegian','danish',
        'finnish','russian','polish','turkish','arabic','hindi','greek',
        'czech','hungarian','romanian','hebrew','persian','thai','vietnamese',
        'indonesian','ukrainian','cantonese','mandarin','tamil','telugu',
        'bengali','bulgarian','croatian','serbian','slovak','slovenian',
        'estonian','latvian','lithuanian','icelandic','catalan','galician',
        'basque','welsh','irish','maltese','albanian','macedonian'
    )
    if ($fullNames -contains $candidate) {
        return $candidate
    }
    return ''
}

<#
.SYNOPSIS
    Removes subtitle files whose detected language isn't in the user's
    preferred list.
.DESCRIPTION
    Walks Path for .srt / .sub / .idx / .ass / .ssa / .vtt (or whatever
    SubtitleExtensions specifies), runs Get-SubtitleLanguageCode on each,
    and deletes the ones whose language isn't a member of
    $PreferredLanguages.

    Three special handling rules:
      - Files in folders with a .subs_ok marker are skipped by default
        (the user signed off on those subs as a set — including any non-
        English entries). Override with -IgnoreVerified.
      - Files with no detected language code are kept by default (the
        most common case where there's only one .srt for a movie, and
        it's almost always the audio's original or the user's primary).
        Override with -DeleteUnknown to also remove those.
      - Files in .actors / .extras / Subs subfolder structures are still
        processed; release-side foreign-language litter often lives in
        these places.
.PARAMETER Path
    Library root to walk.
.PARAMETER PreferredLanguages
    Array of accepted language codes / names (case-insensitive). Typical:
    @('eng','en','english').
.PARAMETER SubtitleExtensions
    File extensions treated as subtitle files.
.PARAMETER IgnoreVerified
    Process folders even if they have a .subs_ok marker. Off by default
    so a verified non-English subtitle set isn't accidentally wiped.
.PARAMETER DeleteUnknown
    Also delete subtitles whose language code couldn't be parsed from the
    filename. Aggressive — off by default. Many releases name their single
    English sub just "Movie.srt" with no code.
.PARAMETER WhatIf
    Preview without deleting. Recommended for the first run.
.OUTPUTS
    Hashtable: ScannedCount, KeptPreferred, KeptUnknown, KeptVerified,
    Deleted (count), DeletedFiles (array of full paths), BytesFreed.
#>
function Invoke-SubtitleLanguagePrune {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string[]]$PreferredLanguages,
        [string[]]$SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt'),
        [switch]$IgnoreVerified,
        [switch]$DeleteUnknown,
        [switch]$WhatIf
    )

    $result = @{
        ScannedCount    = 0
        KeptPreferred   = 0
        KeptUnknown     = 0
        KeptVerified    = 0
        Deleted         = 0
        DeletedFiles    = New-Object 'System.Collections.Generic.List[object]'
        BytesFreed      = [long]0
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Host "Path does not exist: $Path" -ForegroundColor Yellow
        return $result
    }

    # Normalize preferred languages to lowercase for case-insensitive match.
    $prefSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $PreferredLanguages) {
        if ($p) { [void]$prefSet.Add($p.ToLower().Trim()) }
    }

    # Cache verified-folder lookups so we don't Test-Path each marker per
    # subtitle in the same folder.
    $verifiedCache = @{}
    $isVerified = {
        param($folder)
        if ($verifiedCache.ContainsKey($folder)) { return $verifiedCache[$folder] }
        $v = Test-SubtitlesVerified -FolderPath $folder
        $verifiedCache[$folder] = $v
        return $v
    }

    Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $SubtitleExtensions -contains $_.Extension.ToLower() } |
        ForEach-Object {
            $result.ScannedCount++
            $file = $_
            $folder = $file.DirectoryName

            if (-not $IgnoreVerified -and (& $isVerified $folder)) {
                $result.KeptVerified++
                return
            }

            $lang = Get-SubtitleLanguageCode -FileName $file.Name

            if ([string]::IsNullOrWhiteSpace($lang)) {
                if (-not $DeleteUnknown) {
                    $result.KeptUnknown++
                    return
                }
            } elseif ($prefSet.Contains($lang)) {
                $result.KeptPreferred++
                return
            }

            # Deletion candidate.
            $result.DeletedFiles.Add([PSCustomObject]@{
                Path     = $file.FullName
                Size     = $file.Length
                Language = if ($lang) { $lang } else { '(unknown)' }
            })
            $result.BytesFreed += $file.Length

            if (-not $WhatIf) {
                try {
                    Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                    $result.Deleted++
                } catch {
                    Write-Host "  Failed to delete $($file.FullName): $_" -ForegroundColor Red
                }
            }
        }

    return $result
}

<#
.SYNOPSIS
    Walks a library and reports the subtitle-verification state of every
    folder that contains at least one video file.
.DESCRIPTION
    The .subs_ok marker is folder-level, but a video file's effective
    verification state can also come from an ancestor folder (a show-level
    marker covering all seasons, for example). This walker:

    - Finds every directory under $Path that holds at least one video file
      (movie folders, TV season folders, etc.)
    - Reads the marker at that folder, OR if absent, walks ancestor dirs
      up to (but not past) $Path looking for an inherited marker
    - Counts external subtitle files in the folder itself

    Returns one PSCustomObject per video-bearing folder so the caller can
    filter (verified vs not, has-subs vs no-subs, by source, etc.) and
    present audit reports or batch operations.
.PARAMETER Path
    Library root to walk. Typically MoviesLibraryPath or TVShowsLibraryPath.
.PARAMETER VideoExtensions
    Recognized video extensions used to identify "video-bearing" folders.
.PARAMETER SubtitleExtensions
    Recognized subtitle extensions counted into ExternalSubCount.
.OUTPUTS
    Array of PSCustomObject: FolderPath, RelativePath, IsVerified, Source,
    VerifiedDate, VerifiedBy, InheritedFrom (null if marker is on the
    folder itself), ExternalSubCount.
#>
function Get-VerifiedSubtitleStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [string[]]$VideoExtensions = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov'),
        [string[]]$SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt')
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    # Build a unique set of folders that contain at least one video file.
    # One Get-ChildItem -Recurse beats walking the tree manually.
    $videoDirs = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $VideoExtensions -contains $_.Extension.ToLower() } |
        ForEach-Object { [void]$videoDirs.Add($_.DirectoryName) }

    $rootNorm = $Path.TrimEnd('\','/').ToLower()
    $results = New-Object 'System.Collections.Generic.List[object]'

    foreach ($folder in $videoDirs) {
        $markerPath = Join-Path $folder ".subs_ok"
        $isVerified = $false
        $source = $null
        $verifiedDate = $null
        $verifiedBy = $null
        $inheritedFrom = $null

        if (Test-Path -LiteralPath $markerPath) {
            $isVerified = $true
            try {
                $marker = Get-Content -LiteralPath $markerPath -Raw -ErrorAction Stop | ConvertFrom-Json
                $source       = $marker.Source
                $verifiedDate = $marker.VerifiedDate
                $verifiedBy   = $marker.VerifiedBy
            } catch {
                # Marker exists but isn't readable JSON — still counts as
                # verified (the file is the signal); leave fields null.
            }
        } else {
            # Walk ancestors looking for an inherited marker. Stops at the
            # library root to avoid reading markers from siblings of $Path.
            $parent = Split-Path $folder -Parent
            while ($parent -and $parent.TrimEnd('\','/').ToLower().StartsWith($rootNorm)) {
                if ($parent.TrimEnd('\','/').ToLower() -eq $rootNorm) { break }
                $parentMarker = Join-Path $parent ".subs_ok"
                if (Test-Path -LiteralPath $parentMarker) {
                    $isVerified = $true
                    $inheritedFrom = $parent
                    try {
                        $marker = Get-Content -LiteralPath $parentMarker -Raw -ErrorAction Stop | ConvertFrom-Json
                        $source       = $marker.Source
                        $verifiedDate = $marker.VerifiedDate
                        $verifiedBy   = $marker.VerifiedBy
                    } catch {}
                    break
                }
                $parent = Split-Path $parent -Parent
            }
        }

        $externalSubs = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue |
            Where-Object { $SubtitleExtensions -contains $_.Extension.ToLower() }).Count

        $rel = if ($folder.Length -gt $Path.Length) {
            $folder.Substring($Path.Length).TrimStart('\','/')
        } else { '' }

        $results.Add([PSCustomObject]@{
            FolderPath       = $folder
            RelativePath     = $rel
            IsVerified       = $isVerified
            Source           = $source
            VerifiedDate     = $verifiedDate
            VerifiedBy       = $verifiedBy
            InheritedFrom    = $inheritedFrom
            ExternalSubCount = $externalSubs
        })
    }

    return @($results | Sort-Object RelativePath)
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
    Parses an SRT file's first cue start time, returned as seconds.
.DESCRIPTION
    Used by the sync audit to compute "how much did ffsubsync shift this
    sub" — the diff between the original's first-cue start and the synced
    output's first-cue start IS the shift, approximately. Both SRT comma
    (HH:MM:SS,mmm) and dot (HH:MM:SS.mmm) decimal separators are accepted.

    Reads only the first ~8KB of the file — enough to find the first cue
    even with BOM/header noise, and cheap regardless of subtitle length.
.OUTPUTS
    Double (seconds from start of media), or $null if no cue could be parsed.
#>
function Get-SrtFirstCueStart {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$Path)

    try {
        $stream = [System.IO.File]::OpenRead($Path)
        try {
            $buffer = New-Object byte[] 8192
            $read = $stream.Read($buffer, 0, $buffer.Length)
        } finally {
            $stream.Close()
        }
        if ($read -le 0) { return $null }
        # SRT is typically UTF-8 (sometimes Latin-1). UTF-8 decoding tolerates
        # both for the ASCII timestamp characters we're matching on.
        $text = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $read)
        if ($text -match '(\d{1,2}):(\d{2}):(\d{2})[,.](\d{1,3})\s*-->') {
            $h  = [int]$Matches[1]
            $m  = [int]$Matches[2]
            $s  = [int]$Matches[3]
            $ms = [int]$Matches[4]
            return ($h * 3600.0) + ($m * 60.0) + $s + ($ms / 1000.0)
        }
    } catch {
        # Unreadable / permission denied / missing — caller treats as null
    }
    return $null
}

<#
.SYNOPSIS
    Returns the first N cues of an SRT file as plain-text blocks for the
    audit's before/after preview UI.
.DESCRIPTION
    Splits on the standard SRT cue separator (blank line) and keeps blocks
    that contain a timestamp arrow. Each returned string is a complete cue
    including its index number, timestamp line, and dialogue lines.

    Reads up to ~16KB which covers the first ~50 cues comfortably; we only
    show 3 anyway. Bounding the read keeps the preview cheap even for
    multi-MB subtitle files.
.OUTPUTS
    String[] — one entry per cue, in original order.
#>
function Get-SrtCueSample {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [int]$Count = 3
    )

    try {
        $stream = [System.IO.File]::OpenRead($Path)
        try {
            $buffer = New-Object byte[] 16384
            $read = $stream.Read($buffer, 0, $buffer.Length)
        } finally {
            $stream.Close()
        }
        if ($read -le 0) { return @() }
        $text = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $read)
    } catch {
        return @()
    }

    $cues = $text -split "(?:\r?\n){2,}"
    $samples = New-Object 'System.Collections.Generic.List[string]'
    foreach ($cue in $cues) {
        if ($cue -notmatch '\d{1,2}:\d{2}:\d{2}[,.]\d{1,3}\s*-->') { continue }
        $samples.Add($cue.Trim())
        if ($samples.Count -ge $Count) { break }
    }
    return @($samples)
}

<#
.SYNOPSIS
    Audits subtitle/video pairs in a library by running ffsubsync to a temp
    file for each pair and reporting how far the sync would shift each.
.DESCRIPTION
    Two-phase: (1) walk the library for video-bearing folders and match each
    sub to its video (TV-aware basename matching), (2) run ffsubsync to a
    temp output for every pair and capture the predicted offset (synced
    first-cue start - original first-cue start).

    Phase 2 is expensive — 10-60s per pair — so the function supports a
    -Limit cap and skips folders with a .subs_ok marker by default. The
    audit produces NO writes to the user's library; only temp files in a
    session-specific directory under %TEMP%. The caller decides which fixes
    to commit (by copying the temp file over the original).

    Sorted output puts the worst-synced subs first, so a focused review
    session can hit the actionable problems before they fade into a sea
    of trivially-already-good entries.
.PARAMETER Path
    Library root to walk.
.PARAMETER Limit
    Cap on the number of pairs to process. 0 = unlimited.
.PARAMETER IgnoreVerified
    Process folders even if they have a .subs_ok marker.
.OUTPUTS
    Hashtable: Candidates (sorted PSCustomObject[]), TempDir, Failed, Scanned.
    Each Candidate carries: Folder, FolderLeaf, VideoPath, SubPath, SubName,
    TempPath, Offset (seconds, signed), AbsOffset.
#>
function Invoke-SubtitleSyncAudit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [int]$Limit = 0,
        [switch]$IgnoreVerified,
        [string[]]$VideoExtensions = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov')
    )

    if (-not (Test-FFSubSyncInstallation)) {
        Write-Host "ffsubsync is not installed. Install with: pip install ffsubsync" -ForegroundColor Red
        return $null
    }

    # Phase 1: walk for video-bearing folders. Same TV-aware pattern as
    # Invoke-SubtitleSync — movie folders and TV season folders both surface.
    $videoExtPattern = '\.(mkv|mp4|avi|mov|wmv|m4v)$'
    $videoDirs = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match $videoExtPattern } |
        ForEach-Object { [void]$videoDirs.Add($_.DirectoryName) }

    # Inherited-verification check — show-level .subs_ok covers all seasons.
    $rootNorm = $Path.TrimEnd('\','/').ToLower()
    $verifiedCache = @{}
    $isVerifiedInherited = {
        param($folder)
        if ($verifiedCache.ContainsKey($folder)) { return $verifiedCache[$folder] }
        $node = $folder
        while ($node) {
            if (Test-SubtitlesVerified -FolderPath $node) {
                $verifiedCache[$folder] = $true; return $true
            }
            $nodeNorm = $node.TrimEnd('\','/').ToLower()
            if ($nodeNorm -eq $rootNorm) { break }
            $parent = Split-Path $node -Parent
            if (-not $parent -or $parent -eq $node) { break }
            $node = $parent
        }
        $verifiedCache[$folder] = $false; return $false
    }

    # Build (video, sub) pairs with per-video basename matching.
    $pairs = New-Object 'System.Collections.Generic.List[object]'
    foreach ($folder in $videoDirs) {
        if (-not $IgnoreVerified -and (& $isVerifiedInherited $folder)) { continue }

        $videos = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match $videoExtPattern -and $_.Name -notmatch '-trailer\.' })
        if ($videos.Count -eq 0) { continue }

        # ffsubsync supports .srt and .ass; we only audit .srt for simplicity
        # — those are 99% of real subtitle files in a Kodi-style library.
        $allSubs = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -eq '.srt' })
        if ($allSubs.Count -eq 0) { continue }

        foreach ($video in $videos) {
            $videoBase = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)
            $matching = @($allSubs | Where-Object {
                $sb = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                $sb -eq $videoBase -or $sb -like "$videoBase.*"
            })
            if ($matching.Count -eq 0 -and $videos.Count -eq 1) {
                $matching = $allSubs
            }
            foreach ($sub in $matching) {
                $pairs.Add([PSCustomObject]@{
                    Folder    = $folder
                    VideoPath = $video.FullName
                    SubPath   = $sub.FullName
                    SubName   = $sub.Name
                })
            }
        }
    }

    if ($Limit -gt 0 -and $pairs.Count -gt $Limit) {
        $pairs = $pairs | Select-Object -First $Limit
    }

    if ($pairs.Count -eq 0) {
        Write-Host "No (video, .srt) pairs to audit under $Path." -ForegroundColor Yellow
        return @{ Candidates = @(); TempDir = $null; Failed = 0; Scanned = 0 }
    }

    # Phase 2: per-pair ffsubsync to a session-temp dir.
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "ll-sync-audit-$(Get-Random)"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    Write-Host ""
    Write-Host "Running ffsubsync against $($pairs.Count) (video, subtitle) pair(s)." -ForegroundColor Cyan
    Write-Host "  This is the long step — typically 10-60s per pair." -ForegroundColor DarkGray
    Write-Host "  Temp output: $tempDir" -ForegroundColor DarkGray
    Write-Host ""

    $candidates = New-Object 'System.Collections.Generic.List[object]'
    $failed = 0
    $idx = 0

    foreach ($pair in $pairs) {
        $idx++
        $folderLeaf = Split-Path $pair.Folder -Leaf
        $pct = [math]::Round(($idx / $pairs.Count) * 100)
        Write-Progress -Activity "Subtitle sync audit" -Status "$idx/$($pairs.Count) - $folderLeaf - $($pair.SubName)" -PercentComplete $pct

        $tempPath  = Join-Path $tempDir "audit_$idx.srt"
        $stdoutTmp = Join-Path $tempDir "ffsubsync_out_$idx.txt"
        $stderrTmp = Join-Path $tempDir "ffsubsync_err_$idx.txt"

        $ok = $false
        try {
            $proc = Start-Process -FilePath "ffsubsync" `
                -ArgumentList "`"$($pair.VideoPath)`"", "-i", "`"$($pair.SubPath)`"", "-o", "`"$tempPath`"" `
                -Wait -PassThru -NoNewWindow `
                -RedirectStandardOutput $stdoutTmp -RedirectStandardError $stderrTmp
            $ok = ($proc.ExitCode -eq 0 -and (Test-Path -LiteralPath $tempPath))
        } catch {
            $ok = $false
        }

        if (-not $ok) {
            $failed++
            Remove-Item -LiteralPath $stdoutTmp, $stderrTmp -Force -ErrorAction SilentlyContinue
            continue
        }

        $origStart = Get-SrtFirstCueStart -Path $pair.SubPath
        $newStart  = Get-SrtFirstCueStart -Path $tempPath
        Remove-Item -LiteralPath $stdoutTmp, $stderrTmp -Force -ErrorAction SilentlyContinue

        if ($null -ne $origStart -and $null -ne $newStart) {
            $offset = [math]::Round($newStart - $origStart, 3)
            $candidates.Add([PSCustomObject]@{
                Folder     = $pair.Folder
                FolderLeaf = $folderLeaf
                VideoPath  = $pair.VideoPath
                SubPath    = $pair.SubPath
                SubName    = $pair.SubName
                TempPath   = $tempPath
                Offset     = $offset
                AbsOffset  = [math]::Abs($offset)
            })
        } else {
            $failed++
        }
    }

    Write-Progress -Activity "Subtitle sync audit" -Completed

    $sorted = $candidates | Sort-Object AbsOffset -Descending

    return @{
        Candidates = @($sorted)
        TempDir    = $tempDir
        Failed     = $failed
        Scanned    = $pairs.Count
    }
}

<#
.SYNOPSIS
    Phase-4 verification gate: measures every unverified external sub's
    sync drift and acts on the verdict.
.DESCRIPTION
    Builds on Invoke-SubtitleSyncAudit (pairing + ffsubsync-to-temp +
    offset measurement), then classifies each pair:

      |offset| <  InSyncThreshold  -> IN SYNC. Mark .subs_ok; the file was
                                      never the problem.
      |offset| <= FixThreshold     -> FIXABLE DRIFT. Back up the original
                                      (.bak), swap in the ffsubsync output,
                                      mark .subs_ok.
      |offset| >  FixThreshold     -> LIKELY WRONG RELEASE. Re-syncing a
                                      sub cut for a different release
                                      produces garbage mid-film even when
                                      the first cue lines up — flag for
                                      REPLACEMENT (the Phase-3 acquisition
                                      list) instead of pretending to fix.

    Verified/corrected subs missing a language tag are renamed to include
    ".en" so the census recognizes them. -WhatIf reports verdicts without
    touching anything.
.OUTPUTS
    Hashtable: Scanned, InSync, Corrected, WrongRelease (list), Failed,
    Tagged, TempDir cleanup is handled internally.
#>
function Invoke-SubtitleVerificationPass {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [int]$Limit = 0,
        [double]$InSyncThreshold = 0.5,
        [double]$FixThreshold = 15.0,
        [string[]]$VideoExtensions = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov')
    )

    # The measurement itself must ALWAYS run — it's read-only against the
    # library (writes only to temp). Without this suppression, an inherited
    # -WhatIf turns the audit's internal Start-Process/New-Item calls into
    # what-ifs and the "dry run" measures nothing at all. Only the apply
    # steps below are gated on ShouldProcess.
    $prevWhatIf = $WhatIfPreference
    $WhatIfPreference = $false
    try {
        $audit = Invoke-SubtitleSyncAudit -Path $Path -Limit $Limit -VideoExtensions $VideoExtensions
    } finally {
        $WhatIfPreference = $prevWhatIf
    }
    if (-not $audit) { return $null }

    $stats = @{
        Scanned      = $audit.Scanned
        InSync       = 0
        Corrected    = 0
        WrongRelease = @()
        Failed       = $audit.Failed
        Tagged       = 0
    }

    try {
        foreach ($cand in $audit.Candidates) {
            $leaf = $cand.FolderLeaf
            if ($cand.AbsOffset -lt $InSyncThreshold) {
                # Already in sync — the measurement IS the verification.
                if ($PSCmdlet.ShouldProcess($cand.Folder, "Mark verified (offset $($cand.Offset)s)")) {
                    $null = Set-SubtitlesVerified -FolderPath $cand.Folder -Source "sync-verified (offset $($cand.Offset)s)"
                    Write-Host "  = $leaf — in sync (offset $($cand.Offset)s), marked verified" -ForegroundColor Green
                } else {
                    Write-Host "  [DRY] $leaf — in sync (offset $($cand.Offset)s) -> would mark verified" -ForegroundColor Green
                }
                $stats.InSync++
            } elseif ($cand.AbsOffset -le $FixThreshold) {
                # Fixable drift: ffsubsync's output is already sitting in
                # temp — swap it in with a backup of the original.
                if ($PSCmdlet.ShouldProcess($cand.SubPath, "Apply ffsubsync correction (offset $($cand.Offset)s)")) {
                    try {
                        Copy-Item -LiteralPath $cand.SubPath -Destination "$($cand.SubPath).bak" -Force
                        Move-Item -LiteralPath $cand.TempPath -Destination $cand.SubPath -Force
                        $null = Set-SubtitlesVerified -FolderPath $cand.Folder -Source "ffsubsync-corrected (offset $($cand.Offset)s)"
                        Write-Host "  ~ $leaf — corrected $($cand.Offset)s drift, marked verified (backup kept)" -ForegroundColor Cyan
                        $stats.Corrected++
                    } catch {
                        Write-Host "  X $leaf — correction failed: $_" -ForegroundColor Red
                        $stats.Failed++
                        continue
                    }
                } else {
                    Write-Host "  [DRY] $leaf — offset $($cand.Offset)s -> would apply correction + mark verified" -ForegroundColor Cyan
                    $stats.Corrected++
                }
            } else {
                # Beyond plausible drift — a sub cut for a different release.
                # Repairing these is how mid-film desync horrors happen.
                Write-Host "  ! $leaf — offset $($cand.Offset)s exceeds fix threshold; flagged for replacement" -ForegroundColor Yellow
                $stats.WrongRelease += [PSCustomObject]@{
                    Folder  = $cand.Folder
                    SubName = $cand.SubName
                    Offset  = $cand.Offset
                }
                continue
            }

            # Tag untagged-but-good subs so the census recognizes the
            # language without re-probing content.
            $subLang = Get-SubtitleLanguageCode -FileName ([System.IO.Path]::GetFileName($cand.SubPath))
            if (-not $subLang -and (Test-Path -LiteralPath $cand.SubPath)) {
                $newName = [System.IO.Path]::GetFileNameWithoutExtension($cand.SubPath) + '.en.srt'
                if ($PSCmdlet.ShouldProcess($cand.SubPath, "Rename to $newName (language tag)")) {
                    try {
                        Rename-Item -LiteralPath $cand.SubPath -NewName $newName -ErrorAction Stop
                        $stats.Tagged++
                    } catch {}
                }
            }
        }
    } finally {
        if ($audit.TempDir -and (Test-Path -LiteralPath $audit.TempDir)) {
            Remove-Item -LiteralPath $audit.TempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    return $stats
}

#region OpenSubtitles (hash-matched acquisition)

<#
.SYNOPSIS
    Computes the OSDb 64-bit moviehash for a video file.
.DESCRIPTION
    filesize + sum of the first 64KB and last 64KB read as little-endian
    uint64 chunks, mod 2^64 — the classic OpenSubtitles hash. A hash match
    identifies a subtitle uploaded for the EXACT same file, which makes
    the sub synced by construction (no verification pass needed).
    BigInteger arithmetic avoids PS overflow semantics; masked to 64 bits.
#>
function Get-OpenSubtitlesHash {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$Path)

    try {
        $stream = [System.IO.File]::OpenRead($Path)
        try {
            $fileSize = $stream.Length
            if ($fileSize -lt 131072) { return $null }   # OSDb requires >= 128KB

            $mask = [System.Numerics.BigInteger]::Pow(2, 64) - 1
            $hash = [System.Numerics.BigInteger]$fileSize

            $reader = New-Object System.IO.BinaryReader($stream)
            for ($i = 0; $i -lt 8192; $i++) {
                $hash = ($hash + [System.Numerics.BigInteger]$reader.ReadUInt64()) -band $mask
            }
            $stream.Position = $fileSize - 65536
            for ($i = 0; $i -lt 8192; $i++) {
                $hash = ($hash + [System.Numerics.BigInteger]$reader.ReadUInt64()) -band $mask
            }
            return $hash.ToString('x16').PadLeft(16, '0').Substring([Math]::Max(0, $hash.ToString('x16').Length - 16))
        } finally {
            $stream.Dispose()
        }
    } catch {
        return $null
    }
}

<#
.SYNOPSIS
    Logs into the OpenSubtitles REST API and returns a bearer token.
.DESCRIPTION
    Downloads require a JWT from /login (the Api-Key alone only covers
    search). Free accounts get a daily download quota (~20); the caller
    handles quota-exhausted responses from Save-OpenSubtitlesFile.
#>
function Connect-OpenSubtitles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$ApiKey,
        [Parameter(Mandatory)] [string]$Username,
        [Parameter(Mandatory)] [string]$Password
    )
    try {
        $headers = @{ 'Api-Key' = $ApiKey; 'User-Agent' = 'LibraryLint v5.7'; 'Accept' = 'application/json' }
        $body = @{ username = $Username; password = $Password } | ConvertTo-Json
        $resp = Invoke-RestMethod -Uri 'https://api.opensubtitles.com/api/v1/login' -Method Post `
            -Headers $headers -Body $body -ContentType 'application/json' -TimeoutSec 20 -ErrorAction Stop
        return @{ Success = $true; Token = $resp.token }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

<#
.SYNOPSIS
    Searches OpenSubtitles for an English sub matching a moviehash.
.OUTPUTS
    Best match as @{FileId; Release; HearingImpaired} or $null. Only
    results whose moviehash actually matched are considered — that's the
    entire point of this path.
#>
function Search-OpenSubtitlesByHash {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$ApiKey,
        [Parameter(Mandatory)] [string]$MovieHash,
        [string]$Languages = 'en'
    )
    try {
        $headers = @{ 'Api-Key' = $ApiKey; 'User-Agent' = 'LibraryLint v5.7'; 'Accept' = 'application/json' }
        $uri = "https://api.opensubtitles.com/api/v1/subtitles?moviehash=$MovieHash&languages=$Languages"
        $resp = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 20 -ErrorAction Stop

        $matches2 = @($resp.data | Where-Object {
            $_.attributes.moviehash_match -eq $true -and $_.attributes.files.Count -gt 0
        })
        if ($matches2.Count -eq 0) { return $null }

        # Prefer non-HI, then most downloads (community-proven copies).
        $best = $matches2 | Sort-Object `
            @{ Expression = { $_.attributes.hearing_impaired -eq $true } }, `
            @{ Expression = { [int]$_.attributes.download_count }; Descending = $true } |
            Select-Object -First 1
        return @{
            FileId          = $best.attributes.files[0].file_id
            Release         = $best.attributes.release
            HearingImpaired = ($best.attributes.hearing_impaired -eq $true)
        }
    } catch {
        return $null
    }
}

<#
.SYNOPSIS
    Downloads an OpenSubtitles file by id. Reports remaining daily quota.
.OUTPUTS
    @{Success; QuotaExhausted; Remaining; Error}
#>
function Save-OpenSubtitlesFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$ApiKey,
        [Parameter(Mandatory)] [string]$Token,
        [Parameter(Mandatory)] $FileId,
        [Parameter(Mandatory)] [string]$OutPath
    )
    try {
        # 'Accept: application/json' is REQUIRED here: /download answers
        # its absence with a bare 503 (not a 4xx), which is indistinguishable
        # from a service outage — that misdirection cost a week of assuming
        # OpenSubtitles was down while every download silently failed.
        $headers = @{ 'Api-Key' = $ApiKey; 'User-Agent' = 'LibraryLint v5.7'; 'Accept' = 'application/json'; 'Authorization' = "Bearer $Token" }
        $body = @{ file_id = $FileId } | ConvertTo-Json
        $resp = Invoke-RestMethod -Uri 'https://api.opensubtitles.com/api/v1/download' -Method Post `
            -Headers $headers -Body $body -ContentType 'application/json' -TimeoutSec 20 -ErrorAction Stop
        if (-not $resp.link) {
            return @{ Success = $false; Error = ($resp.message ?? 'no download link returned') }
        }
        Invoke-WebRequest -Uri $resp.link -OutFile $OutPath -TimeoutSec 60 -ErrorAction Stop
        return @{ Success = $true; Remaining = $resp.remaining }
    } catch {
        $status = $null
        try { $status = [int]$_.Exception.Response.StatusCode } catch {}
        # 406 = download quota exhausted for the day.
        if ($status -eq 406) {
            return @{ Success = $false; QuotaExhausted = $true; Error = 'daily download quota exhausted' }
        }
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

<#
.SYNOPSIS
    Phase-3 acquisition: fills missing English subs — hash-matched
    OpenSubtitles first (sync-guaranteed, auto-verified), Subdl second
    (name-matched, left for the verification gate).
.PARAMETER Candidates
    Census MissingList entries (objects with Folder / RelativePath).
.OUTPUTS
    Hashtable: HashMatched, SubdlDownloaded, QuotaHit (bool), Unresolved
    (list), Failed, Skipped, QuotaRemaining.
#>
function Invoke-SubtitleAcquisition {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [object[]]$Candidates,
        [string[]]$VideoExtensions    = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov'),
        [string]$OpenSubtitlesApiKey,
        [string]$OpenSubtitlesUsername,
        [string]$OpenSubtitlesPassword,
        [string]$SubdlApiKey,
        # Subdl is DEMOTED to explicit opt-in: name-matched subs are the
        # wrong-release factory, so they only download when the caller
        # asked for the fallback this run (and they still face the
        # verification gate afterwards).
        [switch]$EnableSubdlFallback,
        [int]$Limit = 0
    )

    $stats = @{
        HashMatched    = 0
        SubdlDownloaded = 0
        QuotaHit       = $false
        Unresolved     = @()
        Failed         = 0
        Skipped        = 0
        QuotaRemaining = $null
    }
    $junkNameRegex = '(?i)(^|[\.\-_\s])(trailer|sample|featurette|extras?)($|[\.\-_\s])'

    # A requested-but-keyless Subdl fallback must announce itself: without
    # this, every candidate silently fell through to "no sub found" and a
    # 268-movie pass produced pure fiction in seconds.
    if ($EnableSubdlFallback -and -not $SubdlApiKey) {
        Write-Host "  Subdl fallback requested but no SubdlApiKey configured — Subdl stage will be SKIPPED (Settings > Manage API Keys > 4)." -ForegroundColor Yellow
    }

    # OpenSubtitles session — optional; without it we go straight to Subdl.
    $osToken = $null
    $osEnabled = $false
    if ($OpenSubtitlesApiKey -and $OpenSubtitlesUsername -and $OpenSubtitlesPassword) {
        $login = Connect-OpenSubtitles -ApiKey $OpenSubtitlesApiKey -Username $OpenSubtitlesUsername -Password $OpenSubtitlesPassword
        if ($login.Success) {
            $osToken = $login.Token
            $osEnabled = $true
            Write-Host "  OpenSubtitles: logged in (hash-matching active)" -ForegroundColor Green
        } else {
            Write-Host "  OpenSubtitles login failed ($($login.Error)) — falling back to Subdl only" -ForegroundColor Yellow
        }
    } elseif ($OpenSubtitlesApiKey) {
        Write-Host "  OpenSubtitles username/password missing — hash search available but downloads need login; using Subdl for downloads" -ForegroundColor Yellow
    }

    $processed = 0
    $total = if ($Limit -gt 0) { [Math]::Min($Limit, $Candidates.Count) } else { $Candidates.Count }

    foreach ($cand in $Candidates) {
        if ($Limit -gt 0 -and $processed -ge $Limit) { break }
        $processed++
        $leaf = Split-Path $cand.Folder -Leaf

        $video = Get-ChildItem -LiteralPath $cand.Folder -File -ErrorAction SilentlyContinue |
            Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and $_.Name -notmatch $junkNameRegex } |
            Sort-Object Length -Descending | Select-Object -First 1
        if (-not $video) { $stats.Skipped++; continue }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)
        $outPath = Join-Path $cand.Folder "$baseName.en.srt"

        # --- Stage 1: hash match (sync-guaranteed) ---
        if ($osEnabled -and -not $stats.QuotaHit) {
            $hash = Get-OpenSubtitlesHash -Path $video.FullName
            if ($hash) {
                $match = Search-OpenSubtitlesByHash -ApiKey $OpenSubtitlesApiKey -MovieHash $hash
                if ($match) {
                    if (-not $PSCmdlet.ShouldProcess($video.FullName, "Download hash-matched sub (file_id $($match.FileId))")) {
                        Write-Host "  [DRY] [$processed/$total] $leaf — hash match found ($($match.Release))" -ForegroundColor Green
                        $stats.HashMatched++
                        continue
                    }
                    $dl = Save-OpenSubtitlesFile -ApiKey $OpenSubtitlesApiKey -Token $osToken -FileId $match.FileId -OutPath $outPath
                    if ($dl.Success) {
                        $null = Set-SubtitlesVerified -FolderPath $cand.Folder -Source "opensubtitles-hashmatch ($($match.Release))" -Provider 'opensubtitles'
                        $hiTag = if ($match.HearingImpaired) { ' [HI]' } else { '' }
                        Write-Host "  + [$processed/$total] $leaf — hash-matched$hiTag, verified by construction" -ForegroundColor Green
                        $stats.HashMatched++
                        if ($null -ne $dl.Remaining) { $stats.QuotaRemaining = $dl.Remaining }
                        # Gentle pacing for the API.
                        Start-Sleep -Milliseconds 500
                        continue
                    } elseif ($dl.QuotaExhausted) {
                        $quotaFallbackNote = if ($EnableSubdlFallback -and $SubdlApiKey) { 'remaining movies use Subdl' } else { 'remaining movies stay queued' }
                        Write-Host "  ! OpenSubtitles daily quota exhausted — $quotaFallbackNote (re-run tomorrow for more hash matches)" -ForegroundColor Yellow
                        $stats.QuotaHit = $true
                        # fall through to Subdl for THIS movie too
                    } else {
                        Write-Host "  X [$processed/$total] $leaf — hash-match download failed: $($dl.Error)" -ForegroundColor Red
                        # fall through to Subdl
                    }
                }
            }
        }

        # --- Stage 2: Subdl by IMDB/title (needs the verification gate) ---
        if ($EnableSubdlFallback -and $SubdlApiKey) {
            # IMDB id from any NFO in the folder; title/year from folder name.
            $imdbId = $null
            $nfo = Get-ChildItem -LiteralPath $cand.Folder -Filter '*.nfo' -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($nfo) {
                $nfoRaw = Get-Content -LiteralPath $nfo.FullName -Raw -ErrorAction SilentlyContinue
                if ($nfoRaw -match '(tt\d{7,9})') { $imdbId = $Matches[1] }
            }
            $title = $leaf; $year = $null
            if ($leaf -match '^(.+?)\s*\((\d{4})\)') { $title = $Matches[1].Trim(); $year = $Matches[2] }

            if (-not $PSCmdlet.ShouldProcess($cand.Folder, "Subdl download ($title $year)")) {
                Write-Host "  [DRY] [$processed/$total] $leaf — would try Subdl ($(if ($imdbId) { $imdbId } else { "$title $year" }))" -ForegroundColor Cyan
                continue
            }
            $got = Save-MovieSubtitle -IMDBID $imdbId -Title $title -Year $year `
                -MovieFolder $cand.Folder -MovieTitle $title -Language 'en' -ApiKey $SubdlApiKey
            if ($got) {
                # Name-matched: NOT auto-verified — the verification pass
                # decides. Save-MovieSubtitle dropped a .sub_pending note,
                # so when the gate eventually verifies this sub the marker
                # still records Provider = subdl.
                Write-Host "  ~ [$processed/$total] $leaf — Subdl download (needs verification pass)" -ForegroundColor Cyan
                $stats.SubdlDownloaded++
                continue
            }
        }

        $stats.Unresolved += [PSCustomObject]@{ Folder = $cand.Folder; Name = $leaf }
        Write-Host "  - [$processed/$total] $leaf — no sub found (Whisper-tail candidate)" -ForegroundColor DarkGray
    }

    return $stats
}

#region Whisper (generate-from-audio subtitles)

<#
.SYNOPSIS
    Finds a Python interpreter with faster-whisper importable.
.OUTPUTS
    Argument array to invoke it (e.g. @('python') or @('py','-3.11')),
    or $null when no interpreter has the package.
#>
function Get-WhisperPython {
    foreach ($candidate in @(@('python'), @('py', '-3.11'), @('py', '-3.12'), @('py', '-3.13'))) {
        $exe = Get-Command $candidate[0] -ErrorAction SilentlyContinue
        if (-not $exe) { continue }
        try {
            # [1..0] on a one-element array wraps around instead of slicing
            # empty, so guard the launcher-args slice explicitly.
            $launcherArgs = if ($candidate.Length -gt 1) { $candidate[1..($candidate.Length - 1)] } else { @() }
            $importArgs = @($launcherArgs) + @('-c', 'import faster_whisper')
            $null = & $candidate[0] @importArgs 2>$null
            # Comma operator: a one-element array would otherwise unroll
            # to a bare string on return, making $py[0] index a CHAR.
            if ($LASTEXITCODE -eq 0) { return , $candidate }
        } catch {}
    }
    return $null
}

function Test-WhisperInstallation {
    return ($null -ne (Get-WhisperPython))
}

# Python driver for one transcription. Single-quoted here-string: no PS
# interpolation. Args: <video> <out_srt> <model>. Attempts CUDA first
# (int8_float16 fits large-v3 in ~6GB VRAM), falls back to CPU int8.
# task='translate' emits ENGLISH subs regardless of audio language —
# for English audio it behaves as transcription, for foreign films it
# translates. vad_filter tames Whisper's silence-hallucination habit.
$script:WhisperDriverSource = @'
import sys, os, site

# Make pip-installed NVIDIA runtime DLLs (nvidia-cublas-cu12 / cudnn)
# loadable on Windows without a system CUDA install. Both mechanisms are
# needed: add_dll_directory covers LoadLibraryEx lookups, but ctranslate2
# loads cublas64_12.dll lazily via plain LoadLibrary, which only searches
# PATH — without the PATH prepend the CUDA attempt fails at transcribe
# time (not model load) and everything silently lands on CPU.
try:
    dirs = []
    try:
        dirs.extend(site.getsitepackages())
    except Exception:
        pass
    try:
        dirs.append(site.getusersitepackages())
    except Exception:
        pass
    for sp in [d for d in dirs if d]:
        for sub in (os.path.join('nvidia', 'cublas', 'bin'), os.path.join('nvidia', 'cudnn', 'bin')):
            d = os.path.join(sp, sub)
            if os.path.isdir(d):
                os.add_dll_directory(d)
                os.environ['PATH'] = d + os.pathsep + os.environ.get('PATH', '')
except Exception:
    pass

from faster_whisper import WhisperModel

video, out_path, model_name = sys.argv[1], sys.argv[2], sys.argv[3]

def fmt(t):
    if t < 0:
        t = 0
    h = int(t // 3600); m = int(t % 3600 // 60); s = int(t % 60)
    ms = int(round((t - int(t)) * 1000))
    if ms == 1000:
        ms = 999
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"

last_err = None
for device, ctype in (("cuda", "int8_float16"), ("cpu", "int8")):
    try:
        model = WhisperModel(model_name, device=device, compute_type=ctype)
        segments, info = model.transcribe(video, task="translate", vad_filter=True, beam_size=5)
        n = 0
        with open(out_path, "w", encoding="utf-8") as f:
            for seg in segments:
                text = seg.text.strip()
                if not text:
                    continue
                n += 1
                f.write(f"{n}\n{fmt(seg.start)} --> {fmt(seg.end)}\n{text}\n\n")
        print(f"OK device={device} language={info.language} prob={info.language_probability:.2f} segments={n}")
        sys.exit(0)
    except Exception as e:
        last_err = e
        continue

print(f"FAILED: {last_err}", file=sys.stderr)
sys.exit(1)
'@

<#
.SYNOPSIS
    Generates an English .srt from a video's audio via faster-whisper.
.DESCRIPTION
    The generated timestamps derive from THIS file's audio, so the output
    is synced by construction — the same guarantee class as a hash match.
    Quality is the variable (machine transcript vs human sub), which is
    why the .subs_ok Source records whisper provenance: a later human sub
    can knowingly replace it.
.OUTPUTS
    Hashtable: Success, Device, Language, Segments, Error.
#>
function Invoke-WhisperSubtitle {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [string]$VideoPath,
        [Parameter(Mandatory)] [string]$OutPath,
        [string]$Model = 'large-v3'
    )

    $result = @{ Success = $false; Device = $null; Language = $null; Segments = 0; Error = $null }

    if (-not $PSCmdlet.ShouldProcess($VideoPath, "Generate English subtitle via Whisper ($Model)")) {
        $result.Error = 'skipped (WhatIf)'
        return $result
    }

    $py = @(Get-WhisperPython)
    if (-not $py -or $py.Count -eq 0) {
        $result.Error = 'faster-whisper not installed (Settings > Install/Update Dependencies)'
        return $result
    }

    # Driver persisted per session; rewrite if missing.
    if (-not $script:WhisperDriverPath -or -not (Test-Path -LiteralPath $script:WhisperDriverPath)) {
        $script:WhisperDriverPath = Join-Path ([System.IO.Path]::GetTempPath()) "ll_whisper_driver.py"
        Set-Content -LiteralPath $script:WhisperDriverPath -Value $script:WhisperDriverSource -Encoding UTF8
    }

    $errLog = Join-Path ([System.IO.Path]::GetTempPath()) "ll_whisper_err_$(Get-Random).log"
    try {
        $launcherArgs = if ($py.Length -gt 1) { $py[1..($py.Length - 1)] } else { @() }
        $invokeArgs = @($launcherArgs) + @($script:WhisperDriverPath, $VideoPath, $OutPath, $Model)
        $stdout = & $py[0] @invokeArgs 2>$errLog
        $exit = $LASTEXITCODE

        if ($exit -eq 0 -and $stdout -match 'OK device=(\S+) language=(\S+) prob=\S+ segments=(\d+)') {
            $result.Device   = $Matches[1]
            $result.Language = $Matches[2]
            $result.Segments = [int]$Matches[3]
            # Validate real cues landed on disk.
            if ((Test-Path -LiteralPath $OutPath) -and $result.Segments -gt 0) {
                $head = Get-Content -LiteralPath $OutPath -TotalCount 10 -ErrorAction SilentlyContinue | Out-String
                if ($head -match '-->') { $result.Success = $true }
            }
            if (-not $result.Success) { $result.Error = 'driver reported OK but output has no cues' }
        } else {
            $errText = (Get-Content -LiteralPath $errLog -Raw -ErrorAction SilentlyContinue)
            $result.Error = if ($errText -match 'FAILED: (.+)') { $Matches[1].Trim() } else { "exit $exit" }
        }
    } catch {
        $result.Error = $_.Exception.Message
    } finally {
        Remove-Item -LiteralPath $errLog -Force -ErrorAction SilentlyContinue
    }

    if (-not $result.Success -and (Test-Path -LiteralPath $OutPath)) {
        Remove-Item -LiteralPath $OutPath -Force -ErrorAction SilentlyContinue
    }
    return $result
}

<#
.SYNOPSIS
    Whisper batch pass over movies still missing English subs.
.DESCRIPTION
    The expensive tier: minutes of GPU per movie, but works for anything
    with an audio track and cannot produce out-of-sync output. Naturally
    resumable — folders that gained a sub since the list was built are
    skipped, so re-running continues where the last session stopped.
.OUTPUTS
    Hashtable: Generated, Failed, Skipped, plus per-movie console output.
#>
function Invoke-WhisperSubtitlePass {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [object[]]$Candidates,
        [string[]]$VideoExtensions    = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov'),
        [string[]]$SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt'),
        [string]$Model = 'large-v3',
        [int]$Limit = 0
    )

    $stats = @{ Generated = 0; Failed = 0; Skipped = 0 }
    $junkNameRegex = '(?i)(^|[\.\-_\s])(trailer|sample|featurette|extras?)($|[\.\-_\s])'
    $processed = 0
    $total = if ($Limit -gt 0) { [Math]::Min($Limit, $Candidates.Count) } else { $Candidates.Count }

    Write-Host "  Model '$Model' downloads (~3 GB) on first use, then caches locally." -ForegroundColor DarkGray
    Write-Host ""

    foreach ($cand in $Candidates) {
        if ($Limit -gt 0 -and $processed -ge $Limit) { break }
        $processed++
        $leaf = Split-Path $cand.Folder -Leaf

        # Resumability + belt: anything that acquired a sub since the
        # candidate list was built gets skipped without GPU cost.
        $hasSub = Get-ChildItem -LiteralPath $cand.Folder -File -ErrorAction SilentlyContinue |
            Where-Object { $SubtitleExtensions -contains $_.Extension.ToLower() } | Select-Object -First 1
        if ($hasSub) { $stats.Skipped++; continue }

        $video = Get-ChildItem -LiteralPath $cand.Folder -File -ErrorAction SilentlyContinue |
            Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and $_.Name -notmatch $junkNameRegex } |
            Sort-Object Length -Descending | Select-Object -First 1
        if (-not $video) { $stats.Skipped++; continue }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)
        $outPath = Join-Path $cand.Folder "$baseName.en.srt"

        Write-Host "  [$processed/$total] $leaf — transcribing..." -ForegroundColor Gray
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $r = Invoke-WhisperSubtitle -VideoPath $video.FullName -OutPath $outPath -Model $Model
        $sw.Stop()
        $mins = [math]::Round($sw.Elapsed.TotalMinutes, 1)

        if ($r.Success) {
            $null = Set-SubtitlesVerified -FolderPath $cand.Folder -Source "whisper-$Model ($($r.Language) -> en, $($r.Device))" -Provider 'whisper'
            Write-Host "  + [$processed/$total] $leaf — $($r.Segments) cues in ${mins}m ($($r.Device), audio: $($r.Language))" -ForegroundColor Green
            $stats.Generated++
        } else {
            Write-Host "  X [$processed/$total] $leaf — $($r.Error) (${mins}m)" -ForegroundColor Red
            $stats.Failed++
        }
    }

    return $stats
}

#endregion

# --- Subtitle acquisition queue ----------------------------------------
# The free OpenSubtitles tier allows ~5 downloads per rolling 24h — too
# few for batch runs, ideal for a trickle: every Status run spends the
# day's quota against the backlog. The queue persists the census's
# missing list with per-entry attempt metadata; hash-matched downloads
# are sync-verified by construction, so the trickle needs no follow-up.

function Get-SubtitleQueuePath {
    return "$env:LOCALAPPDATA\LibraryLint\subtitle_acquire_queue.json"
}

function Update-SubtitleAcquireQueue {
    [CmdletBinding()]
    param([object[]]$MissingList = @())

    $path = Get-SubtitleQueuePath
    # Preserve attempt metadata for entries that survive the rebuild —
    # a no-hash-match stamp keeps an entry from burning quota-run time
    # every single day.
    $prior = @{}
    if (Test-Path -LiteralPath $path) {
        try {
            $loaded = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -AsHashtable
            foreach ($e in @($loaded.entries)) { $prior[$e.Folder] = $e }
        } catch {}
    }

    $entries = @()
    foreach ($m in $MissingList) {
        if ($prior.ContainsKey($m.Folder)) {
            $entries += $prior[$m.Folder]
        } else {
            $entries += @{ Folder = $m.Folder; RelativePath = $m.RelativePath; Attempts = 0; NoHashMatchAt = $null }
        }
    }

    $dir = Split-Path $path -Parent
    if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    @{ version = 1; updatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); entries = $entries } |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $path -Encoding UTF8
    return $entries.Count
}

<#
.SYNOPSIS
    Spends the daily OpenSubtitles quota against the acquisition queue.
.DESCRIPTION
    Hash-matched downloads ONLY — each success is saved as .en.srt,
    marked verified, and removed from the queue. Entries with no hash
    match are stamped and skipped on future runs — they're the Whisper
    candidates (Subtitle Health Check offers generation for them) — until
    RetryNoMatchAfterDays lapses — new sub uploads happen, so they get
    another look eventually. Self-healing: entries whose folder vanished
    or that gained a sub some other way are dropped on sight.
.OUTPUTS
    Hashtable: Downloaded, NoMatch, Healed, Pending, DeferredNoMatch,
    QuotaHit, Error.
#>
function Invoke-SubtitleQueueRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$OpenSubtitlesApiKey,
        [Parameter(Mandatory)] [string]$OpenSubtitlesUsername,
        [Parameter(Mandatory)] [string]$OpenSubtitlesPassword,
        [string[]]$VideoExtensions    = @('.mkv', '.mp4', '.avi', '.m4v', '.wmv', '.mov'),
        [string[]]$SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt'),
        [int]$MaxDownloads = 5,
        [int]$RetryNoMatchAfterDays = 30
    )

    $stats = @{ Downloaded = 0; NoMatch = 0; Healed = 0; Pending = 0; DeferredNoMatch = 0; QuotaHit = $false; ServiceUnavailable = $false; Error = $null }
    $path = Get-SubtitleQueuePath
    if (-not (Test-Path -LiteralPath $path)) { return $stats }

    try {
        $queue = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -AsHashtable
    } catch {
        $stats.Error = "queue file unreadable: $($_.Exception.Message)"
        return $stats
    }
    $entries = [System.Collections.Generic.List[object]]::new()
    foreach ($e in @($queue.entries)) { $entries.Add($e) }
    if ($entries.Count -eq 0) { return $stats }

    $junkNameRegex = '(?i)(^|[\.\-_\s])(trailer|sample|featurette|extras?)($|[\.\-_\s])'
    $now = Get-Date
    $eligible = @($entries | Where-Object {
        -not $_.NoHashMatchAt -or
        (($now - [datetime]::Parse($_.NoHashMatchAt)).TotalDays -ge $RetryNoMatchAfterDays)
    })
    $stats.DeferredNoMatch = $entries.Count - $eligible.Count
    if ($eligible.Count -eq 0) {
        $stats.Pending = 0
        return $stats
    }

    # Login only when there's something to try.
    $login = Connect-OpenSubtitles -ApiKey $OpenSubtitlesApiKey -Username $OpenSubtitlesUsername -Password $OpenSubtitlesPassword
    if (-not $login.Success) {
        $stats.Error = "OpenSubtitles login failed: $($login.Error)"
        $stats.Pending = $eligible.Count
        return $stats
    }

    # Lead-in so the per-item lines below have a visible parent — without
    # it, a failure on the first item printed as an orphan line floating
    # under whatever section came before.
    Write-Host ""
    Write-Host "  Subtitle queue : $($eligible.Count) pending — trying up to $MaxDownloads hash-matched download(s)" -ForegroundColor Cyan

    $toRemove = [System.Collections.Generic.List[object]]::new()
    $consecutiveFailures = 0
    foreach ($entry in $eligible) {
        if ($stats.Downloaded -ge $MaxDownloads -or $stats.QuotaHit) { break }

        # Self-heal: folder gone or sub arrived some other way -> drop.
        if (-not (Test-Path -LiteralPath $entry.Folder)) {
            $toRemove.Add($entry); $stats.Healed++; continue
        }
        $hasSub = Get-ChildItem -LiteralPath $entry.Folder -File -ErrorAction SilentlyContinue |
            Where-Object { $SubtitleExtensions -contains $_.Extension.ToLower() } | Select-Object -First 1
        if ($hasSub) {
            $toRemove.Add($entry); $stats.Healed++; continue
        }

        $video = Get-ChildItem -LiteralPath $entry.Folder -File -ErrorAction SilentlyContinue |
            Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and $_.Name -notmatch $junkNameRegex } |
            Sort-Object Length -Descending | Select-Object -First 1
        if (-not $video) { $toRemove.Add($entry); $stats.Healed++; continue }

        $hash = Get-OpenSubtitlesHash -Path $video.FullName
        $match = if ($hash) { Search-OpenSubtitlesByHash -ApiKey $OpenSubtitlesApiKey -MovieHash $hash } else { $null }
        if (-not $match) {
            $entry.NoHashMatchAt = $now.ToString('yyyy-MM-dd HH:mm:ss')
            $entry.Attempts = [int]$entry.Attempts + 1
            $stats.NoMatch++
            Start-Sleep -Milliseconds 400
            continue
        }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)
        $outPath = Join-Path $entry.Folder "$baseName.en.srt"
        $dl = Save-OpenSubtitlesFile -ApiKey $OpenSubtitlesApiKey -Token $login.Token -FileId $match.FileId -OutPath $outPath
        if ($dl.Success) {
            $null = Set-SubtitlesVerified -FolderPath $entry.Folder -Source "opensubtitles-hashmatch ($($match.Release))" -Provider 'opensubtitles'
            Write-Host "    + $($entry.RelativePath) — hash-matched, verified" -ForegroundColor Green
            $toRemove.Add($entry)
            $stats.Downloaded++
            $consecutiveFailures = 0
            Start-Sleep -Milliseconds 500
        } elseif ($dl.QuotaExhausted) {
            $stats.QuotaHit = $true
        } elseif ($dl.Error -match '503|Service Unavailable') {
            # Service-side outage — every further attempt fails identically.
            # Stop immediately instead of marching the whole queue into a
            # dead endpoint; the entry keeps its place (no Attempts bump,
            # the failure isn't its fault) and the next Status run retries.
            Write-Host "    ! OpenSubtitles unavailable (503) — stopping; auto-retry next run" -ForegroundColor Yellow
            $stats.ServiceUnavailable = $true
            break
        } else {
            $entry.Attempts = [int]$entry.Attempts + 1
            Write-Host "    X $($entry.RelativePath) — download failed: $($dl.Error)" -ForegroundColor Red
            $consecutiveFailures++
            if ($consecutiveFailures -ge 3) {
                # Three different entries failing back-to-back is a systemic
                # problem, not three unlucky files — same treatment as 503.
                Write-Host "    ! 3 consecutive failures — stopping; auto-retry next run" -ForegroundColor Yellow
                $stats.ServiceUnavailable = $true
                break
            }
        }
    }

    foreach ($r in $toRemove) { [void]$entries.Remove($r) }
    $stats.Pending = @($entries | Where-Object {
        -not $_.NoHashMatchAt -or (($now - [datetime]::Parse($_.NoHashMatchAt)).TotalDays -ge $RetryNoMatchAfterDays)
    }).Count
    $stats.DeferredNoMatch = $entries.Count - $stats.Pending

    @{ version = 1; updatedAt = $now.ToString('yyyy-MM-dd HH:mm:ss'); entries = $entries.ToArray() } |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $path -Encoding UTF8

    return $stats
}

#endregion

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

    $baseUrl = "https://api.subdl.com/api/v1/subtitles"
    $results = $null

    # The two searches are ALTERNATIVES, so each gets its own try — a
    # transient failure on the IMDB attempt must not skip the title
    # fallback, and a network error should say so rather than present
    # as "no results".

    # Try IMDB ID first (most accurate)
    if ($IMDBID) {
        try {
            $url = "$baseUrl`?api_key=$ApiKey&imdb_id=$IMDBID&languages=$Language&subs_per_page=10"
            $response = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 20 -ErrorAction Stop
            if ($response.status -and $response.subtitles -and $response.subtitles.Count -gt 0) {
                $results = $response.subtitles
            }
        } catch {
            Write-Host "  Subdl IMDB search failed ($($_.Exception.Message)) — trying title search" -ForegroundColor Yellow
        }
    }

    # Fallback to title search
    if (-not $results -and $Title) {
        try {
            $encodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)
            $url = "$baseUrl`?api_key=$ApiKey&film_name=$encodedTitle&languages=$Language&subs_per_page=10"
            if ($Year) {
                $url += "&year=$Year"
            }
            $response = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 20 -ErrorAction Stop
            if ($response.status -and $response.subtitles -and $response.subtitles.Count -gt 0) {
                $results = $response.subtitles
            }
        } catch {
            Write-Host "  Subdl title search failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    return $results
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
            Set-SubtitlesVerified -FolderPath $MovieFolder -Source "included" -Provider 'release'
        } elseif ($subCheck.EmbeddedEnglish -or $subCheck.EmbeddedSubs.Count -gt 0) {
            Set-SubtitlesVerified -FolderPath $MovieFolder -Source "embedded" -Provider 'embedded'
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

            # Provenance note for the verification gate — consumed when
            # the folder is eventually marked verified, so "came from
            # Subdl" survives even though no .subs_ok is written yet.
            Set-SubtitlePendingProvenance -FolderPath $MovieFolder -Provider 'subdl' -Detail (Split-Path $destPath -Leaf)

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
                        Set-SubtitlesVerified -FolderPath $MovieFolder -Source "subdl-synced" -Provider 'subdl'
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

    # Get all subtitle files. Skip anything still living in a Subs/ /
    # Subtitles/ subfolder — Repair-SubtitlePlacement owns those (it
    # either moves them out or deliberately leaves them when ambiguous).
    # Without this exclusion, an ambiguous sub Step 1 chose to keep gets
    # silently deleted here as "no matching video in this folder."
    $subFolderNamesExclude = @('subs', 'sub', 'subtitles', 'subtitle', 'srt')
    $subtitleFiles = Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object {
            $SubtitleExtensions -contains $_.Extension.ToLower() -and
            $subFolderNamesExclude -notcontains (Split-Path (Split-Path $_.FullName -Parent) -Leaf).ToLower()
        }

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
        $hasMatchingVideo = Get-ChildItem -LiteralPath $sub.DirectoryName -File -ErrorAction SilentlyContinue |
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
        $videoFiles = @(Get-ChildItem -LiteralPath $sub.DirectoryName -File -ErrorAction SilentlyContinue |
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
        $existingMatch = Get-ChildItem -LiteralPath $sub.DirectoryName -File -ErrorAction SilentlyContinue |
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
    $subFolders = Get-ChildItem -LiteralPath $Path -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -in $subFolderNames }

    if ($subFolders.Count -eq 0) {
        Write-Host "  No subtitle subfolders found" -ForegroundColor Gray
    } else {
        Write-Host "  Found $($subFolders.Count) subtitle subfolder(s)" -ForegroundColor Cyan

        foreach ($subFolder in $subFolders) {
            $stats.SubfoldersFound++
            $parentFolder = $subFolder.Parent.FullName

            # Enumerate ALL videos in the parent. For a movie folder this is
            # typically 1; for a TV season folder it's the episode list.
            $videosInParent = @(Get-ChildItem -LiteralPath $parentFolder -File -ErrorAction SilentlyContinue |
                Where-Object { $VideoExtensions -contains $_.Extension.ToLower() -and $_.Name -notmatch '-trailer\.' })

            if ($videosInParent.Count -eq 0) {
                Write-Host "  [skip]         $($subFolder.Parent.Name)\$($subFolder.Name) — parent folder has no video" -ForegroundColor DarkGray
                continue
            }

            # Index videos by basename (lowercase) for per-sub matching in
            # multi-video parents (TV season). For single-video parents the
            # index has one entry and any sub falls back to it cleanly.
            $videosByBase = @{}
            foreach ($v in $videosInParent) {
                $videosByBase[[System.IO.Path]::GetFileNameWithoutExtension($v.Name).ToLower()] = $v
            }

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

                # Pick the video this sub belongs to. Try a prefix match
                # against the basename index: a sub like "S01E03.eng" should
                # land on video "S01E03.mkv". This is the TV-shape fix —
                # previously every sub in Subs/ got renamed onto whichever
                # video was biggest in the parent.
                $targetVideo = $null
                foreach ($videoBase in $videosByBase.Keys) {
                    # Sub's basename starts with the video's basename (sub may
                    # add ".eng" / ".forced" etc., or match exactly).
                    if ($subNameLower -eq $videoBase -or $subNameLower.StartsWith("$videoBase.")) {
                        $targetVideo = $videosByBase[$videoBase]
                        break
                    }
                }

                if (-not $targetVideo) {
                    if ($videosInParent.Count -eq 1) {
                        # Single-video parent: fall back to that video. Covers
                        # movies where the sub is just "subtitle.srt" with no
                        # naming relationship to the video.
                        $targetVideo = $videosInParent[0]
                    } else {
                        Write-Host "  [skip]         $($subFolder.Name)\$($sub.Name) — no episode in parent matches this sub's basename" -ForegroundColor DarkGray
                        $stats.SubsSkipped++
                        continue
                    }
                }

                $videoBaseName = [System.IO.Path]::GetFileNameWithoutExtension($targetVideo.Name)
                $newSubName = "$videoBaseName$detectedLang$($sub.Extension)"
                $newSubPath = Join-Path $parentFolder $newSubName

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

    # Walk recursively for every folder that contains at least one video file.
    # Movie libraries → one folder per movie. TV libraries → one folder per
    # season (where the episode files live). Both shapes work without the
    # caller having to know the library type.
    $videoExtPattern = '\.(mkv|mp4|avi|mov|wmv|m4v)$'
    $videoDirs = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match $videoExtPattern } |
        ForEach-Object { [void]$videoDirs.Add($_.DirectoryName) }

    $totalFolders = $videoDirs.Count
    $processed = 0

    # Inherited-verification helper: a TV show might have its .subs_ok marker
    # at the show root rather than per-season. Walk ancestors up to (but not
    # past) $Path looking for any marker that covers this folder.
    $rootNorm = $Path.TrimEnd('\','/').ToLower()
    $verifiedCache = @{}
    $isVerifiedInherited = {
        param($folder)
        if ($verifiedCache.ContainsKey($folder)) { return $verifiedCache[$folder] }
        $node = $folder
        while ($node) {
            if (Test-SubtitlesVerified -FolderPath $node) {
                $verifiedCache[$folder] = $true
                return $true
            }
            $nodeNorm = $node.TrimEnd('\','/').ToLower()
            if ($nodeNorm -eq $rootNorm) { break }
            $parent = Split-Path $node -Parent
            if (-not $parent -or $parent -eq $node) { break }
            $node = $parent
        }
        $verifiedCache[$folder] = $false
        return $false
    }

    Write-Host "Scanning $totalFolders video-bearing folder(s)..." -ForegroundColor Cyan

    foreach ($folder in $videoDirs) {
        $processed++
        $folderLeaf = Split-Path $folder -Leaf
        $percentComplete = [math]::Round(($processed / [Math]::Max(1, $totalFolders)) * 100)
        Write-Progress -Activity "Syncing Subtitles" -Status "$processed of $totalFolders - $folderLeaf" -PercentComplete $percentComplete

        # Skip if verified (own marker OR ancestor inherits) unless -Force.
        if (-not $Force -and (& $isVerifiedInherited $folder)) {
            $stats.AlreadyVerified++
            continue
        }

        $videos = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match $videoExtPattern -and $_.Name -notmatch '-trailer\.' })

        if ($videos.Count -eq 0) {
            $stats.NoVideo++
            continue
        }

        # Find all subtitle files in this folder up front; we'll match per
        # video below. Doing it once avoids re-enumerating per episode.
        $allSubs = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '\.(srt|sub|ass|ssa)$' })

        if ($allSubs.Count -eq 0) { continue }

        $folderHadFailure = $false

        foreach ($video in $videos) {
            $videoBase = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)

            # Match subs to this video by basename, with optional .lang or
            # .lang.modifier suffix. Movie shape (one video, plain "Movie.srt"
            # next to "Movie.mkv") → exact basename match. TV shape (each
            # episode has its own .srt) → only that episode's subs match.
            $matchingSubs = @($allSubs | Where-Object {
                $subBase = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                $subBase -eq $videoBase -or $subBase -like "$videoBase.*"
            })

            # Movie-shape fallback: if no per-video match found AND there's
            # only one video in this folder, accept any sub in the folder
            # (typical "Movie.mkv" + "Movie.eng.srt" or "subtitle.srt").
            if ($matchingSubs.Count -eq 0 -and $videos.Count -eq 1) {
                $matchingSubs = $allSubs
            }

            foreach ($sub in $matchingSubs) {
                $stats.Total++

                Write-Host "  [$processed/$totalFolders] $folderLeaf - $($sub.Name)" -ForegroundColor Gray -NoNewline

                if ($WhatIf) {
                    Write-Host " [would sync]" -ForegroundColor Cyan
                    $stats.Synced++
                    continue
                }

                if ($Backup) {
                    $result = Invoke-FFSubSync -VideoPath $video.FullName -SubtitlePath $sub.FullName -Backup
                    $stats.BackupsCreated++
                } else {
                    $result = Invoke-FFSubSync -VideoPath $video.FullName -SubtitlePath $sub.FullName
                }

                if ($result) {
                    Write-Host " [synced]" -ForegroundColor Green
                    $stats.Synced++
                } else {
                    Write-Host " [failed]" -ForegroundColor Red
                    $stats.Failed++
                    $folderHadFailure = $true
                }
            }
        }

        # Mark this folder verified only if every sub processed THIS pass
        # succeeded. The original tracked $stats.Failed globally, which
        # meant one failure in episode A poisoned the marker for episode B
        # in a different folder. Local tracking fixes that.
        if (-not $WhatIf -and -not $folderHadFailure) {
            Set-SubtitlesVerified -FolderPath $folder -Source "ffsubsync"
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
    Remove-SubtitlesVerified, Get-VerifiedSubtitleStatus, Get-SubtitleLanguageCode, Invoke-SubtitleLanguagePrune,
    Get-SubtitleHealthSnapshot, Get-SrtFirstCueStart, Get-SrtCueSample, Invoke-SubtitleSyncAudit,
    Test-FFSubSyncInstallation, Invoke-FFSubSync,
    Search-SubdlSubtitle, Save-MovieSubtitle,
    Repair-OrphanedSubtitles, Repair-SubtitlePlacement, Invoke-SubtitleSync, Restore-SubtitleBackups,
    Get-EmbeddedSubtitleInfo, Get-SubtitleCensus, Get-SubtitleCensusSummaryPath, Invoke-EmbeddedSubtitleExtraction, Invoke-SubtitleVerificationPass,
    Get-OpenSubtitlesHash, Connect-OpenSubtitles, Search-OpenSubtitlesByHash, Save-OpenSubtitlesFile, Invoke-SubtitleAcquisition,
    Get-SubtitleQueuePath, Update-SubtitleAcquireQueue, Invoke-SubtitleQueueRun,
    Get-WhisperPython, Test-WhisperInstallation, Invoke-WhisperSubtitle, Invoke-WhisperSubtitlePass
