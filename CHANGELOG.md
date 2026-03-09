# Changelog

All notable changes to LibraryLint will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [5.5.1] - 2026-03-09

### Added
- **Incomplete Collections utility** — scans NFO files for set membership, queries TMDB for full collection parts, and displays owned vs. missing movies highlighted by color
- **`Get-TMDBCollectionParts`** function in TMDB module — retrieves all movies in a TMDB collection by ID
- **NFO quality check** now validates IMDB ID, rating, premiered date, genres, director, actors, set, and trailer (conditional on DownloadTrailers + local trailer)
- **New auto-clean tags** — EC, IMAX, REPACK2, REPACK3, and anniversary variants (10th–100th)
- **`Test-NFOMatchesFolder` validation** in metadata refresh — removes NFO and logs error when TMDB returns a wrong match
- **Ampersand normalization** — `&` in folder names cleaned to ` & ` during folder cleanup and to spaces during title normalization
- **Fix & Repair path defaults** — now offers Movies, TV Shows, or Browse instead of requiring folder dialog every time

### Improved
- **TMDB search reliability** — results are now scored by title similarity (exact match, containment, word overlap) instead of blindly accepting the first result; rejects matches below quality threshold
- **TMDB search title variations** — automatically retries with hyphens→slashes for numeric titles (e.g., "50-50" → "50/50") and removes "Chapter One/Two/..." suffixes (e.g., "It Chapter One" → "It")
- **Year extraction** — prefers parenthesized years `(YYYY)` over bare year numbers to protect titles like "Blade Runner 2049", "2001", and "1917"
- **Tag stripping safety** — preserves year through tag removal and requires letters in result to prevent title wipeouts (e.g., "Real Steel" no longer destroyed by REAL tag)
- **Standalone movie NFOs** now include `<set/>` empty tag to prevent repeated quality-check re-flagging

### Fixed
- **Saved config overriding tags** — Tags array is now always sourced from script defaults; saved config tags are ignored during import and excluded from export
- **Folder name cleanup preprocessing** — `Rename-CleanFolderNames` runs before metadata refresh to break the chicken-and-egg cycle of dirty names preventing TMDB lookup

## [5.5.0] - 2026-03-07

### Added
- **Quality module** (`modules/Quality.psm1`) — extracted 6 quality/codec functions (Get-QualityConcerns, Get-QualityScore, Get-VideoCodecInfo, Invoke-CodecAnalysis, Invoke-Transcode, New-TranscodeScript) into a standalone module
- **`Invoke-QualityScore` wrapper** in main script to pre-compute MediaInfo/ReleaseInfo/HDR data for the Quality module
- **`-QualityScorer` scriptblock parameter** for `Invoke-CodecAnalysis` to enable full MediaInfo-backed quality scoring from the module
- **`RERip` tag** added to the auto-removed release tags list
- **Automatic yt-dlp update check** — runs `yt-dlp -U` once per session before trailer downloads to prevent failures from outdated versions
- **Dependency update checker** (Utilities > Check for Dependency Updates) — checks winget for available updates to 7-Zip, MediaInfo, FFmpeg, and yt-dlp with optional one-click upgrade
- **SFTP bandwidth throttling** — configurable download speed limit (`SFTPSpeedLimitKBps`) to avoid saturating connections on shared seedboxes
- **SFTP file exclusion patterns** — skip files by wildcard during sync (e.g., `*.nfo`, `*.txt`, `Sample*`) via `SFTPExcludePatterns` config
- **Python script v5.1** — major update to match core PowerShell functionality:
  - Interactive menu system with persistent loop
  - Inbox processing with auto-detection (Movie vs TV Show)
  - Library transfer (inbox to main collection)
  - Configuration persistence (`~/.librarylint/config.json`)
  - First-run setup wizard
  - Settings menu with live toggle support

### Improved
- **Module loader** — generic loop over `$script:ModuleNames` array replaces hardcoded per-module imports; automatically initializes `*ModuleLoaded` flags for all modules
- **Main script reduced** to ~15,500 lines (down from ~18,900) through module extraction
- **SFTP server configuration wizard** — auto-discovers SSH keys in `~/.ssh/`, warns about OpenSSH vs PuTTY .ppk format, shows `ssh-keygen` / `ssh-copy-id` commands for new setups

### Changed
- **Python script** — remains experimental; updated core functionality to gauge cross-platform interest

## [5.4.7] - 2026-03-05

### Added
- **"X. Exit" option** in all top-level submenus (Help, Fix & Repair, Enhancements, Utilities, Settings) to exit the app directly without navigating back to the main menu first

### Improved
- **SFTP Prune scan output** - files grouped by folder instead of listed individually; shows folder name, file count, size, and age
- **SFTP Prune deletion output** - reports results per folder with deleted/failed/already-gone counts instead of per-file status
- **SFTP Prune size summary** moved to bottom of scan listing for better readability

## [5.4.6] - 2026-03-05

### Added
- **`DownloadArtwork` config setting** - global toggle to enable/disable artwork downloads (posters, fanart, logos) during inbox processing and metadata refresh; defaults to enabled for backward compatibility
- **Artwork toggle in processing setup** - "Download artwork?" prompt shown alongside existing trailer/subtitle options when customizing settings before a run
- **Artwork toggle in persistent settings** - configurable in Settings > Quality Preferences alongside trailer and NFO options
- **Logo** - added project logo to `assets/logo.png` and README header

### Improved
- **README branding** - centered header with logo, updated version badge to 5.4.5
- **GitHub Releases** - release notes now reference CHANGELOG.md instead of duplicating content

## [5.4.5] - 2026-03-03

### Added
- **Movie Set Artwork (MSIF)** - scan movie library for TMDB collection membership and download set artwork (poster, fanart, clearlogo, banner, landscape, disc, clearart) to a dedicated Kodi Movie Set Information Folder
- **MSIF first-run setup** - inline guided setup when path not configured: suggests location based on library, offers to add to mirror backup folders, shows Kodi configuration hint
- **`MovieSetArtworkPath` config** - new library path setting for Kodi MSIF, configurable in Settings > Configure Library Paths with browse/status support
- **Collection info from TMDB** - `Get-TMDBMovieDetails` now extracts `CollectionId` and `CollectionName` from `belongs_to_collection`
- **`set.nfo` generation** - creates Kodi-compatible NFO with collection name and TMDB collection ID for each movie set

### Improved
- **Manage Artwork menu** - redesigned with API key status display, explicit delete/download sections, and new Movie Sets section (option 5)
- **Menu reorganization** - moved Manage Artwork from Fix & Repair to Enhancements to eliminate overlap; renumbered Fix & Repair options
- **`-Quiet` parameter** for `Invoke-ArtworkManagement` - silent operation during bulk refresh flows
- **Changelog** - backfilled entries for v5.3.0 through v5.4.4

### Fixed
- **Robocopy "Removed X of Y" spam** - re-added `/NP` flag to suppress console progress lines that bypassed `ReadLine()` capture
- **Misaligned box-drawing borders** - Processing Summary, Help, Duplicate Report, Health Check, and Codec Analysis boxes had borders 2 chars narrower than content rows; widened and re-centered all

## [5.4.2] - 2026-03-03

### Improved
- **Mirror output** - fix ghost text from scan line, consolidate file copy errors to one line, add file counter to progress bar, show expected file count and size in folder header, summarize per-folder copy errors
- **Deferred duplicate scan** - health check no longer scans twice; defers to fix menu, runs enhanced detection once when user opts in
- **Duplicate resolution menu** - clarified wording to "lower-quality copies" to make clear the best version is kept

## [5.4.1] - 2026-03-02

### Added
- **Season thumbs** - download `season##-landscape.jpg` from fanart.tv `seasonthumb`
- **Keyart** - download `keyart.jpg` (theatrical poster) for movies from fanart.tv (Kodi 21+ Omega)
- **Season-all artwork** - `season-all-poster.jpg`, `season-all-banner.jpg`, `season-all-landscape.jpg` for TV shows
- **Plex `logo.png`** - copy of `clearlogo.png` for Plex compatibility (movies and TV shows)

### Fixed
- **Colon-safe folder naming** - `Get-SafeFileName` now converts colons to ` - ` instead of bare `-` (e.g., "Dune: Part Two" becomes "Dune - Part Two")
- **Duplicate detection evasion** - `Get-NormalizedTitle` now normalizes hyphens adjacent to spaces on either side, catching folders like "Dune- Part Two" vs "Dune Part Two"

## [5.4.0] - 2026-03-02

### Added
- **Clean Junk Files** - new Fix & Repair option to remove trailer NFOs, sample videos, and orphaned NFO files
- **VC-1/WMV transcode** - auto-detect legacy VC-1 codec and offer H.265 transcode (CRF 28)
- **Sortable HTML reports** - library report columns are now clickable for ascending/descending sort

### Improved
- **Mirror backup overhaul** (Samba/network share compatibility):
  - Fix robocopy exit code 16: remove `/Z`, add `/COPY:DT /DCOPY:T /FFT`
  - Pre-flight path validation with clear OK/NOT FOUND output
  - Streaming line-by-line progress replaces blocking `ReadToEnd()`
  - Per-file progress with transfer speed, ETA, and current filename
  - Per-folder and overall average speed summaries
  - Bump `/MT:16`, reduce retries `/R:2 /W:5`
  - Early skip when source and dest are fully in sync
- **NFO mismatch handling** - interactive menu (match NFO, match folder, review, transfer) replaces Y/N prompt; fixes stale FolderPath after renames
- **Stats reset** between processing runs
- **Trailer/sample skipping** during NFO generation

## [5.3.3] - 2026-02-19

### Added
- **NFO validation** - new `Test-NFOMatchesFolder` function validates that NFO metadata matches the folder it's in using three-tier matching (exact, containment, Jaccard similarity) with year verification
- **NFO mismatch detection in health check** - health check now identifies folders where the NFO file contains metadata for the wrong movie and offers to delete and regenerate from TMDB
- **Transfer block on bad NFOs** - inbox processing blocks library transfer when TMDB returned the wrong movie, preventing library pollution; lists each mismatch with details
- **Duplicate check opt-out** - health check now asks before running the duplicate movie check (can be slow on large libraries)

### Improved
- **TMDB search accuracy** - `Get-NormalizedTitle -Strict` now used in all TMDB search paths to preserve leading articles ("The", "A", "An"); added retry-without-year fallback in `Invoke-MetadataRefresh` and `Repair-MovieFolderYears`
- **Metadata refresh** - replaced hand-rolled title extraction in `Invoke-MetadataRefresh` with `Get-NormalizedTitle`, fixing failures for titles with release tags
- **Folder rename safety** - NFO mismatch detection now uses containment check: only flags folder for rename when NFO title is a shorter/cleaner version of the folder title (folder has extra tags); rejects when NFO title is longer or completely different (bad NFO, not bad folder)
- **Fix & Repair menu reordered** - follows dependency chain: metadata refresh → folder names → file names → subtitles → artwork → empty folders → duplicates
- **Health check fix actions reordered** - follows same dependency chain: cleanup → NFO fixes → folder names → file names → subtitles → duplicates
- **Tag stripping word boundaries** - `Rename-CleanFolderNames` now enforces word boundaries when matching tags, preventing partial matches like "ENG" inside "Vengeance"
- **Tag list expanded** - added 'UNCUT', 'Directors Cut' (without apostrophe), and 'DANISH' to release tag list
- **Duplicate detection report** - box-drawing characters now use Unicode char codes for consistent rendering across terminal encodings
- **Trailer key validation** - filters out invalid YouTube video IDs (too short) before attempting download

### Fixed
- **Double-named file fix** - rewrote suffix extraction to scan backwards from end of filename using a whitelist of known extensions; immune to dots in titles like "When Harry Met Sally..." and "Say Anything..."
- **Orphaned file rename creating duplication** - orphan code was matching files already correctly named when the old basename from `release-info.json` was a prefix of the current name (e.g., "Terminator 3" matched "Terminator 3- Rise of the Machines (2003)"); now skips files that already start with the correct basename
- **Aggressive tag stripping destroyed movie titles** - tags like REAL, LiMiTED, DC, WEB matched words in actual titles ("Real Steel" → empty, "Blade Runner 2049" → "Blade Runner"); removed tag/year post-processing from display titles — TMDB and NFO titles are now used as-is
- **Bad NFO causing wrong renames** - folders like "X-Men: First Class" were being renamed to match wrong NFO metadata from bad TMDB matches; containment check now protects correctly-named folders
- **Wildcard escaping in file matching** - `[WildcardPattern]::Escape()` now used for all `-like` comparisons with folder/file basenames, preventing characters like `[` and `]` in names from being interpreted as wildcards

## [5.3.2] - 2026-02-18

### Added
- **Fix File Names tool** - new Fix & Repair option that renames video files to match their folder name (Kodi/Plex/Jellyfin standard naming); also renames associated files (NFO, subtitles, keyart) to match
- **Structured release info** - new `release-info.json` format stores parsed metadata (resolution, source, codec, audio, HDR, edition, release group, streaming service) extracted from original filenames; replaces plain-text `release-info.txt`
- **Release info migration** - existing `release-info.txt` files are automatically converted to JSON when folders are re-processed; `Read-ReleaseInfo` falls back to `.txt` for unprocessed folders
- **Orphaned file fix** - "Fix File Names" detects and renames leftover associated files (e.g., keyart) that weren't renamed with the video, using the original filename from `release-info.json`
- **Dismiss duplicates in health check** - option to dismiss duplicate groups that aren't actual duplicates, clearing them from the action menu

### Improved
- **Error details at transfer** - transfer warnings now list each specific error instead of just a count
- **Non-critical errors downgraded** - artwork downloads, subtitle search/download, trailer processing, and ffsubsync failures no longer count as errors (changed from ERROR to WARNING)
- **Fix & Repair menu** - added "Fix File Names" as option 3; renumbered remaining options accordingly
- **Health check** - now detects video files that don't match their folder name and offers a fix action with dry-run preview

## [5.3.1] - 2026-02-18

### Added
- **Full-package auto-update** - updater now downloads the portable ZIP (script + modules + bat) instead of just the .ps1 file; preserves user config, backs up entire install directory, lists all updated files
- **Metadata refresh summary** - lists all corrected folder names at the end of a metadata refresh run (movies and TV shows)
- **Duplicate video file cleanup** - detects multiple video files in the same folder (e.g., BRRip + Criterion DVDRip), keeps the best quality, removes the rest along with their NFO/subtitle files
- **Folder merge on rename conflict** - when folder renaming produces a name that already exists, merges contents intelligently instead of silently failing (compares video quality, keeps best)
- **Duplicate detection in health check** - Full Health Check now includes duplicate movie detection alongside existing checks

### Improved
- **Fix & Repair menu** - redesigned layout with visual separators; option 1 (Full Health Check) clearly runs all individual tools, options 2-7 are standalone, Codec Analysis separated as option 8
- **Health check fix menu** - now loops after each action instead of exiting; auto-exits when all issues are resolved
- **Health check orphaned subtitles** - changed from delete to repair; now renames subtitles to match video files using `Repair-SubtitlePlacement`
- **Quality concern thresholds** - adjusted all bitrate and file-size thresholds for HEVC/x265, which achieves the same quality at ~40-50% lower bitrates than x264 (e.g., 1080p x265 "very low" is now 1 Mbps instead of 2 Mbps)
- **Sample file detection** - catches common misspellings like "SAMPE" and "SMAPLE" in addition to "SAMPLE"

### Fixed
- **Duplicate folder transfer** - folders that failed to rename (because target already existed) were silently kept and transferred alongside the correct folder
- **WinSCP console output** - disabled transfer progress output that persisted in the terminal during unrelated operations
- **Redundant duplicate check** - removed Step 11 duplicate folder scan from inbox processing (now handled by merge-on-rename and duplicate video cleanup)

## [5.3.0] - 2026-02-18

### Added
- **Uninstall utility** - new Utilities menu option with three modes: uninstall dependencies via winget, clean up LibraryLint data (config, logs, undo manifests), or full uninstall
- **`Uninstall-WingetPackage`** helper function mirroring `Install-WingetPackage`
- Selectable dependency removal with status display and confirmation prompts
- Data cleanup shows folder sizes and requires `YES` confirmation for safety

## [5.2.8] - 2026-02-15

### Added
- **Unified Process Inbox** - single menu option replaces "Process Movies", "Process TV Shows", and "Process All"; auto-detects media type per folder using episode patterns, season folders, year detection, and file count heuristics
- **Smart defaults** - Process Inbox shows a settings summary and lets you press Enter to go, or N to customize each option
- **API key management** - new "Manage API Keys" option in Settings menu to add, edit, validate, and clear API keys (TMDB, TVDB, Fanart.tv, Subdl)
- **Actionable health check** - health check results now offer an action menu to fix findings (delete empty folders, zero-byte files, orphaned subtitles, small videos, generate missing NFOs, fix naming issues)

### Improved
- **HTML-only exports** - replaced CSV exports with styled HTML reports that open natively in any browser, with search and dark theme
- **Cancel options** - all processing menus now show a clear "Cancelled." message when folder selection is dismissed
- **MediaInfo install** - wildcard path matching for winget versioned directories, PATH refresh after install, reuses detection function instead of duplicating paths
- **Dependency installer** - added `--exact` flag to winget installs, handles "already installed" exit code, refreshes PATH after every install
- **Help text** - corrected MediaInfo winget package ID to `MediaArea.MediaInfo.CLI`

### Changed
- **Menu simplified** - main menu reduced from 7 options to 5 (renumbered: Fix & Repair=2, Enhancements=3, Utilities=4, Settings=5)
- **Single inbox** - setup wizard now asks for one shared inbox folder instead of separate Movies/TV inboxes

### Fixed
- **PowerShell 5.1 compatibility** - fixed UTF-8 encoding for box-drawing characters in stock PowerShell console
- **Batch launcher** - replaced ANSI escape codes with plain ASCII for reliable rendering across all terminals
- **Dependency install** - shows manual download URLs when winget is unavailable

## [5.2.7] - 2026-02-15

### Added
- **Artwork management** - new "Manage Artwork" option in Fix & Repair menu to clear, refresh, or selectively remove artwork types (poster, fanart, clearlogo, etc.)
- **Artwork sync** - intelligent artwork download that only fetches missing artwork without replacing existing files
- **Version in main menu** - version number now displayed in the menu header for easy identification

### Improved
- **Folder dialog** - now opens to the configured library path instead of starting at root
- **Tag stripping** - added spaced variants of common tags (e.g., "HD Rip", "WEB DL", "Blu Ray", "DVD Rip", "BD Rip")
- **Release group tags** - added "NW" streaming tag

### Fixed
- **Mirror backup** - improved time formatting and drive space reporting

## [5.2.6] - 2026-02-07

### Added
- **Drive capacity checking** - warns before SFTP sync and Mirror backup when space is low, blocks when insufficient
- **Year±1 duplicate matching** - catches year typos in duplicate detection (e.g., 2023 vs 2024)
- **Bonus content tags** - title normalization strips "Behind the Scenes", "Making Of", "Featurette", etc.

### Fixed
- **Metadata refresh** - strips bonus content tags before TMDB search to prevent mismatches

## [5.2.5] - 2026-02-02

### Added
- **Movie Discovery** - cross-reference library against TMDB/IMDb lists to find movies you don't have
- **Seedbox Management** - restart Deluge, rTorrent, Transmission services remotely via SSH
- **Combined Seedbox menu** - unified SFTP Sync, Manage Services, and Configure Server

### Improved
- Renamed "Strip to Video Only" to "Strip Folders Down to Video Only" for clarity
- SFTP Sync now processes largest video files first to fix companion file routing
- Added .tmp to companion file extensions
- Added "German", "US Cut", "UK Cut", "International Cut" to title normalization tags

### Fixed
- Settings menu not returning to main menu on "0"
- Utilities menu loop issue

## [5.2.4] - 2026-01-28

### Added
- **Color splash on launch** - batch launcher now displays LibraryLint branding on startup

### Changed
- Removed duplicate `LibraryLint.bat` (keeping `Run-LibraryLint.bat`)

## [5.2.3] - 2026-01-28

### Added
- **Interactive help menu** - press `?` from main menu for help on getting started, menu options, dependencies, API keys, file locations, and command-line options
- **Mirror backup ETA** - shows estimated time remaining during mirror operations
- **Inno Setup installer script** - ready for building Windows installer packages
- **Build script** - `Build-Release.ps1` for creating installer and portable ZIP
- **PowerShell module manifest** - `LibraryLint.psd1` for future PowerShell Gallery publishing

### Fixed
- **Mirror backup performance** - restored multi-threaded copying (was accidentally running single-threaded)

## [5.2.2] - 2026-01-28

### Added
- **Auto-updater** - check for new versions from GitHub with `-Update` flag or Settings menu
- **Batch launcher** - `Run-LibraryLint.bat` for double-click execution
- **Extended movie artwork** - clearart, extrafanart folder (from fanart.tv)
- **Extended TV show artwork** - clearlogo, banner, clearart, characterart, landscape (from fanart.tv)
- **TV show actor images** - downloads actor photos from TVDB
- **Season banners** - from fanart.tv
- **Improved NFO validation** - checks for empty fields and minimum plot length
- **Rename episodes preference** - can now be saved to config

### Fixed
- **HTML entity decoding** - fixed `&apos;` appearing in TV show folder names
- **Media type selection** - Fix/Repair menus now require explicit choice (no saved default)

## [5.2.1] - 2026-01-15

### Fixed
- **SFTP sync improvements** - better handling of folder structure

## [5.2.0] - 2026-01-12

### Added
- **TVDB API integration** for TV shows (search, metadata, episodes)
- **TV show NFO generation** (`tvshow.nfo` and episode-level NFOs)
- **TV show artwork downloads** (poster, fanart, season posters, episode thumbnails)
- **TV show metadata refresh** support in Fix & Repair menu
- **TV show folder name repair** - fix casing, years, and title mismatches for TV shows
- **TV show auto-move** - transfer processed shows from inbox to library with season merging
- **Screener/CAM/TS detection** in quality concerns (DVDScr, HDCAM, Telesync, etc.)
- **Trailer fallback support** - tries alternate trailers if primary is unavailable/private
- **Auto-install dependencies** via winget (7-Zip, MediaInfo, FFmpeg, yt-dlp)
- **SFTP folder structure preservation** - maintains remote folder hierarchy when downloading
- **Improved TV show detection** in SFTP sync (E01 patterns, "Complete Series", etc.)
- **Version management** - `-Version` flag and version display in header
- **Plex compatibility** - `background.jpg` alongside `fanart.jpg` for TV shows
- **Setup wizard improvements**:
  - `-Setup` flag to re-run setup wizard anytime
  - FFmpeg path prompt with auto-detection
  - Quality preferences (trailer quality, NFO generation, download trailers toggle)
  - Existing values shown as defaults when re-running setup
- **Subtitle download modes** - choose between All movies, Foreign films only, or None
- **Auto subtitle sync** - optional ffsubsync integration to fix timing after download
- **Failed trailer list** - shows which movies failed trailer download at end of run
- **Per-session auto-move toggle** - choose whether to move to library before processing starts

### Fixed
- **Transcode preserves NFO/images** when video extension changes (e.g., .avi to .mkv)
- **Trailer IAMF codec warning** - now prefers AAC audio to avoid experimental codecs
- **Movie title priority** - NFO title now takes precedence over TMDB search results
- **Folder name mismatch detection** - Fix tool now catches when folder name doesn't match NFO
- **No-op rename proposals** - removed confusing identical current/new name suggestions
- **SFTP destination paths** - fixed double "Inbox" in download paths
- **Duplicate Remove-EmptyFolders function** - consolidated into single robust implementation

### Changed
- Progress display added to folder scanning operations

## [5.1.0] - 2026-01-10

### Added
- **Modular architecture** with `modules/` folder for better organization
- **SFTP Sync module** (`modules/Sync.psm1`) - download files from seedbox/remote server
- **Mirror Backup module** (`modules/Mirror.psm1`) - robocopy-based mirroring to external drives
- **First-run setup wizard** - guided configuration for new users
- **Library transfer feature** - move processed movies from inbox to main collection
- **Example config file** (`config/config.example.json`)

### Fixed
- **Bracket path handling** - fixed issues with brackets `[]` in filenames using `-LiteralPath`
- **File renaming** to remove brackets from filenames for better compatibility
- **Improved bitrate warnings** - better detection of encodes vs source files

## [5.0.0] - 2026-01-01

### Added
- **Rebranded** from MediaCleaner to LibraryLint
- **Dry run reports** - auto-generated HTML report showing proposed changes
- **Configuration file** - save/load settings to JSON file
- **MediaInfo integration** - accurate codec detection from file headers
- **Undo/rollback support** - manifest-based rollback of changes
- **Enhanced duplicate detection** - file hashing for exact duplicates
- **Export reports** - CSV, HTML, and JSON library exports
- **Retry logic** - automatic retry of failed operations with backoff
- **Verbose mode** - optional debug output with `-Verbose` flag
- **Safe filename sanitization** for Windows compatibility

### Changed
- Complete rewrite with improved architecture
- Better error handling throughout
