# Changelog

All notable changes to LibraryLint will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
