# Changelog

All notable changes to LibraryLint will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [5.2.3] - 2026-01-28

### Added
- **Interactive help menu** - press `?` from main menu for help on getting started, menu options, dependencies, API keys, file locations, and command-line options

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
