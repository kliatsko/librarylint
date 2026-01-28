# TODO

## Branding

- [ ] Design logo for LibraryLint (magnifying glass + film reel, lint roller concept, etc.)

## Features

### TV Show Improvements
- [x] Add TVDB API key to config
- [x] Create TVDB API validation function
- [x] Create TV show search/metadata fetch functions
- [x] Add TV show NFO generation with TVDB data
- [x] Add TV show artwork downloads (poster, fanart, season posters)
- [x] Update Process New TV Shows to use TVDB
- [x] Update Refresh/Repair Metadata to support TV shows
- [x] Add Plex/Jellyfin/Emby compatibility for TV shows:
  - [x] Support `tvshow.nfo` at show level
  - [x] Support episode-level `.nfo` files
  - [x] Add `background.jpg` alongside `fanart.jpg` (Plex compatibility)
- [x] Add "Fix Folder Names" support for TV shows (casing fixes like movies)

### Auto-Move from Inbox to Main Library
- [x] Add `MoviesLibraryPath` and `TVShowsLibraryPath` to config
- [x] Create move function with duplicate checking
- [x] Add confirmation prompt before moving
- [x] Integrate into Process New Movies workflow
- [x] Option to enable/disable auto-move per session
- [x] Integrate into Process New TV Shows workflow

### Combined Processing
- [x] Add "Process All (Movies + TV)" menu option (option 3)
- [x] Shared settings prompt (dry-run, keep subtitles, auto-move)
- [x] Run movie processing on MoviesInboxPath, then TV processing on TVShowsInboxPath
- [x] Save previous inbox path for TV shows (like movies does)

### Initial Setup Wizard
- [x] Create first-run setup process that detects missing config
- [x] Prompt user for media library paths
- [x] Prompt for API keys (TMDB, Subdl)
- [x] Save answers to config file so script doesn't ask on every run
- [x] Prompt for preferred codecs and quality settings
- [x] Prompt for FFmpeg path (or auto-detect)
- [x] Add `-Setup` flag to re-run setup wizard

### User-Friendly Distribution
- [ ] Create installer package (MSI or self-extracting exe)
- [ ] Bundle FFmpeg or provide guided download
- [ ] Add auto-updater functionality to check for new versions
- [ ] Create simple "double-click to run" experience (no PowerShell knowledge required)
- [ ] Add uninstaller
- [ ] Consider a simple GUI wrapper for non-technical users

### Version Management
- [x] Add `$Version` variable at top of script as single source of truth
- [x] Add `--version` / `-Version` flag to display current version
- [x] Use Semantic Versioning (MAJOR.MINOR.PATCH) format
- [x] Create CHANGELOG.md to document changes per release
- [ ] Tag GitHub releases to match script version
- [ ] (Optional) Add PowerShell module manifest (.psd1) for future PowerShell Gallery publishing

## Improvements

- [x] Improve README with step-by-step first-timer instructions
- [x] Create a "Quick Start" guide (GETTING_STARTED.md)
- [ ] Add screenshots/examples to documentation

## Bugs

- (none currently)

## Technical Debt

- (none currently)

## Completed in v5.2.2

- [x] Download Missing Artwork enhancement option
- [x] Extended movie artwork: clearart, extrafanart folder (fanart.tv)
- [x] TV show artwork from fanart.tv: clearlogo, banner, clearart, character art, landscape
- [x] TV show actor images from TVDB
- [x] Season banners from fanart.tv
- [x] Improved NFO quality validation (checks for empty fields, minimum plot length)
- [x] Fixed HTML entity decoding in TV show folder names (`&apos;` → `'`)
- [x] Media type selection now requires explicit choice in Fix/Repair menus
- [x] Rename episodes preference can be saved to config
- [x] Interactive TV library path setup when not configured

## Completed in v5.2

- [x] TVDB API integration for TV shows (search, metadata, episodes)
- [x] TV show NFO generation (`tvshow.nfo` and episode NFOs)
- [x] TV show artwork downloads (poster, fanart, season posters, episode thumbs)
- [x] TV show metadata refresh support
- [x] Screener/CAM/TS detection in quality concerns
- [x] Trailer fallback support (tries alternate trailers if primary is unavailable)
- [x] Auto-install dependencies via winget (7-Zip, MediaInfo, FFmpeg, yt-dlp)
- [x] SFTP sync folder structure preservation
- [x] Improved TV show detection in SFTP sync (E01 patterns, "Complete Series", etc.)
- [x] Transcode preserves associated NFO/image files when extension changes
- [x] Fixed trailer IAMF codec warning (prefer AAC audio)
- [x] Fixed movie title priority (NFO > TMDB > folder name)
- [x] Fixed folder name tool to detect title mismatches from NFO
- [x] Fixed no-op rename proposals in folder fix tool
- [x] Added progress display to folder scanning
- [x] Fix Folder Names support for TV shows (casing, years, title mismatches)
- [x] TV show auto-move from inbox to library (with season merging)
- [x] `-Setup` flag to re-run setup wizard anytime
- [x] FFmpeg path prompt in setup (with auto-detection)
- [x] Quality preferences in setup (trailer quality, NFO generation)
- [x] Subtitle download modes (All/Foreign/None) with foreign film detection
- [x] Auto subtitle sync option (ffsubsync after download)
- [x] Failed trailer list at end of trailer download run
- [x] Per-session auto-move toggle (Y/N prompt before processing)
- [x] Combined "Process All (Movies + TV)" menu option
- [x] Release info preservation in quality scoring
- [x] TV show folder name cleanup (strips release info, saves to release-info.txt)

## Completed in v5.1

- [x] Modular architecture with `modules/` folder
- [x] SFTP Sync module (`modules/Sync.psm1`)
- [x] Mirror Backup module (`modules/Mirror.psm1`)
- [x] Example config file (`config/config.example.json`)
- [x] First-run setup wizard
- [x] Library transfer feature (inbox to main collection)
- [x] Bracket path handling fix (using `-LiteralPath`)
- [x] File renaming to remove brackets from filenames
- [x] Improved bitrate warnings (better encode detection)
