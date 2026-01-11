# TODO

## Branding

- [ ] Design logo for LibraryLint (magnifying glass + film reel, lint roller concept, etc.)

## Features

### TV Show Improvements
- [ ] Add TVDB API key to config
- [ ] Create TVDB API validation function
- [ ] Create TV show search/metadata fetch functions
- [ ] Add TV show NFO generation with TVDB data
- [ ] Add TV show artwork downloads (poster, fanart, banner, etc.)
- [ ] Update Process New TV Shows to use TVDB
- [ ] Update Refresh/Repair Metadata to support TV shows
- [ ] Add Plex/Jellyfin/Emby compatibility for TV shows:
  - [ ] Verify folder structure: `Show Name/Season 01/Show Name - S01E01 - Episode Title.ext`
  - [ ] Support `tvshow.nfo` at show level
  - [ ] Support episode-level `.nfo` files
  - [ ] Add `background.jpg` alongside `fanart.jpg` (Plex compatibility)
- [ ] Add "Fix Folder Names" support for TV shows (casing fixes like movies)

### Auto-Move from Inbox to Main Library
- [x] Add `MoviesLibraryPath` and `TVShowsLibraryPath` to config
- [x] Create move function with duplicate checking
- [x] Add confirmation prompt before moving
- [x] Integrate into Process New Movies workflow
- [ ] Option to enable/disable auto-move per session
- [ ] Integrate into Process New TV Shows workflow

### Initial Setup Wizard
- [x] Create first-run setup process that detects missing config
- [x] Prompt user for media library paths
- [x] Prompt for API keys (TMDB, Subdl)
- [x] Save answers to config file so script doesn't ask on every run
- [ ] Prompt for preferred codecs and quality settings
- [ ] Prompt for FFmpeg path (or auto-detect)
- [ ] Add `--setup` or `--reconfigure` flag to re-run setup

### User-Friendly Distribution
- [ ] Create installer package (MSI or self-extracting exe)
- [ ] Bundle FFmpeg or provide guided download
- [ ] Add auto-updater functionality to check for new versions
- [ ] Create simple "double-click to run" experience (no PowerShell knowledge required)
- [ ] Add uninstaller
- [ ] Consider a simple GUI wrapper for non-technical users

### Version Management
- [ ] Add `$Version` variable at top of script as single source of truth
- [ ] Add `--version` / `-Version` flag to display current version
- [ ] Use Semantic Versioning (MAJOR.MINOR.PATCH) format
- [ ] Create CHANGELOG.md to document changes per release
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
