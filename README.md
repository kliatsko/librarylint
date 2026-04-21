<p align="center">
  <img src="assets/logo.png" alt="LibraryLint" width="200">
</p>

<h1 align="center">LibraryLint</h1>

<p align="center">
A modular toolkit for media library organization, cleanup, and management.<br>
Designed for managing movie and TV show collections with support for Kodi/Plex/Jellyfin-compatible naming and metadata.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-7+-blue.svg" alt="PowerShell">
  <img src="https://img.shields.io/badge/Windows-10+-green.svg" alt="Windows">
  <img src="https://img.shields.io/badge/Version-5.6.2-orange.svg" alt="Version">
</p>

> **New to LibraryLint?** Check out the [Getting Started Guide](GETTING_STARTED.md) for setup instructions.

## Features

### Core Functionality
- **Dry-run mode** - Preview all changes before applying them
- **Comprehensive logging** - All operations logged with timestamps
- **Progress tracking** - Visual progress indicators with ETA for long operations
- **Auto-updater** - Check for new versions from GitHub
- **First-run setup wizard** - Guided configuration on first launch
- **Automatic dependency installation** - Installs 7-Zip, FFmpeg, yt-dlp via winget

### Movie Processing
- Extract archives (.rar, .zip, .7z, .tar, .gz, .bz2)
- Remove unnecessary files (samples, proofs, screenshots)
- Process trailers (move to `_Trailers` folder or delete)
- Process subtitles (keep preferred language, delete others)
- Create individual folders for loose video files
- Clean folder names by removing quality/codec/release tags
- Format movie years with parentheses (`Movie 2024` → `Movie (2024)`)
- Generate Kodi-compatible NFO files with TMDB metadata
- Download artwork (poster, fanart, clearlogo, clearart)
- Auto-move processed movies from inbox to main library

### TV Show Processing
- Extract all archives
- Parse episode info (S01E01, 1x01, multi-episode S01E01-E03)
- Organize episodes into Season folders
- Rename episodes to standard format
- Detect missing episodes (gap detection)
- Remove empty folders
- Generate NFO files with TVDB metadata (tvshow.nfo + episode NFOs)
- Download artwork (poster, fanart, season posters, actor images)
- Auto-move processed shows from inbox to main library (with season merging)

### Advanced Features
- **Duplicate Detection** - Find duplicates using file hashing and quality scoring
- **TMDB Integration** - Fetch movie metadata from The Movie Database
- **TVDB Integration** - Fetch TV show metadata from TheTVDB
- **Fanart.tv Integration** - Extended artwork (clearlogo, banner, clearart, extrafanart)
- **Codec Analysis** - Analyze video codecs and generate FFmpeg transcode scripts
- **Health Check** - Validate library for issues (empty folders, missing files, etc.)
- **MediaInfo Integration** - Accurate codec detection from file headers
- **Export Reports** - Generate CSV, HTML, and JSON library reports
- **Undo/Rollback** - Manifest-based rollback of changes
- **Configuration Files** - Save/load settings to JSON
- **Radarr Integration** - Re-acquisition utility, import verification for SFTP pruning
- **TMDB ID Deduplication** - Detect duplicates by metadata, not just folder name
- **NFO-only Refresh** - Regenerate NFOs without re-downloading artwork/trailers

### Sync & Backup Modules
- **SFTP Sync** - Download new files from seedbox/remote server (requires WinSCP)
- **Mirror Backup** - Robocopy-based mirroring to external drives with ETA
- **Integrated Workflow** - Sync → Process → Transfer → Mirror

## Requirements

- Windows 10 or later
- PowerShell 7 or later
- 7-Zip (automatically installed if not present)

### Optional Dependencies
| Tool | Purpose | Install |
|------|---------|---------|
| [MediaInfo](https://mediaarea.net/en/MediaInfo) | Accurate codec detection | `winget install MediaArea.MediaInfo` |
| [FFmpeg](https://ffmpeg.org/) | Video transcoding | `winget install ffmpeg` |
| [yt-dlp](https://github.com/yt-dlp/yt-dlp) | Trailer downloads | `winget install yt-dlp` |
| [WinSCP](https://winscp.net/) | SFTP sync module | `winget install WinSCP` |

### Optional API Keys (Free)
| Service | Purpose | Get Key |
|---------|---------|---------|
| TMDB | Movie metadata & posters | [themoviedb.org](https://www.themoviedb.org/settings/api) |
| TVDB | TV show metadata & artwork | [thetvdb.com](https://thetvdb.com/api-information) |
| Fanart.tv | Clearlogos, banners, clearart | [fanart.tv](https://fanart.tv/get-an-api-key/) |
| Subdl | Subtitle downloads | [subdl.com](https://subdl.com/panel/api) |

## Installation

### Option 1: Double-click launcher
1. Download or clone the repository
2. Double-click `Run-LibraryLint.bat`

### Option 2: PowerShell
```powershell
git clone https://github.com/kliatsko/librarylint.git
cd librarylint
.\LibraryLint.ps1
```

### Option 3: First-time setup
Run with `-Setup` to launch the configuration wizard:
```powershell
.\LibraryLint.ps1 -Setup
```

## Usage

### Interactive Mode
Simply run the script and follow the prompts:
```powershell
.\LibraryLint.ps1
```

### With Verbose Output
```powershell
.\LibraryLint.ps1 -Verbose
```

### With Custom Config File
```powershell
.\LibraryLint.ps1 -ConfigFile "C:\path\to\config.json"
```

### Check for Updates
```powershell
.\LibraryLint.ps1 -Update
```

## Main Menu Options

| Option | Description |
|--------|-------------|
| **New Content** ||
| 1 | **Process New Movies** - Full movie cleanup, metadata, and organization |
| 2 | **Process New TV Shows** - Organize episodes into season folders |
| 3 | **Process All** - Run both movies and TV shows in sequence |
| **Library Maintenance** ||
| 4 | **Fix & Repair** - Fix folder names, refresh metadata, find duplicates |
| 5 | **Enhancements** - Download trailers, subtitles, artwork |
| 6 | **Utilities** - SFTP sync, mirror backup, export reports, undo |
| **Other** ||
| 7 | **Settings** - Configure paths, API keys, preferences |
| ? | **Help** - Interactive help menu |
| 0 | **Exit** |

## Supported Formats

### Video
`.mp4`, `.mkv`, `.avi`, `.mov`, `.wmv`, `.flv`, `.m4v`

### Subtitles
`.srt`, `.sub`, `.idx`, `.ass`, `.ssa`, `.vtt`

### Archives
`.rar`, `.zip`, `.7z`, `.tar`, `.gz`, `.bz2`

## Configuration

Settings are stored in `%LOCALAPPDATA%\LibraryLint\LibraryLint.config.json`

Key configuration options:
- `DryRun` - Preview mode (no changes made)
- `KeepSubtitles` - Keep subtitle files
- `KeepTrailers` - Move trailers to `_Trailers` folder
- `PreferredSubtitleLanguages` - Languages to keep (default: English)
- `GenerateNFO` - Auto-generate Kodi NFO files
- `TMDBApiKey` - Your TMDB API key
- `DownloadTrailers` - Download movie trailers from YouTube
- `TrailerQuality` - Trailer quality: 1080p, 720p, or 480p
- `DownloadSubtitles` - Download subtitles from Subdl.com
- `SubtitleLanguage` - Subtitle language code (default: en)
- `RetryCount` - Number of retries for failed operations
- `EnableUndo` - Enable undo manifest creation

## Trailer Downloads (Optional)

LibraryLint can automatically download movie trailers from YouTube and save them in Kodi-compatible format.

### Setup

1. **Install yt-dlp** (choose one method):
   ```powershell
   # Using winget (recommended)
   winget install yt-dlp

   # Using pip
   pip install yt-dlp

   # Manual download
   # Download yt-dlp.exe from https://github.com/yt-dlp/yt-dlp/releases
   # Place it in your PATH or C:\Program Files\yt-dlp\
   ```

2. **Run LibraryLint** - If yt-dlp is detected, you'll be prompted:
   ```
   Download movie trailers from YouTube? (Y/N) [N]
   ```

3. **Trailers are saved as** `MovieTitle-trailer.mp4` in each movie folder

### Configuration

| Option | Default | Description |
|--------|---------|-------------|
| `DownloadTrailers` | `false` | Enable/disable trailer downloads |
| `TrailerQuality` | `720p` | Video quality: 1080p, 720p, or 480p |

### Storage Estimates

| Quality | Per Trailer | 100 Movies | 500 Movies |
|---------|-------------|------------|------------|
| 1080p | 50-150 MB | 5-15 GB | 25-75 GB |
| 720p | 20-60 MB | 2-6 GB | 10-30 GB |
| 480p | 10-30 MB | 1-3 GB | 5-15 GB |

## Subtitle Downloads (Optional)

LibraryLint can automatically download subtitles from [Subdl.com](https://subdl.com) and save them in Kodi-compatible format.

### How It Works

1. **Uses IMDB ID** - Searches by IMDB ID first (from TMDB metadata) for accurate matches
2. **Fallback to title** - If no IMDB match, searches by movie title and year
3. **Skips existing** - Won't download if subtitle already exists in the folder
4. **Rate limited** - Respects Subdl.com's 1 request/second rate limit

### Setup

1. **Get a free API key** at [https://subdl.com/panel/api](https://subdl.com/panel/api)
   - Create a free account on Subdl.com
   - Go to your panel and generate an API key

2. **Run LibraryLint** - You'll be prompted for your API key the first time

3. **Subtitles are saved as** `MovieTitle.en.srt` in each movie folder

### Configuration

| Option | Default | Description |
|--------|---------|-------------|
| `DownloadSubtitles` | `false` | Enable/disable subtitle downloads |
| `SubtitleLanguage` | `en` | Language code (en, es, fr, de, it, pt, etc.) |
| `SubdlApiKey` | `null` | Your Subdl.com API key |

### Supported Languages

Common language codes: `en` (English), `es` (Spanish), `fr` (French), `de` (German), `it` (Italian), `pt` (Portuguese), `ru` (Russian), `ja` (Japanese), `ko` (Korean), `zh` (Chinese)

### Storage

Subtitles are very small - typically 50-150 KB each. A library of 500 movies would only use ~25-75 MB for subtitles.

## File Locations

| File Type | Location |
|-----------|----------|
| Logs | `%LOCALAPPDATA%\LibraryLint\Logs\` |
| Config | `%LOCALAPPDATA%\LibraryLint\LibraryLint.config.json` |
| Undo Manifests | `%LOCALAPPDATA%\LibraryLint\Undo\` |

## Modules

LibraryLint includes optional modules for sync and backup workflows. These are loaded automatically if present in the `modules/` folder.

### SFTP Sync (`modules/Sync.psm1`)

Downloads new files from a remote SFTP server (seedbox, NAS, etc.) with tracking to avoid re-downloading.

**Features:**
- Recursive scanning of remote directories
- Automatic file categorization (Movies vs TV Shows based on size and naming)
- Download tracking to skip already-synced files
- Optional deletion of files after download
- Separate sync and prune paths (download from one location, clean up another)
- Radarr-verified pruning (only delete files that Radarr has imported)
- Dry-run and list-only modes

**Requirements:** [WinSCP](https://winscp.net/) with .NET assembly (`winget install WinSCP`)

**Configuration:**
```json
{
  "SFTPHost": "your-server.com",
  "SFTPPort": 22,
  "SFTPUsername": "username",
  "SFTPPassword": "password",
  "SFTPRemotePaths": ["/downloads"],
  "SFTPPrunePaths": [],
  "SFTPDeleteAfterDownload": false
}
```

### Mirror Backup (`modules/Mirror.psm1`)

Mirrors media folders to a backup drive using robocopy with `/MIR` flag for exact synchronization.

**Features:**
- Multi-threaded copying (8 threads by default)
- Progress tracking with file counts and ETA
- Detailed summary of copied/skipped/deleted files
- Dry-run mode for preview

**Configuration:**
```json
{
  "MirrorSourceDrive": "G:",
  "MirrorDestDrive": "F:",
  "MirrorFolders": ["Movies", "Shows"]
}
```

### Quality Analysis (`modules/Quality.psm1`)

Scores video files by resolution, codec, source, and audio quality. Identifies files needing transcoding.

**Features:**
- Quality scoring (resolution, codec, source, audio, HDR)
- Codec analysis with centralized caching
- FFmpeg transcode script generation
- Force rescan option to bypass cache

### Subtitles (`modules/Subtitles.psm1`)

Manages subtitle files — detection, downloading, renaming, and timing correction.

**Features:**
- Language detection and filtering
- Subtitle downloading via Subdl API
- Timing correction with ffsubsync
- Orphaned subtitle cleanup

### TMDB/TVDB (`modules/TMDB.psm1`)

API client for metadata lookups from The Movie Database and TheTVDB.

**Features:**
- TMDB movie search with Jaccard similarity scoring
- TVDB show/episode lookup with cached tokens
- Year tolerance and title normalization
- Collection and set information

## Quality Scoring

When detecting duplicates, files are scored based on:
- **Resolution**: 2160p (100) > 1080p (80) > 720p (60) > 480p (40)
- **Source**: BluRay (50) > WEB-DL (40) > WEBRip (35) > HDRip (30)
- **Codec**: x265/HEVC (30) > x264 (20) > XviD (10)
- **Audio**: Atmos (25) > DTS-HD (20) > TrueHD (20) > DTS (15) > AC3 (10)
- **HDR**: +20 bonus points

## Screenshots

### Processing Summary
```
╔══════════════════════════════════════════════════════════════════╗
║                      PROCESSING SUMMARY                         ║
╠══════════════════════════════════════════════════════════════════╣
║  Duration:              00:02:34                                ║
║  Files Deleted:         47                                      ║
║  Space Reclaimed:       2.3 GB                                  ║
║  Archives Extracted:    12                                      ║
║  Folders Created:       8                                       ║
║  Folders Renamed:       23                                      ║
╚══════════════════════════════════════════════════════════════════╝
```

## Contributing

Thanks for your interest in contributing!

LibraryLint is a personal project that I maintain in my spare time, so please keep the following in mind.

Setting Expectations
This is a passion project, not a full-time endeavor. Response times on issues and pull requests will vary — sometimes I'll get back to you in a day, sometimes it might take a few weeks. I appreciate your patience.

I'm also selective about scope. LibraryLint is intentionally a lightweight, run-it-when-you-want-it toolkit. Feature requests that would push it toward becoming a background service or duplicating what Radarr/Sonarr already do well will likely be declined. That's not a reflection on the quality of the idea — it's about keeping the project focused and maintainable.

How to Contribute

Reporting Bugs

Before opening an issue, please:
Make sure you're running the latest version (.\LibraryLint.ps1 -Update)

Check existing issues to avoid duplicates
Include your PowerShell version ($PSVersionTable.PSVersion), Windows version, and the relevant log output from %LOCALAPPDATA%\LibraryLint\Logs\

If possible, describe the steps to reproduce the issue

Suggesting Features

Open an issue with the "enhancement" label. A good feature request includes:

A clear description of the problem you're trying to solve
Why existing functionality doesn't cover it
Whether you'd be willing to implement it yourself

Submitting Pull Requests
Fork the repo and create a branch from main
Test your changes with -DryRun mode before submitting
Keep PRs focused — one feature or fix per PR
Update the README if your change adds or modifies user-facing functionality
Include a brief description of what the PR does and why

Code Style
Follow existing PowerShell conventions used throughout the project
Use meaningful variable and function names
Add comments for non-obvious logic
Test with both -DryRun and live runs before submitting

What I'm Most Interested In
Bug fixes and edge case handling
Improvements to parsing logic (folder names, episode patterns)
Better error handling and user-facing messages
Documentation improvements
What's Probably Out of Scope
Cross-platform support (this is a Windows/PowerShell tool by design)
Background service or daemon mode
Integration with download clients (that's Radarr/Sonarr territory)
GUI or web interface

License
By contributing, you agree that your contributions will be licensed under the MIT License.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

**Nick Kliatsko**

## Acknowledgments

- [7-Zip](https://www.7-zip.org/) for archive extraction
- [MediaInfo](https://mediaarea.net/en/MediaInfo) for codec detection
- [The Movie Database (TMDB)](https://www.themoviedb.org/) for movie metadata
- [TheTVDB](https://thetvdb.com/) for TV show metadata
- [Fanart.tv](https://fanart.tv/) for extended artwork
- [Subdl.com](https://subdl.com) for subtitle downloads
- [WinSCP](https://winscp.net/) for SFTP transfers
- [yt-dlp](https://github.com/yt-dlp/yt-dlp) for trailer downloads
