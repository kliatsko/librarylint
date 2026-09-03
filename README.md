<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="assets/logo-dark.png">
    <source media="(prefers-color-scheme: light)" srcset="assets/logo.png">
    <img src="assets/logo.png" alt="LibraryLint" width="200">
  </picture>
</p>

<h1 align="center">LibraryLint</h1>

<p align="center">
A PowerShell toolkit for curating a Kodi-style media library<br>
when downloads happen on a remote seedbox and you process them manually.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-7+-blue.svg" alt="PowerShell">
  <img src="https://img.shields.io/badge/Windows-10+-green.svg" alt="Windows">
  <img src="https://img.shields.io/badge/Version-5.8.0-orange.svg" alt="Version">
</p>

> **New to LibraryLint?** Check out the [Getting Started Guide](GETTING_STARTED.md) for setup instructions.

## Who is this for?

LibraryLint solves a specific problem: **curating a Kodi-style local media library from a remote seedbox where the *arr stack already handles imports**. The pull-down, the Kodi-side enrichment, and the seedbox cleanup are LibraryLint's territory; *arr is left to do what it does well (search, import to the seedbox library mirror).

### The pipeline

```
 ┌──────────────────┐         ┌──────────────────┐         ┌──────────────────┐
 │   Seedbox        │  SFTP   │   Local PC       │ robocopy│   HTPC / NAS     │
 │  rTorrent +      │ ──────▶ │  LibraryLint:    │ ──────▶ │  Kodi / Emby /   │
 │  Radarr/Sonarr   │         │  NFO, artwork,   │         │  Jellyfin        │
 │  + CDH renames   │         │  subs, dedup,    │         │  (reads NFOs)    │
 │  to media/       │         │  prune the seed  │         │                  │
 └──────────────────┘         └──────────────────┘         └──────────────────┘
   downloads + auto-import     Kodi-ready local copy        watch from here
   hit-and-run aware           local library is canonical   mirrored over LAN
```

### Fit

- You run a seedbox for downloads (Ultra.cc, Whatbox, Feralhosting, your own VPS)
- You use Radarr/Sonarr (typically on the seedbox) for search and import — CDH builds a canonical library mirror at e.g. `media/Movies/` and `media/TV Shows/`
- You watch on a **Kodi-based front-end** (LibreELEC, CoreELEC, OSMC) or Emby/Jellyfin in NFO-library mode
- You want a *local* curated copy with per-folder NFOs, fetched artwork, and consistent naming — not just the seedbox copy
- You want a deterministic manual processing step locally rather than a continuous background daemon
- You care about hit-and-run windows on private trackers (working-dir prune respects an N-day age guard before touching anything still seeding)

### Not a fit

- **Plex Media Server users.** Plex scans by metadata agents and largely ignores NFOs. *arr CDH already handles imports, and Plex's own scraper handles metadata — most of LibraryLint's enrichment is wasted effort for this audience.
- **All-local *arr stack (no remote downloader).** Most of LibraryLint's value is in pulling content from a remote seedbox and pruning the seedbox afterwards. If your *arr stack runs on the same host as your library, CDH plus your front-end's own scraper covers most of the same needs.
- **Transmission or qBittorrent without a working/complete split.** LibraryLint's three-tier prune model (working dir → complete folder → library mirror) is built around rTorrent's move-on-completion semantics.
- **Daemon-style expectations.** LibraryLint is invoked manually or via a scheduled task; there's no service mode or always-on processing.

### What's swappable

- **Mirror destination** — any path robocopy can reach (UNC share, local drive, external drive, mapped network drive)
- **Final consumer** — anything that reads NFOs at the folder level (Kodi is the primary target; Emby and Jellyfin work when configured to use NFOs as the metadata source)

Other components (download client semantics, local library naming) are currently tied to the assumed workflow.

## Features

### Core Functionality
- **Pipeline workflow (`S`)** - One command runs the whole chain: seedbox scan, free-space checks, inbox processing, subtitle census + daily subtitle queue, library health (Radarr/Sonarr), HTPC wake, mirror, HTPC shutdown. Enter at the main menu defaults to it; `-Status` runs it from the CLI
- **Dry-run mode** - Preview all changes before applying them
- **Comprehensive logging** - All operations logged with timestamps
- **Progress tracking** - Visual progress indicators with ETA for long operations
- **Auto-updater** - Check for new versions from GitHub
- **First-run setup wizard** - Guided configuration on first launch
- **Automatic dependency installation** - Installs 7-Zip, FFmpeg, yt-dlp, Tesseract via winget; ffsubsync and faster-whisper via pip

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

### Subtitle Pipeline
Every movie ends up with a verified, in-sync English soft subtitle — via the cheapest tier that works:
- **Census** - MediaInfo scan bucketing the library: has subs / embedded text tracks / missing
- **Embedded extraction** - English text tracks pulled out with FFmpeg (zero sync risk)
- **OpenSubtitles hash matching** - OSDb moviehash finds subs for the byte-identical release; synced by construction. A daily quota queue in the `S` workflow works through the backlog automatically
- **Verification gate** - ffsubsync measures every unverified sub against the audio: in-sync gets marked, drift gets corrected, wrong-cut gets flagged for replacement
- **Whisper generation** - faster-whisper transcribes/translates the audio itself on GPU (CUDA with CPU fallback) for movies no provider matches
- **Provenance markers** - each movie folder's `.subs_ok` records where its subs came from (opensubtitles / whisper / embedded / hardsub / manual...) with an append-only history
- **Hardsub audit** - OCR frame sampling finds burned-in subtitles, with a visual HTML report (evidence frame pairs embedded) and one-key marking so hardsubbed movies leave the acquisition pool

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
- **Sonarr Integration** - Missing-episode counts in Status, episode upgrade re-acquisition
- **HTPC Control** - Wake-on-LAN, Kodi JSON-RPC shutdown/keep-alive; the `S` workflow wakes the HTPC for mirroring and puts it back to sleep after
- **TMDB ID Deduplication** - Detect duplicates by metadata, not just folder name
- **NFO-only Refresh** - Regenerate NFOs without re-downloading artwork/trailers

### Sync & Backup Modules
- **SFTP Sync** - Download new files from seedbox/remote server (requires WinSCP), with per-user quota display and remote RAR extraction
- **Seedbox Prune** - Hit-and-run-aware cleanup of the seedbox once content is confirmed local, including rTorrent dead-torrent erasure over SSH
- **Mirror Backup** - Robocopy-based mirroring with ETA, a dead-destination watchdog, Kodi keep-alive during long copies, and timestamp repair on cancel
- **Integrated Workflow** - the `S` pipeline chains sync → process → subtitles → transfer → mirror

## Requirements

- Windows 10 or later
- PowerShell 7 or later
- 7-Zip (automatically installed if not present)

### Optional Dependencies
| Tool | Purpose | Install |
|------|---------|---------|
| [MediaInfo](https://mediaarea.net/en/MediaInfo) | Accurate codec detection, subtitle census | `winget install MediaArea.MediaInfo` |
| [FFmpeg](https://ffmpeg.org/) | Video transcoding, subtitle extraction, hardsub audit | `winget install ffmpeg` |
| [yt-dlp](https://github.com/yt-dlp/yt-dlp) | Trailer downloads | `winget install yt-dlp` |
| [WinSCP](https://winscp.net/) | SFTP sync module | `winget install WinSCP` |
| [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) | Hardsub audit (burned-in subtitle detection) | `winget install UB-Mannheim.TesseractOCR` |
| [ffsubsync](https://github.com/smacke/ffsubsync) | Subtitle sync verification & correction | auto-installed via pip (needs Python 3.11) |
| [faster-whisper](https://github.com/SYSTRAN/faster-whisper) | Whisper subtitle generation (GPU-accelerated) | auto-installed via pip (needs Python) |

### Optional API Keys (Free)
| Service | Purpose | Get Key |
|---------|---------|---------|
| TMDB | Movie metadata & posters | [themoviedb.org](https://www.themoviedb.org/settings/api) |
| TVDB | TV show metadata & artwork | [thetvdb.com](https://thetvdb.com/api-information) |
| Fanart.tv | Clearlogos, banners, clearart | [fanart.tv](https://fanart.tv/get-an-api-key/) |
| OpenSubtitles | Hash-matched subtitle downloads (primary subtitle source) | [opensubtitles.com](https://www.opensubtitles.com/en/consumers) |
| Subdl | Name-matched subtitle fallback (opt-in) | [subdl.com](https://subdl.com/panel/api) |

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

### Run the Pipeline Directly (no menu)
```powershell
.\LibraryLint.ps1 -Status
```

### Check for Updates
```powershell
.\LibraryLint.ps1 -Update
```

## Main Menu Options

| Option | Description |
|--------|-------------|
| **Pipeline** ||
| S | **Run Pipeline (Status)** - Snapshot + guided punch list: seedbox → inbox → library → mirror. Enter defaults to this |
| **Maintenance** ||
| 1 | **Process Inbox** - Auto-detect and organize new downloads (also part of S) |
| 2 | **Library Maintenance** - Health, metadata, artwork, subtitles, quality, cleanup |
| 3 | **Utilities & Recovery** - Seedbox tools, mirror, undo, quarantine, exports, HTPC |
| **Other** ||
| 4 | **Settings** - Configuration, API keys, updates |
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
- `OpenSubtitlesApiKey` / `OpenSubtitlesUsername` / `OpenSubtitlesPassword` - Hash-matched subtitle downloads (see Subtitle Management)
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

## Subtitle Management

LibraryLint's goal is a verified, in-sync English soft subtitle on every movie. The Subtitle Health Check (Library Maintenance) runs a tiered pipeline, cheapest tier first:

1. **Census** - MediaInfo scans every movie folder and buckets it: has external subs / has embedded English text tracks / missing. New arrivals are censused in the same `S` run that imports them.
2. **Embedded extraction** - English SRT/ASS tracks are extracted with FFmpeg. Zero sync risk: the track was authored for that exact file.
3. **Hash-matched acquisition** - OpenSubtitles is searched by OSDb moviehash (filesize + first/last 64KB), which only matches subs uploaded for the byte-identical release — synced by construction, auto-verified. The `S` workflow spends up to `OpenSubtitlesDailyLimit` downloads per run against a persistent queue, so the backlog drains automatically day by day. Subdl name-search remains as an explicit opt-in fallback whose downloads must pass the verification gate.
4. **Verification gate** - ffsubsync measures each unverified external sub against the audio. In sync → marked verified; small constant drift → corrected (original backed up); gross offset → flagged wrong-release for replacement, never "fixed" into garbage.
5. **Whisper generation** - for the tail no provider matches (uncommon encodes, obscure titles), faster-whisper transcribes the audio itself on the GPU (`large-v3`, CUDA with CPU fallback, translate mode for foreign audio). Output timestamps derive from the file's own audio, so the result is synced by construction.

Progress is recorded per movie folder in a `.subs_ok` marker with a machine-readable `Provider` field (`opensubtitles`, `whisper`, `embedded`, `subdl`, `release`, `manual`, `hardsub`) and an append-only history, so you can always answer "where did this sub come from?"

### Hardsub Audit

A separate audit samples frames across each movie (skipping intros/credits), OCRs the subtitle region with Tesseract, and uses a temporal check (subtitles change ~2.5s later; scene text persists) to find burned-in subtitles. Results come with a self-contained HTML report showing the actual flagged frame pairs, so verdicts can be confirmed at a glance. Confirmed hardsubs can be marked in the movie's `.subs_ok` (`Provider: hardsub`) — English is already on screen, so they leave the acquisition and Whisper pools.

### Configuration

| Option | Default | Description |
|--------|---------|-------------|
| `OpenSubtitlesApiKey` / `Username` / `Password` | `null` | OpenSubtitles account (all three needed for downloads) |
| `OpenSubtitlesDailyLimit` | `5` | Hash-matched downloads per `S` run (free tier ~5/day; raise if VIP) |
| `SubdlApiKey` | `null` | Subdl fallback API key (opt-in per run) |
| `PreferredSubtitleLanguages` | `eng, en, english` | Languages kept by cleanup passes |

## File Locations

| File Type | Location |
|-----------|----------|
| Logs | `%LOCALAPPDATA%\LibraryLint\Logs\` |
| Config | `%LOCALAPPDATA%\LibraryLint\LibraryLint.config.json` |
| Undo Manifests | `%LOCALAPPDATA%\LibraryLint\Undo\` |
| Reports (CSV / hardsub HTML) | `%LOCALAPPDATA%\LibraryLint\Reports\` |
| Subtitle census & queue | `%LOCALAPPDATA%\LibraryLint\subtitle_census.json`, `subtitle_acquire_queue.json` |

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
- Progress tracking with file counts and ETA (NIC-sampled transfer rate)
- Dead-destination watchdog: detects a mid-copy HTPC/NAS death via throughput stall + TCP probe instead of hanging
- Timestamp repair on cancel, so an interrupted run doesn't re-copy everything next time
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
- Hardsub audit: OCR frame sampling with temporal sub-vs-scene-text discrimination and a visual HTML evidence report
- "Best available" markers (`.quality_ok`) to exempt movies with no better release from quality flagging
- Force rescan option to bypass cache

### Subtitles (`modules/Subtitles.psm1`)

The subtitle pipeline: census, extraction, acquisition, verification, generation.

**Features:**
- Embedded-track census and English SRT/ASS extraction (MediaInfo + FFmpeg)
- OpenSubtitles OSDb hash matching with a quota-aware daily download queue
- Subdl name-search fallback (opt-in)
- Sync verification and drift correction with ffsubsync (with backup/restore)
- Whisper subtitle generation via faster-whisper (CUDA, CPU fallback)
- Provenance tracking in `.subs_ok` markers (provider + history)
- Language filtering, placement fixes, orphaned subtitle cleanup

### HTPC Control (`modules/Htpc.psm1`)

Wakes and sleeps the playback box around mirror runs.

**Features:**
- Wake-on-LAN magic packets with TCP readiness polling
- Kodi JSON-RPC shutdown / suspend / reboot with validated stored credentials
- Idle-shutdown keep-alive during long mirror copies (restored afterwards)

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
- [OpenSubtitles](https://www.opensubtitles.com/) for hash-matched subtitle downloads
- [Subdl.com](https://subdl.com) for subtitle downloads
- [ffsubsync](https://github.com/smacke/ffsubsync) for subtitle sync measurement and correction
- [faster-whisper](https://github.com/SYSTRAN/faster-whisper) for GPU subtitle generation
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) for hardsub detection
- [WinSCP](https://winscp.net/) for SFTP transfers
- [yt-dlp](https://github.com/yt-dlp/yt-dlp) for trailer downloads
