# Getting Started with LibraryLint

LibraryLint is a PowerShell toolkit for organizing and managing your media library. This guide will help you get set up.

## Quick Start

1. **Clone or download** the repository
2. **Run the script**: `.\LibraryLint.ps1`
3. **First run** will prompt you to configure paths and install dependencies

## System Requirements

- Windows 10/11
- PowerShell 5.1 or later
- ~500MB disk space for dependencies

## Dependencies

LibraryLint uses several external tools. Most are optional but recommended.

### Required
| Tool | Purpose | Install Command |
|------|---------|-----------------|
| 7-Zip | Extract archives | Auto-installed on first run |

### Recommended
| Tool | Purpose | Install Command |
|------|---------|-----------------|
| MediaInfo | Analyze video files | `winget install MediaArea.MediaInfo` |
| FFmpeg | Transcode videos | `winget install ffmpeg` |

### Optional (for specific features)
| Tool | Purpose | Install Command |
|------|---------|-----------------|
| yt-dlp | Download trailers | `winget install yt-dlp` |
| WinSCP | SFTP sync from seedbox | `winget install WinSCP` |

## Configuration

On first run, LibraryLint creates a config file at:
```
%LOCALAPPDATA%\LibraryLint\LibraryLint.config.json
```

### Key Settings

**Paths** (set these first):
```json
{
  "MoviesInboxPath": "D:\\Downloads\\Movies",
  "MoviesLibraryPath": "D:\\Media\\Movies",
  "TVShowsInboxPath": "D:\\Downloads\\Shows",
  "TVShowsLibraryPath": "D:\\Media\\Shows"
}
```

**Mirror/Backup** (optional):
```json
{
  "MirrorSourceDrive": "D:",
  "MirrorDestDrive": "E:",
  "MirrorFolders": ["Movies", "Shows"]
}
```

**SFTP Sync** (optional - for seedbox users):
```json
{
  "SFTPHost": "your-seedbox.com",
  "SFTPPort": 22,
  "SFTPUsername": "your-username",
  "SFTPPassword": "your-password",
  "SFTPRemotePaths": ["/downloads/complete"]
}
```

### API Keys (Free)

For metadata, artwork, and subtitles, you'll want these free API keys:

1. **TMDB** (movie metadata & posters)
   - Sign up at https://www.themoviedb.org/signup
   - Go to Settings > API > Request API Key
   - Add to config: `"TMDBApiKey": "your-key-here"`

2. **Fanart.tv** (clearlogos, banners)
   - Sign up at https://fanart.tv/
   - Get your API key from your profile
   - Add to config: `"FanartTVApiKey": "your-key-here"`

3. **Subdl** (subtitles) - *optional*
   - Sign up at https://subdl.com/
   - Get API key from your account
   - Add to config: `"SubdlApiKey": "your-key-here"`

## Folder Structure

LibraryLint expects this general structure:

```
D:\
├── Downloads\          (or "Inbox")
│   ├── Movies\        <- Put downloaded movies here
│   └── Shows\         <- Put downloaded shows here
│
└── Media\             (or your library root)
    ├── Movies\        <- Organized movies go here
    │   └── Movie Name (2024)\
    │       ├── Movie Name (2024).mkv
    │       ├── Movie Name (2024).nfo
    │       ├── poster.jpg
    │       └── fanart.jpg
    │
    └── Shows\         <- Organized shows go here
        └── Show Name\
            └── Season 01\
                └── Show Name - S01E01.mkv
```

## Common Workflows

### Process New Movies
1. Download movies to your inbox folder
2. Run LibraryLint > Option 1 "Process New Movies"
3. Script will:
   - Extract archives
   - Remove junk files (samples, proofs)
   - Clean folder names
   - Fetch metadata from TMDB
   - Download posters/artwork
   - Generate NFO files
4. Optionally transfer to main library

### Mirror to Backup Drive
1. Connect your backup drive
2. Run LibraryLint > Option M "Mirror to Backup"
3. Uses robocopy to sync Movies/Shows folders

### SFTP Sync (Seedbox)
1. Configure SFTP settings in config
2. Run LibraryLint > Option S "SFTP Sync"
3. Downloads new files from seedbox
4. Optionally auto-processes downloaded files

## Tips

- **Dry Run**: Most operations offer a dry-run mode to preview changes
- **Undo**: Enable `"EnableUndo": true` to allow undoing recent changes
- **Logs**: Check `%LOCALAPPDATA%\LibraryLint\Logs\` for detailed logs

## Troubleshooting

### "7-Zip not found"
The script will attempt to auto-install. If it fails:
```powershell
winget install 7zip
```

### "MediaInfo not found"
Quality scoring will use filename parsing instead. For better results:
```powershell
winget install MediaArea.MediaInfo
```

### "TMDB API key not configured"
NFO files will be created with basic info only. Get a free key from themoviedb.org.

### Paths with special characters
If your paths have brackets `[]`, the script handles them with `-LiteralPath`.

## Getting Help

- Check the README.md for full documentation
- View logs in `%LOCALAPPDATA%\LibraryLint\Logs\`
- Report issues at the GitHub repository
