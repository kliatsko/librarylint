#!/usr/bin/env python3
"""
LibraryLint - Cross-platform media library organization and cleanup tool

================================================================================
                              STATUS: EXPERIMENTAL
================================================================================
This Python version is UNMAINTAINED and may be missing features from the primary
PowerShell version. It exists as a proof-of-concept for potential cross-platform
support.

For the full-featured, actively maintained version, use:
    LibraryLint.ps1 (Windows PowerShell)

If you need cross-platform support (Linux/Mac), please open an issue on GitHub
to express interest. Development will resume if there's sufficient demand.
================================================================================

This Python script automates the cleanup and organization of downloaded media files.
It's a cross-platform port of the PowerShell LibraryLint script.

Features (may be incomplete):
    - Archive extraction (rar, zip, 7z, tar, gz, bz2)
    - Unnecessary file removal (samples, proofs, screenshots)
    - Subtitle handling (keep preferred language)
    - Trailer management (move to _Trailers folder)
    - Folder organization and name cleaning
    - TV show episode detection and organization
    - Duplicate detection with quality scoring and file hashing
    - NFO file parsing and generation
    - TMDB metadata integration
    - Health check and codec analysis
    - MediaInfo integration for accurate codec detection
    - FFmpeg integration for transcoding
    - HDR detection (Dolby Vision, HDR10+, HDR10, HLG)
    - Export functions (CSV, HTML, JSON)
    - Retry logic with exponential backoff
    - Undo/rollback support

NOT implemented in Python version:
    - SFTP Sync module
    - Mirror Backup module
    - First-run setup wizard
    - Library transfer (inbox to main collection)
    - Trailer downloads (yt-dlp integration)
    - Subtitle downloads (Subdl.com integration)
    - Fanart.tv artwork downloads

Requirements:
    - Python 3.8+
    - 7-Zip or p7zip installed
    - Optional: requests library for TMDB integration
    - Optional: MediaInfo CLI for accurate codec detection
    - Optional: FFmpeg for transcoding

Author: Nick Kliatsko
Version: 5.0 (experimental, unmaintained)
"""

import os
import re
import sys
import json
import shutil
import logging
import argparse
import subprocess
import hashlib
import csv
import time
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any, Tuple
from xml.etree import ElementTree as ET
from xml.dom import minidom
from html import escape as html_escape

# Optional imports
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False


# =============================================================================
# CONFIGURATION
# =============================================================================

@dataclass
class Config:
    """Global configuration settings"""
    dry_run: bool = False
    log_file: str = ""
    seven_zip_path: str = ""
    mediainfo_path: str = ""
    ffmpeg_path: str = ""

    video_extensions: tuple = ('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.m4v', '.webm')
    subtitle_extensions: tuple = ('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt')
    archive_extensions: tuple = ('.rar', '.zip', '.7z', '.tar', '.gz', '.bz2')

    unnecessary_patterns: tuple = ('*Sample*', '*Proof*', '*Screens*', '*RARBG*', '*WWW*')
    trailer_patterns: tuple = ('*Trailer*', '*trailer*', '*TRAILER*', '*Teaser*', '*teaser*')
    preferred_subtitle_languages: tuple = ('eng', 'en', 'english')

    keep_subtitles: bool = True
    keep_trailers: bool = True
    generate_nfo: bool = False
    organize_seasons: bool = True
    rename_episodes: bool = False
    check_duplicates: bool = False
    tmdb_api_key: str = ""

    # Retry settings
    retry_count: int = 3
    retry_delay_seconds: int = 2

    # Reports folder
    reports_folder: str = ""

    # Quality/release tags to remove from folder names
    tags: tuple = (
        '1080p', '2160p', '720p', '480p', '4K', '2K', 'UHD',
        'HDRip', 'DVDRip', 'BRRip', 'BR-Rip', 'BDRip', 'BD-Rip', 'WEB-DL', 'WEBRip', 'BluRay',
        'x264', 'x265', 'X265', 'H264', 'H265', 'HEVC', 'XviD', 'DivX', 'AVC', 'AV1', 'VP9',
        'AAC', 'AC3', 'DTS', 'Atmos', 'TrueHD', 'DD5.1', '5.1', '7.1', 'EAC3', 'FLAC',
        'HDR', 'HDR10', 'HDR10+', 'DolbyVision', 'DoVi', 'SDR', '10bit', 'HLG',
        'Extended', 'Unrated', 'Remastered', 'REPACK', 'PROPER', 'Remux',
        'YIFY', 'YTS', 'RARBG', 'SPARKS', 'NF', 'AMZN', 'HULU', 'WEB', 'DSNP', 'MAX'
    )

    def __post_init__(self):
        # Set up log file
        if not self.log_file:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.log_file = f"LibraryLint_{timestamp}.log"

        # Find 7-zip
        if not self.seven_zip_path:
            self.seven_zip_path = self._find_executable('7z', [
                r"C:\Program Files\7-Zip\7z.exe",
                r"C:\Program Files (x86)\7-Zip\7z.exe",
                '/usr/bin/7z', '/usr/local/bin/7z', '/opt/homebrew/bin/7z'
            ])

        # Find MediaInfo
        if not self.mediainfo_path:
            self.mediainfo_path = self._find_executable('mediainfo', [
                r"C:\Program Files\MediaInfo\MediaInfo.exe",
                r"C:\Program Files (x86)\MediaInfo\MediaInfo.exe",
                '/usr/bin/mediainfo', '/usr/local/bin/mediainfo', '/opt/homebrew/bin/mediainfo'
            ])

        # Find FFmpeg
        if not self.ffmpeg_path:
            self.ffmpeg_path = self._find_executable('ffmpeg', [
                r"C:\ffmpeg\bin\ffmpeg.exe",
                r"C:\Program Files\ffmpeg\bin\ffmpeg.exe",
                '/usr/bin/ffmpeg', '/usr/local/bin/ffmpeg', '/opt/homebrew/bin/ffmpeg'
            ])

        # Set up reports folder
        if not self.reports_folder:
            self.reports_folder = os.path.join(os.path.dirname(__file__), 'Reports')

    def _find_executable(self, name: str, paths: List[str]) -> str:
        """Find executable in common paths or PATH"""
        for path in paths:
            if os.path.exists(path):
                return path

        # Try to find in PATH
        try:
            cmd = ['where', name] if sys.platform == 'win32' else ['which', name]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                return result.stdout.strip().split('\n')[0]
        except:
            pass

        return ""


@dataclass
class Stats:
    """Statistics tracking"""
    start_time: datetime = field(default_factory=datetime.now)
    end_time: Optional[datetime] = None
    files_deleted: int = 0
    bytes_deleted: int = 0
    archives_extracted: int = 0
    archives_failed: int = 0
    folders_created: int = 0
    folders_renamed: int = 0
    files_moved: int = 0
    empty_folders_removed: int = 0
    subtitles_processed: int = 0
    subtitles_deleted: int = 0
    trailers_moved: int = 0
    nfo_files_created: int = 0
    nfo_files_read: int = 0
    operations_retried: int = 0
    errors: int = 0
    warnings: int = 0


# Global instances
config = Config()
stats = Stats()
logger: Optional[logging.Logger] = None
failed_operations: List[Dict] = []
undo_log: List[Dict] = []


# =============================================================================
# LOGGING
# =============================================================================

def setup_logging(log_file: str) -> logging.Logger:
    """Set up logging to file and console"""
    log = logging.getLogger('LibraryLint')
    log.setLevel(logging.DEBUG)

    # File handler
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
    log.addHandler(fh)

    return log


def log_info(message: str):
    """Log info message"""
    if logger:
        logger.info(message)


def log_warning(message: str):
    """Log warning message"""
    if logger:
        logger.warning(message)
    stats.warnings += 1


def log_error(message: str):
    """Log error message"""
    if logger:
        logger.error(message)
    stats.errors += 1


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def format_file_size(bytes_size: int) -> str:
    """Format bytes to human-readable string"""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if bytes_size < 1024:
            return f"{bytes_size:.2f} {unit}"
        bytes_size /= 1024
    return f"{bytes_size:.2f} PB"


def print_color(message: str, color: str = 'white', end: str = '\n'):
    """Print colored output (cross-platform)"""
    colors = {
        'red': '\033[91m',
        'green': '\033[92m',
        'yellow': '\033[93m',
        'blue': '\033[94m',
        'magenta': '\033[95m',
        'cyan': '\033[96m',
        'white': '\033[97m',
        'gray': '\033[90m',
        'reset': '\033[0m'
    }

    # Enable ANSI on Windows
    if sys.platform == 'win32':
        os.system('')

    print(f"{colors.get(color, '')}{message}{colors['reset']}", end=end)


def matches_pattern(name: str, patterns: tuple) -> bool:
    """Check if name matches any pattern (supports wildcards)"""
    import fnmatch
    for pattern in patterns:
        if fnmatch.fnmatch(name, pattern) or fnmatch.fnmatch(name.lower(), pattern.lower()):
            return True
    return False


def add_failed_operation(op_type: str, params: Dict, error: str):
    """Add a failed operation for retry"""
    failed_operations.append({
        'type': op_type,
        'params': params,
        'error': error,
        'retry_count': 0,
        'timestamp': datetime.now()
    })


def add_undo_entry(action: str, params: Dict):
    """Add an entry to the undo log"""
    undo_log.append({
        'action': action,
        'params': params,
        'timestamp': datetime.now()
    })


def retry_failed_operations():
    """Retry all failed operations with exponential backoff"""
    global failed_operations

    if not failed_operations:
        return

    print_color(f"\nRetrying {len(failed_operations)} failed operation(s)...", 'yellow')
    log_info(f"Starting retry of {len(failed_operations)} failed operations")

    still_failed = []

    for op in failed_operations:
        if op['retry_count'] >= config.retry_count:
            print_color(f"Max retries reached for: {op['type']} - {op['params']}", 'red')
            still_failed.append(op)
            continue

        op['retry_count'] += 1
        delay = config.retry_delay_seconds * (2 ** (op['retry_count'] - 1))
        print_color(f"Retry {op['retry_count']}/{config.retry_count} (waiting {delay}s): {op['type']}", 'cyan')

        time.sleep(delay)

        try:
            if op['type'] == 'Move':
                shutil.move(op['params']['source'], op['params']['destination'])
                print_color(f"Retry successful: Move {op['params']['source']}", 'green')
                stats.operations_retried += 1
            elif op['type'] == 'Delete':
                if os.path.isdir(op['params']['path']):
                    shutil.rmtree(op['params']['path'])
                else:
                    os.remove(op['params']['path'])
                print_color(f"Retry successful: Delete {op['params']['path']}", 'green')
                stats.operations_retried += 1
            elif op['type'] == 'Rename':
                os.rename(op['params']['path'], op['params']['new_path'])
                print_color(f"Retry successful: Rename {op['params']['path']}", 'green')
                stats.operations_retried += 1
        except Exception as e:
            print_color(f"Retry failed: {e}", 'yellow')
            op['error'] = str(e)
            still_failed.append(op)

    failed_operations = still_failed

    if still_failed:
        print_color(f"{len(still_failed)} operation(s) still failed after retries", 'yellow')
        log_warning(f"{len(still_failed)} operations failed after all retries")
    else:
        print_color("All retry operations completed successfully", 'green')


# =============================================================================
# MEDIAINFO INTEGRATION
# =============================================================================

def test_mediainfo_installation() -> bool:
    """Check if MediaInfo CLI is available"""
    if config.mediainfo_path and os.path.exists(config.mediainfo_path):
        return True
    return False


def test_ffmpeg_installation() -> bool:
    """Check if FFmpeg is available"""
    if config.ffmpeg_path and os.path.exists(config.ffmpeg_path):
        return True
    return False


def get_mediainfo_details(file_path: str) -> Optional[Dict[str, Any]]:
    """Get detailed video info using MediaInfo CLI"""
    info = {
        'video_codec': 'Unknown',
        'audio_codec': 'Unknown',
        'resolution': 'Unknown',
        'width': 0,
        'height': 0,
        'duration': 0,
        'bitrate': 0,
        'frame_rate': 0,
        'hdr': False,
        'audio_channels': 0,
        'container': 'Unknown'
    }

    if not test_mediainfo_installation():
        return None

    try:
        # Use JSON output for reliable parsing
        result = subprocess.run(
            [config.mediainfo_path, '--Output=JSON', file_path],
            capture_output=True, text=True
        )

        if result.returncode != 0:
            return None

        data = json.loads(result.stdout)

        if not data.get('media', {}).get('track'):
            return None

        tracks = data['media']['track']

        # Get General track info
        for track in tracks:
            if track.get('@type') == 'General':
                if track.get('Duration'):
                    info['duration'] = int(float(track['Duration']) * 1000)
                if track.get('OverallBitRate'):
                    info['bitrate'] = int(track['OverallBitRate'])
                if track.get('Format'):
                    info['container'] = track['Format']

        # Get Video track info
        for track in tracks:
            if track.get('@type') == 'Video':
                if track.get('Format'):
                    info['video_codec'] = track['Format']
                if track.get('Width'):
                    info['width'] = int(track['Width'])
                if track.get('Height'):
                    info['height'] = int(track['Height'])
                if track.get('FrameRate'):
                    info['frame_rate'] = float(track['FrameRate'])

                # Check for HDR
                hdr_format = track.get('HDR_Format', '')
                color_primaries = track.get('colour_primaries', '')
                if hdr_format or 'BT.2020' in color_primaries:
                    info['hdr'] = True

        # Get Audio track info
        for track in tracks:
            if track.get('@type') == 'Audio':
                if track.get('Format'):
                    info['audio_codec'] = track['Format']
                if track.get('Channels'):
                    info['audio_channels'] = int(track['Channels'])
                break  # Just get first audio track

        # Determine resolution label
        if info['height'] >= 2160:
            info['resolution'] = '2160p'
        elif info['height'] >= 1080:
            info['resolution'] = '1080p'
        elif info['height'] >= 720:
            info['resolution'] = '720p'
        elif info['height'] >= 480:
            info['resolution'] = '480p'
        elif info['height'] > 0:
            info['resolution'] = f"{info['height']}p"

        return info
    except Exception as e:
        log_warning(f"Error getting MediaInfo: {e}")
        return None


def get_mediainfo_hdr_format(file_path: str) -> Optional[Dict[str, Any]]:
    """Get detailed HDR format info using MediaInfo"""
    if not test_mediainfo_installation():
        return None

    try:
        result = {
            'format': None,
            'profile': None,
            'compatibility': None,
            'color_primaries': None,
            'transfer_characteristics': None
        }

        # Get HDR_Format
        proc = subprocess.run(
            [config.mediainfo_path, '--Inform=Video;%HDR_Format%', file_path],
            capture_output=True, text=True
        )
        if proc.stdout.strip():
            result['format'] = proc.stdout.strip()

        # Get transfer characteristics
        proc = subprocess.run(
            [config.mediainfo_path, '--Inform=Video;%transfer_characteristics%', file_path],
            capture_output=True, text=True
        )
        if proc.stdout.strip():
            result['transfer_characteristics'] = proc.stdout.strip()

        # Get color primaries
        proc = subprocess.run(
            [config.mediainfo_path, '--Inform=Video;%colour_primaries%', file_path],
            capture_output=True, text=True
        )
        if proc.stdout.strip():
            result['color_primaries'] = proc.stdout.strip()

        # Normalize HDR format names
        if result['format']:
            if re.search(r'Dolby Vision|DOVI', result['format'], re.IGNORECASE):
                result['format'] = 'Dolby Vision'
            elif re.search(r'HDR10\+|SMPTE ST 2094', result['format'], re.IGNORECASE):
                result['format'] = 'HDR10+'
            elif re.search(r'SMPTE ST 2086|HDR10', result['format'], re.IGNORECASE):
                result['format'] = 'HDR10'
            elif re.search(r'HLG|ARIB STD-B67', result['format'], re.IGNORECASE):
                result['format'] = 'HLG'
        elif result['transfer_characteristics']:
            if 'PQ' in result['transfer_characteristics'] or 'SMPTE ST 2084' in result['transfer_characteristics']:
                result['format'] = 'HDR10'
            elif 'HLG' in result['transfer_characteristics']:
                result['format'] = 'HLG'
        elif result['color_primaries'] and 'BT.2020' in result['color_primaries']:
            result['format'] = 'HDR'

        if result['format']:
            return result
        return None
    except Exception as e:
        log_warning(f"Error getting HDR format: {e}")
        return None


# =============================================================================
# ENHANCED DUPLICATE DETECTION WITH HASHING
# =============================================================================

def get_file_partial_hash(file_path: str, sample_size: int = 1024 * 1024) -> Optional[str]:
    """Calculate a partial hash for duplicate detection"""
    try:
        file_size = os.path.getsize(file_path)

        # For small files, hash the entire file
        if file_size <= sample_size * 2:
            hasher = hashlib.md5()
            with open(file_path, 'rb') as f:
                hasher.update(f.read())
            return hasher.hexdigest()

        # For large files, hash first and last chunks plus size
        hasher = hashlib.md5()
        with open(file_path, 'rb') as f:
            # Read first chunk
            hasher.update(f.read(sample_size))

            # Seek to end and read last chunk
            f.seek(-sample_size, 2)
            hasher.update(f.read(sample_size))

            # Add file size for uniqueness
            hasher.update(str(file_size).encode())

        return hasher.hexdigest()
    except Exception as e:
        log_warning(f"Error calculating hash for {file_path}: {e}")
        return None


def find_duplicate_movies_enhanced(path: str) -> List[List[Dict]]:
    """Enhanced duplicate detection using file hashing"""
    print_color("\nScanning for duplicate movies (enhanced)...", 'yellow')
    log_info(f"Starting enhanced duplicate scan in: {path}")

    try:
        root_path = Path(path)
        movie_folders = [f for f in root_path.iterdir()
                        if f.is_dir() and f.name != '_Trailers']

        if not movie_folders:
            print_color("No movie folders found", 'cyan')
            return []

        title_lookup = {}
        hash_lookup = {}
        total = len(movie_folders)

        for i, folder in enumerate(movie_folders, 1):
            print(f"\rScanning: {i}/{total} - {folder.name[:40]}...".ljust(60), end='', flush=True)

            # Find main video file
            video_files = [f for f in folder.iterdir()
                          if f.is_file() and f.suffix.lower() in config.video_extensions]
            if not video_files:
                continue

            video_file = max(video_files, key=lambda f: f.stat().st_size)

            title_info = get_normalized_title(folder.name)
            quality = get_quality_score(folder.name, str(video_file))

            # Calculate partial hash
            file_hash = get_file_partial_hash(str(video_file))

            entry = {
                'path': str(folder),
                'original_name': folder.name,
                'year': title_info['year'],
                'quality': quality,
                'file_size': video_file.stat().st_size,
                'file_path': str(video_file),
                'file_hash': file_hash,
                'match_type': []
            }

            # Add to title lookup
            key = title_info['normalized_title']
            if title_info['year']:
                key = f"{key}|{title_info['year']}"

            if key not in title_lookup:
                title_lookup[key] = []
            title_lookup[key].append(entry)

            # Add to hash lookup
            if file_hash:
                if file_hash not in hash_lookup:
                    hash_lookup[file_hash] = []
                hash_lookup[file_hash].append(entry)

        print()  # New line after progress

        # Find duplicates
        duplicates = []
        processed_paths = set()

        # First, find exact duplicates by hash
        for hash_val, entries in hash_lookup.items():
            if len(entries) > 1:
                for entry in entries:
                    entry['match_type'].append('ExactHash')
                    processed_paths.add(entry['path'])
                duplicates.append(entries)

        # Then find title-based duplicates
        for key, entries in title_lookup.items():
            unprocessed = [e for e in entries if e['path'] not in processed_paths]
            if len(unprocessed) > 1:
                for entry in unprocessed:
                    entry['match_type'].append('TitleMatch')

                    # Check if file sizes are similar (within 10%)
                    avg_size = sum(e['file_size'] for e in unprocessed) / len(unprocessed)
                    if abs(entry['file_size'] - avg_size) / avg_size < 0.1:
                        entry['match_type'].append('SimilarSize')

                duplicates.append(unprocessed)

        log_info(f"Enhanced duplicate scan found {len(duplicates)} duplicate groups")
        return duplicates
    except Exception as e:
        log_error(f"Error in enhanced duplicate scan: {e}")
        return []


def show_enhanced_duplicate_report(path: str):
    """Display enhanced duplicate report with match type info"""
    duplicates = find_duplicate_movies_enhanced(path)

    if not duplicates:
        print_color("No duplicate movies found!", 'green')
        return

    print_color("\n+" + "=" * 64 + "+", 'yellow')
    print_color("|" + " ENHANCED DUPLICATE REPORT ".center(64) + "|", 'yellow')
    print_color("+" + "=" * 64 + "+", 'yellow')
    print_color(f"|  Found {len(duplicates)} potential duplicate group(s)".ljust(64) + "|", 'yellow')
    print_color("+" + "=" * 64 + "+", 'yellow')

    for group_num, group in enumerate(duplicates, 1):
        match_types = set()
        for entry in group:
            match_types.update(entry.get('match_type', []))

        print_color(f"\n[{group_num}] Duplicates (Match: {', '.join(match_types)}):", 'cyan')

        sorted_group = sorted(group, key=lambda x: x['quality']['score'], reverse=True)

        for i, movie in enumerate(sorted_group):
            size_str = format_file_size(movie['file_size'])
            score = movie['quality']['score']
            hash_str = movie['file_hash'][:8] + '...' if movie['file_hash'] else 'N/A'

            if i == 0:
                print_color("  [KEEP] ", 'green', end='')
            else:
                print_color("  [DEL?] ", 'red', end='')

            print_color(movie['original_name'], 'white')
            print_color(f"         Score: {score} | {size_str} | Hash: {hash_str}", 'gray')
            print_color(f"         {movie['quality']['resolution']} {movie['quality']['source']} {movie['quality']['codec']}", 'gray')


# =============================================================================
# QUALITY SCORING
# =============================================================================

def get_quality_score(filename: str, file_path: str = None) -> Dict[str, Any]:
    """Calculate quality score based on filename and MediaInfo"""
    quality = {
        'score': 0,
        'resolution': 'Unknown',
        'codec': 'Unknown',
        'source': 'Unknown',
        'audio': 'Unknown',
        'hdr': False,
        'hdr_format': None,
        'bitrate': 0,
        'width': 0,
        'height': 0,
        'audio_channels': 0,
        'details': [],
        'data_source': 'Filename'
    }

    # Try MediaInfo first
    media_info = None
    if file_path and os.path.isfile(file_path):
        media_info = get_mediainfo_details(file_path)

    if media_info:
        quality['data_source'] = 'MediaInfo'
        quality['width'] = media_info['width']
        quality['height'] = media_info['height']
        quality['bitrate'] = media_info['bitrate']
        quality['audio_channels'] = media_info['audio_channels']

        # Resolution from actual dimensions
        height = media_info['height']
        if height >= 2160:
            quality['resolution'] = '2160p'
            quality['score'] += 100
            quality['details'].append(f"4K/2160p [MediaInfo: {media_info['width']}x{height}] (+100)")
        elif height >= 1080:
            quality['resolution'] = '1080p'
            quality['score'] += 80
            quality['details'].append(f"1080p [MediaInfo: {media_info['width']}x{height}] (+80)")
        elif height >= 720:
            quality['resolution'] = '720p'
            quality['score'] += 60
            quality['details'].append(f"720p [MediaInfo: {media_info['width']}x{height}] (+60)")
        elif height >= 480:
            quality['resolution'] = '480p'
            quality['score'] += 40
            quality['details'].append(f"480p [MediaInfo: {media_info['width']}x{height}] (+40)")
        elif height > 0:
            quality['resolution'] = f"{height}p"
            quality['score'] += 20
            quality['details'].append(f"{height}p [MediaInfo] (+20)")

        # Video codec from MediaInfo
        video_codec = media_info['video_codec']
        if video_codec:
            codec_map = {
                r'HEVC|H\.?265': ('HEVC/x265', 20),
                r'AVC|H\.?264': ('x264', 15),
                r'AV1': ('AV1', 25),
                r'VP9': ('VP9', 18),
                r'MPEG-4|DivX|XviD': ('XviD', 5),
                r'VC-1|WMV': ('VC-1', 8)
            }
            for pattern, (name, points) in codec_map.items():
                if re.search(pattern, video_codec, re.IGNORECASE):
                    quality['codec'] = name
                    quality['score'] += points
                    quality['details'].append(f"{name} [MediaInfo: {video_codec}] (+{points})")
                    break
            else:
                quality['codec'] = video_codec
                quality['score'] += 10
                quality['details'].append(f"{video_codec} [MediaInfo] (+10)")

        # Audio codec from MediaInfo
        audio_codec = media_info['audio_codec']
        channels = media_info['audio_channels']
        if audio_codec:
            channel_info = f" {channels}ch" if channels > 0 else ""
            audio_map = {
                r'Atmos': ('Atmos', 15),
                r'TrueHD': ('TrueHD', 12),
                r'DTS-HD|DTS.*HD': ('DTS-HD', 10),
                r'DTS.*X|DTS:X': ('DTS:X', 14),
                r'^DTS$': ('DTS', 8),
                r'E-AC-3|EAC3': ('EAC3', 7),
                r'AC-3|AC3': ('AC3', 5),
                r'AAC': ('AAC', 3),
                r'FLAC': ('FLAC', 6),
                r'Opus': ('Opus', 4),
            }
            for pattern, (name, points) in audio_map.items():
                if re.search(pattern, audio_codec, re.IGNORECASE):
                    quality['audio'] = name
                    quality['score'] += points
                    quality['details'].append(f"{name} [MediaInfo{channel_info}] (+{points})")
                    break

        # HDR detection from MediaInfo
        if media_info['hdr']:
            quality['hdr'] = True
            hdr_format = get_mediainfo_hdr_format(file_path)
            if hdr_format and hdr_format['format']:
                quality['hdr_format'] = hdr_format['format']
                hdr_scores = {
                    'Dolby Vision': 18,
                    'HDR10+': 16,
                    'HDR10': 12,
                    'HLG': 10,
                    'HDR': 10
                }
                points = hdr_scores.get(hdr_format['format'], 10)
                quality['score'] += points
                quality['details'].append(f"{hdr_format['format']} [MediaInfo] (+{points})")
            else:
                quality['score'] += 10
                quality['details'].append("HDR [MediaInfo] (+10)")

        # Bitrate bonus
        if quality['bitrate'] > 0:
            bitrate_mbps = quality['bitrate'] / 1000000
            if bitrate_mbps >= 40:
                quality['score'] += 20
                quality['details'].append(f"High Bitrate [{bitrate_mbps:.1f} Mbps] (+20)")
            elif bitrate_mbps >= 20:
                quality['score'] += 15
                quality['details'].append(f"Good Bitrate [{bitrate_mbps:.1f} Mbps] (+15)")
            elif bitrate_mbps >= 10:
                quality['score'] += 10
                quality['details'].append(f"Moderate Bitrate [{bitrate_mbps:.1f} Mbps] (+10)")
            elif bitrate_mbps >= 5:
                quality['score'] += 5
                quality['details'].append(f"Low Bitrate [{bitrate_mbps:.1f} Mbps] (+5)")

    # Fallback to filename parsing
    if not media_info:
        quality['data_source'] = 'Filename'
        filename_lower = filename.lower()

        # Resolution scoring
        if any(x in filename_lower for x in ['2160p', '4k', 'uhd']):
            quality['resolution'] = '2160p'
            quality['score'] += 100
            quality['details'].append('4K/2160p (+100)')
        elif '1080p' in filename_lower:
            quality['resolution'] = '1080p'
            quality['score'] += 80
            quality['details'].append('1080p (+80)')
        elif '720p' in filename_lower:
            quality['resolution'] = '720p'
            quality['score'] += 60
            quality['details'].append('720p (+60)')
        elif any(x in filename_lower for x in ['480p', 'dvd']):
            quality['resolution'] = '480p'
            quality['score'] += 40
            quality['details'].append('480p (+40)')

        # Codec scoring
        if 'av1' in filename_lower:
            quality['codec'] = 'AV1'
            quality['score'] += 25
            quality['details'].append('AV1 (+25)')
        elif any(x in filename_lower for x in ['x265', 'h265', 'h.265', 'hevc']):
            quality['codec'] = 'HEVC/x265'
            quality['score'] += 20
            quality['details'].append('HEVC/x265 (+20)')
        elif 'vp9' in filename_lower:
            quality['codec'] = 'VP9'
            quality['score'] += 18
            quality['details'].append('VP9 (+18)')
        elif any(x in filename_lower for x in ['x264', 'h264', 'h.264', 'avc']):
            quality['codec'] = 'x264'
            quality['score'] += 15
            quality['details'].append('x264 (+15)')
        elif any(x in filename_lower for x in ['xvid', 'divx']):
            quality['codec'] = 'XviD'
            quality['score'] += 5
            quality['details'].append('XviD (+5)')

        # Audio scoring
        if 'atmos' in filename_lower:
            quality['audio'] = 'Atmos'
            quality['score'] += 15
            quality['details'].append('Atmos (+15)')
        elif re.search(r'dts[\s.\-]?x|dtsx', filename_lower):
            quality['audio'] = 'DTS:X'
            quality['score'] += 14
            quality['details'].append('DTS:X (+14)')
        elif 'truehd' in filename_lower:
            quality['audio'] = 'TrueHD'
            quality['score'] += 12
            quality['details'].append('TrueHD (+12)')
        elif re.search(r'dts-hd|dtshd|dts[\s.\-]?hd[\s.\-]?ma', filename_lower):
            quality['audio'] = 'DTS-HD'
            quality['score'] += 10
            quality['details'].append('DTS-HD (+10)')
        elif 'dts' in filename_lower:
            quality['audio'] = 'DTS'
            quality['score'] += 8
            quality['details'].append('DTS (+8)')
        elif any(x in filename_lower for x in ['eac3', 'ddp', 'dd+']):
            quality['audio'] = 'EAC3'
            quality['score'] += 7
            quality['details'].append('EAC3/DD+ (+7)')
        elif any(x in filename_lower for x in ['ac3', 'dd5.1']):
            quality['audio'] = 'AC3'
            quality['score'] += 5
            quality['details'].append('AC3 (+5)')
        elif 'flac' in filename_lower:
            quality['audio'] = 'FLAC'
            quality['score'] += 6
            quality['details'].append('FLAC (+6)')
        elif 'aac' in filename_lower:
            quality['audio'] = 'AAC'
            quality['score'] += 3
            quality['details'].append('AAC (+3)')

        # HDR scoring
        if re.search(r'dolby[\s.\-]?vision|dovi|dv[\s.\-]hdr|\.dv\.', filename_lower):
            quality['hdr'] = True
            quality['hdr_format'] = 'Dolby Vision'
            quality['score'] += 18
            quality['details'].append('Dolby Vision (+18)')
        elif re.search(r'hdr10\+|hdr10plus', filename_lower):
            quality['hdr'] = True
            quality['hdr_format'] = 'HDR10+'
            quality['score'] += 16
            quality['details'].append('HDR10+ (+16)')
        elif 'hdr10' in filename_lower:
            quality['hdr'] = True
            quality['hdr_format'] = 'HDR10'
            quality['score'] += 12
            quality['details'].append('HDR10 (+12)')
        elif 'hlg' in filename_lower:
            quality['hdr'] = True
            quality['hdr_format'] = 'HLG'
            quality['score'] += 10
            quality['details'].append('HLG (+10)')
        elif 'hdr' in filename_lower:
            quality['hdr'] = True
            quality['hdr_format'] = 'HDR'
            quality['score'] += 10
            quality['details'].append('HDR (+10)')

    # Source scoring (always from filename)
    filename_lower = filename.lower()
    if 'remux' in filename_lower:
        quality['source'] = 'Remux'
        quality['score'] += 35
        quality['details'].append('Remux (+35)')
    elif any(x in filename_lower for x in ['bluray', 'blu-ray', 'bdrip', 'brrip']):
        quality['source'] = 'BluRay'
        quality['score'] += 30
        quality['details'].append('BluRay (+30)')
    elif any(x in filename_lower for x in ['web-dl', 'webdl']):
        quality['source'] = 'WEB-DL'
        quality['score'] += 25
        quality['details'].append('WEB-DL (+25)')
    elif 'webrip' in filename_lower:
        quality['source'] = 'WEBRip'
        quality['score'] += 20
        quality['details'].append('WEBRip (+20)')
    elif 'hdtv' in filename_lower:
        quality['source'] = 'HDTV'
        quality['score'] += 15
        quality['details'].append('HDTV (+15)')
    elif 'dvdrip' in filename_lower:
        quality['source'] = 'DVDRip'
        quality['score'] += 10
        quality['details'].append('DVDRip (+10)')

    return quality


# =============================================================================
# EXPORT FUNCTIONS
# =============================================================================

def export_library_to_csv(path: str, output_path: str = None, media_type: str = "Movies") -> Optional[str]:
    """Export library data to CSV format"""
    if not output_path:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(os.path.dirname(path), f"MediaLibrary_{timestamp}.csv")

    print_color("\nExporting library to CSV...", 'yellow')
    log_info(f"Starting CSV export for: {path}")

    try:
        items = []
        root_path = Path(path)

        if media_type == "Movies":
            folders = [f for f in root_path.iterdir() if f.is_dir() and f.name != '_Trailers']

            for folder in folders:
                video_files = [f for f in folder.iterdir()
                              if f.is_file() and f.suffix.lower() in config.video_extensions]
                if not video_files:
                    continue

                video_file = max(video_files, key=lambda f: f.stat().st_size)
                title_info = get_normalized_title(folder.name)
                quality = get_quality_score(folder.name, str(video_file))

                items.append({
                    'Title': title_info['normalized_title'].title() if title_info['normalized_title'] else '',
                    'Year': title_info['year'] or '',
                    'FolderName': folder.name,
                    'FileName': video_file.name,
                    'FileSizeMB': round(video_file.stat().st_size / (1024 * 1024), 2),
                    'Resolution': quality['resolution'],
                    'VideoCodec': quality['codec'],
                    'AudioCodec': quality['audio'],
                    'Source': quality['source'],
                    'QualityScore': quality['score'],
                    'HDR': quality['hdr'],
                    'HDRFormat': quality.get('hdr_format', ''),
                    'Bitrate': round(quality['bitrate'] / 1000000, 1) if quality['bitrate'] > 0 else '',
                    'Container': video_file.suffix.lstrip('.').upper(),
                    'Path': str(folder),
                    'DataSource': quality['data_source']
                })

        if items:
            with open(output_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=items[0].keys())
                writer.writeheader()
                writer.writerows(items)

            print_color(f"Exported {len(items)} items to: {output_path}", 'green')
            log_info(f"CSV export completed: {len(items)} items to {output_path}")
        else:
            print_color("No items found to export", 'yellow')

        return output_path
    except Exception as e:
        log_error(f"Error exporting to CSV: {e}")
        return None


def export_library_to_json(path: str, output_path: str = None, media_type: str = "Movies") -> Optional[str]:
    """Export library data to JSON format"""
    if not output_path:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(os.path.dirname(path), f"MediaLibrary_{timestamp}.json")

    print_color("\nExporting library to JSON...", 'yellow')

    try:
        library = {
            'export_date': datetime.now().isoformat(),
            'media_type': media_type,
            'source_path': path,
            'items': []
        }

        root_path = Path(path)

        if media_type == "Movies":
            folders = [f for f in root_path.iterdir() if f.is_dir() and f.name != '_Trailers']

            for folder in folders:
                video_files = [f for f in folder.iterdir()
                              if f.is_file() and f.suffix.lower() in config.video_extensions]
                if not video_files:
                    continue

                video_file = max(video_files, key=lambda f: f.stat().st_size)
                title_info = get_normalized_title(folder.name)
                quality = get_quality_score(folder.name, str(video_file))

                library['items'].append({
                    'title': title_info['normalized_title'],
                    'year': title_info['year'],
                    'folder_name': folder.name,
                    'file_name': video_file.name,
                    'file_size': video_file.stat().st_size,
                    'quality': quality,
                    'path': str(folder)
                })

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(library, f, indent=2, default=str)

        print_color(f"Exported {len(library['items'])} items to: {output_path}", 'green')
        log_info(f"JSON export completed: {len(library['items'])} items")

        return output_path
    except Exception as e:
        log_error(f"Error exporting to JSON: {e}")
        return None


def export_library_to_html(path: str, output_path: str = None, media_type: str = "Movies") -> Optional[str]:
    """Export library data to HTML report"""
    if not output_path:
        if not os.path.exists(config.reports_folder):
            os.makedirs(config.reports_folder)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(config.reports_folder, f"MediaLibrary_{timestamp}.html")

    print_color("\nGenerating HTML report...", 'yellow')
    log_info(f"Starting HTML report for: {path}")

    try:
        items = []
        stats_data = {
            'total_items': 0,
            'total_size_gb': 0,
            'by_resolution': {},
            'by_codec': {},
            'by_year': {}
        }

        root_path = Path(path)
        folders = [f for f in root_path.iterdir() if f.is_dir() and f.name != '_Trailers']
        total = len(folders)

        for i, folder in enumerate(folders, 1):
            print(f"\rScanning: {i}/{total} - {folder.name[:40]}...".ljust(60), end='', flush=True)

            video_files = [f for f in folder.iterdir()
                          if f.is_file() and f.suffix.lower() in config.video_extensions]
            if not video_files:
                continue

            video_file = max(video_files, key=lambda f: f.stat().st_size)
            title_info = get_normalized_title(folder.name)
            quality = get_quality_score(folder.name, str(video_file))
            file_size = video_file.stat().st_size

            items.append({
                'title': title_info['normalized_title'].title() if title_info['normalized_title'] else folder.name,
                'year': title_info['year'],
                'resolution': quality['resolution'],
                'codec': quality['codec'],
                'size': file_size,
                'score': quality['score'],
                'hdr': quality['hdr'],
                'hdr_format': quality.get('hdr_format', '')
            })

            stats_data['total_items'] += 1
            stats_data['total_size_gb'] += file_size / (1024 ** 3)

            res = quality['resolution']
            stats_data['by_resolution'][res] = stats_data['by_resolution'].get(res, 0) + 1

            codec = quality['codec']
            stats_data['by_codec'][codec] = stats_data['by_codec'].get(codec, 0) + 1

            if title_info['year']:
                year = title_info['year']
                stats_data['by_year'][year] = stats_data['by_year'].get(year, 0) + 1

        print(f"\rScanning complete: {total} items processed".ljust(60))

        # Generate HTML
        html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LibraryLint Library Report</title>
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background: #1a1a2e; color: #eee; padding: 20px; }}
        h1 {{ color: #00d4ff; margin-bottom: 20px; text-align: center; }}
        h2 {{ color: #00d4ff; margin: 20px 0 10px; border-bottom: 2px solid #00d4ff; padding-bottom: 5px; }}
        .stats-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }}
        .stat-card {{ background: #16213e; padding: 20px; border-radius: 10px; text-align: center; }}
        .stat-value {{ font-size: 2.5em; color: #00d4ff; font-weight: bold; }}
        .stat-label {{ color: #888; margin-top: 5px; }}
        .chart-container {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 30px; }}
        .chart {{ background: #16213e; padding: 20px; border-radius: 10px; }}
        .bar-chart {{ margin-top: 10px; }}
        .bar-item {{ display: flex; align-items: center; margin: 8px 0; }}
        .bar-label {{ width: 100px; font-size: 0.9em; }}
        .bar-container {{ flex: 1; background: #0f3460; height: 25px; border-radius: 5px; overflow: hidden; }}
        .bar {{ height: 100%; background: linear-gradient(90deg, #00d4ff, #0097b2); transition: width 0.5s; }}
        .bar-value {{ margin-left: 10px; font-size: 0.9em; color: #888; }}
        table {{ width: 100%; border-collapse: collapse; background: #16213e; border-radius: 10px; overflow: hidden; }}
        th {{ background: #0f3460; color: #00d4ff; padding: 15px; text-align: left; }}
        td {{ padding: 12px 15px; border-bottom: 1px solid #0f3460; }}
        tr:hover {{ background: #1f4068; }}
        .quality-high {{ color: #00ff88; }}
        .quality-medium {{ color: #ffaa00; }}
        .quality-low {{ color: #ff4444; }}
        .search-box {{ width: 100%; padding: 12px; margin-bottom: 20px; background: #16213e; border: 2px solid #0f3460; border-radius: 8px; color: #eee; font-size: 1em; }}
        .search-box:focus {{ outline: none; border-color: #00d4ff; }}
        .footer {{ text-align: center; margin-top: 30px; color: #666; font-size: 0.9em; }}
    </style>
</head>
<body>
    <h1>LibraryLint Library Report</h1>
    <p style="text-align: center; color: #888; margin-bottom: 30px;">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

    <div class="stats-grid">
        <div class="stat-card">
            <div class="stat-value">{stats_data['total_items']}</div>
            <div class="stat-label">Total {media_type}</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">{stats_data['total_size_gb']:.1f} GB</div>
            <div class="stat-label">Total Size</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">{len(stats_data['by_resolution'])}</div>
            <div class="stat-label">Resolution Types</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">{len(stats_data['by_codec'])}</div>
            <div class="stat-label">Codec Types</div>
        </div>
    </div>

    <div class="chart-container">
        <div class="chart">
            <h2>By Resolution</h2>
            <div class="bar-chart">'''

        # Add resolution bars
        max_res = max(stats_data['by_resolution'].values()) if stats_data['by_resolution'] else 1
        for res, count in sorted(stats_data['by_resolution'].items(), key=lambda x: x[1], reverse=True):
            pct = round((count / max_res) * 100)
            html += f'''
                <div class="bar-item">
                    <span class="bar-label">{res}</span>
                    <div class="bar-container"><div class="bar" style="width: {pct}%"></div></div>
                    <span class="bar-value">{count}</span>
                </div>'''

        html += '''
            </div>
        </div>
        <div class="chart">
            <h2>By Codec</h2>
            <div class="bar-chart">'''

        # Add codec bars
        max_codec = max(stats_data['by_codec'].values()) if stats_data['by_codec'] else 1
        for codec, count in sorted(stats_data['by_codec'].items(), key=lambda x: x[1], reverse=True):
            pct = round((count / max_codec) * 100)
            html += f'''
                <div class="bar-item">
                    <span class="bar-label">{codec}</span>
                    <div class="bar-container"><div class="bar" style="width: {pct}%"></div></div>
                    <span class="bar-value">{count}</span>
                </div>'''

        html += '''
            </div>
        </div>
    </div>

    <h2>Library Contents</h2>
    <input type="text" class="search-box" placeholder="Search movies..." onkeyup="filterTable(this.value)">
    <table id="libraryTable">
        <thead>
            <tr>
                <th>Title</th>
                <th>Year</th>
                <th>Resolution</th>
                <th>Codec</th>
                <th>Size</th>
                <th>Score</th>
            </tr>
        </thead>
        <tbody>'''

        # Add table rows
        for item in sorted(items, key=lambda x: x['title']):
            size_str = format_file_size(item['size'])
            score_class = 'quality-high' if item['score'] >= 100 else ('quality-medium' if item['score'] >= 50 else 'quality-low')
            html += f'''
            <tr>
                <td>{html_escape(item['title'])}</td>
                <td>{item['year'] or ''}</td>
                <td>{item['resolution']}</td>
                <td>{item['codec']}</td>
                <td>{size_str}</td>
                <td class="{score_class}">{item['score']}</td>
            </tr>'''

        html += f'''
        </tbody>
    </table>

    <div class="footer">
        <p>Generated by LibraryLint v5.0</p>
    </div>

    <script>
        function filterTable(query) {{
            const table = document.getElementById('libraryTable');
            const rows = table.getElementsByTagName('tr');
            query = query.toLowerCase();
            for (let i = 1; i < rows.length; i++) {{
                const cells = rows[i].getElementsByTagName('td');
                let found = false;
                for (let j = 0; j < cells.length; j++) {{
                    if (cells[j].textContent.toLowerCase().includes(query)) {{
                        found = true;
                        break;
                    }}
                }}
                rows[i].style.display = found ? '' : 'none';
            }}
        }}
    </script>
</body>
</html>'''

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)

        print_color(f"HTML report generated: {output_path}", 'green')
        log_info(f"HTML report generated: {output_path}")

        return output_path
    except Exception as e:
        log_error(f"Error generating HTML report: {e}")
        return None


# =============================================================================
# EPISODE PARSING
# =============================================================================

def get_episode_info(filename: str) -> Dict[str, Any]:
    """Parse season/episode info from filename"""
    info = {
        'season': None,
        'episode': None,
        'episodes': [],
        'show_title': None,
        'episode_title': None,
        'is_multi_episode': False
    }

    basename = Path(filename).stem

    # Pattern 1: S01E01 or S01E01E02 or S01E01-E03
    match = re.match(r'^(.+?)[.\s_-]+[Ss](\d{1,2})[Ee](\d{1,2})(?:[Ee-](\d{1,2}))?(?:[Ee-](\d{1,2}))?(.*)$', basename)
    if match:
        info['show_title'] = re.sub(r'\.', ' ', match.group(1)).strip()
        info['season'] = int(match.group(2))
        info['episode'] = int(match.group(3))
        info['episodes'].append(info['episode'])

        if match.group(4):
            info['episodes'].append(int(match.group(4)))
            info['is_multi_episode'] = True
        if match.group(5):
            info['episodes'].append(int(match.group(5)))

        if match.group(6):
            remainder = re.sub(r'^[.\s_-]+', '', match.group(6))
            remainder = re.sub(r'\.', ' ', remainder)
            remainder = re.sub(r'\s*(720p|1080p|2160p|4K|HDTV|WEB-DL|WEBRip|BluRay|x264|x265|HEVC|AAC|AC3).*$', '', remainder, flags=re.IGNORECASE)
            if remainder.strip():
                info['episode_title'] = remainder.strip()

        return info

    # Pattern 2: 1x01 format
    match = re.match(r'^(.+?)[.\s_-]+(\d{1,2})x(\d{1,2})(.*)$', basename)
    if match:
        info['show_title'] = re.sub(r'\.', ' ', match.group(1)).strip()
        info['season'] = int(match.group(2))
        info['episode'] = int(match.group(3))
        info['episodes'].append(info['episode'])
        return info

    # Pattern 3: Season 1 Episode 1
    match = re.match(r'^(.+?)[.\s_-]+Season\s*(\d{1,2})[.\s_-]+Episode\s*(\d{1,2})(.*)$', basename, re.IGNORECASE)
    if match:
        info['show_title'] = re.sub(r'\.', ' ', match.group(1)).strip()
        info['season'] = int(match.group(2))
        info['episode'] = int(match.group(3))
        info['episodes'].append(info['episode'])
        return info

    return info


def get_normalized_title(name: str) -> Dict[str, Any]:
    """Extract normalized title and year from folder/file name"""
    result = {'normalized_title': None, 'year': None}

    basename = Path(name).stem

    # Try to extract year
    year_match = re.search(r'(19|20)\d{2}', basename)
    if year_match:
        result['year'] = year_match.group()

    # Remove everything after year or quality tags
    title = re.sub(r'[\(\[]?(19|20)\d{2}[\)\]]?.*$', '', basename)
    title = re.sub(r'\s*(720p|1080p|2160p|4K|HDRip|DVDRip|BRRip|BluRay|WEB-DL|WEBRip|x264|x265|HEVC).*$', '', title, flags=re.IGNORECASE)

    # Normalize
    title = re.sub(r'\.', ' ', title)
    title = re.sub(r'[_-]', ' ', title)
    title = re.sub(r'\s+', ' ', title)
    title = title.strip().lower()

    # Remove common articles
    title = re.sub(r'^(the|a|an)\s+', '', title)

    result['normalized_title'] = title
    return result


# =============================================================================
# FILE OPERATIONS
# =============================================================================

def remove_unnecessary_files(path: str):
    """Remove sample, proof, and screenshot files"""
    print_color("Cleaning unnecessary files...", 'yellow')
    log_info(f"Starting unnecessary file cleanup in: {path}")

    try:
        for root, dirs, files in os.walk(path):
            # Check directories
            for d in dirs[:]:
                if matches_pattern(d, config.unnecessary_patterns):
                    dir_path = os.path.join(root, d)
                    if config.dry_run:
                        print_color(f"[DRY-RUN] Would delete: {dir_path}", 'yellow')
                        log_info(f"Would delete: {dir_path}")
                    else:
                        try:
                            dir_size = sum(f.stat().st_size for f in Path(dir_path).rglob('*') if f.is_file())
                            shutil.rmtree(dir_path)
                            print_color(f"Deleted: {d}", 'gray')
                            log_info(f"Deleted: {dir_path}")
                            stats.files_deleted += 1
                            stats.bytes_deleted += dir_size
                            dirs.remove(d)
                        except Exception as e:
                            add_failed_operation('Delete', {'path': dir_path}, str(e))
                            log_error(f"Error deleting {dir_path}: {e}")

            # Check files
            for f in files:
                if matches_pattern(f, config.unnecessary_patterns):
                    file_path = os.path.join(root, f)
                    if config.dry_run:
                        print_color(f"[DRY-RUN] Would delete: {file_path}", 'yellow')
                    else:
                        try:
                            file_size = os.path.getsize(file_path)
                            os.remove(file_path)
                            print_color(f"Deleted: {f}", 'gray')
                            stats.files_deleted += 1
                            stats.bytes_deleted += file_size
                        except Exception as e:
                            add_failed_operation('Delete', {'path': file_path}, str(e))
                            log_error(f"Error deleting {file_path}: {e}")

        print_color("Unnecessary files cleaned", 'green')
    except Exception as e:
        log_error(f"Error during file cleanup: {e}")


def process_subtitles(path: str):
    """Process subtitle files - keep preferred languages"""
    print_color("Processing subtitle files...", 'yellow')
    log_info(f"Starting subtitle processing in: {path}")

    try:
        subtitle_files = []
        for ext in config.subtitle_extensions:
            subtitle_files.extend(Path(path).rglob(f'*{ext}'))

        if not subtitle_files:
            print_color("No subtitle files found", 'cyan')
            return

        print_color(f"Found {len(subtitle_files)} subtitle file(s)", 'cyan')

        for sub in subtitle_files:
            stats.subtitles_processed += 1
            sub_name_lower = sub.stem.lower()

            # Check if preferred language
            is_preferred = False
            for lang in config.preferred_subtitle_languages:
                if re.search(rf'\.{lang}$|\.{lang}\.|_{lang}$|_{lang}[_.]', sub_name_lower):
                    is_preferred = True
                    break

            # If no language tag, assume default language
            has_lang_tag = bool(re.search(r'\.(eng|en|english|spa|es|spanish|fre|fr|french|ger|de|german)(\.|$|_)', sub_name_lower))
            if not has_lang_tag:
                is_preferred = True

            if config.keep_subtitles and is_preferred:
                print_color(f"Keeping subtitle: {sub.name}", 'green')
            else:
                if config.dry_run:
                    print_color(f"[DRY-RUN] Would delete subtitle: {sub.name}", 'yellow')
                else:
                    try:
                        sub_size = sub.stat().st_size
                        sub.unlink()
                        print_color(f"Deleted subtitle: {sub.name}", 'gray')
                        stats.subtitles_deleted += 1
                        stats.bytes_deleted += sub_size
                    except Exception as e:
                        log_error(f"Error deleting subtitle {sub}: {e}")

        print_color("Subtitle processing completed", 'green')
    except Exception as e:
        log_error(f"Error processing subtitles: {e}")


def move_trailers_to_folder(path: str):
    """Move trailer files to _Trailers folder"""
    print_color("Processing trailer files...", 'yellow')
    log_info(f"Starting trailer processing in: {path}")

    try:
        trailer_files = []

        for pattern in config.trailer_patterns:
            for ext in config.video_extensions:
                trailer_files.extend(Path(path).rglob(f'*{pattern.strip("*")}*{ext}'))

        # Remove duplicates
        trailer_files = list(set(trailer_files))

        if not trailer_files:
            print_color("No trailer files found", 'cyan')
            return

        print_color(f"Found {len(trailer_files)} trailer file(s)", 'cyan')

        trailers_folder = Path(path) / '_Trailers'

        if not config.dry_run and not trailers_folder.exists():
            trailers_folder.mkdir(parents=True)
            print_color("Created _Trailers folder", 'green')
            stats.folders_created += 1

        for trailer in trailer_files:
            dest_path = trailers_folder / trailer.name

            # Handle duplicates
            counter = 1
            while dest_path.exists():
                dest_path = trailers_folder / f"{trailer.stem}_{counter}{trailer.suffix}"
                counter += 1

            if config.dry_run:
                print_color(f"[DRY-RUN] Would move trailer: {trailer.name} -> _Trailers/", 'yellow')
            else:
                try:
                    shutil.move(str(trailer), str(dest_path))
                    print_color(f"Moved trailer: {trailer.name}", 'green')
                    stats.trailers_moved += 1
                except Exception as e:
                    add_failed_operation('Move', {'source': str(trailer), 'destination': str(dest_path)}, str(e))
                    log_warning(f"Error moving trailer {trailer.name}: {e}")

        print_color("Trailer processing completed", 'green')
    except Exception as e:
        log_error(f"Error processing trailers: {e}")


def extract_archives(path: str, delete_after: bool = True):
    """Extract all archives in the path"""
    print_color("Extracting archives...", 'yellow')
    log_info(f"Starting archive extraction in: {path}")

    if not config.seven_zip_path:
        print_color("7-Zip not found. Please install 7-Zip.", 'red')
        log_error("7-Zip not found")
        return

    try:
        archive_files = []
        for ext in config.archive_extensions:
            archive_files.extend(Path(path).rglob(f'*{ext}'))

        if not archive_files:
            print_color("No archives found to extract", 'cyan')
            return

        print_color(f"Found {len(archive_files)} archive(s) to extract", 'cyan')

        for i, archive in enumerate(archive_files, 1):
            if config.dry_run:
                print_color(f"[DRY-RUN] [{i}/{len(archive_files)}] Would extract: {archive.name}", 'yellow')
            else:
                print_color(f"[{i}/{len(archive_files)}] Extracting: {archive.name}", 'cyan')

                try:
                    result = subprocess.run(
                        [config.seven_zip_path, 'x', f'-o{path}', str(archive), '-r', '-y'],
                        capture_output=True, text=True
                    )

                    if result.returncode != 0:
                        print_color(f"Warning: Failed to extract {archive.name}", 'yellow')
                        log_warning(f"Failed to extract: {archive}")
                        stats.archives_failed += 1
                    else:
                        print_color(f"Successfully extracted {archive.name}", 'green')
                        stats.archives_extracted += 1
                except Exception as e:
                    log_error(f"Error extracting {archive}: {e}")
                    stats.archives_failed += 1

        # Delete archives after extraction
        if delete_after and not config.dry_run:
            print_color("Deleting archive files...", 'yellow')
            for ext in config.archive_extensions:
                for archive in Path(path).rglob(f'*{ext}'):
                    try:
                        archive_size = archive.stat().st_size
                        archive.unlink()
                        stats.files_deleted += 1
                        stats.bytes_deleted += archive_size
                    except:
                        pass

            # Also delete .r00, .r01, etc.
            for archive in Path(path).rglob('*.r[0-9][0-9]'):
                try:
                    archive_size = archive.stat().st_size
                    archive.unlink()
                    stats.files_deleted += 1
                    stats.bytes_deleted += archive_size
                except:
                    pass

        print_color("Archives processed", 'green')
    except Exception as e:
        log_error(f"Error processing archives: {e}")


def create_folders_for_loose_files(path: str):
    """Create individual folders for loose video files"""
    print_color("Creating folders for loose video files...", 'yellow')
    log_info(f"Starting folder creation for loose files in: {path}")

    try:
        root_path = Path(path)
        loose_files = [f for f in root_path.iterdir()
                      if f.is_file() and f.suffix.lower() in config.video_extensions]

        if not loose_files:
            print_color("No loose video files found", 'cyan')
            return

        for file in loose_files:
            folder_name = file.stem
            folder_path = root_path / folder_name

            if config.dry_run:
                print_color(f"[DRY-RUN] Would create folder '{folder_name}' and move: {file.name}", 'yellow')
            else:
                print_color(f"Processing: {file.name}", 'cyan')

                if not folder_path.exists():
                    folder_path.mkdir(parents=True)
                    stats.folders_created += 1

                try:
                    shutil.move(str(file), str(folder_path / file.name))
                    print_color(f"Moved to: {folder_name}", 'green')
                    stats.files_moved += 1
                except Exception as e:
                    add_failed_operation('Move', {'source': str(file), 'destination': str(folder_path / file.name)}, str(e))
                    log_error(f"Error moving {file.name}: {e}")
    except Exception as e:
        log_error(f"Error organizing loose files: {e}")


def clean_folder_names(path: str):
    """Clean folder names by removing tags and formatting"""
    print_color("Cleaning folder names...", 'yellow')
    log_info(f"Starting folder name cleaning in: {path}")

    try:
        root_path = Path(path)

        for folder in root_path.iterdir():
            if not folder.is_dir() or folder.name == '_Trailers':
                continue

            new_name = folder.name

            # Remove tags
            for tag in config.tags:
                if tag.lower() in new_name.lower():
                    pattern = re.compile(re.escape(tag), re.IGNORECASE)
                    parts = pattern.split(new_name)
                    if parts[0].strip(' .-'):
                        new_name = parts[0].strip(' .-')

            # Replace dots with spaces
            if '.' in new_name:
                new_name = re.sub(r'\.', ' ', new_name)
                new_name = re.sub(r'\s+', ' ', new_name).strip()

            # Format year with parentheses
            year_match = re.search(r'\s(19|20)(\d{2})$', new_name)
            if year_match:
                new_name = re.sub(r'\s((19|20)\d{2})$', r' (\1)', new_name)

            if new_name != folder.name:
                new_path = root_path / new_name

                if config.dry_run:
                    print_color(f"[DRY-RUN] Would rename '{folder.name}' to '{new_name}'", 'yellow')
                else:
                    try:
                        if not new_path.exists():
                            folder.rename(new_path)
                            log_info(f"Renamed '{folder.name}' to '{new_name}'")
                            stats.folders_renamed += 1
                    except Exception as e:
                        add_failed_operation('Rename', {'path': str(folder), 'new_path': str(new_path)}, str(e))
                        log_error(f"Error renaming {folder.name}: {e}")

        print_color("Folder names cleaned", 'green')
    except Exception as e:
        log_error(f"Error cleaning folder names: {e}")


def remove_empty_folders(path: str):
    """Remove empty folders recursively"""
    print_color("Cleaning up empty folders...", 'yellow')
    log_info(f"Starting empty folder cleanup in: {path}")

    try:
        removed = 0
        for root, dirs, files in os.walk(path, topdown=False):
            for d in dirs:
                dir_path = os.path.join(root, d)
                try:
                    if not os.listdir(dir_path):
                        if config.dry_run:
                            print_color(f"[DRY-RUN] Would remove: {dir_path}", 'yellow')
                        else:
                            os.rmdir(dir_path)
                            stats.empty_folders_removed += 1
                            removed += 1
                except:
                    pass

        if removed > 0:
            print_color(f"Removed {removed} empty folder(s)", 'green')
        else:
            print_color("No empty folders found", 'cyan')
    except Exception as e:
        log_error(f"Error removing empty folders: {e}")


def organize_seasons(path: str):
    """Organize TV episodes into season folders"""
    print_color("Organizing episodes into season folders...", 'yellow')
    log_info(f"Starting season organization in: {path}")

    try:
        video_files = []
        for ext in config.video_extensions:
            video_files.extend(Path(path).rglob(f'*{ext}'))

        if not video_files:
            print_color("No video files found", 'cyan')
            return

        print_color(f"Found {len(video_files)} video file(s)", 'cyan')
        organized = 0

        for file in video_files:
            ep_info = get_episode_info(file.name)

            if ep_info['season'] is not None:
                season_folder = f"Season {ep_info['season']:02d}"
                season_path = Path(path) / season_folder

                # Skip if already in correct folder
                if file.parent.name == season_folder:
                    print_color(f"Already organized: {file.name}", 'gray')
                    continue

                if config.dry_run:
                    print_color(f"[DRY-RUN] Would move to {season_folder}: {file.name}", 'yellow')
                else:
                    if not season_path.exists():
                        season_path.mkdir(parents=True)
                        print_color(f"Created folder: {season_folder}", 'green')
                        stats.folders_created += 1

                    try:
                        dest_path = season_path / file.name
                        shutil.move(str(file), str(dest_path))
                        print_color(f"Moved to {season_folder}: {file.name}", 'cyan')
                        stats.files_moved += 1
                        organized += 1
                    except Exception as e:
                        log_warning(f"Error moving {file.name}: {e}")
            else:
                print_color(f"Could not parse: {file.name}", 'yellow')

        if organized > 0:
            print_color(f"Organized {organized} episode(s) into season folders", 'green')
    except Exception as e:
        log_error(f"Error organizing seasons: {e}")


# =============================================================================
# NFO FUNCTIONS
# =============================================================================

def read_nfo_file(nfo_path: str) -> Optional[Dict[str, Any]]:
    """Parse an existing NFO file"""
    metadata = {
        'title': None,
        'original_title': None,
        'year': None,
        'plot': None,
        'rating': None,
        'imdb_id': None,
        'tmdb_id': None,
        'genres': [],
        'directors': [],
        'actors': []
    }

    try:
        tree = ET.parse(nfo_path)
        root = tree.getroot()
        stats.nfo_files_read += 1

        # Movie NFO
        if root.tag == 'movie':
            metadata['title'] = root.findtext('title')
            metadata['original_title'] = root.findtext('originaltitle')
            metadata['year'] = root.findtext('year')
            metadata['plot'] = root.findtext('plot')
            metadata['rating'] = root.findtext('rating')

            # Get IDs
            for uniqueid in root.findall('uniqueid'):
                if uniqueid.get('type') == 'imdb':
                    metadata['imdb_id'] = uniqueid.text
                elif uniqueid.get('type') == 'tmdb':
                    metadata['tmdb_id'] = uniqueid.text

            metadata['genres'] = [g.text for g in root.findall('genre') if g.text]
            metadata['directors'] = [d.text for d in root.findall('director') if d.text]

        return metadata
    except Exception as e:
        log_error(f"Error parsing NFO {nfo_path}: {e}")
        return None


def create_movie_nfo(video_path: str, title: str = None, year: str = None):
    """Create a basic Kodi NFO file for a movie"""
    try:
        video_file = Path(video_path)
        nfo_path = video_file.parent / f"{video_file.stem}.nfo"

        if nfo_path.exists():
            print_color(f"NFO already exists: {video_file.name}", 'cyan')
            return

        # Extract title/year from folder name if not provided
        if not title or not year:
            folder_name = video_file.parent.name
            title_info = get_normalized_title(folder_name)

            if not title:
                title = folder_name
                if title_info['year']:
                    title = re.sub(rf'\s*[\(\[]?{title_info["year"]}[\)\]]?.*$', '', title)
                title = re.sub(r'\.', ' ', title).strip()

            if not year:
                year = title_info['year']

        if config.dry_run:
            print_color(f"[DRY-RUN] Would create NFO for: {title} ({year})", 'yellow')
            return

        # Create NFO XML
        movie = ET.Element('movie')
        ET.SubElement(movie, 'title').text = title
        if year:
            ET.SubElement(movie, 'year').text = year
        ET.SubElement(movie, 'plot')
        ET.SubElement(movie, 'outline')
        ET.SubElement(movie, 'tagline')
        ET.SubElement(movie, 'runtime')
        ET.SubElement(movie, 'thumb')
        ET.SubElement(movie, 'fanart')

        # Pretty print
        xml_str = minidom.parseString(ET.tostring(movie)).toprettyxml(indent='    ')
        xml_str = '\n'.join(xml_str.split('\n')[1:])
        xml_str = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_str

        with open(nfo_path, 'w', encoding='utf-8') as f:
            f.write(xml_str)

        print_color(f"Created NFO: {title} ({year})", 'green')
        stats.nfo_files_created += 1
    except Exception as e:
        log_error(f"Error creating NFO for {video_path}: {e}")


# =============================================================================
# DUPLICATE DETECTION
# =============================================================================

def find_duplicate_movies(path: str) -> List[List[Dict]]:
    """Find potential duplicate movies"""
    print_color("\nScanning for duplicate movies...", 'yellow')
    log_info(f"Starting duplicate scan in: {path}")

    try:
        root_path = Path(path)
        movie_folders = [f for f in root_path.iterdir()
                        if f.is_dir() and f.name != '_Trailers']

        if not movie_folders:
            print_color("No movie folders found", 'cyan')
            return []

        # Build lookup
        title_lookup = {}

        for folder in movie_folders:
            title_info = get_normalized_title(folder.name)
            quality = get_quality_score(folder.name)

            # Find main video file
            video_files = [f for f in folder.iterdir()
                         if f.is_file() and f.suffix.lower() in config.video_extensions]
            file_size = max((f.stat().st_size for f in video_files), default=0) if video_files else 0

            entry = {
                'path': str(folder),
                'original_name': folder.name,
                'year': title_info['year'],
                'quality': quality,
                'file_size': file_size
            }

            key = title_info['normalized_title']
            if title_info['year']:
                key = f"{key}|{title_info['year']}"

            if key not in title_lookup:
                title_lookup[key] = []
            title_lookup[key].append(entry)

        # Find duplicates
        duplicates = [entries for entries in title_lookup.values() if len(entries) > 1]

        return duplicates
    except Exception as e:
        log_error(f"Error scanning for duplicates: {e}")
        return []


def show_duplicate_report(path: str):
    """Display duplicate movies report"""
    duplicates = find_duplicate_movies(path)

    if not duplicates:
        print_color("No duplicate movies found!", 'green')
        return

    print_color("\n╔══════════════════════════════════════════════════════════════╗", 'yellow')
    print_color("║                    DUPLICATE REPORT                          ║", 'yellow')
    print_color("╠══════════════════════════════════════════════════════════════╣", 'yellow')
    print_color(f"║  Found {len(duplicates)} potential duplicate group(s)                        ║", 'yellow')
    print_color("╚══════════════════════════════════════════════════════════════╝", 'yellow')

    for group_num, group in enumerate(duplicates, 1):
        print_color(f"\n[{group_num}] Potential Duplicates:", 'cyan')

        # Sort by quality score
        sorted_group = sorted(group, key=lambda x: x['quality']['score'], reverse=True)

        for i, movie in enumerate(sorted_group):
            size_str = format_file_size(movie['file_size'])
            score = movie['quality']['score']
            resolution = movie['quality']['resolution']
            source = movie['quality']['source']

            if i == 0:
                print_color(f"  [KEEP] ", 'green', end='')
            else:
                print_color(f"  [DEL?] ", 'red', end='')

            print_color(movie['original_name'], 'white')
            print_color(f"         Score: {score} | {size_str} | {resolution} {source}", 'gray')


# =============================================================================
# HEALTH CHECK
# =============================================================================

def health_check(path: str, media_type: str = "Movies"):
    """Perform library health check"""
    print_color("\n╔══════════════════════════════════════════════════════════════╗", 'cyan')
    print_color("║                  LIBRARY HEALTH CHECK                        ║", 'cyan')
    print_color("╚══════════════════════════════════════════════════════════════╝", 'cyan')

    log_info(f"Starting health check for: {path}")

    issues = {
        'empty_folders': [],
        'no_video_files': [],
        'zero_byte_files': [],
        'orphaned_subtitles': [],
        'missing_nfo': [],
        'small_videos': [],
        'naming_issues': []
    }

    root_path = Path(path)

    # Check empty folders
    print_color("\nChecking for empty folders...", 'yellow')
    for folder in root_path.rglob('*'):
        if folder.is_dir() and not list(folder.iterdir()):
            issues['empty_folders'].append(folder)

    # Check folders without video files
    print_color("Checking for folders without video files...", 'yellow')
    for folder in root_path.iterdir():
        if folder.is_dir() and folder.name != '_Trailers':
            has_video = any(f.suffix.lower() in config.video_extensions
                          for f in folder.iterdir() if f.is_file())
            if not has_video:
                issues['no_video_files'].append(folder)

    # Check zero-byte files
    print_color("Checking for corrupted/zero-byte files...", 'yellow')
    for file in root_path.rglob('*'):
        if file.is_file() and file.stat().st_size == 0:
            issues['zero_byte_files'].append(file)

    # Check small videos (under 50MB, not samples/trailers)
    print_color("Checking for suspiciously small video files...", 'yellow')
    for file in root_path.rglob('*'):
        if (file.is_file() and
            file.suffix.lower() in config.video_extensions and
            file.stat().st_size < 50 * 1024 * 1024 and
            not any(x in file.name.lower() for x in ['sample', 'trailer', 'teaser'])):
            issues['small_videos'].append(file)

    # Check for naming issues
    print_color("Checking for naming issues...", 'yellow')
    for folder in root_path.iterdir():
        if folder.is_dir() and folder.name != '_Trailers':
            # Check for dots in folder names
            if re.search(r'\..*\.', folder.name) and not re.search(r'\(\d{4}\)', folder.name):
                issues['naming_issues'].append({
                    'path': str(folder),
                    'issue': 'Contains dots (should be spaces)'
                })
            # Check for missing year
            if media_type == "Movies" and not re.search(r'(19|20)\d{2}', folder.name):
                issues['naming_issues'].append({
                    'path': str(folder),
                    'issue': 'Missing year'
                })

    # Display results
    print_color("\n=== Health Check Results ===", 'cyan')

    total_issues = 0

    if issues['empty_folders']:
        print_color(f"\nEmpty Folders ({len(issues['empty_folders'])}):", 'yellow')
        for f in issues['empty_folders'][:10]:
            print_color(f"  - {f}", 'gray')
        total_issues += len(issues['empty_folders'])

    if issues['no_video_files']:
        print_color(f"\nFolders Without Video Files ({len(issues['no_video_files'])}):", 'yellow')
        for f in issues['no_video_files'][:10]:
            print_color(f"  - {f.name}", 'gray')
        total_issues += len(issues['no_video_files'])

    if issues['zero_byte_files']:
        print_color(f"\nZero-Byte Files ({len(issues['zero_byte_files'])}):", 'red')
        for f in issues['zero_byte_files'][:10]:
            print_color(f"  - {f}", 'gray')
        total_issues += len(issues['zero_byte_files'])

    if issues['small_videos']:
        print_color(f"\nSuspiciously Small Videos ({len(issues['small_videos'])}):", 'yellow')
        for f in issues['small_videos'][:10]:
            print_color(f"  - {f.name} ({format_file_size(f.stat().st_size)})", 'gray')
        total_issues += len(issues['small_videos'])

    if issues['naming_issues']:
        print_color(f"\nNaming Issues ({len(issues['naming_issues'])}):", 'yellow')
        for issue in issues['naming_issues'][:10]:
            print_color(f"  - {issue['path']}: {issue['issue']}", 'gray')
        if len(issues['naming_issues']) > 10:
            print_color(f"  ... and {len(issues['naming_issues']) - 10} more", 'gray')
        total_issues += len(issues['naming_issues'])

    print()
    if total_issues == 0:
        print_color("Library is healthy! No issues found.", 'green')
    else:
        print_color(f"Found {total_issues} issue(s) in the library.", 'yellow')


# =============================================================================
# CODEC ANALYSIS
# =============================================================================

def codec_analysis(path: str):
    """Analyze video codecs in library"""
    print_color("\n╔══════════════════════════════════════════════════════════════╗", 'cyan')
    print_color("║                    CODEC ANALYSIS                            ║", 'cyan')
    print_color("╚══════════════════════════════════════════════════════════════╝", 'cyan')

    log_info(f"Starting codec analysis for: {path}")

    video_files = []
    for ext in config.video_extensions:
        video_files.extend(Path(path).rglob(f'*{ext}'))

    if not video_files:
        print_color("No video files found", 'cyan')
        return

    print_color(f"\nAnalyzing {len(video_files)} video file(s)...", 'yellow')

    analysis = {
        'total_files': len(video_files),
        'total_size': 0,
        'by_resolution': {},
        'by_codec': {},
        'by_container': {},
        'need_transcode': []
    }

    for i, file in enumerate(video_files, 1):
        print(f"\rAnalyzing: {i}/{len(video_files)}".ljust(40), end='', flush=True)

        file_size = file.stat().st_size
        analysis['total_size'] += file_size

        quality = get_quality_score(file.name, str(file))
        container = file.suffix.upper().lstrip('.')

        # Count by resolution
        res = quality['resolution']
        analysis['by_resolution'][res] = analysis['by_resolution'].get(res, 0) + 1

        # Count by codec
        codec = quality['codec']
        analysis['by_codec'][codec] = analysis['by_codec'].get(codec, 0) + 1

        # Count by container
        analysis['by_container'][container] = analysis['by_container'].get(container, 0) + 1

        # Check if needs transcoding (only legacy codecs/containers)
        needs_transcode = False
        reasons = []
        transcode_mode = 'none'

        if quality['codec'] in ['XviD', 'DivX', 'MPEG-4']:
            needs_transcode = True
            transcode_mode = 'transcode'
            reasons.append("Legacy codec - will transcode to H.264")
        elif container == 'AVI':
            needs_transcode = True
            if quality['codec'] in ['x264', 'H.264', 'AVC']:
                transcode_mode = 'remux'
                reasons.append("H.264 in AVI - will remux to MKV (no re-encoding)")
            else:
                transcode_mode = 'transcode'
                reasons.append("AVI with legacy codec - will transcode to H.264")

        if needs_transcode:
            analysis['need_transcode'].append({
                'path': str(file),
                'filename': file.name,
                'size': file_size,
                'resolution': res,
                'codec': codec,
                'container': container,
                'reasons': '; '.join(reasons),
                'transcode_mode': transcode_mode
            })

    print()  # New line after progress

    # Display results
    print_color("\n=== Library Statistics ===", 'cyan')
    print_color(f"Total Files: {analysis['total_files']}", 'white')
    print_color(f"Total Size: {format_file_size(analysis['total_size'])}", 'white')

    print_color("\n=== By Resolution ===", 'cyan')
    for res, count in sorted(analysis['by_resolution'].items(), key=lambda x: x[1], reverse=True):
        pct = round((count / analysis['total_files']) * 100, 1)
        print_color(f"  {res}: {count} ({pct}%)", 'white')

    print_color("\n=== By Video Codec ===", 'cyan')
    for codec, count in sorted(analysis['by_codec'].items(), key=lambda x: x[1], reverse=True):
        pct = round((count / analysis['total_files']) * 100, 1)
        print_color(f"  {codec}: {count} ({pct}%)", 'white')

    print_color("\n=== By Container ===", 'cyan')
    for container, count in sorted(analysis['by_container'].items(), key=lambda x: x[1], reverse=True):
        pct = round((count / analysis['total_files']) * 100, 1)
        print_color(f"  {container}: {count} ({pct}%)", 'white')

    if analysis['need_transcode']:
        print_color(f"\n=== Potential Transcoding Queue ===", 'yellow')
        print_color(f"Found {len(analysis['need_transcode'])} file(s) that may benefit from transcoding:", 'white')

        for item in analysis['need_transcode'][:10]:
            print_color(f"\n  {item['filename']}", 'white')
            print_color(f"    Size: {format_file_size(item['size'])} | {item['resolution']} | {item['codec']}", 'gray')
            print_color(f"    Mode: {item['transcode_mode']} | Reason: {item['reasons']}", 'yellow')

        if len(analysis['need_transcode']) > 10:
            print_color(f"\n  ... and {len(analysis['need_transcode']) - 10} more files", 'gray')
    else:
        print_color("\nNo files require transcoding for compatibility.", 'green')


# =============================================================================
# TMDB INTEGRATION
# =============================================================================

def search_tmdb_movie(title: str, year: str = None, api_key: str = None) -> Optional[Dict]:
    """Search TMDB for a movie"""
    if not HAS_REQUESTS or not api_key:
        return None

    try:
        url = "https://api.themoviedb.org/3/search/movie"
        params = {'api_key': api_key, 'query': title}
        if year:
            params['year'] = year

        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()

        if data.get('results'):
            movie = data['results'][0]
            return {
                'id': movie['id'],
                'title': movie['title'],
                'original_title': movie.get('original_title'),
                'year': movie.get('release_date', '')[:4] if movie.get('release_date') else None,
                'overview': movie.get('overview'),
                'rating': movie.get('vote_average'),
                'poster_path': f"https://image.tmdb.org/t/p/w500{movie['poster_path']}" if movie.get('poster_path') else None
            }
        return None
    except Exception as e:
        log_error(f"Error searching TMDB: {e}")
        return None


def get_tmdb_movie_details(movie_id: int, api_key: str) -> Optional[Dict]:
    """Get detailed movie info from TMDB"""
    if not HAS_REQUESTS or not api_key:
        return None

    try:
        url = f"https://api.themoviedb.org/3/movie/{movie_id}"
        params = {'api_key': api_key, 'append_to_response': 'credits,external_ids'}

        response = requests.get(url, params=params)
        response.raise_for_status()
        movie = response.json()

        directors = [c['name'] for c in movie.get('credits', {}).get('crew', []) if c.get('job') == 'Director']
        cast = [{'name': c['name'], 'role': c.get('character', '')}
                for c in movie.get('credits', {}).get('cast', [])[:10]]

        return {
            'id': movie['id'],
            'title': movie['title'],
            'original_title': movie.get('original_title'),
            'tagline': movie.get('tagline'),
            'year': movie.get('release_date', '')[:4] if movie.get('release_date') else None,
            'overview': movie.get('overview'),
            'rating': movie.get('vote_average'),
            'votes': movie.get('vote_count'),
            'runtime': movie.get('runtime'),
            'genres': [g['name'] for g in movie.get('genres', [])],
            'studios': [s['name'] for s in movie.get('production_companies', [])],
            'directors': directors,
            'cast': cast,
            'imdb_id': movie.get('external_ids', {}).get('imdb_id'),
            'poster_path': f"https://image.tmdb.org/t/p/w500{movie['poster_path']}" if movie.get('poster_path') else None,
            'backdrop_path': f"https://image.tmdb.org/t/p/original{movie['backdrop_path']}" if movie.get('backdrop_path') else None
        }
    except Exception as e:
        log_error(f"Error getting TMDB details: {e}")
        return None


# =============================================================================
# MAIN PROCESSING
# =============================================================================

def process_movies(path: str):
    """Process a movie library"""
    print_color("\nMovie Routine", 'magenta')
    log_info(f"Movie routine started for path: {path}")

    if config.dry_run:
        print_color("\n*** DRY-RUN MODE - Previewing changes only ***\n", 'yellow')

    # Step 1: Process trailers
    if config.keep_trailers:
        move_trailers_to_folder(path)

    # Step 2: Remove unnecessary files
    remove_unnecessary_files(path)

    # Step 3: Process subtitles
    process_subtitles(path)

    # Step 4: Extract archives
    extract_archives(path, delete_after=True)

    # Step 5: Create folders for loose files
    create_folders_for_loose_files(path)

    # Step 6: Clean folder names
    clean_folder_names(path)

    # Step 7: Remove empty folders
    remove_empty_folders(path)

    # Step 8: Retry failed operations
    if failed_operations:
        retry_failed_operations()

    # Step 9: Check for duplicates
    if config.check_duplicates:
        show_enhanced_duplicate_report(path)

    print_color("\nMovie routine completed!", 'magenta')


def process_tv_shows(path: str):
    """Process a TV show library"""
    print_color("\nTV Show Routine", 'magenta')
    log_info(f"TV Show routine started for path: {path}")

    if config.dry_run:
        print_color("\n*** DRY-RUN MODE - Previewing changes only ***\n", 'yellow')

    # Step 1: Extract archives
    extract_archives(path, delete_after=True)

    # Step 2: Remove unnecessary files
    remove_unnecessary_files(path)

    # Step 3: Process subtitles
    process_subtitles(path)

    # Step 4: Organize into season folders
    if config.organize_seasons:
        organize_seasons(path)

    # Step 5: Remove empty folders
    remove_empty_folders(path)

    # Step 6: Retry failed operations
    if failed_operations:
        retry_failed_operations()

    print_color("\nTV Show routine completed!", 'magenta')


def show_statistics():
    """Display processing statistics"""
    stats.end_time = datetime.now()
    duration = stats.end_time - stats.start_time

    print_color("\n╔══════════════════════════════════════════════════════════════╗", 'cyan')
    print_color("║                    PROCESSING SUMMARY                        ║", 'cyan')
    print_color("╠══════════════════════════════════════════════════════════════╣", 'cyan')

    print_color(f"║  Duration:              {str(duration).split('.')[0]:<39}║", 'cyan')

    print_color("╠══════════════════════════════════════════════════════════════╣", 'cyan')

    if stats.files_deleted > 0:
        print_color(f"║  Files Deleted:         {stats.files_deleted:<39}║", 'cyan')
    if stats.bytes_deleted > 0:
        print_color(f"║  Space Reclaimed:       {format_file_size(stats.bytes_deleted):<39}║", 'cyan')
    if stats.archives_extracted > 0:
        print_color(f"║  Archives Extracted:    {stats.archives_extracted:<39}║", 'cyan')
    if stats.folders_created > 0:
        print_color(f"║  Folders Created:       {stats.folders_created:<39}║", 'cyan')
    if stats.folders_renamed > 0:
        print_color(f"║  Folders Renamed:       {stats.folders_renamed:<39}║", 'cyan')
    if stats.files_moved > 0:
        print_color(f"║  Files Moved:           {stats.files_moved:<39}║", 'cyan')
    if stats.trailers_moved > 0:
        print_color(f"║  Trailers Moved:        {stats.trailers_moved:<39}║", 'cyan')
    if stats.nfo_files_created > 0:
        print_color(f"║  NFO Files Created:     {stats.nfo_files_created:<39}║", 'cyan')
    if stats.operations_retried > 0:
        print_color(f"║  Operations Retried:    {stats.operations_retried:<39}║", 'cyan')

    print_color("╠══════════════════════════════════════════════════════════════╣", 'cyan')

    print_color(f"║  Errors:                {stats.errors:<39}║", 'cyan')
    print_color(f"║  Warnings:              {stats.warnings:<39}║", 'cyan')

    print_color("╚══════════════════════════════════════════════════════════════╝", 'cyan')


# =============================================================================
# CLI INTERFACE
# =============================================================================

def main():
    """Main entry point"""
    global config, stats, logger

    parser = argparse.ArgumentParser(
        description='LibraryLint - Cross-platform media library organization tool',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  %(prog)s --movies /path/to/movies
  %(prog)s --tvshows /path/to/shows --dry-run
  %(prog)s --health-check /path/to/library
  %(prog)s --codec-analysis /path/to/library
  %(prog)s --export-csv /path/to/library
  %(prog)s --export-html /path/to/library
        '''
    )

    parser.add_argument('--movies', metavar='PATH', help='Process movie library at PATH')
    parser.add_argument('--tvshows', metavar='PATH', help='Process TV show library at PATH')
    parser.add_argument('--health-check', metavar='PATH', help='Run health check on library')
    parser.add_argument('--codec-analysis', metavar='PATH', help='Analyze codecs in library')
    parser.add_argument('--duplicates', metavar='PATH', help='Find duplicate movies')
    parser.add_argument('--duplicates-enhanced', metavar='PATH', help='Find duplicates with file hashing')

    parser.add_argument('--export-csv', metavar='PATH', help='Export library to CSV')
    parser.add_argument('--export-json', metavar='PATH', help='Export library to JSON')
    parser.add_argument('--export-html', metavar='PATH', help='Export library to HTML report')

    parser.add_argument('--dry-run', action='store_true', help='Preview changes without making them')
    parser.add_argument('--no-subtitles', action='store_true', help='Delete all subtitle files')
    parser.add_argument('--no-trailers', action='store_true', help='Delete trailers instead of moving')
    parser.add_argument('--no-organize-seasons', action='store_true', help='Skip season folder organization')
    parser.add_argument('--check-duplicates', action='store_true', help='Check for duplicate movies')
    parser.add_argument('--generate-nfo', action='store_true', help='Generate NFO files')

    parser.add_argument('--tmdb-key', metavar='KEY', help='TMDB API key for metadata fetching')
    parser.add_argument('--log-file', metavar='FILE', help='Custom log file path')

    parser.add_argument('--version', action='version', version='LibraryLint 5.0')

    args = parser.parse_args()

    # Configure
    config.dry_run = args.dry_run
    config.keep_subtitles = not args.no_subtitles
    config.keep_trailers = not args.no_trailers
    config.organize_seasons = not args.no_organize_seasons
    config.check_duplicates = args.check_duplicates
    config.generate_nfo = args.generate_nfo
    config.tmdb_api_key = args.tmdb_key or ''

    if args.log_file:
        config.log_file = args.log_file

    # Setup logging
    logger = setup_logging(config.log_file)

    # Display header
    print_color("\n=== LibraryLint v5.0 (Python) ===", 'cyan')
    print_color("Cross-platform media library organization tool", 'gray')
    print()
    print_color("=" * 64, 'yellow')
    print_color("  WARNING: This Python version is EXPERIMENTAL/UNMAINTAINED", 'yellow')
    print_color("  For the full-featured version, use: LibraryLint.ps1", 'yellow')
    print_color("  Request cross-platform support: github.com/kliatsko/librarylint/issues", 'yellow')
    print_color("=" * 64, 'yellow')

    if test_mediainfo_installation():
        print_color("MediaInfo: Available", 'green')
    else:
        print_color("MediaInfo: Not found (using filename parsing)", 'yellow')

    if test_ffmpeg_installation():
        print_color("FFmpeg: Available", 'green')
    else:
        print_color("FFmpeg: Not found", 'yellow')

    print()

    if config.dry_run:
        print_color("DRY-RUN MODE ENABLED - No changes will be made\n", 'yellow')

    # Process based on arguments
    if args.movies:
        if not os.path.isdir(args.movies):
            print_color(f"Error: '{args.movies}' is not a valid directory", 'red')
            sys.exit(1)
        process_movies(args.movies)
        show_statistics()

    elif args.tvshows:
        if not os.path.isdir(args.tvshows):
            print_color(f"Error: '{args.tvshows}' is not a valid directory", 'red')
            sys.exit(1)
        process_tv_shows(args.tvshows)
        show_statistics()

    elif args.health_check:
        if not os.path.isdir(args.health_check):
            print_color(f"Error: '{args.health_check}' is not a valid directory", 'red')
            sys.exit(1)
        health_check(args.health_check)

    elif args.codec_analysis:
        if not os.path.isdir(args.codec_analysis):
            print_color(f"Error: '{args.codec_analysis}' is not a valid directory", 'red')
            sys.exit(1)
        codec_analysis(args.codec_analysis)

    elif args.duplicates:
        if not os.path.isdir(args.duplicates):
            print_color(f"Error: '{args.duplicates}' is not a valid directory", 'red')
            sys.exit(1)
        show_duplicate_report(args.duplicates)

    elif args.duplicates_enhanced:
        if not os.path.isdir(args.duplicates_enhanced):
            print_color(f"Error: '{args.duplicates_enhanced}' is not a valid directory", 'red')
            sys.exit(1)
        show_enhanced_duplicate_report(args.duplicates_enhanced)

    elif args.export_csv:
        if not os.path.isdir(args.export_csv):
            print_color(f"Error: '{args.export_csv}' is not a valid directory", 'red')
            sys.exit(1)
        export_library_to_csv(args.export_csv)

    elif args.export_json:
        if not os.path.isdir(args.export_json):
            print_color(f"Error: '{args.export_json}' is not a valid directory", 'red')
            sys.exit(1)
        export_library_to_json(args.export_json)

    elif args.export_html:
        if not os.path.isdir(args.export_html):
            print_color(f"Error: '{args.export_html}' is not a valid directory", 'red')
            sys.exit(1)
        export_library_to_html(args.export_html)

    else:
        # Interactive mode
        print_color("Select an option:", 'cyan')
        print_color("  1. Process Movies", 'white')
        print_color("  2. Process TV Shows", 'white')
        print_color("  3. Health Check", 'white')
        print_color("  4. Codec Analysis", 'white')
        print_color("  5. Find Duplicates", 'white')
        print_color("  6. Find Duplicates (Enhanced)", 'white')
        print_color("  7. Export to CSV", 'white')
        print_color("  8. Export to HTML", 'white')
        print_color("  9. Export to JSON", 'white')

        choice = input("\nEnter choice (1-9): ").strip()

        if choice in ['1', '2', '3', '4', '5', '6', '7', '8', '9']:
            path = input("Enter path to media library: ").strip()

            if not os.path.isdir(path):
                print_color(f"Error: '{path}' is not a valid directory", 'red')
                sys.exit(1)

            dry_run_input = input("Enable dry-run mode? (y/N): ").strip().lower()
            config.dry_run = dry_run_input == 'y'

            if choice == '1':
                process_movies(path)
                show_statistics()
            elif choice == '2':
                process_tv_shows(path)
                show_statistics()
            elif choice == '3':
                health_check(path)
            elif choice == '4':
                codec_analysis(path)
            elif choice == '5':
                show_duplicate_report(path)
            elif choice == '6':
                show_enhanced_duplicate_report(path)
            elif choice == '7':
                export_library_to_csv(path)
            elif choice == '8':
                export_library_to_html(path)
            elif choice == '9':
                export_library_to_json(path)
        else:
            print_color("Invalid choice", 'red')
            sys.exit(1)

    print_color(f"\nLog file saved to: {config.log_file}", 'cyan')


if __name__ == '__main__':
    main()
