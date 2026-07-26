# Changelog

All notable changes to LibraryLint will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [5.7.0] - 2026-07-26

### Added
- **Status dashboard (main menu `S`).** One-screen pipeline snapshot: inbox / `_Review` / quarantine contents (with per-item listings), seedbox untracked files, Radarr/Sonarr library health (missing items, health warnings, queue depth), HTPC reachability, and mirror delta via `robocopy /L` pre-flight. Each actionable bucket then offers a Y/N prompt so the whole punch list can be worked from one place. Auto-wakes the HTPC when the mirror destination needs it and shuts it back down afterwards — only if Status woke it.
- **HTPC control module (`modules/Htpc.psm1`).** Wake-on-LAN magic packets (`Send-Wol`), TCP readiness polling (`Wait-Htpc`), `Start-Htpc`, and Kodi JSON-RPC shutdown/suspend/reboot (`Stop-Htpc`, `Invoke-KodiJsonRpc`). New `Htpc*` config keys; menu entries under Utilities (Wake HTPC / Stop HTPC).
- **Sonarr integration.** `SonarrUrl`/`SonarrApiKey` config keys with setup-wizard and Manage-API-Keys flows; missing-episode counts on the Status dashboard; combined **Radarr / Sonarr Status** utility (version, health messages, queue, disk space, missing breakdown).
- **Sonarr Re-acquisition (Utilities 16).** TV counterpart to the Radarr flow: scans the TV library for episodes below a chosen resolution, matches shows against Sonarr's series list, and queues an `EpisodeSearch` so Sonarr hunts for upgrades.
- **Persistent RAR-extraction tracking.** Extracted releases are recorded in `rar_extraction_tracking.json` keyed by content signature (leaf + part count + first archive), so releases no longer re-extract on every pass after the extracted-sync cleanup removes the remote `.extracted/` folder. Also blocks same-run duplicate extraction of hardlinked working-dir/complete copies. Menu offers a re-extract override.
- **Movie-set artwork scan cache.** Per-collection MSIF snapshots (`movie_set_artwork_cache.json`, 30-day TTL, invalidated by on-disk changes) skip the TMDB/Fanart.tv round-trips for already-complete sets — full-library set-artwork passes drop from minutes to seconds.
- **Hardsub Audit utility.** Samples frames via ffmpeg and flags movies likely to have burned-in subtitles (the Deadpool-KORSUB class of bad rips).
- **Quality-concern coverage for sub-HD files.** 480p bitrate thresholds plus a blanket "Sub-HD resolution — re-acquisition candidate" concern; Radarr Re-acquisition gained a "by quality concerns" scan mode.
- **TV episode upgrade contest.** Merging a season now detects the same episode arriving under a different filename, quality-scores both, keeps the winner in the library, and demotes the loser to the inbox (nothing deleted).
- **`<actor>` elements in `tvshow.nfo`.** Kodi cast pages populate from TVDB data; downloaded `.actors/` images are finally referenced.
- **Undo system is live.** `Add-UndoOperation` is now wired into all library moves (movies, TV, upgrades, alternate cuts), duplicate moves, and merge renames/moves — "Undo Previous Session" finds real manifests for the first time.
- **ffsubsync as a managed dependency** with auto-install (including a Python 3.11 fallback when wheel builds fail on newer Pythons); Subtitle Sync Audit and related tooling ship in the Library Tools menu.

### Changed
- **Mirror speed display** now samples the NIC bytes-sent counter (the same source Task Manager reads) instead of parsing robocopy's bursty output — the displayed rate no longer decays toward zero on long copies.
- **Transcode output quieted** (`-hide_banner -loglevel error -stats`): no more per-frame codec warning floods; live progress line retained. Transcode prompt now defaults to "Run now".
- **Transcode replacement is crash-safe**: original renames to `.bak`, the temp swaps in, and the `.bak` restores on any failure — a folder can never be left without a playable file. `Invoke-Transcode` supports `-WhatIf`.
- **SFTP sync and prune survive session death**: one automatic reconnect + retry mid-loop instead of every remaining item failing.
- **TMDB error handling**: outages print a visible warning instead of silently registering as "no match"; all API calls carry timeouts; deleted TMDB collections log one quiet line instead of a JSON dump.
- **`Clean Up LibraryLint Data`** itemizes the tracking files and caches it deletes (SFTP tracking, extraction tracking, collection/set-artwork caches) instead of hiding them in "other files".
- **Test suite tests production code.** The harness AST-extracts main-script functions and imports `Quality.psm1` directly; the five drifted inline copies are gone. 76 tests, all green, fully offline.
- **`config.example.json` rebuilt** around the modern key set (`InboxPath`, Radarr/Sonarr/Trakt, quarantine, MSIF, mirror credentials, SFTP additions) with legacy split-inbox keys documented as such.

### Fixed
- **Span-format episode parsing**: `S01E01-E03` now parses as a multi-episode file covering episodes 1–3 (previously parsed as episode 1 alone with "E03" leaking into the episode title).
- **Merge deletions can no longer destroy data**: `Rename-OrMergeFolder` keeps the source folder when any merge move failed or when video files remain inside it, instead of unconditionally `Remove-Item -Recurse`.
- **Wildcard-unsafe subtitle repair**: folders with `[brackets]` in their names no longer break `Repair-OrphanedSubtitles` listings (which could delete perfectly matched subtitles).
- **Corrupt SFTP tracking file recovery**: an unreadable `sftp_downloaded.json` is backed up to `.corrupt` with a prominent warning instead of being silently replaced by an empty file on the next save.
- **`CheckDuplicates` honored**: the inbox-processing answer now actually triggers a post-transfer duplicate scan (it was previously saved and ignored).
- **Status inbox counting** excludes LibraryLint-managed `_`-prefixed scratch folders, and the mirror pre-flight authenticates SMB with stored credentials (no more false "dest unreachable" against credential-gated shares, with the real error surfaced when it does fail).
- **Unreachable TV NFO-mismatch block removed** — it could never execute, and its fallback would have written a movie NFO into a TV show folder.
- **Sync Extracted to Inbox** resolves the unified `InboxPath` correctly instead of failing on the legacy `MoviesInboxPath`.

### Removed
- ~1,240 lines of dead code: eleven never-called functions (legacy duplicate trio, `Invoke-ArtworkSync`/`Invoke-ArtworkDownload`, legacy metadata fetchers, and others), four phantom module names in the loader, and the unused `EnableParallelProcessing`/`MaxParallelJobs` config keys. Duplicated regexes and scoring chains consolidated into shared helpers.

## [5.6.8] - 2026-05-21

### Added
- **rTorrent torrent removal — full three-phase feature.** LibraryLint can now reach into rTorrent on the seedbox over SSH to clean up torrent entries whose data has been pruned, removing the dead 0%/errored rows that used to accumulate in the rTorrent client.
  - **Phase 1 — SCGI client core.** New `Invoke-RTorrentCommand`, `Get-SeedboxTorrents`, `Remove-SeedboxTorrent` in `modules/Sync.psm1`. The transport is an embedded Python3 SCGI client (sent over the existing SSH channel via WinSCP's `Session.ExecuteCommand`, base64-encoded for clean shell quoting) that speaks XML-RPC-over-SCGI to rTorrent's local Unix socket at `~/.config/rtorrent/socket`. No dependency on ruTorrent's HTTP layer or its auth. Each listed torrent carries hash, name, base_path, completion/state, and a `path_exists` flag computed server-side so dead torrents are identifiable in a single round trip.
  - **Phase 2 — standalone "Clean dead torrents" pass.** New SFTP submenu option (9) that lists every torrent whose `path_exists` is false (data already gone) and erases just those entries. Inherently hit-and-run-safe: you cannot seed data that doesn't exist. Standard dry-run preview + confirm.
  - **Phase 3 — prune-integrated erasure.** New `EraseTorrents` switch on `Invoke-SFTPPrune`. Before deletions begin, snapshots every torrent whose data currently lives under a `RemotePaths` (prune folder) root. After deletions + folder cleanup, re-queries rTorrent and erases any snapshotted torrent whose data has now gone missing — those are exactly the torrents this prune killed. Scoped tightly: working-dir torrents and anything outside the prune folder are never in the snapshot, so the erase can't touch a torrent still seeding within its hit-and-run window. Menu prompts `(Y/N) [Y]` only when not in dry-run mode.
  - New helper `Test-RTorrentPathUnderRoots` collapses the seedbox's two namespace projections (`/home<digits>/` ↔ `/home/`) before path comparisons, so torrents' rTorrent-reported `/home/...` paths match LibraryLint's `/home16/...` configured roots.

### Fixed
- **Movie collection cache was never actually hitting.** The cache file in 5.6.5/5.6.6 was saving fine (NFO mtimes written as ISO 8601 strings) but `Read-MovieCollectionCache`'s `ConvertFrom-Json -AsHashtable` auto-parses ISO 8601 strings into `DateTime` objects. The downstream comparison code computed the current file's mtime as a string and used `-eq` to compare against the cached value — string-vs-DateTime always returns false, so every entry was treated as a cache miss → re-query TMDB. Net effect: a 1500-movie library hit TMDB for every movie on every run, defeating the entire purpose of the cache. Fix in `Read-MovieCollectionCache` re-stringifies `NFOMtime` and `CheckedAt` back to ISO 8601 immediately after loading, so the comparison is apples-to-apples. Verified: cache hits now register correctly on identical re-runs.
- **Prune's Radarr library fetch now retries transient SQLite-busy errors.** Radarr's SQLite is single-writer; background tasks (RSS sync, queue processing, refresh) hold write locks that briefly block reads. The prune's `GET /api/v3/movie` was a bare `Invoke-RestMethod` with no retry, so one transient `database is locked` would throw straight through, dump the entire 30-line SQL exception + stack trace to the console, and fall back to age-only pruning. Now wraps the call in a 2-attempt loop with a 2-second pause between (matching the helper already used in the re-acquisition workflow), and compresses the failure message to a single line — `Radarr DB busy (SQLite lock contention) after retry` for the common case, first line of the exception otherwise. Also clarifies the fallback wording: only prune-folder files lose Radarr verification; library-mirror files use the local-verified path and are unaffected.
- **Mirror progress display now adapts to terminal resize.** The fixed `PadRight(120)` was wider than narrow terminals — each progress write wrapped to multiple visual rows, and `\r` only returned cursor to the last wrapped row. When resized mid-job, subsequent writes left dozens of orphan copies scattered across the screen. Now reads `[Console]::BufferWidth - 1` on every tick, pads the visible portion to current width, and truncates rather than wraps if the line is too long. Future resizes don't accumulate orphans.
- **Mirror speed display falls back to cumulative average instead of "warming up" mid-transfer.** Robocopy outputs in bursts and the rolling 5-second byte counter goes flat between bursts (no new completion lines parsed, even though the network transfer is continuous). Previously displayed `warming up` during these gaps, which was misleading mid-job. Now falls back to `xxx MB/s (avg)` (cumulative average over the whole folder) when the rolling window is flat but we have meaningful elapsed time and bytes. The `(avg)` tag distinguishes it from a live measurement. `warming up` only appears in the first 5 seconds, before any meaningful average exists.

### Improved
- **yt-dlp trailer download surfaces an actionable hint on age-gate failures.** When a trailer download fails with `Sign in to confirm your age` (or yt-dlp's `cookies-from-browser` suggestion), LibraryLint detects the signature in stderr and prints a labeled HINT after the failure block — branched on whether `YtDlpCookieBrowser` is already configured. If unset: "set YtDlpCookieBrowser in your config to your browser (firefox/chrome/edge/...) and be logged into YouTube there." If set but the auth still failed: "browser must be closed (cookie DB lock), be logged in, same Windows user (Chromium browsers encrypt cookies per-user)." yt-dlp's stderr is captured to a temp file and parsed for the age-gate signature; on non-age-gate failures, the first line is logged at DEBUG without spamming the console.

## [5.6.7] - 2026-05-14

### Changed
- **Positioning correction: CDH is enabled.** Earlier docs framed LibraryLint as serving an *arr-with-Completed-Download-Handling-disabled workflow. In practice, the *arr stack on the seedbox uses CDH to rename and import finished downloads into the canonical library mirror at `media/Movies/` and `media/TV Shows/`; LibraryLint's role is the seedbox-to-local-PC pull, Kodi-side enrichment (NFOs, artwork, dedup), and seedbox cleanup once content lands locally. README's pipeline diagram, "Fit" and "Not a fit" bullets, `GETTING_STARTED.md`'s pre-install callout, `CLAUDE.md`'s Context section, and `LibraryLint.psd1`'s Description field all updated to match.
- **CLAUDE.md namespace note** added — documents that the seedbox presents the same filesystem through two namespace projections (`/home/...` inside the chroot where *arr/ruTorrent run; `/home16/...` via SFTP where LibraryLint connects), and that these are not symlinks but separate views. This informs why Sonarr's Root Folder paths use `/home/` while LibraryLint's SFTP-side config uses `/home16/` — both are correct for their respective tools.

### Added
- **Dark-mode README logo.** New `assets/logo-dark.png` (RGB-inverted variant of the original — preserves hue, flips luminance, so dark outlines/text become light and cream highlights become navy). README now wraps the logo in an HTML `<picture>` element with `prefers-color-scheme` media queries so GitHub serves the right asset per theme. Auto-generated; can be replaced with a hand-crafted dark variant later without further README changes.

## [5.6.6] - 2026-05-14

### Added
- **Auto-discovery of untracked seedbox library files** — `Invoke-SFTPPrune` now optionally walks `LibraryPaths` on the seedbox and matches untracked files to local libraries by folder name + main-video file size, then adds them to tracking as `AutoDiscovered=true` entries so the local-verified prune path can clean them. Solves the symptom where files that existed on the seedbox before LibraryLint started syncing (or were uploaded by Radarr directly) were invisible to the prune. New `LocalLibraryPaths` parameter on `Invoke-SFTPPrune`; wired in the menu from `MoviesLibraryPath` + `TVShowsLibraryPath`.
- **Working-dir prune now matches against local libraries too** — `Invoke-SFTPPruneWorkingDir` previously only checked against the seedbox library mirror. With aggressive local-verified pruning emptying that mirror, the working-dir prune had nothing to match against and reported `With library size match: 0`. Now also builds a local-library size index, and a working folder is eligible if EITHER the seedbox or the local library has a primary-video size hit. New `LocalLibraryPaths` parameter.
- **Loose-video promotion in working-dir prune** — videos sitting directly inside category-named folders (like `/.../rtorrent/movie.mkv` where the parent is on the category-container blacklist) were silently skipped along with the parent. Now each video is promoted to its own candidate; the parent folder is still protected from deletion, but individual files get matched and deleted with a clean `RemoveFiles` against just the file path.
- **Age-fallback confirmation prompt** — when the regular prune finds library files with no local copy but old enough to fall through the age threshold, surfaces them in a yellow warning block (oldest first, with each entry's last-known LocalPath) and requires explicit `Y` confirmation before deletion. Default-N so a fat-fingered Enter doesn't delete files you haven't synced yet. LocalVerified and Age-tier files still delete without the prompt.
- **Multi-part RAR completeness pre-check** — `Get-RarReleaseInfo` now flags incomplete sets via two cheap checks: zero-byte parts (sparse-allocated paused downloads) and sequential numbering gaps in `.r##`/`.partN.rar` chains. Incomplete releases are listed separately with their reason and skipped from extraction, avoiding wall-of-failure noise from unrar.

### Improved
- **Mirror progress speed display: rolling 5s window** — replaced the EMA with a true rolling-window calculation. The EMA over the 200ms display tick was decaying toward zero between robocopy's bursty output, causing displays like `18 bytes/s | ETA 572166h 43m` during active 300 MB/s transfers. The rolling window absorbs bursts uniformly and converges on the actual throughput within ~5s.
- **Mirror summary counters correct after long runs** — async `OutputDataReceived` events can fire after `WaitForExit()` returns; the dequeue loop's last `IsEmpty` check could exit before robocopy's `Files:`/`Bytes:` summary lines were delivered, leaving the parser with nothing and reporting `Copied: 0` after multi-hour transfers. Added a final queue-drain after `WaitForExit()` (the documented signal that async handlers have flushed) so the summary always makes it through.
- **Mirror per-folder header counters no longer leak across iterations** — `$folderTotalFiles` was checked before being assigned each iteration, so iteration 1 worked but iteration 2 inherited iteration 1's value ("Shows" displayed Movies' 2,494 files / 0 bytes). Moved the assignment ahead of the print.
- **Mirror folder name spacing in scan output** — `Write-Host "`r  $folder".PadRight(20)` counted the `\r` toward the 20-char width; folder names of length ≥17 (like "Movie Set Artwork") got no separator before the count ("Movie Set Artwork5 to copy"). Now pads the visible portion only.
- **Working-dir prune "Too recent" list is now oldest-first** — sorted by `NewestMtime` ascending so items closest to becoming prune-eligible appear at the top of the list, not items most recently downloaded.
- **Regular prune menu prompt rewritten** — the old prompt said "delete old files from the download directory" with one path. The local-verified path made that wording obsolete; the prompt now lists both eligibility paths (prune folder by age, library mirror by local match) and the Radarr-verification description notes it only applies to prune-folder files (library files with a verified local copy are already proven).

### Fixed
- **Working-dir prune finding 0 library matches** — root cause was that aggressive seedbox-library pruning (via 5.6.5's local-verified path) had emptied the library mirror that working-dir prune was using as its sole match source. The new local-library indexing solves the case structurally; the user-visible symptom (everything in working/ sitting indefinitely) is gone as soon as local libraries are configured.
- **Untracked seedbox library files were invisible to prune** — the prune iterated only the tracking file (`$downloaded`), so files that existed on the seedbox before LibraryLint started syncing (or arrived via other paths) were never considered. The new auto-discovery scan fixes this without permanent corruption to the tracking file — entries are written through the standard save path so the second run already sees them as tracked.

### Changed
- **README positioning** — replaced the generic "Kodi/Plex/Jellyfin-compatible" tagline with explicit niche positioning. New `Who is this for?` section near the top with a workflow diagram (Seedbox → Local PC → HTPC), `Fit` and `Not a fit` bullet lists, and a `What's swappable` callout. The codebase is unchanged; the positioning is honest about the specific workflow LibraryLint optimizes for (Kodi-style library + remote seedbox + *arr CDH disabled) rather than trying to appear universal.
- **Positioning aligned across all docs** — `GETTING_STARTED.md` opens with a callout pointing users to the README's `Who is this for?` section before they install (so users who would be wasting their time bail out before configuring). `CLAUDE.md` notes the README section as the canonical positioning source for future AI contributors. `LibraryLint.psd1` `Description` field rewritten to match.

## [5.6.5] - 2026-05-13

### Added
- **Multi-part RAR extraction on the seedbox (TEST)** — new SFTP submenu option 8 detects `.rar` / `.r##` / `.partN.rar` chains in `SFTPPrunePaths` ∪ `SFTPWorkingPaths` and runs `unrar` remotely via WinSCP's SSH exec channel. Output goes to a sibling `<release>.extracted/` folder so the seeding torrent files are untouched. New public functions: `Get-RarReleaseInfo`, `Find-SFTPRarReleases`, `Invoke-SFTPRemoteUnrar`, `Invoke-SFTPExtractRarReleases`, `Invoke-RadarrDownloadedScan`. New config keys: `SFTPUnrarCommand`, `SFTPExtractedSuffix`. After a successful extraction, optionally POSTs Radarr's `DownloadedMoviesScan` command so the queue item self-clears.
- **Movie set artwork local cache** — TMDB collection lookups for each library folder are now cached at `%LOCALAPPDATA%\LibraryLint\movie_collection_cache.json`. Re-runs skip both the NFO read and the TMDB API call for folders whose NFO mtime hasn't changed (90-day TTL); negative results (no TMDB ID, not in a collection) are cached too. New `-ForceRefresh` parameter on `Invoke-MovieSetArtworkDownload` to bypass the cache. New public helpers: `Get-MovieCollectionCachePath`, `Read-MovieCollectionCache`, `Save-MovieCollectionCache`.
- **Working-dir prune age guard** — `Invoke-SFTPPruneWorkingDir` now accepts a `DaysOld` parameter (default 14) that skips releases whose newest file mtime is within the threshold. Prevents nuking still-seeding fresh downloads just because their library counterpart was already synced. New config key: `SFTPWorkingPruneDaysOld`.
- **Test-TMDBHealth diagnostic** — new exported function in `modules/TMDB.psm1` that distinguishes environmental TMDB failures (bad/revoked key, 429 rate limit, 5xx, network) from per-title parse issues. Probes both the configuration and search endpoints with a known-good query (Inception 2010) and returns a structured result (`Healthy`, `KeyConfigured`, `ConfigEndpointOk`, `SearchEndpointOk`, `Reason`, `SampleHit`).

### Improved
- **Mirror progress (robocopy) — async output + size-poll fallback** — `Invoke-Mirror`'s output loop replaced `StandardOutput.ReadLine()` with `OutputDataReceived` into a `ConcurrentQueue`, so the loop no longer blocks waiting for robocopy to emit a line. Between lines, the destination file's on-disk size is polled to compute in-flight percentage — the progress bar now ticks smoothly during multi-GB copies instead of staying frozen at the size announcement and only jumping when the file completes. Event subscription cleanup added in `finally` to avoid accumulating leaked subscriptions across mirror runs.
- **Inbox-to-library quality compare picks the main feature** — was using `Select-Object -First 1` after `Get-ChildItem` (alphabetical), letting trailers/samples gate import decisions. Now sorts by `Length -Descending` and excludes `-trailer/.trailer/_trailer/sample/featurette/extras/behindthescenes/deleted scenes` patterns. Also surfaces the comparison details on `[EXISTS - skipped]` (previously only on `[UPGRADE AVAILABLE]`) so wrong skip decisions are visible. Same fix applied to the quarantine-restore flow.
- **Codec analysis report shows filename** — the Lowest Quality Files report now prints the actual filename under the folder name when they differ, so it's obvious when a flagged folder's "low quality" is from a trailer/sample/extras file rather than the main movie.
- **Strengthened trailer/sample filter** in both codec analysis and re-acquisition workflows — catches dot-separated (`.trailer.`), underscore-separated, and bare (`trailer.mkv`) variants; also excludes content under `Trailers/`, `Samples/`, `Extras/`, `Featurettes/`, `Behind The Scenes/`, `Deleted Scenes/`, `Bonus/` subfolders that the recursive scan walks into.
- **Robust unrar probing** — uses `command -v` (POSIX shell builtin) instead of `unrar -?` (which writes help to stderr and exits non-zero, making both Output and ExitStatus unreliable). Falls back to common absolute paths (`/usr/bin/unrar`, `/usr/local/bin/unrar`, `~/bin/unrar`, etc.) when not on PATH. Whichever resolves becomes the binary used for the rest of the run.
- **Reconnect between extractions** — each remote unrar call disposes its WinSCP session and opens a fresh one before the next, avoiding the "Element session@0 already read to the end" IPC corruption that follows long-running `ExecuteCommand` calls. Success is now verified by listing the destination for a video file (more reliable than exit status, which WinSCP doesn't always receive).
- **Unrar switches** — replaced `-inul` (silent-everything) with `-idq -o+ -y` so error messages surface while progress noise is still suppressed, and missing-volume prompts fail fast.
- **Dedup overlapping RAR scan paths** — when `SFTPPrunePaths` and `SFTPWorkingPaths` both root inside the same physical directory (Ultra.cc's `/home` and `/home16` mount aliases, etc.), the same release was scanned and extracted twice. Now collapses by basename + part count + first-archive name, keeping the shortest folder path.
- **Empty-folder cascade in regular SFTP prune** — only the immediate parent of each delete-group was being checked for emptiness, with no deepest-first sort or grandparent cascade. Mirrors the sync's queue-based pattern: collect touched parents, sort deepest-first, push grandparents on each successful removal, bounded by configured roots.

### Fixed
- **SFTP prune missed the library mirror** — `Invoke-SFTPPrune` was passed only `SFTPPrunePaths` as valid roots. Tracked paths under `SFTPRemotePaths` (Radarr's renamed library) didn't match, the remap fallback then silently rewrote them onto a non-existent twin under the prune root, and `FileExists` returned false so they were classified as "already gone" — clearing the tracking entry while the real file stayed on the seedbox forever. New `LibraryPaths` parameter accepts both sets of roots; tracked paths under either are pruned in place. Processed release folders no longer pile up under `/media/Movies/` after pruning.
- **Re-acquisition resolution filter was score-based** — "Below 720p" was filtering on `Score < 60`, but a 480p BluRay x264 5.1 file easily scores past 60 (40 base + 15 codec + 30 source + audio). 480p BluRays at scores 60–95 were slipping past the "low quality" filter despite being clearly below 720p. Now filters on the parsed resolution height directly (`< 720`, etc.), matching the codec analysis classification and the user's mental model.



### Improved
- **Radarr re-acquisition throttle, retry, and breather** — the per-movie loop in option 11 now sleeps 250ms between iterations, takes a 5-second breather every 25 movies, and wraps every Radarr API call (lookup-by-TMDB, lookup-by-title, PUT update, POST search, POST add, plus the pre-loop library fetch) in a transient-error retry helper. Catches "database is locked" / 429 / 5xx / network errors with one automatic retry after a 2s pause; non-transient errors still propagate immediately so "already exists" classification is preserved. Prevents Radarr SQLite single-writer contention under burst load (a 100-movie re-acquisition was reliably DB-locking before).
- **Resolution bucketing in Get-QualityScore now uses max(width, height)** — wide-aspect releases (2.39:1, 2.76:1, etc.) are no longer downgraded by the height-only check. Example: The Creator (2023) at 1920×696 (2.76:1, Ultra Panavision-style framing) was being classified as 480p (+40); now correctly 1080p (+80). Same fix catches 2.39:1 anamorphic 1080p (1920×800), and the symmetric height-based check still catches rare 4:3 1080p (1440×1080).

### Fixed
- **SFTP `-Force` flag was too coarse** — the "Re-download files even if already present locally?" prompt was bypassing every skip tier including the tracking-file filter and byte-identical name+size match, which made entire libraries appear new on a re-acquisition run. `-Force` now only suppresses the folder-name match (the imprecise heuristic that false-positives on quality upgrades). Tracking-file matches and name+size matches always apply, so files we literally already have stay skipped.

## [5.6.3] - 2026-04-22

### Added
- **Restore-SubtitleBackups handles lone .bak files** — previously only chained `.srt.bak` / `.sub.bak` etc. patterns were detected; now also restores release-included `Movie.bak` backups by SRT-content sniffing and resolving the single subtitle sibling in the folder. Skips ambiguous cases (multiple sibling subs) with a clear reason instead of guessing.
- **`-KeepBackup` switch on Restore-SubtitleBackups** — preserves `.bak` files after restore as permanent reference. Default still removes them, matching prior behavior.
- **Trailer rename catch-all in Rename-VideoToMatchFolder** — finds files matching `*-trailer.<videoExt>$` and renames the single trailer in a folder to `<expectedBaseName>-trailer.<ext>` when the basename diverges. Catches release-named trailers (e.g. `Ong-Bak The Thai Warrior (2003) ... -trailer.mp4`) that didn't share basename history with the video and so escaped the existing startswith-current-basename / orphan-by-old-basename heuristics.
- **Folder/NFO title-mismatch safety in Rename-VideoToMatchFolder** — refuses to propagate a broken folder name onto files when the folder's title portion clearly differs from the NFO `<title>`. Loose comparison (lowercase, alphanumerics-only), tolerates prefix matches in either direction. Reports skipped folders at the end with their NFO title and a route to "Fix Folder Names first." Prevents the historic Greek/Ong-Bak failure mode where a previously-mangled folder name was being faithfully copied into otherwise-correct files.

### Improved
- **Restore-SubtitleBackups output** — per-file `[would restore -> NewName]` / `[restored -> NewName]` / `[skip: reason]` preview matches the orphan-fix convention. Single consolidated summary at the end.
- **Restore Subtitle Backups menu (option 11)** — dry-run path now previews then prompts to apply, mirroring the option 12/13 pattern. Auto-printed Format-Table from the function's return value suppressed.
- **Fix File Names menu (option 6)** — captures `Rename-VideoToMatchFolder`'s return count, suppresses the Format-Table that was leaking to the host, and only prompts "Apply these changes?" when dry-run actually found changes pending. Skips the prompt entirely on a clean run.

### Fixed
- **Bracket-path sweep — round 3** — fixed ~21 more wildcard-vulnerable cmdlet sites across `LibraryLint.ps1` (duplicate-handling Move/Remove/Get-ChildItem, Remove-UnnecessaryFiles cleanup paths, subtitle delete, trailer move, empty-folder delete, archive cleanup, file rename, loose-file move-up, log file readers, and `Invoke-QualityScore`'s file-existence check), `modules/Quality.psm1` (Test-Path / Get-Item in two `Get-QualityScore` paths), `modules/Subtitles.psm1` (SubDL download Copy-Item, orphan delete x2, orphan rename), and `modules/Sync.psm1` (WinSCP DLL install Copy-Item). Folders like `[REC] (2007)` now process correctly through deletion, rename, quality scoring, and subtitle-orphan repair flows.

## [5.6.2] - 2026-04-21

### Added
- **Update-SFTPTrackingFromLocal** — new SFTP submenu option 7 walks the seedbox and local library, then marks remote files you already have locally (by name+size or release folder match) as already-downloaded so future syncs skip them
- **LibraryPaths on Invoke-SFTPSync** — skip-detection index now walks both the inbox and the configured MoviesLibraryPath/TVShowsLibraryPath, catching files that were long-since moved out of inbox
- **Quarantine collision handling** — Restore-FromQuarantine now mirrors the inbox→library chooser pattern: per-folder quality comparison, Y/N/Review batch prompts for upgrades and redundants, TMDB-ID-aware collision detection, bounded by configured library roots
- **SFTP empty-folder cleanup** — sync (delete-after-download) and prune now remove release folders on the seedbox after their files are deleted, cascading up to parent folders, bounded by configured roots
- **Legacy codec sidecar cleanup** — Remove-UnnecessaryFiles sweeps deprecated per-folder codec-info.json files during movie processing; also available as an action in the Health Check Phase 3 results

### Improved
- **TMDB scorer — compact-form match** — "War Games" now matches "WarGames" on TMDB and "Winters Bone" matches "Winter's Bone"; whitespace-collapsed title comparison bridges compound-vs-split words and apostrophe gaps
- **TMDB scorer — vote-count-aware ranking** — year-exact bonus now requires vote_count ≥ 5, preventing zero-vote obscure candidates from outscoring famous off-by-one films; popularity tiebreaker replaced with log-scaled vote_count bonus
- **NFO gate — TMDB or IMDB** — folders with only a TMDB ID (typical for foreign/obscure films with no IMDB entry) now pass the gate instead of being quarantined
- **Mirror to Backup diagnostics** — net use output captured on failure, "Connected" only prints on actual exit 0, UNC-specific troubleshooting block lists remote-drive-offline / credential / reachability causes with ready-to-paste diagnostic commands and a bare-username warning
- **Subtitle repair output** — three overlapping summaries collapsed into one consolidated block, consistent `[would X]` / `[X]` per-file action preview, auto-printed Format-Table suppressed at all 7 call sites

### Fixed
- **Video rename post-NFO** — Rename-VideoToMatchFolder now always runs after Invoke-NFOGeneration (was only conditional on a year-fix), so fresh imports with canonical folder names get the rename + release-info.json instead of being skipped because no NFO existed at Step 8e
- **Bracket-path sweep** — 15 wildcard-vulnerable Test-Path / Get-Content / New-MovieNFOFromTMDB sites converted to -LiteralPath; folders like `[REC] (2007)` now process correctly end-to-end (NFO read, NFO write, metadata refresh, Read-NFOFile, Invoke-MetadataRefresh passes 1 and 2, etc.)
- **Duplicate parameter in Fix Subtitles dry-run** — LibraryLint.ps1 Option 12 passed `-VideoExtensions`, `-SubtitleExtensions`, `-PreferredSubtitleLanguages`, `-KeepSubtitles` twice, erroring at runtime

## [5.6.1] - 2026-04-10

### Added
- **Centralized codec cache** — codec analysis results stored in a single `codec-cache.json` in AppData instead of per-folder sidecar files; includes cache versioning for automatic invalidation
- **Radarr-verified SFTP pruning** — prune operation cross-references Radarr's library to confirm movies were imported before deleting from seedbox
- **Separate SFTP prune paths** — new `SFTPPrunePaths` config option separates "sync from" and "prune from" locations on the seedbox
- **TMDB ID duplicate detection** — library transfer catches duplicates by comparing TMDB IDs in NFO files, even when folder names differ
- **NFO pre-check on transfer** — movies with missing or broken NFOs are held back from library transfer until metadata is fixed
- **NFO-only metadata refresh** — health check NFO repairs skip artwork/trailer/subtitle re-downloads
- **Single-folder detection** — Fix & Repair tools detect when a single movie folder is selected instead of a library root and offer to handle it
- **Sidecar cleanup function** — `Remove-CodecSidecarFiles` cleans up old per-folder `codec-info.json` files

### Improved
- **TMDB title matching** — uses Jaccard similarity for longer titles and requires exact match for short titles (1-2 words) to reduce false positives
- **SFTP new file display** — groups files by movie folder instead of by date
- **winget dependency upgrades** — added `--force` flag to handle stalled upgrades

### Fixed
- **LiteralPath in folder creation** — `New-FoldersForLooseFiles` now uses `-LiteralPath` for folders with brackets or special characters

### Changed
- **Module manifest** — real GUID, PowerShell 7.0 minimum, all 5 modules declared in NestedModules and FileList
- **AI context** — replaced outdated `.github/copilot-instructions.md` with `CLAUDE.md`; added Claude Code project settings and custom commands
- **Python stub** — header updated from "EXPERIMENTAL" to "STUB" with clearer call-to-action
- **Documentation** — updated PowerShell requirement to 7+, version badges, installer version, module documentation for Quality/Subtitles/TMDB, config example with SFTPPrunePaths

## [5.6.0] - 2026-04-06

### Added
- **Merged Library Tools menu** — combined Fix & Repair and Enhancements into a single organized menu with Health, Metadata, Media, Cleanup, and Analysis sections
- **Radarr re-acquisition utility** — scan library for low-quality movies or import from CSV/Trakt lists, add to Radarr with automatic searching; re-monitors existing movies and triggers upgrades
- **Trakt public list integration** — import movie lists from Trakt for Radarr re-acquisition using public API (no OAuth required)
- **Seedbox vs Library comparison** — compare seedbox contents against local library to find movies not yet transferred, with optional download of missing items
- **Export Movie List utility** — generates a text file of all movie titles and years from the library
- **Foreign Language Check utility** — scans NFO files and flags movies not in the user's preferred language, grouped by language
- **Export Debug Log utility** — creates sanitized log files for troubleshooting (strips usernames, paths, API keys, hostnames, IPs)
- **Codec analysis force rescan** — option to ignore cached codec-info.json files and re-analyze with MediaInfo
- **Network share authentication** — mirror backup supports authenticated UNC paths with saved credentials
- **Radarr and Trakt API keys** in Settings > Manage API Keys with setup instructions
- **SFTP first-sync detection** — shows "Remote content" instead of misleading "New content" labels when no download history exists
- **SFTP sync date breakdown** — shows already-synced files grouped by download date, and new files grouped by remote file modification date
- **Auto-move prompt option** — inbox processing now offers Yes/No/Prompt for library transfer instead of just Yes/No
- **Duplicate handling on rename collision** — when folder rename targets an existing folder, compares quality scores and prompts to delete the lower-quality version

### Improved
- **WinSCP .NET compatibility** — auto-downloads netstandard2.0 assembly from NuGet when only .NET Framework version is installed; skips .NET version check for .NET 10+ compatibility; reuses loaded assemblies across multiple SFTP operations
- **NFO title matching** — normalizes `&` to `and` before comparison (fixes "Bill & Ted's" vs "Bill and Teds" mismatches)
- **Language tag protection** — `Rename-CleanFolderNames` now skips language tags (ENGLISH, ITALIAN, etc.) preventing "The English Patient" from being renamed to "The"
- **Metadata refresh performance** — pre-scans library to identify folders needing work, skips healthy NFOs silently instead of printing 1300+ "skipping" lines
- **SFTP incomplete file detection** — filters out trailers, samples, and non-media files (NFO, SRT, JPG) from the suspiciously-small and zero-byte checks
- **SFTP prune path remapping** — handles tracking file paths that no longer match current remote path config (e.g., after fixing double-nested paths)
- **SFTP category routing** — recognizes `radarr` and `sonarr` as category folder names for proper routing to `_Movies`/`_Shows`
- **Mirror backup first-run** — defaults to unconfigured state instead of hardcoded drive letters; auto-redirects to setup when running without configuration
- **Comprehensive release tag list** — added ~100 new tags including streaming services (AMZN, NF, DSNP, HMAX, ATVP, etc.), audio formats (EAC3, DTS-X, FLAC, Opus, 8CH), HDR variants (DV, HLG, 12bit), source types (CAM, SCREENER, REMUX), and 50+ common release groups
- **Subtitle download marked experimental** — all subtitle auto-download features clearly labeled with warnings about accuracy

### Fixed
- **MediaInfo winget package ID** — corrected from `MediaArea.MediaInfo.CLI` to `MediaArea.MediaInfo`
- **PowerShell `-replace` error** — removed invalid third argument `'IgnoreCase'` in TV show episode title parsing that caused "allows only two elements" error
- **Folder count mismatch** — "Folders to fix" count now matches confirmation prompt by filtering no-op renames before display
- **Mirror setup prompts** — first-run shows contextual prompts ("Enter path" with examples) instead of fake defaults like "G:\"
- **Codec analysis triple header** — removed duplicate "=== CODEC ANALYSIS ===" banner
- **Standalone hyphen cleanup** — removes separator hyphens ("Mission - Impossible" → "Mission Impossible") while preserving compound words ("Spider-Man")
- **Scratch file cleanup** — removed 6 leftover debug/test files and added `_*.ps1` to .gitignore

## [5.5.2] - 2026-03-26

### Added
- **Loose file auto-wrapping** — inbox processor automatically wraps loose video files in `_Movies`/`_Shows` containers into individual folders before scanning
- **Mismatched trailer detection** in health check — finds trailers whose filename doesn't match the movie file, with auto-rename fix action
- **Incomplete NFO detection** in health check — flags NFOs missing required metadata fields (title, year, TMDB/IMDB IDs, plot, rating, premiered, genre, director, actors)

### Improved
- **Title comparison** — added compact normalization (strips all punctuation) for matching titles with apostrophes ("Pan's" vs "Pans"), periods ("E.T." vs "ET"), and hyphens ("Spider-Man" vs "Spiderman")
- **Roman numeral matching** — title comparison normalizes roman numerals to arabic ("Gladiator II" matches "Gladiator 2")
- **Year tolerance** — high-confidence title matches now allow up to ±3 year difference (handles folders with slightly wrong years like "Glengarry Glen Ross (1990)" → TMDB 1992)
- **Language tag safety** — full language names (ENGLISH, ITALIAN, FRENCH, etc.) now require non-space separators to strip, protecting titles like "The English Patient" and "The Italian Job"
- **SFTP sync category routing** — files from remote `movies/` or `tv/` directories now route correctly to `_Movies`/`_Shows` regardless of file size, using the remote directory structure as a classification hint
- **Mirror cleanup** — robocopy's direct console output ("Removed X of Y") is now properly cleared after each folder sync

### Fixed
- **Loose files misrouted in sync** — files in remote category directories (e.g., `movies/file.srt`) no longer end up in `_Downloads/movies` instead of `_Movies`
- **Category directory nesting** — SFTP sync strips known category directory names (`movies`, `tv`, `shows`, etc.) from all path depths, not just 3+ part paths
- **Duplicate parameter error** — `Repair-SubtitlePlacement` and `Repair-OrphanedSubtitles` no longer pass `-VideoExtensions`/`-SubtitleExtensions` twice in dry-run mode
- **Title normalization punctuation** — `Normalize-Title` replaces punctuation with spaces instead of removing it, so "L.A." normalizes to "L A" (not "LA") matching folder names correctly
- **MediaInfo JSON parsing** — handles empty-key properties in MediaInfo JSON output that caused `ConvertFrom-Json` failures

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
