# LibraryLint

PowerShell toolkit for post-download media library processing. Handles metadata
lookup, file renaming, artwork fetching, SFTP sync from seedbox to local storage,
and duplicate detection. Runs manually from VSCode — not a daemon, not a service.

## Context

LibraryLint sits in the middle of a pipeline that spans two machines:

- **ruTorrent** (on the Ultra.cc seedbox) does the actual downloading. Working
  files live under `/home16/nullpointr/downloads/rtorrent/`, completed
  torrents land in the per-category subfolders under `complete/radarr/` and
  `complete/sonarr/`.
- **Radarr / Sonarr** (on the seedbox) handle search and import. **Completed
  Download Handling is enabled** — *arr renames finished downloads and
  produces a canonical library mirror at `/home16/nullpointr/media/Movies/`
  and `/home16/nullpointr/media/TV Shows/`. *arr is the authoritative
  importer on the seedbox side; LibraryLint should not duplicate that work.
- **LibraryLint** runs on the local Windows PC and: pulls from the seedbox
  library mirror via SFTP, layers richer NFOs and artwork onto each title
  for Kodi-style consumption, deduplicates, organizes into the local
  library, and prunes the seedbox once content is confirmed local.
- Final destination is the HTPC (Dell OptiPlex 3060 running LibreELEC/Kodi)
  via an external drive mounted at `/var/media/`. The HTPC reads NFOs and
  artwork written by LibraryLint.

The split is deliberate: *arr handles the import side from rTorrent into the
seedbox library mirror; LibraryLint handles seedbox → local → Kodi-ready.
Do not suggest moving the import responsibility from *arr into LibraryLint,
and do not suggest disabling *arr's CDH — both are how the seedbox library
mirror gets populated in the first place.

**Namespace note:** the seedbox presents the same filesystem through two
paths. Inside the seedbox's chroot (where *arr, ruTorrent run), paths look
like `/home/nullpointr/...`. Via SFTP (where LibraryLint connects from the
local PC), the same files appear as `/home16/nullpointr/...`. They are not
symlinks or aliases in the usual sense — they are two namespace projections
of one disk, and tools running inside the chroot cannot reach `/home16/`.
This is why Sonarr's Root Folder paths use `/home/` while LibraryLint's
SFTP-side config uses `/home16/`. Both are correct for their respective
namespaces.

The public-facing positioning of this niche (for README, install docs, and
package metadata) lives in the [Who is this for?](README.md#who-is-this-for)
section of `README.md`. When changing any user-visible language about who
LibraryLint serves — tagline, GETTING_STARTED.md intro, `.psd1` Description,
release notes — keep that README section as the source of truth and update it
first. Don't introduce new tagline language elsewhere that contradicts it.

## Tech stack

- **PowerShell 7+** (cross-platform, not Windows PowerShell 5.1)
- **Pester 5** for testing
- **Posh-SSH** (or equivalent) for SFTP operations
- Standard module layout: `LibraryLint.psd1` manifest + `LibraryLint.psm1`
  root module, public functions in `Public/`, private helpers in `Private/`

## Conventions

- Use approved PowerShell verbs (`Get-`, `Set-`, `Invoke-`, `Sync-`, etc.).
  Run `Get-Verb` if unsure.
- Functions use `[CmdletBinding()]` and fully typed parameters.
- Prefer advanced functions over simple ones — always support `-WhatIf` and
  `-Confirm` for anything that writes, renames, deletes, or transfers.
- Verbose variable names. No one-letter names outside of tight loops.
- Comment-based help on every public function: `.SYNOPSIS`, `.DESCRIPTION`,
  `.PARAMETER`, `.EXAMPLE`.
- Inline comments explain **why**, not **what**. The code already says what.
- Pipeline-friendly where it makes sense (`ValueFromPipeline`,
  `ValueFromPipelineByPropertyName`).
- Error handling: `throw` for unrecoverable, `Write-Error` for recoverable,
  never swallow exceptions silently.

## Testing

- All tests live in `Tests/` and use Pester 5 syntax.
- Test files mirror source: `Public/Sync-Library.ps1` → `Tests/Sync-Library.Tests.ps1`.
- Mock external dependencies (SFTP, filesystem writes, network calls). Tests
  must run offline with no seedbox access.
- Run the full suite with `Invoke-Pester` from the repo root.
- Do not write tests that hit the real seedbox or touch `/var/media/`.

## Version bump checklist

When releasing a new version, **all** of the following must be updated:

1. `LibraryLint.ps1` — `$script:AppVersion` (line ~148)
2. `LibraryLint.psd1` — `ModuleVersion`
3. `installer/LibraryLint.iss` — `#define MyAppVersion`
4. `installer/README.md` — version appears in 3 places (build output, command
   example, portable ZIP script)
5. `README.md` — version badge (`img.shields.io/badge/Version-...`)
6. `CHANGELOG.md` — new release section at the top

Also verify that these stay accurate and haven't drifted:

- `README.md` — PowerShell version badge and requirement text
- `GETTING_STARTED.md` — PowerShell version requirement
- `config/config.example.json` — any new config keys added since last release

## Security

- **Never** commit credentials, SFTP keys, API tokens, or seedbox URLs.
- Secrets come from environment variables or a gitignored
  `config.local.psd1` — never from source.
- If you see a hardcoded credential anywhere, stop and flag it.

## What to ask before doing

- Any destructive operation (delete, overwrite, mass rename) — confirm scope
  and show a dry-run first.
- Any change to the module manifest (`.psd1`) version, exports, or dependencies.
- Any new external dependency before adding it to `RequiredModules`.