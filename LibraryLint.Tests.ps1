#Requires -Modules Pester

<#
.SYNOPSIS
    Pester unit tests for LibraryLint.ps1

.DESCRIPTION
    Test suite for LibraryLint functionality including:
    - Quality scoring (modules\Quality.psm1)
    - Episode parsing
    - Title normalization
    - Statistics tracking
    - File matching patterns

    All tests run offline: no TMDB, SFTP, or seedbox access, and module
    functions are called filename-only so MediaInfo probing is skipped.

.NOTES
    Run with: Invoke-Pester -Path .\LibraryLint.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # ----------------------------------------------------------------------
    # Test fixtures: minimal $script:Config / $script:Stats stubs.
    # Production functions read these through $script: scope at CALL time
    # (e.g. Get-NormalizedTitle reads Config.Tags, Write-Log reads
    # Config.LogFile and appends to Stats.ErrorDetails), so the stubs must
    # exist before any extracted function executes. Keep the keys aligned
    # with what production expects.
    # ----------------------------------------------------------------------
    $script:Config = @{
        LogFile = Join-Path $TestDrive "test.log"
        DryRun = $true
        SevenZipPath = "C:\Program Files\7-Zip\7z.exe"
        VideoExtensions = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.m4v')
        SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt')
        ArchiveExtensions = @('*.rar', '*.zip', '*.7z', '*.tar', '*.gz', '*.bz2')
        UnnecessaryPatterns = @('*Sample*', '*Proof*', '*Screens*')
        TrailerPatterns = @('*Trailer*', '*trailer*', '*TRAILER*', '*Teaser*', '*teaser*')
        PreferredSubtitleLanguages = @('eng', 'en', 'english')
        KeepSubtitles = $true
        KeepTrailers = $true
        GenerateNFO = $false
        OrganizeSeasons = $true
        RenameEpisodes = $false
        CheckDuplicates = $false
        TMDBApiKey = $null
        Tags = @('1080p', '2160p', '720p', '480p', '4K', 'HDRip', 'BluRay', 'x264', 'x265', 'HEVC')
    }

    $script:Stats = @{
        StartTime = Get-Date
        EndTime = $null
        FilesDeleted = 0
        BytesDeleted = 0
        ArchivesExtracted = 0
        ArchivesFailed = 0
        FoldersCreated = 0
        FoldersRenamed = 0
        FilesMoved = 0
        EmptyFoldersRemoved = 0
        SubtitlesProcessed = 0
        SubtitlesDeleted = 0
        TrailersMoved = 0
        NFOFilesCreated = 0
        NFOFilesRead = 0
        Errors = 0
        ErrorDetails = @()
        Warnings = 0
    }

    # ----------------------------------------------------------------------
    # Bring PRODUCTION main-script functions into scope.
    #
    # LibraryLint.ps1 cannot be dot-sourced whole — it launches an
    # interactive menu loop at the bottom. Instead, parse the script with
    # the PowerShell AST, locate each named FunctionDefinitionAst, and
    # define it in this test scope by Invoke-Expression of its exact source
    # text. That way the tests always exercise the real production code —
    # never an inline copy that can drift.
    #
    # To put another main-script function under test, add its name to
    # $mainScriptFunctionsUnderTest below. Module functions are NOT
    # extracted this way — import their .psm1 directly (see Quality.psm1
    # import further down).
    # ----------------------------------------------------------------------
    $scriptPath = Join-Path $PSScriptRoot "LibraryLint.ps1"
    $parseTokens = $null
    $parseErrors = $null
    $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath, [ref]$parseTokens, [ref]$parseErrors)

    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "LibraryLint.ps1 has $($parseErrors.Count) parse error(s); first: $($parseErrors[0].Message)"
    }

    $mainScriptFunctionsUnderTest = @(
        'Format-FileSize'
        'Get-EpisodeInfo'
        'Get-NormalizedTitle'
        'Write-Log'
    )

    $allFunctionAsts = $scriptAst.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
    }, $true)

    foreach ($functionName in $mainScriptFunctionsUnderTest) {
        $functionAst = $allFunctionAsts |
            Where-Object { $_.Name -eq $functionName } |
            Select-Object -First 1
        if (-not $functionAst) {
            throw "Function '$functionName' not found in LibraryLint.ps1 — was it renamed or moved to a module? Update the test harness."
        }
        Invoke-Expression $functionAst.Extent.Text
    }

    # ----------------------------------------------------------------------
    # Module functions: import the real module.
    #
    # Get-QualityScore lives in modules\Quality.psm1. Tests call it with a
    # filename only (no -FilePath / -MediaInfo), so the MediaInfo branch is
    # skipped and the suite stays offline. Quality.psm1 does not read
    # $script:Config, so no extra module-scope setup is needed.
    # ----------------------------------------------------------------------
    Import-Module (Join-Path $PSScriptRoot "modules\Quality.psm1") -Force
}

Describe "Format-FileSize" {
    It "Formats bytes correctly" {
        Format-FileSize 500 | Should -Be "500 bytes"
    }

    It "Formats kilobytes correctly" {
        Format-FileSize 1024 | Should -Be "1.00 KB"
        Format-FileSize 2048 | Should -Be "2.00 KB"
    }

    It "Formats megabytes correctly" {
        Format-FileSize (1024 * 1024) | Should -Be "1.00 MB"
        Format-FileSize (500 * 1024 * 1024) | Should -Be "500.00 MB"
    }

    It "Formats gigabytes correctly" {
        Format-FileSize (1024 * 1024 * 1024) | Should -Be "1.00 GB"
        Format-FileSize (4.5 * 1024 * 1024 * 1024) | Should -Be "4.50 GB"
    }

    It "Formats terabytes correctly" {
        Format-FileSize (1024 * 1024 * 1024 * 1024) | Should -Be "1.00 TB"
    }
}

Describe "Get-QualityScore" {
    Context "Resolution Detection" {
        It "Detects 2160p/4K resolution" {
            $result = Get-QualityScore "Movie.2160p.BluRay.mkv"
            $result.Resolution | Should -Be "2160p"
            $result.Score | Should -BeGreaterOrEqual 100
        }

        It "Detects 1080p resolution" {
            $result = Get-QualityScore "Movie.1080p.WEB-DL.mkv"
            $result.Resolution | Should -Be "1080p"
            $result.Score | Should -BeGreaterOrEqual 80
        }

        It "Detects 720p resolution" {
            $result = Get-QualityScore "Movie.720p.HDTV.mkv"
            $result.Resolution | Should -Be "720p"
            $result.Score | Should -BeGreaterOrEqual 60
        }

        It "Detects 480p resolution" {
            $result = Get-QualityScore "Movie.480p.DVDRip.mkv"
            $result.Resolution | Should -Be "480p"
            $result.Score | Should -BeGreaterOrEqual 40
        }
    }

    Context "Source Detection" {
        It "Detects BluRay source" {
            $result = Get-QualityScore "Movie.1080p.BluRay.x264.mkv"
            $result.Source | Should -Be "BluRay"
        }

        It "Detects WEB-DL source" {
            $result = Get-QualityScore "Movie.1080p.WEB-DL.x264.mkv"
            $result.Source | Should -Be "WEB-DL"
        }

        It "Detects WEBRip source" {
            $result = Get-QualityScore "Movie.1080p.WEBRip.x264.mkv"
            $result.Source | Should -Be "WEBRip"
        }

        It "Detects HDTV source" {
            $result = Get-QualityScore "Movie.720p.HDTV.x264.mkv"
            $result.Source | Should -Be "HDTV"
        }
    }

    Context "Codec Detection" {
        It "Detects HEVC/x265 codec" {
            $result = Get-QualityScore "Movie.2160p.BluRay.x265.mkv"
            $result.Codec | Should -Be "HEVC/x265"
        }

        It "Detects x264 codec" {
            $result = Get-QualityScore "Movie.1080p.BluRay.x264.mkv"
            $result.Codec | Should -Be "x264"
        }
    }

    Context "HDR Detection" {
        It "Detects HDR content" {
            $result = Get-QualityScore "Movie.2160p.BluRay.HDR.x265.mkv"
            $result.HDR | Should -Be $true
        }

        It "Detects Dolby Vision" {
            $result = Get-QualityScore "Movie.2160p.BluRay.DoVi.x265.mkv"
            $result.HDR | Should -Be $true
            $result.HDRFormat | Should -Be "Dolby Vision"
        }

        It "Detects HDR10+" {
            $result = Get-QualityScore "Movie.2160p.BluRay.HDR10Plus.x265.mkv"
            $result.HDR | Should -Be $true
            $result.HDRFormat | Should -Be "HDR10+"
        }

        It "Detects HDR10" {
            $result = Get-QualityScore "Movie.2160p.BluRay.HDR10.x265.mkv"
            $result.HDR | Should -Be $true
            $result.HDRFormat | Should -Be "HDR10"
        }

        It "Detects HLG" {
            $result = Get-QualityScore "Movie.2160p.BluRay.HLG.x265.mkv"
            $result.HDR | Should -Be $true
            $result.HDRFormat | Should -Be "HLG"
        }

        It "Non-HDR content returns false" {
            $result = Get-QualityScore "Movie.1080p.BluRay.x264.mkv"
            $result.HDR | Should -Be $false
        }
    }

    Context "Audio Codec Detection" {
        It "Detects Atmos" {
            $result = Get-QualityScore "Movie.2160p.BluRay.TrueHD.Atmos.mkv"
            $result.Audio | Should -Be "Atmos"
        }

        It "Detects TrueHD" {
            $result = Get-QualityScore "Movie.2160p.BluRay.TrueHD.mkv"
            $result.Audio | Should -Be "TrueHD"
        }

        It "Detects DTS-HD" {
            $result = Get-QualityScore "Movie.1080p.BluRay.DTS-HD.mkv"
            $result.Audio | Should -Be "DTS-HD"
        }

        It "Detects DTS:X" {
            $result = Get-QualityScore "Movie.2160p.BluRay.DTS-X.mkv"
            $result.Audio | Should -Be "DTS:X"
        }

        It "Detects EAC3/DD+" {
            $result = Get-QualityScore "Movie.1080p.WEB-DL.DDP.mkv"
            $result.Audio | Should -Be "EAC3"
        }
    }

    Context "New Codecs" {
        It "Detects AV1 codec" {
            $result = Get-QualityScore "Movie.2160p.WEB-DL.AV1.mkv"
            $result.Codec | Should -Be "AV1"
            $result.Score | Should -BeGreaterThan (Get-QualityScore "Movie.2160p.WEB-DL.x265.mkv").Score
        }
    }

    Context "Remux Source" {
        It "Detects Remux source" {
            $result = Get-QualityScore "Movie.2160p.BluRay.Remux.mkv"
            $result.Source | Should -Be "Remux"
        }

        It "Remux has higher score than BluRay" {
            # Hold resolution and codec constant so only the source differs —
            # the comparison isolates the Remux-vs-BluRay source bonus.
            $remuxScore = (Get-QualityScore "Movie.2160p.BluRay.Remux.x265.mkv").Score
            $blurayScore = (Get-QualityScore "Movie.2160p.BluRay.x265.mkv").Score
            $remuxScore | Should -BeGreaterThan $blurayScore
        }
    }

    Context "New Properties" {
        It "Has DataSource property" {
            $result = Get-QualityScore "Movie.1080p.BluRay.x264.mkv"
            $result.DataSource | Should -Be "Filename"
        }

        It "Has HDRFormat property" {
            $result = Get-QualityScore "Movie.2160p.BluRay.HDR10.x265.mkv"
            $result.HDRFormat | Should -Be "HDR10"
        }

        It "Accepts optional FilePath parameter" {
            { Get-QualityScore -FileName "Movie.mkv" -FilePath "C:\nonexistent\Movie.mkv" } | Should -Not -Throw
        }
    }

    Context "Score Calculation" {
        It "Higher quality gets higher score" {
            $score4k = (Get-QualityScore "Movie.2160p.BluRay.x265.mkv").Score
            $score1080 = (Get-QualityScore "Movie.1080p.BluRay.x264.mkv").Score
            $score720 = (Get-QualityScore "Movie.720p.HDTV.x264.mkv").Score

            $score4k | Should -BeGreaterThan $score1080
            $score1080 | Should -BeGreaterThan $score720
        }
    }
}

Describe "Get-EpisodeInfo" {
    Context "S01E01 Format" {
        It "Parses standard S01E01 format" {
            $result = Get-EpisodeInfo "Breaking.Bad.S01E01.Pilot.720p.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 1
            $result.ShowTitle | Should -Be "Breaking Bad"
            $result.IsMultiEpisode | Should -Be $false
        }

        It "Parses lowercase s01e01 format" {
            $result = Get-EpisodeInfo "the.office.s02e05.720p.mkv"
            $result.Season | Should -Be 2
            $result.Episode | Should -Be 5
            $result.ShowTitle | Should -Be "the office"
        }

        It "Parses double-digit episode numbers" {
            $result = Get-EpisodeInfo "Show.Name.S01E15.Episode.Title.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 15
        }

        It "Parses double-digit season numbers" {
            $result = Get-EpisodeInfo "Show.Name.S12E03.mkv"
            $result.Season | Should -Be 12
            $result.Episode | Should -Be 3
        }
    }

    Context "Multi-Episode Format" {
        It "Parses S01E01E02 format" {
            $result = Get-EpisodeInfo "Show.Name.S01E01E02.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 1
            $result.Episodes | Should -Contain 1
            $result.Episodes | Should -Contain 2
            $result.IsMultiEpisode | Should -Be $true
        }

        It "Parses S01E01-E03 format" {
            # Span releases name only their endpoints; production range-fills
            # the middle episodes so downstream organizers see every episode
            # the file covers.
            $result = Get-EpisodeInfo "Show.Name.S01E01-E03.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 1
            $result.IsMultiEpisode | Should -Be $true
            $result.Episodes | Should -Contain 1
            $result.Episodes | Should -Contain 2
            $result.Episodes | Should -Contain 3
        }
    }

    Context "1x01 Format" {
        It "Parses 1x01 format" {
            $result = Get-EpisodeInfo "Show.Name.1x05.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 5
        }

        It "Parses double-digit 10x15 format" {
            $result = Get-EpisodeInfo "Show.Name.10x15.mkv"
            $result.Season | Should -Be 10
            $result.Episode | Should -Be 15
        }
    }

    Context "Invalid Formats" {
        It "Returns null for unrecognized format" {
            $result = Get-EpisodeInfo "Random.Movie.2024.mkv"
            $result.Season | Should -Be $null
            $result.Episode | Should -Be $null
        }
    }
}

Describe "Get-NormalizedTitle" {
    Context "Year Extraction" {
        It "Extracts year in parentheses" {
            $result = Get-NormalizedTitle "The Movie (2024)"
            $result.Year | Should -Be "2024"
        }

        It "Extracts year without parentheses" {
            $result = Get-NormalizedTitle "The Movie 2024"
            $result.Year | Should -Be "2024"
        }

        It "Extracts year from complex filename" {
            $result = Get-NormalizedTitle "The.Movie.2024.1080p.BluRay.x264"
            $result.Year | Should -Be "2024"
        }

        It "Handles movies from 1900s" {
            $result = Get-NormalizedTitle "Classic Film (1985)"
            $result.Year | Should -Be "1985"
        }
    }

    Context "Title Normalization" {
        It "Removes quality tags" {
            $result = Get-NormalizedTitle "Movie.Name.1080p.BluRay.x264"
            $result.NormalizedTitle | Should -Not -Match "1080p"
            $result.NormalizedTitle | Should -Not -Match "bluray"
        }

        It "Replaces dots with spaces" {
            $result = Get-NormalizedTitle "Movie.Name.2024"
            $result.NormalizedTitle | Should -Not -Match "\."
        }

        It "Converts to lowercase" {
            $result = Get-NormalizedTitle "THE MOVIE NAME"
            $result.NormalizedTitle | Should -Be "movie name"
        }

        It "Removes leading articles" {
            $result = Get-NormalizedTitle "The Great Movie (2024)"
            $result.NormalizedTitle | Should -Be "great movie"

            $result = Get-NormalizedTitle "A Good Film (2024)"
            $result.NormalizedTitle | Should -Be "good film"
        }

        It "Trims whitespace" {
            $result = Get-NormalizedTitle "  Movie Name  (2024)"
            $result.NormalizedTitle | Should -Not -Match "^\s"
            $result.NormalizedTitle | Should -Not -Match "\s$"
        }
    }

    Context "Edge Cases" {
        It "Handles empty input" {
            $result = Get-NormalizedTitle ""
            $result.NormalizedTitle | Should -Be ""
        }

        It "Handles no year" {
            $result = Get-NormalizedTitle "Movie Without Year"
            $result.Year | Should -Be $null
            $result.NormalizedTitle | Should -Be "movie without year"
        }
    }
}

Describe "Configuration" {
    It "Has valid video extensions" {
        $script:Config.VideoExtensions | Should -Contain ".mp4"
        $script:Config.VideoExtensions | Should -Contain ".mkv"
        $script:Config.VideoExtensions | Should -Contain ".avi"
    }

    It "Has valid subtitle extensions" {
        $script:Config.SubtitleExtensions | Should -Contain ".srt"
        $script:Config.SubtitleExtensions | Should -Contain ".sub"
    }

    It "Has valid archive extensions" {
        $script:Config.ArchiveExtensions | Should -Contain "*.rar"
        $script:Config.ArchiveExtensions | Should -Contain "*.zip"
        $script:Config.ArchiveExtensions | Should -Contain "*.7z"
    }

    It "Has preferred subtitle languages" {
        $script:Config.PreferredSubtitleLanguages | Should -Contain "eng"
        $script:Config.PreferredSubtitleLanguages | Should -Contain "en"
    }
}

Describe "Statistics Tracking" {
    BeforeEach {
        # Reset stats
        $script:Stats.FilesDeleted = 0
        $script:Stats.BytesDeleted = 0
        $script:Stats.Errors = 0
        $script:Stats.ErrorDetails = @()
        $script:Stats.Warnings = 0
    }

    It "Tracks errors via Write-Log" {
        Write-Log "Test error" "ERROR"
        $script:Stats.Errors | Should -Be 1
        $script:Stats.ErrorDetails | Should -Contain "Test error"
    }

    It "Tracks warnings via Write-Log" {
        Write-Log "Test warning" "WARNING"
        $script:Stats.Warnings | Should -Be 1
    }

    It "Tracks multiple errors" {
        Write-Log "Error 1" "ERROR"
        Write-Log "Error 2" "ERROR"
        Write-Log "Error 3" "ERROR"
        $script:Stats.Errors | Should -Be 3
    }

    It "Writes log entries to the configured log file" {
        Write-Log "Log file smoke test" "INFO"
        $script:Config.LogFile | Should -Exist
        Get-Content $script:Config.LogFile -Raw | Should -Match "\[INFO\] Log file smoke test"
    }
}

Describe "File Matching Patterns" {
    Context "Unnecessary Files" {
        It "Matches sample files" {
            "Sample.mkv" -like "*Sample*" | Should -Be $true
            "movie-sample.mkv" -like "*sample*" | Should -Be $true
        }

        It "Matches proof files" {
            "Proof.jpg" -like "*Proof*" | Should -Be $true
        }

        It "Matches screenshots" {
            "Screens" -like "*Screens*" | Should -Be $true
            "Screenshots" -like "*Screens*" | Should -Be $true
        }
    }

    Context "Trailer Files" {
        It "Matches trailer files" {
            "Movie-Trailer.mkv" -like "*Trailer*" | Should -Be $true
            "movie-trailer.mkv" -like "*trailer*" | Should -Be $true
        }

        It "Matches teaser files" {
            "Movie-Teaser.mkv" -like "*Teaser*" | Should -Be $true
        }
    }
}

Describe "Integration Tests" -Tag "Integration" {
    BeforeAll {
        # Create test directory structure
        $testRoot = Join-Path $TestDrive "MediaLibrary"
        $movieFolder = Join-Path $testRoot "Movies"
        $tvFolder = Join-Path $testRoot "TVShows"

        New-Item -Path $movieFolder -ItemType Directory -Force | Out-Null
        New-Item -Path $tvFolder -ItemType Directory -Force | Out-Null
    }

    Context "Movie Library Structure" {
        BeforeAll {
            $moviePath = Join-Path $TestDrive "MediaLibrary\Movies"

            # Create test movie folders
            $movie1 = Join-Path $moviePath "The.Matrix.1999.1080p.BluRay.x264"
            $movie2 = Join-Path $moviePath "Inception (2010)"

            New-Item -Path $movie1 -ItemType Directory -Force | Out-Null
            New-Item -Path $movie2 -ItemType Directory -Force | Out-Null

            # Create dummy video files
            New-Item -Path (Join-Path $movie1 "movie.mkv") -ItemType File -Force | Out-Null
            New-Item -Path (Join-Path $movie2 "movie.mkv") -ItemType File -Force | Out-Null
        }

        It "Identifies movie folders" {
            $moviePath = Join-Path $TestDrive "MediaLibrary\Movies"
            $folders = Get-ChildItem -Path $moviePath -Directory
            $folders.Count | Should -Be 2
        }

        It "Parses movie names correctly" {
            $result1 = Get-NormalizedTitle "The.Matrix.1999.1080p.BluRay.x264"
            $result1.Year | Should -Be "1999"

            $result2 = Get-NormalizedTitle "Inception (2010)"
            $result2.Year | Should -Be "2010"
        }
    }

    Context "TV Show Library Structure" {
        BeforeAll {
            $tvPath = Join-Path $TestDrive "MediaLibrary\TVShows"
            $showPath = Join-Path $tvPath "Breaking Bad"

            New-Item -Path $showPath -ItemType Directory -Force | Out-Null

            # Create test episode files
            $episodes = @(
                "Breaking.Bad.S01E01.Pilot.720p.mkv",
                "Breaking.Bad.S01E02.720p.mkv",
                "Breaking.Bad.S02E01.720p.mkv"
            )

            foreach ($ep in $episodes) {
                New-Item -Path (Join-Path $showPath $ep) -ItemType File -Force | Out-Null
            }
        }

        It "Identifies episode files" {
            $showPath = Join-Path $TestDrive "MediaLibrary\TVShows\Breaking Bad"
            $files = Get-ChildItem -Path $showPath -File
            $files.Count | Should -Be 3
        }

        It "Parses episode info for all files" {
            $showPath = Join-Path $TestDrive "MediaLibrary\TVShows\Breaking Bad"
            $files = Get-ChildItem -Path $showPath -File

            foreach ($file in $files) {
                $info = Get-EpisodeInfo $file.Name
                $info.Season | Should -Not -Be $null
                $info.Episode | Should -Not -Be $null
            }
        }

        It "Identifies multiple seasons" {
            $showPath = Join-Path $TestDrive "MediaLibrary\TVShows\Breaking Bad"
            $files = Get-ChildItem -Path $showPath -File

            $seasons = $files | ForEach-Object {
                $info = Get-EpisodeInfo $_.Name
                $info.Season
            } | Sort-Object -Unique

            $seasons.Count | Should -Be 2
            $seasons | Should -Contain 1
            $seasons | Should -Contain 2
        }
    }
}

Describe "Edge Cases and Error Handling" {
    Context "Special Characters in Filenames" {
        It "Handles filenames with apostrophes" {
            $result = Get-NormalizedTitle "Marvel's Avengers (2012)"
            $result.Year | Should -Be "2012"
        }

        It "Handles filenames with colons" {
            $result = Get-NormalizedTitle "Movie - The Sequel (2024)"
            $result.NormalizedTitle | Should -Not -Be $null
        }

        It "Handles filenames with brackets" {
            $result = Get-NormalizedTitle "Movie [2024] 1080p"
            $result.Year | Should -Be "2024"
        }
    }

    Context "Empty and Null Inputs" {
        It "Handles null filename in Get-QualityScore" {
            { Get-QualityScore $null } | Should -Not -Throw
        }

        It "Handles empty filename in Get-EpisodeInfo" {
            $result = Get-EpisodeInfo ""
            $result.Season | Should -Be $null
        }
    }
}
