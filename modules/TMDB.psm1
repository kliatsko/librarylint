# TMDB.psm1
# TMDB and TVDB API client functions for LibraryLint
# Provides movie/TV show search, metadata retrieval, and collection info
# Part of LibraryLint suite

#region Private State

# Cache TVDB auth token at module scope
$script:TVDBToken = $null
$script:TVDBTokenExpiry = $null

#endregion

#region TMDB Functions

<#
.SYNOPSIS
    Tests if a TMDB API key is valid
.PARAMETER ApiKey
    TMDB API key to validate
.OUTPUTS
    Boolean indicating if the key is valid
#>
function Test-TMDBApiKey {
    param(
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $false
    }

    try {
        # Use the configuration endpoint which requires a valid API key
        $url = "https://api.themoviedb.org/3/configuration?api_key=$ApiKey"
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        # If we get here, the key is valid
        if ($response.images) {
            return $true
        }
        return $false
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 401) {
            Write-Host "Invalid API key" -ForegroundColor Red
        } else {
            Write-Host "Error validating API key: $_" -ForegroundColor Red
        }
        return $false
    }
}

<#
.SYNOPSIS
    Searches TMDB for a movie by title and year
.PARAMETER Title
    The movie title to search for
.PARAMETER Year
    The movie year (optional but recommended)
.PARAMETER ApiKey
    TMDB API key (get one free at https://www.themoviedb.org/settings/api)
.OUTPUTS
    Hashtable with movie metadata from TMDB
#>
function Search-TMDBMovie {
    param(
        [string]$Title,
        [string]$Year = $null,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        Write-Host "TMDB API key not provided" -ForegroundColor Yellow
        return $null
    }

    try {
        # Build list of title variations to try
        $variations = [System.Collections.Generic.List[string]]::new()
        $variations.Add($Title)

        # Variation: replace numeric hyphens with slashes (e.g., "50-50" -> "50/50")
        if ($Title -match '\d-\d') {
            $slashVar = $Title -replace '(\d)-(\d)', '$1/$2'
            if ($slashVar -ne $Title) { $variations.Add($slashVar) }
        }

        # Variation: remove "Chapter One/Two/Three/..." suffixes (TMDB uses "IT" not "IT Chapter One")
        if ($Title -match '(?i)\s+Chapter\s+(One|Two|Three|Four|Five|1|2|3|4|5)\s*$') {
            $noChapter = ($Title -replace '(?i)\s+Chapter\s+(One|Two|Three|Four|Five|1|2|3|4|5)\s*$', '').Trim()
            if ($noChapter) { $variations.Add($noChapter) }
        }

        # Normalize the original title for comparison scoring
        $queryNorm = ($Title -replace '[^\w\s]', ' ' -replace '\s+', ' ').Trim().ToLower()
        $queryWords = @($queryNorm -split '\s+' | Where-Object { $_.Length -ge 1 } | Select-Object -Unique)
        # Whitespace-collapsed form for the compact-match check below — bridges
        # compound-vs-split words ("War Games" vs "WarGames") and apostrophe
        # gaps ("Winters Bone" vs "Winter's Bone") that the space-preserving
        # normalization can't equate.
        $queryCompact = $queryNorm -replace '\s+', ''

        $overallBest = $null
        $overallBestScore = -1

        foreach ($searchTitle in $variations) {
            $encodedTitle = [System.Web.HttpUtility]::UrlEncode($searchTitle)
            $url = "https://api.themoviedb.org/3/search/movie?api_key=$ApiKey&query=$encodedTitle"
            if ($Year) { $url += "&year=$Year" }

            $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

            if (-not $response.results -or $response.results.Count -eq 0) { continue }

            # Score each candidate by title similarity to the original query
            foreach ($candidate in $response.results) {
                # Hard year gate: when the caller specified a target year, any
                # candidate whose release year is more than ±1 off is a
                # different movie — regardless of title similarity. A ±1
                # tolerance absorbs legitimate discrepancies (release dates
                # differ across countries, premiere vs wide release). This
                # kills wrong matches like "War Games (1983)" -> "War Games
                # (2009)" and "The Town (2009)" -> "The Town That Was (2007)"
                # that otherwise coast through on title alone because the
                # year bonus is just a tiebreaker, not a filter.
                if ($Year -and $candidate.release_date) {
                    $candYearStr = $candidate.release_date.Substring(0, 4)
                    if ($candYearStr -match '^\d{4}$') {
                        $yearDiff = [math]::Abs([int]$candYearStr - [int]$Year)
                        if ($yearDiff -gt 1) { continue }
                    }
                }

                $score = 0

                foreach ($titleToCheck in @($candidate.title, $candidate.original_title)) {
                    if (-not $titleToCheck) { continue }
                    $candNorm = ($titleToCheck -replace '[^\w\s]', ' ' -replace '\s+', ' ').Trim().ToLower()

                    # Exact normalized match
                    if ($candNorm -eq $queryNorm) {
                        $score = [math]::Max($score, 100)
                        continue
                    }

                    # Compact-form match: collapse all whitespace and compare.
                    # Catches "War Games" <-> "WarGames" and "Winters Bone" <->
                    # "Winter's Bone" (apostrophe collapses to a stray 's' in
                    # the space-preserving normalization). Scored just below
                    # the exact match so a true exact-match candidate still
                    # wins when both forms are present in the result set.
                    $candCompact = $candNorm -replace '\s+', ''
                    if ($queryCompact.Length -gt 0 -and $candCompact -eq $queryCompact) {
                        $score = [math]::Max($score, 95)
                        continue
                    }

                    # Prefix-anchored containment. A candidate title that CONTAINS
                    # the query is only a match if the query appears at the START
                    # of the candidate (allowing a leading article). Without this
                    # anchor, a query like "the town" silently matches a long
                    # candidate like "lesson movie from michel deville nude in the
                    # town and village" purely because the substring "the town"
                    # happens to appear mid-title. The same guard applies in the
                    # reverse direction when the user's folder name is longer
                    # than the canonical title (e.g. "Dune Part Two" vs "Dune").
                    $queryStripped = $queryNorm -replace '^(the|a|an)\s+', ''
                    $candStripped  = $candNorm  -replace '^(the|a|an)\s+', ''
                    if ($candNorm.Contains($queryNorm)) {
                        if ($candStripped.StartsWith($queryStripped)) {
                            $score = [math]::Max($score, 80)
                        }
                        continue
                    }
                    if ($queryNorm.Contains($candNorm)) {
                        if ($queryStripped.StartsWith($candStripped)) {
                            $score = [math]::Max($score, 80)
                        }
                        continue
                    }

                    # Word overlap scoring
                    $candWords = @($candNorm -split '\s+' | Where-Object { $_.Length -ge 2 })
                    $querySignificant = @($queryWords | Where-Object { $_.Length -ge 2 })
                    if ($querySignificant.Count -gt 0 -and $candWords.Count -gt 0) {
                        $intersection = @($querySignificant | Where-Object { $_ -in $candWords }).Count
                        # For short titles (1-2 words), require exact or near-exact match
                        if ($querySignificant.Count -le 2) {
                            # Short title path: every query word must appear in the
                            # candidate AND the candidate must not be dramatically
                            # longer than the query. Without the size cap, a 2-word
                            # query like "the town" falsely matches an 11-word
                            # candidate that happens to contain both words.
                            if ($intersection -eq $querySignificant.Count -and
                                $candWords.Count -le $querySignificant.Count + 2) {
                                $score = [math]::Max($score, 80)
                            }
                        } else {
                            # Longer title: use Jaccard similarity
                            $union = ($querySignificant + $candWords | Select-Object -Unique).Count
                            if ($union -gt 0) {
                                $jaccard = $intersection / $union
                                $wordScore = [int](70 * $jaccard)
                                $score = [math]::Max($score, $wordScore)
                            }
                        }
                    }
                }

                # Year match bonus, gated on a minimum vote count. Without the
                # gate, an obscure zero-vote candidate that exactly matches the
                # query year (e.g. a 2022 Lebanese short titled "Talk to Me")
                # outscores the famous off-by-one candidate (the 2023 Australian
                # horror) purely on year, since the popularity tiebreaker below
                # can't make up the +15. The gate keeps year as a real
                # disambiguator between established releases without letting
                # near-anonymous TMDB entries weaponize it.
                $voteCount = if ($candidate.vote_count) { [int]$candidate.vote_count } else { 0 }
                if ($Year -and $candidate.release_date -and
                    $candidate.release_date.StartsWith($Year) -and $voteCount -ge 5) {
                    $score += 15
                }

                # Recognition bonus from vote_count (max 12 points). vote_count
                # is a stable signal of how well-known a film is; popularity is
                # spiky and trend-driven. The cap is calibrated so this bonus
                # alone can flip a tied title-score pair toward the well-known
                # candidate without overwhelming a real title-score difference.
                if ($voteCount -gt 0) {
                    $score += [math]::Min(12, [int]([math]::Log10($voteCount + 1) * 3))
                }

                if ($score -gt $overallBestScore) {
                    $overallBestScore = $score
                    $overallBest = $candidate
                }
            }
        }

        # Final sanity gate: the winner must share at least one distinctive query
        # word (length >= 4) with the candidate's title or original_title. Short
        # words like 'of', 'no', 'the' don't prove a real match — 'beasts' vs
        # 'beast' also fails, which is the correct behavior (plural stemming
        # would pick up wrong movies). Without this gate the scorer can accept
        # matches that only passed the 60-point threshold via year + popularity
        # bonuses on garbage candidates.
        $querySignificantLong = @($queryWords | Where-Object { $_.Length -ge 4 })
        if ($overallBest -and $querySignificantLong.Count -gt 0) {
            $bestTitleNorm = if ($overallBest.title) {
                ($overallBest.title -replace '[^\w\s]', ' ' -replace '\s+', ' ').Trim().ToLower()
            } else { '' }
            $bestOrigNorm = if ($overallBest.original_title) {
                ($overallBest.original_title -replace '[^\w\s]', ' ' -replace '\s+', ' ').Trim().ToLower()
            } else { '' }
            $bestWords = @((($bestTitleNorm + ' ' + $bestOrigNorm) -split '\s+') | Where-Object { $_ })
            $hasLongWordMatch = $false
            foreach ($word in $querySignificantLong) {
                if ($bestWords -contains $word) {
                    $hasLongWordMatch = $true
                    break
                }
            }
            # Compact-form equality is an alternate way to satisfy the gate:
            # a candidate whose space-collapsed title exactly equals the
            # query's space-collapsed form (War Games <-> WarGames, Winters
            # Bone <-> Winter's Bone) won't share length-4+ tokens with the
            # query but is unambiguously the same title. Equality (not
            # substring) keeps this from re-introducing the "the town" inside
            # a long candidate sentence false-positive that the gate exists
            # to block.
            $bestTitleCompact = $bestTitleNorm -replace '\s+', ''
            $bestOrigCompact  = $bestOrigNorm  -replace '\s+', ''
            $compactExactMatch = $queryCompact.Length -gt 0 -and (
                $bestTitleCompact -eq $queryCompact -or $bestOrigCompact -eq $queryCompact
            )
            if (-not $hasLongWordMatch -and -not $compactExactMatch) {
                Write-Log "TMDB scorer rejected '$($overallBest.title)' for query '$Title': no length-4+ word overlap or compact-form match" "DEBUG"
                return $null
            }
        }

        # Require minimum title similarity (65 = word overlap + year/popularity bonuses).
        # Progression: 50 (original) -> 60 (after initial tightening) -> 65 now,
        # so the year+popularity bonuses (max ~20) alone can't carry a candidate
        # that scored zero on title. Short-title word-overlap now gives 80, so
        # legitimate short matches clear 65 easily.
        if ($overallBest -and $overallBestScore -ge 65) {
            # Debug: log the winning match so we can audit scorer decisions when
            # a mismatch slips through in the wild. Includes query + winner +
            # final score + which candidate was examined, enough to replay the
            # scoring by hand.
            Write-Log "TMDB match accepted: query='$Title' (year=$Year) -> '$($overallBest.title)' (id=$($overallBest.id), release=$($overallBest.release_date)) score=$overallBestScore" "DEBUG"
            return @{
                Id = $overallBest.id
                Title = $overallBest.title
                OriginalTitle = $overallBest.original_title
                Year = if ($overallBest.release_date) { $overallBest.release_date.Substring(0,4) } else { $null }
                Overview = $overallBest.overview
                Rating = $overallBest.vote_average
                Votes = $overallBest.vote_count
                PosterPath = if ($overallBest.poster_path) { "https://image.tmdb.org/t/p/w500$($overallBest.poster_path)" } else { $null }
                BackdropPath = if ($overallBest.backdrop_path) { "https://image.tmdb.org/t/p/original$($overallBest.backdrop_path)" } else { $null }
            }
        }

        if ($overallBest) {
            Write-Log "TMDB match rejected: query='$Title' (year=$Year) -> best was '$($overallBest.title)' (id=$($overallBest.id)) score=$overallBestScore (threshold 65)" "DEBUG"
        }
        return $null
    }
    catch {
        Write-Host "Error searching TMDB: $_" -ForegroundColor Red
        return $null
    }
}

<#
.SYNOPSIS
    Gets detailed movie information from TMDB by ID
.PARAMETER MovieId
    The TMDB movie ID
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with detailed movie metadata
#>
function Get-TMDBMovieDetails {
    param(
        [int]$MovieId,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    try {
        $url = "https://api.themoviedb.org/3/movie/$MovieId`?api_key=$ApiKey&append_to_response=credits,external_ids,videos,release_dates"
        $movie = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        $directors = @()
        $cast = @()

        $writers = @()

        if ($movie.credits) {
            $directors = $movie.credits.crew | Where-Object { $_.job -eq 'Director' } | Select-Object -ExpandProperty name
            $writers = $movie.credits.crew | Where-Object { $_.job -in @('Writer', 'Screenplay', 'Story') } | Select-Object -ExpandProperty name -Unique
            $cast = $movie.credits.cast | Select-Object -First 10 | ForEach-Object {
                @{
                    Name = $_.name
                    Role = $_.character
                    Thumb = if ($_.profile_path) { "https://image.tmdb.org/t/p/w185$($_.profile_path)" } else { $null }
                }
            }
        }

        # Get US certification (MPAA rating)
        $certification = $null
        if ($movie.release_dates -and $movie.release_dates.results) {
            $usRelease = $movie.release_dates.results | Where-Object { $_.iso_3166_1 -eq 'US' }
            if ($usRelease -and $usRelease.release_dates) {
                $cert = $usRelease.release_dates | Where-Object { $_.certification } | Select-Object -First 1
                if ($cert) { $certification = $cert.certification }
            }
        }

        # Get production countries
        $countries = @()
        if ($movie.production_countries) {
            $countries = $movie.production_countries | Select-Object -ExpandProperty name
        }

        # Get YouTube trailer keys (prefer official trailers, then teasers)
        # Return ALL keys sorted by priority so we can try fallbacks if one is unavailable
        $trailerKey = $null
        $trailerKeys = @()
        if ($movie.videos -and $movie.videos.results) {
            # Priority 1: Official trailers
            $officialTrailers = $movie.videos.results | Where-Object {
                $_.site -eq 'YouTube' -and $_.type -eq 'Trailer' -and $_.official -eq $true
            } | Select-Object -ExpandProperty key

            # Priority 2: Non-official trailers
            $otherTrailers = $movie.videos.results | Where-Object {
                $_.site -eq 'YouTube' -and $_.type -eq 'Trailer' -and $_.official -ne $true
            } | Select-Object -ExpandProperty key

            # Priority 3: Teasers
            $teasers = $movie.videos.results | Where-Object {
                $_.site -eq 'YouTube' -and $_.type -eq 'Teaser'
            } | Select-Object -ExpandProperty key

            # Combine all keys in priority order
            $trailerKeys = @($officialTrailers) + @($otherTrailers) + @($teasers) | Where-Object { $_ }

            # First key is the primary trailer
            if ($trailerKeys.Count -gt 0) {
                $trailerKey = $trailerKeys[0]
            }
        }

        return @{
            Id = $movie.id
            TMDBID = $movie.id
            Title = $movie.title
            OriginalTitle = $movie.original_title
            OriginalLanguage = $movie.original_language  # ISO 639-1 code (en, fr, es, ja, etc.)
            Tagline = $movie.tagline
            Year = if ($movie.release_date) { $movie.release_date.Substring(0,4) } else { $null }
            ReleaseDate = $movie.release_date
            Overview = $movie.overview
            VoteAverage = $movie.vote_average
            VoteCount = $movie.vote_count
            Runtime = $movie.runtime
            Genres = $movie.genres | Select-Object -ExpandProperty name
            Studios = $movie.production_companies | Select-Object -ExpandProperty name
            Directors = $directors
            Actors = $cast
            IMDBID = $movie.external_ids.imdb_id
            PosterPath = if ($movie.poster_path) { "https://image.tmdb.org/t/p/w500$($movie.poster_path)" } else { $null }
            BackdropPath = if ($movie.backdrop_path) { "https://image.tmdb.org/t/p/original$($movie.backdrop_path)" } else { $null }
            TrailerKey = $trailerKey
            TrailerKeys = $trailerKeys
            CollectionId = if ($movie.belongs_to_collection) { $movie.belongs_to_collection.id } else { $null }
            CollectionName = if ($movie.belongs_to_collection) { $movie.belongs_to_collection.name } else { $null }
            MPAA = $certification
            Countries = $countries
            Writers = $writers
            Premiered = $movie.release_date
        }
    }
    catch {
        Write-Host "Error getting TMDB movie details: $_" -ForegroundColor Red
        return $null
    }
}

<#
.SYNOPSIS
    Gets collection/set artwork URLs from TMDB
.PARAMETER CollectionId
    The TMDB collection ID
.PARAMETER ApiKey
    TMDB API key
#>
function Get-TMDBCollectionImages {
    param(
        [int]$CollectionId,
        [string]$ApiKey
    )

    if (-not $ApiKey -or -not $CollectionId) {
        return $null
    }

    try {
        $url = "https://api.themoviedb.org/3/collection/$CollectionId`?api_key=$ApiKey"
        $collection = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        return @{
            Name = $collection.name
            PosterPath = if ($collection.poster_path) { "https://image.tmdb.org/t/p/original$($collection.poster_path)" } else { $null }
            BackdropPath = if ($collection.backdrop_path) { "https://image.tmdb.org/t/p/original$($collection.backdrop_path)" } else { $null }
        }
    }
    catch {
        Write-Host "Error getting TMDB collection images for ID ${CollectionId}: $_" -ForegroundColor Yellow
        return $null
    }
}

<#
.SYNOPSIS
    Gets all movies in a TMDB collection
.PARAMETER CollectionId
    The TMDB collection ID
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with collection Name and Parts (array of movies with Title, Year, TMDBID)
#>
function Get-TMDBCollectionParts {
    param(
        [int]$CollectionId,
        [string]$ApiKey
    )

    if (-not $ApiKey -or -not $CollectionId) {
        return $null
    }

    try {
        $url = "https://api.themoviedb.org/3/collection/$CollectionId`?api_key=$ApiKey"
        $collection = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        $parts = @()
        if ($collection.parts) {
            $parts = $collection.parts | ForEach-Object {
                @{
                    Title = $_.title
                    Year = if ($_.release_date) { $_.release_date.Substring(0, 4) } else { $null }
                    TMDBID = $_.id
                }
            } | Sort-Object { $_.Year }
        }

        return @{
            Name = $collection.name
            Parts = $parts
        }
    }
    catch {
        Write-Host "Error getting TMDB collection for ID ${CollectionId}: $_" -ForegroundColor Yellow
        return $null
    }
}

<#
.SYNOPSIS
    Searches TMDB for a TV show by title
.PARAMETER Title
    The TV show title to search for
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with TV show metadata from TMDB
#>
function Search-TMDBTVShow {
    param(
        [string]$Title,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    try {
        $encodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)
        $url = "https://api.themoviedb.org/3/search/tv?api_key=$ApiKey&query=$encodedTitle"

        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        if ($response.results -and $response.results.Count -gt 0) {
            $show = $response.results[0]

            return @{
                Id = $show.id
                Title = $show.name
                OriginalTitle = $show.original_name
                FirstAirDate = $show.first_air_date
                Year = if ($show.first_air_date) { $show.first_air_date.Substring(0,4) } else { $null }
                Overview = $show.overview
                Rating = $show.vote_average
                PosterPath = if ($show.poster_path) { "https://image.tmdb.org/t/p/w500$($show.poster_path)" } else { $null }
            }
        }

        return $null
    }
    catch {
        Write-Host "Error searching TMDB TV: $_" -ForegroundColor Red
        return $null
    }
}

<#
.SYNOPSIS
    Gets episode details from TMDB
.PARAMETER ShowId
    The TMDB TV show ID
.PARAMETER Season
    The season number
.PARAMETER Episode
    The episode number
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with episode metadata
#>
function Get-TMDBEpisode {
    param(
        [int]$ShowId,
        [int]$Season,
        [int]$Episode,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    try {
        $url = "https://api.themoviedb.org/3/tv/$ShowId/season/$Season/episode/$Episode`?api_key=$ApiKey"
        $ep = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        return @{
            Title = $ep.name
            Overview = $ep.overview
            AirDate = $ep.air_date
            Season = $ep.season_number
            Episode = $ep.episode_number
            Rating = $ep.vote_average
            StillPath = if ($ep.still_path) { "https://image.tmdb.org/t/p/w300$($ep.still_path)" } else { $null }
        }
    }
    catch {
        Write-Host "Error getting TMDB episode: $_" -ForegroundColor Red
        return $null
    }
}

#endregion

#region TVDB Functions

<#
.SYNOPSIS
    Authenticates with TVDB API v4 and returns a bearer token
.PARAMETER ApiKey
    TVDB API key (get one at https://thetvdb.com/api-information)
.OUTPUTS
    Bearer token string or $null if authentication fails
#>
function Get-TVDBToken {
    param(
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    # Check if we have a cached valid token (tokens last 30 days, refresh after 7)
    if ($script:TVDBToken -and $script:TVDBTokenExpiry -and (Get-Date) -lt $script:TVDBTokenExpiry) {
        return $script:TVDBToken
    }

    try {
        $url = "https://api4.thetvdb.com/v4/login"
        $body = @{ apikey = $ApiKey } | ConvertTo-Json

        $response = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json" -ErrorAction Stop

        if ($response.status -eq "success" -and $response.data.token) {
            $script:TVDBToken = $response.data.token
            # Cache for 7 days (tokens last 30 days)
            $script:TVDBTokenExpiry = (Get-Date).AddDays(7)
            return $script:TVDBToken
        }

        Write-Host "TVDB authentication failed: unexpected response" -ForegroundColor Red
        return $null
    }
    catch {
        Write-Host "TVDB authentication error: $_" -ForegroundColor Red
        return $null
    }
}

<#
.SYNOPSIS
    Tests if a TVDB API key is valid
.PARAMETER ApiKey
    The TVDB API key to validate
.OUTPUTS
    Boolean indicating if the key is valid
#>
function Test-TVDBApiKey {
    param(
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $false
    }

    try {
        $token = Get-TVDBToken -ApiKey $ApiKey
        return [bool]$token
    }
    catch {
        return $false
    }
}

<#
.SYNOPSIS
    Searches TVDB for a TV show by title
.PARAMETER Title
    The show title to search for
.PARAMETER Year
    Optional year to filter results
.PARAMETER ApiKey
    TVDB API key
.OUTPUTS
    Hashtable with show metadata from TVDB
#>
function Search-TVDBShow {
    param(
        [string]$Title,
        [string]$Year = $null,
        [string]$ApiKey
    )

    $token = Get-TVDBToken -ApiKey $ApiKey
    if (-not $token) {
        Write-Host "TVDB: No valid token available" -ForegroundColor Yellow
        return $null
    }

    try {
        $encodedTitle = [System.Uri]::EscapeDataString($Title)
        $url = "https://api4.thetvdb.com/v4/search?query=$encodedTitle&type=series"

        if ($Year) {
            $url += "&year=$Year"
        }

        $headers = @{
            "Authorization" = "Bearer $token"
            "Accept" = "application/json"
        }

        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop

        if ($response.status -eq "success" -and $response.data -and $response.data.Count -gt 0) {
            # Return the best match (first result, or filter by year if provided)
            $match = $response.data[0]

            # If year specified, try to find exact match
            if ($Year) {
                $exactMatch = $response.data | Where-Object { $_.year -eq $Year } | Select-Object -First 1
                if ($exactMatch) {
                    $match = $exactMatch
                }
            }

            return @{
                TVDBID = $match.tvdb_id
                Title = $match.name
                Year = $match.year
                Overview = $match.overview
                Status = $match.status
                PosterPath = $match.image_url
                Network = $match.network
                Slug = $match.slug
            }
        }

        return $null
    }
    catch {
        Write-Host "TVDB search error: $_" -ForegroundColor Red
        return $null
    }
}

<#
.SYNOPSIS
    Gets detailed information about a TV show from TVDB
.PARAMETER ShowId
    The TVDB show ID
.PARAMETER ApiKey
    TVDB API key
.OUTPUTS
    Hashtable with detailed show metadata
#>
function Get-TVDBShowDetails {
    param(
        [int]$ShowId,
        [string]$ApiKey
    )

    $token = Get-TVDBToken -ApiKey $ApiKey
    if (-not $token) {
        return $null
    }

    try {
        $url = "https://api4.thetvdb.com/v4/series/$ShowId/extended"
        $headers = @{
            "Authorization" = "Bearer $token"
            "Accept" = "application/json"
        }

        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop

        if ($response.status -eq "success" -and $response.data) {
            $show = $response.data

            # Extract genres
            $genres = @()
            if ($show.genres) {
                $genres = $show.genres | ForEach-Object { $_.name }
            }

            # Find primary poster and fanart
            $poster = $null
            $fanart = $null
            if ($show.artworks) {
                $posterArt = $show.artworks | Where-Object { $_.type -eq 2 -and $_.language -eq "eng" } | Select-Object -First 1
                if (-not $posterArt) {
                    $posterArt = $show.artworks | Where-Object { $_.type -eq 2 } | Select-Object -First 1
                }
                if ($posterArt) { $poster = $posterArt.image }

                $fanartArt = $show.artworks | Where-Object { $_.type -eq 3 -and $_.language -eq "eng" } | Select-Object -First 1
                if (-not $fanartArt) {
                    $fanartArt = $show.artworks | Where-Object { $_.type -eq 3 } | Select-Object -First 1
                }
                if ($fanartArt) { $fanart = $fanartArt.image }
            }

            # Extract studios/networks
            $studios = @()
            if ($show.networks) {
                $studios = $show.networks | ForEach-Object { $_.name }
            } elseif ($show.originalNetwork) {
                $studios = @($show.originalNetwork.name)
            }

            # Extract actors/characters
            $actors = @()
            if ($show.characters) {
                $actors = $show.characters | Where-Object { $_.type -eq 3 } | Select-Object -First 10 | ForEach-Object {
                    @{
                        Name = $_.personName
                        Role = $_.name
                        Thumb = $_.image
                        PersonId = $_.peopleId
                    }
                }
            }

            return @{
                TVDBID = $show.id
                Title = $show.name
                OriginalTitle = $show.originalName
                Year = if ($show.firstAired) { $show.firstAired.Substring(0, 4) } else { $null }
                FirstAired = $show.firstAired
                Overview = $show.overview
                Status = $show.status.name
                Runtime = $show.averageRuntime
                Genres = $genres
                Studios = $studios
                Rating = $show.score
                PosterPath = $poster
                FanartPath = $fanart
                IMDBID = $show.remoteIds | Where-Object { $_.sourceName -eq "IMDB" } | Select-Object -First 1 -ExpandProperty id
                Actors = $actors
                Seasons = $show.seasons | Where-Object { $_.type.type -eq "official" } | ForEach-Object {
                    @{
                        SeasonNumber = $_.number
                        Name = $_.name
                        EpisodeCount = $_.episodeCount
                        Image = $_.image
                    }
                }
            }
        }

        return $null
    }
    catch {
        Write-Host "TVDB show details error: $_" -ForegroundColor Red
        return $null
    }
}

<#
.SYNOPSIS
    Gets episode information from TVDB
.PARAMETER ShowId
    The TVDB show ID
.PARAMETER Season
    Season number
.PARAMETER Episode
    Episode number
.PARAMETER ApiKey
    TVDB API key
.OUTPUTS
    Hashtable with episode metadata
#>
function Get-TVDBEpisode {
    param(
        [int]$ShowId,
        [int]$Season,
        [int]$Episode,
        [string]$ApiKey
    )

    $token = Get-TVDBToken -ApiKey $ApiKey
    if (-not $token) {
        return $null
    }

    try {
        # First get all episodes for the season
        $url = "https://api4.thetvdb.com/v4/series/$ShowId/episodes/default?season=$Season"
        $headers = @{
            "Authorization" = "Bearer $token"
            "Accept" = "application/json"
        }

        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop

        if ($response.status -eq "success" -and $response.data.episodes) {
            $ep = $response.data.episodes | Where-Object { $_.seasonNumber -eq $Season -and $_.number -eq $Episode } | Select-Object -First 1

            if ($ep) {
                return @{
                    TVDBID = $ep.id
                    Title = $ep.name
                    Overview = $ep.overview
                    AirDate = $ep.aired
                    Season = $ep.seasonNumber
                    Episode = $ep.number
                    Runtime = $ep.runtime
                    StillPath = $ep.image
                }
            }
        }

        return $null
    }
    catch {
        Write-Host "TVDB episode error: $_" -ForegroundColor Red
        return $null
    }
}

<#
.SYNOPSIS
    Gets all episodes for a season from TVDB
.PARAMETER ShowId
    The TVDB show ID
.PARAMETER Season
    Season number
.PARAMETER ApiKey
    TVDB API key
.OUTPUTS
    Array of episode metadata hashtables
#>
function Get-TVDBSeasonEpisodes {
    param(
        [int]$ShowId,
        [int]$Season,
        [string]$ApiKey
    )

    $token = Get-TVDBToken -ApiKey $ApiKey
    if (-not $token) {
        return @()
    }

    try {
        $url = "https://api4.thetvdb.com/v4/series/$ShowId/episodes/default?season=$Season"
        $headers = @{
            "Authorization" = "Bearer $token"
            "Accept" = "application/json"
        }

        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop

        if ($response.status -eq "success" -and $response.data.episodes) {
            return $response.data.episodes | Where-Object { $_.seasonNumber -eq $Season } | ForEach-Object {
                @{
                    TVDBID = $_.id
                    Title = $_.name
                    Overview = $_.overview
                    AirDate = $_.aired
                    Season = $_.seasonNumber
                    Episode = $_.number
                    Runtime = $_.runtime
                    StillPath = $_.image
                }
            }
        }

        return @()
    }
    catch {
        Write-Host "TVDB season episodes error: $_" -ForegroundColor Red
        return @()
    }
}

#endregion

# Export public functions
Export-ModuleMember -Function Test-TMDBApiKey, Search-TMDBMovie, Get-TMDBMovieDetails, Get-TMDBCollectionImages, Get-TMDBCollectionParts,
    Search-TMDBTVShow, Get-TMDBEpisode,
    Get-TVDBToken, Test-TVDBApiKey, Search-TVDBShow, Get-TVDBShowDetails,
    Get-TVDBEpisode, Get-TVDBSeasonEpisodes
