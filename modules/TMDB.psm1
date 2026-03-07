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
        $encodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)
        $url = "https://api.themoviedb.org/3/search/movie?api_key=$ApiKey&query=$encodedTitle"

        if ($Year) {
            $url += "&year=$Year"
        }

        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        if ($response.results -and $response.results.Count -gt 0) {
            # If year was provided, try to find an exact year match first
            $movie = $null
            if ($Year) {
                $movie = $response.results | Where-Object {
                    $_.release_date -and $_.release_date.StartsWith($Year)
                } | Select-Object -First 1
            }

            # Fall back to first result if no exact year match
            if (-not $movie) {
                $movie = $response.results[0]
            }

            return @{
                Id = $movie.id
                Title = $movie.title
                OriginalTitle = $movie.original_title
                Year = if ($movie.release_date) { $movie.release_date.Substring(0,4) } else { $null }
                Overview = $movie.overview
                Rating = $movie.vote_average
                Votes = $movie.vote_count
                PosterPath = if ($movie.poster_path) { "https://image.tmdb.org/t/p/w500$($movie.poster_path)" } else { $null }
                BackdropPath = if ($movie.backdrop_path) { "https://image.tmdb.org/t/p/original$($movie.backdrop_path)" } else { $null }
            }
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
        $url = "https://api.themoviedb.org/3/movie/$MovieId`?api_key=$ApiKey&append_to_response=credits,external_ids,videos"
        $movie = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        $directors = @()
        $cast = @()

        if ($movie.credits) {
            $directors = $movie.credits.crew | Where-Object { $_.job -eq 'Director' } | Select-Object -ExpandProperty name
            $cast = $movie.credits.cast | Select-Object -First 10 | ForEach-Object {
                @{
                    Name = $_.name
                    Role = $_.character
                    Thumb = if ($_.profile_path) { "https://image.tmdb.org/t/p/w185$($_.profile_path)" } else { $null }
                }
            }
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
Export-ModuleMember -Function Test-TMDBApiKey, Search-TMDBMovie, Get-TMDBMovieDetails, Get-TMDBCollectionImages,
    Search-TMDBTVShow, Get-TMDBEpisode,
    Get-TVDBToken, Test-TVDBApiKey, Search-TVDBShow, Get-TVDBShowDetails,
    Get-TVDBEpisode, Get-TVDBSeasonEpisodes
