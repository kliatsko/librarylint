#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for Search-TMDBMovie's candidate scoring (modules\TMDB.psm1).

.DESCRIPTION
    Three library folders — Split, Enemy, Passengers — ended up with NFOs for
    the wrong film because an obscure candidate that matched the folder's
    year exactly outscored the famous film dated one year off. The result
    set here is the one TMDB actually returned for "Split" with year 2016 on
    2026-09-10: Shyamalan's Split (2017-01-19, 18,606 votes) and Deborah
    Kampmeier's Split (2016-04-07, 56 votes), plus the smaller entries. HTTP
    is mocked, so this runs offline.

.NOTES
    Run with: Invoke-Pester -Path .\Tests\TMDB.Search.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'modules\TMDB.psm1') -Force

    function New-Candidate {
        param([int]$Id, [string]$Title, [string]$Date, [int]$Votes)
        [PSCustomObject]@{ id = $Id; title = $Title; original_title = $Title; release_date = $Date; vote_count = $Votes; vote_average = 6.5; overview = ''; poster_path = $null; backdrop_path = $null }
    }
    # Captured result set, in TMDB's order.
    $script:splitResults = @(
        New-Candidate -Id 381288 -Title 'Split' -Date '2017-01-19' -Votes 18606
        New-Candidate -Id 358364 -Title 'Split' -Date '2016-04-07' -Votes 56
        New-Candidate -Id 409583 -Title 'Split' -Date '2016-08-02' -Votes 26
        New-Candidate -Id 407223 -Title 'Split' -Date '2016-07-29' -Votes 3
        New-Candidate -Id 425636 -Title 'Split' -Date '2016-11-09' -Votes 8
        New-Candidate -Id 53373  -Title 'Split Estate' -Date '2009-08-07' -Votes 0
    )
}

Describe "Search-TMDBMovie year tiebreak" {
    It "picks Shyamalan's Split for 'Split (2016)' over the 56-vote same-year film" {
        InModuleScope TMDB -Parameters @{ Results = $script:splitResults } {
            param($Results)
            Mock Invoke-RestMethod { [PSCustomObject]@{ results = $Results } }
            $hit = Search-TMDBMovie -Title 'Split' -Year '2016' -ApiKey 'k'
            $hit.Id | Should -Be 381288
            $hit.Year | Should -Be '2017'
        }
    }

    # The year must still decide between two films people actually know:
    # a 200-vote "Prey" from 2021 is a real film, not noise, and a folder
    # named "Prey (2021)" means that one — not the Predator film a year later.
    It "lets an established same-title film win its exact year against a famous one a year off" {
        InModuleScope TMDB {
            Mock Invoke-RestMethod {
                [PSCustomObject]@{ results = @(
                    [PSCustomObject]@{ id = 766507; title = 'Prey'; original_title = 'Prey'; release_date = '2022-08-02'; vote_count = 22000; vote_average = 7.8; overview = ''; poster_path = $null; backdrop_path = $null }
                    [PSCustomObject]@{ id = 781963; title = 'Prey'; original_title = 'Prey'; release_date = '2021-09-10'; vote_count = 200;   vote_average = 5.7; overview = ''; poster_path = $null; backdrop_path = $null }
                ) }
            }
            (Search-TMDBMovie -Title 'Prey' -Year '2021' -ApiKey 'k').Id | Should -Be 781963
            (Search-TMDBMovie -Title 'Prey' -Year '2022' -ApiKey 'k').Id | Should -Be 766507
        }
    }

    It "gives a candidate with fewer than five votes no year credit at all" {
        InModuleScope TMDB {
            Mock Invoke-RestMethod {
                [PSCustomObject]@{ results = @(
                    [PSCustomObject]@{ id = 1; title = 'Talk to Me'; original_title = 'Talk to Me'; release_date = '2023-07-26'; vote_count = 3000; vote_average = 7.2; overview = ''; poster_path = $null; backdrop_path = $null }
                    [PSCustomObject]@{ id = 2; title = 'Talk to Me'; original_title = 'Talk to Me'; release_date = '2022-03-01'; vote_count = 2;    vote_average = 0;   overview = ''; poster_path = $null; backdrop_path = $null }
                ) }
            }
            (Search-TMDBMovie -Title 'Talk to Me' -Year '2022' -ApiKey 'k').Id | Should -Be 1
        }
    }

    It "still applies the hard year gate: a film two or more years off never wins on title alone" {
        InModuleScope TMDB {
            Mock Invoke-RestMethod {
                [PSCustomObject]@{ results = @(
                    [PSCustomObject]@{ id = 1; title = 'The Town'; original_title = 'The Town'; release_date = '2010-09-15'; vote_count = 5000; vote_average = 7.4; overview = ''; poster_path = $null; backdrop_path = $null }
                ) }
            }
            Search-TMDBMovie -Title 'The Town' -Year '2007' -ApiKey 'k' | Should -BeNullOrEmpty
        }
    }
}
