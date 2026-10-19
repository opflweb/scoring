"""Name matching must never hand a rostered player someone else's stats.

Found live in 2026 week 4: players with no stats that week (injured, inactive)
fell through to a different player - "Josh Jacobs" -> "Josh Jobe", "Jordan
Mason" -> "Jordan Stout", "AJ Brown" -> "Mike Brown".
"""

import polars as pl
import pytest

from opfl.data_fetcher import NFLDataFetcher, fuzzy_match_name


@pytest.mark.parametrize(
    ('query', 'candidate'),
    [
        ('Josh Jacobs', 'Josh Jobe'),
        ('Jordan Mason', 'Jordan Stout'),
        ('Kyler Murray', 'Eric Murray'),
    ],
)
def test_fuzzy_rejects_a_different_player(query, candidate):
    assert fuzzy_match_name(query, [candidate], threshold=75) is None


@pytest.mark.parametrize(
    ('query', 'candidate'),
    [
        ('JK Dobbins', 'J.K. Dobbins'),
        ('Luther Burden', 'Luther Burden III'),
        ('Will Lutz', 'Wil Lutz'),
        ('Deebo Samual', 'Deebo Samuel'),
    ],
)
def test_fuzzy_accepts_spelling_variants(query, candidate):
    assert fuzzy_match_name(query, [candidate], threshold=75) == candidate


def _fetcher_with(players):
    fetcher = NFLDataFetcher(2026, 4)
    fetcher._player_stats = pl.DataFrame(
        {
            'player_display_name': [name for name, _ in players],
            'team': [team for _, team in players],
        }
    )
    return fetcher


def _fetcher_with_positions(players):
    fetcher = NFLDataFetcher(2026, 6)
    fetcher._player_stats = pl.DataFrame(
        {
            'player_display_name': [name for name, _, _ in players],
            'team': [team for _, team, _ in players],
            'position': [position for _, _, position in players],
        }
    )
    return fetcher


def test_fuzzy_never_matches_a_defender():
    # 2026 week 6: Detroit was on bye, so the WR fell through to a Titans LB.
    fetcher = _fetcher_with_positions(
        [('James Williams', 'TEN', 'LB'), ('Calvin Ridley', 'TEN', 'WR')]
    )
    assert fetcher.find_player('Jameson Williams', 'DET', 'WR') is None


def test_fuzzy_still_finds_a_traded_player_on_a_new_team():
    fetcher = _fetcher_with_positions([('Deebo Samuel', 'WAS', 'WR')])
    match = fetcher.find_player('Deebo Samual', 'SF', 'WR')
    assert match['player_display_name'] == 'Deebo Samuel'


def test_unique_last_name_on_team_still_needs_the_first_name():
    fetcher = _fetcher_with([('Mike Brown', 'NE'), ('Drake Maye', 'NE')])
    assert fetcher.find_player('AJ Brown', 'NE', 'WR', use_fuzzy=False) is None


def test_unique_last_name_on_team_matches_a_nickname():
    fetcher = _fetcher_with([('Wil Lutz', 'DEN')])
    match = fetcher.find_player('Will Lutz', 'DEN', 'K', use_fuzzy=False)
    assert match['player_display_name'] == 'Wil Lutz'
