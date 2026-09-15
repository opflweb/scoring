"""The web/data.json contract.

`web/data.json` is the only thing the website reads, so its shape is a real
interface. These tests run against the committed artifact — no network, no
nflverse — which means they also catch a bad export getting committed.

The scoring workflow runs these after writing data.json and before pushing.
"""

import json
from pathlib import Path

import pytest

from opfl.constants import ALL_TEAM_CODES

DATA_JSON = Path(__file__).parent.parent / 'web' / 'data.json'

REQUIRED_TOP_LEVEL_KEYS = {
    'updated_at',
    'season',
    'current_week',
    'regular_season_weeks',
    'weeks',
    'schedule',
    'team_number_map',
    'standings',
    'standings_through_week',
    'standings_in_progress',
    'team_stats',
    'playoffs',
    'game_times',
    'trade_deadline_week',
    'taxi_squads',
    'pending_trades',
    'draft_picks',
    'banners',
}

REQUIRED_PLAYER_KEYS = {'name', 'nfl_team', 'position', 'score', 'starter'}
REQUIRED_TEAM_KEYS = {'name', 'owner', 'abbrev', 'roster', 'total_score', 'score_rank'}
REQUIRED_STANDINGS_KEYS = {
    'name',
    'owner',
    'abbrev',
    'rank_points',
    'wins',
    'losses',
    'ties',
    'top_half',
    'points_for',
    'points_against',
}

VALID_POSITIONS = {'QB', 'RB', 'WR', 'TE', 'K', 'DF', 'HC'}


@pytest.fixture(scope='module')
def data():
    if not DATA_JSON.exists():
        pytest.skip('web/data.json has not been generated')
    return json.loads(DATA_JSON.read_text())


def test_top_level_keys_are_present(data):
    missing = REQUIRED_TOP_LEVEL_KEYS - set(data)
    assert not missing, f'web/data.json is missing {sorted(missing)}'


def test_weeks_carry_teams_and_a_scored_flag(data):
    assert data['weeks'], 'no weeks exported'
    for week in data['weeks']:
        assert {'week', 'teams', 'has_scores'} <= set(week)
        assert isinstance(week['week'], int)
        assert len(week['teams']) == len(ALL_TEAM_CODES)


def test_every_team_entry_has_the_expected_fields(data):
    for week in data['weeks']:
        for team in week['teams']:
            missing = REQUIRED_TEAM_KEYS - set(team)
            assert not missing, f'week {week["week"]} {team.get("abbrev")} missing {missing}'
            assert team['abbrev'] in ALL_TEAM_CODES


def test_every_roster_player_has_the_expected_fields(data):
    for week in data['weeks']:
        for team in week['teams']:
            assert team['roster'], f'{team["abbrev"]} has an empty roster'
            for player in team['roster']:
                missing = REQUIRED_PLAYER_KEYS - set(player)
                assert not missing, f'{player.get("name")} missing {missing}'
                assert player['position'] in VALID_POSITIONS
                assert isinstance(player['score'], (int, float))
                assert isinstance(player['starter'], bool)


def test_each_team_starts_exactly_nine_players(data):
    """QB1 / RB2 / WR2 / TE1 / K1 / DF1 / HC1."""
    for week in data['weeks']:
        for team in week['teams']:
            starters = [p for p in team['roster'] if p['starter']]
            assert len(starters) == 9, (
                f'week {week["week"]} {team["abbrev"]} started {len(starters)}'
            )


def test_team_total_equals_the_sum_of_its_starters(data):
    for week in data['weeks']:
        for team in week['teams']:
            expected = sum(p['score'] for p in team['roster'] if p['starter'])
            assert team['total_score'] == pytest.approx(expected, abs=0.05), (
                f'week {week["week"]} {team["abbrev"]}'
            )


def test_standings_cover_every_team_with_the_expected_fields(data):
    """Standings are empty until a week's NFL games are all final, then they
    cover the whole league. A partial league is a bug either way."""
    if not data['standings']:
        assert data['standings_through_week'] == 0
        return

    assert {s['abbrev'] for s in data['standings']} == set(ALL_TEAM_CODES)
    for standing in data['standings']:
        missing = REQUIRED_STANDINGS_KEYS - set(standing)
        assert not missing, f'{standing.get("abbrev")} missing {missing}'


def test_standings_only_count_completed_weeks(data):
    """The week-1-scored-before-Monday-night bug: an unfinished week must not
    post W/L records."""
    final_weeks = [w['week'] for w in data['weeks'] if w.get('final')]
    assert data['standings_through_week'] == max(final_weeks, default=0)

    # Only weeks with schedule pairings contribute W/L (playoff weeks 16-17
    # aren't part of the round-robin schedule and are seeded separately).
    scheduled_final_weeks = [w for w in final_weeks if str(w) in data['schedule']]
    games_played = sum(s['wins'] + s['losses'] + s['ties'] for s in data['standings'])
    assert games_played == len(scheduled_final_weeks) * len(ALL_TEAM_CODES)


def test_every_week_declares_whether_it_is_final(data):
    for week in data['weeks']:
        assert isinstance(week.get('final'), bool), f'week {week["week"]} has no final flag'


def test_schedule_pairs_reference_real_teams_and_nobody_plays_twice(data):
    for week, pairings in data['schedule'].items():
        played = [code for pair in pairings for code in pair]
        assert set(played) <= set(ALL_TEAM_CODES), f'week {week} has an unknown team'
        assert len(played) == len(set(played)), f'week {week} schedules a team twice'


def test_every_scheduled_week_has_matching_week_data(data):
    exported = {w['week'] for w in data['weeks']}
    scheduled = {int(w) for w in data['schedule']}
    assert scheduled <= exported, f'scheduled but not exported: {scheduled - exported}'


def test_team_stats_only_covers_teams_that_appear_in_standings(data):
    """team_stats is empty until a week is final (same rule as standings);
    once populated it should never invent a team standings doesn't have."""
    assert set(data['team_stats']) <= {s['abbrev'] for s in data['standings']}
