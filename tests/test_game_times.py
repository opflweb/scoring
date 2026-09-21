"""NFL kickoff timestamps exported for the score display."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / 'scripts'))

from export_for_web import build_game_opponents, build_game_times  # noqa: E402


def test_game_times_use_eastern_daylight_saving_offsets():
    rows = [
        {
            'week': 2,
            'gameday': '2026-09-20',
            'gametime': '13:00',
            'home_team': 'BUF',
            'away_team': 'MIA',
        },
        {
            'week': 14,
            'gameday': '2026-12-06',
            'gametime': '13:00',
            'home_team': 'NYG',
            'away_team': 'PHI',
        },
    ]

    game_times = build_game_times(rows)

    assert game_times[2]['BUF'] == '2026-09-20T13:00:00-04:00'
    assert game_times[14]['NYG'] == '2026-12-06T13:00:00-05:00'


def test_game_opponents_include_home_away_and_final_status():
    rows = [
        {
            'week': 2,
            'game_type': 'REG',
            'home_team': 'LAR',
            'away_team': 'NYG',
            'result': None,
        },
        {
            'week': 3,
            'game_type': 'REG',
            'home_team': 'BUF',
            'away_team': 'MIA',
            'result': 7,
        },
    ]

    opponents = build_game_opponents(rows)

    assert opponents[2]['LA'] == {'opponent': 'NYG', 'is_home': True, 'final': False}
    assert opponents[2]['NYG'] == {'opponent': 'LA', 'is_home': False, 'final': False}
    assert opponents[3]['BUF']['final'] is True
