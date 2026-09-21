"""Weekly matchup projections: a shrunk blend of player history and opponent
strength, rolled up into a team total and a head-to-head win probability."""

from opfl.projections import calculate_week_projections
from opfl.week_archive import save_week


def archived_week(number, season, final=True):
    """A minimal finished week: two teams, one QB starter each."""
    return {
        'week': number,
        'has_scores': True,
        'final': final,
        'teams': [
            {
                'abbrev': 'AAA',
                'roster': [
                    {'name': 'Big Thrower', 'nfl_team': 'KC', 'position': 'QB', 'score': 25.0, 'starter': True}
                ],
            },
            {
                'abbrev': 'BBB',
                'roster': [
                    {'name': 'Small Thrower', 'nfl_team': 'BUF', 'position': 'QB', 'score': 15.0, 'starter': True}
                ],
            },
        ],
    }


def live_week(number):
    return {
        'week': number,
        'has_scores': False,
        'teams': [
            {
                'abbrev': 'AAA',
                'roster': [
                    {'name': 'Big Thrower', 'nfl_team': 'KC', 'position': 'QB', 'score': 0.0, 'starter': True},
                    {'name': 'Bench Thrower', 'nfl_team': 'KC', 'position': 'QB', 'score': 0.0, 'starter': False},
                ],
            },
            {
                'abbrev': 'BBB',
                'roster': [
                    {'name': 'Small Thrower', 'nfl_team': 'BUF', 'position': 'QB', 'score': 0.0, 'starter': True}
                ],
            },
        ],
    }


SCHEDULE_ROWS = [
    {
        'season': 2026,
        'week': w,
        'game_type': 'REG',
        'home_team': 'KC',
        'away_team': 'BUF',
        'result': '20-17' if w < 3 else None,
    }
    for w in range(1, 4)
]


class TestCalculateWeekProjections:
    def test_projects_toward_players_own_history(self, tmp_path):
        for week_num in range(1, 3):
            save_week(2026, week_num, archived_week(week_num, 2026), [['AAA', 'BBB']], True, data_dir=tmp_path)

        week_data = live_week(3)
        result = calculate_week_projections(
            week_data, [['AAA', 'BBB']], 2026, 3, SCHEDULE_ROWS, data_dir=tmp_path
        )

        aaa_player = week_data['teams'][0]['roster'][0]
        bbb_player = week_data['teams'][1]['roster'][0]
        # Big Thrower has scored more than Small Thrower every prior week, so
        # his projection should follow, shrunk toward their shared position
        # average rather than sitting exactly at his own history.
        assert aaa_player['projected_points'] > bbb_player['projected_points']
        assert 15.0 < aaa_player['projected_points'] < 25.0

        assert result['AAA']['projected_total'] == aaa_player['projected_points']
        assert result['BBB']['projected_total'] == bbb_player['projected_points']
        assert 'projected_points' in week_data['teams'][0]['roster'][1]
        assert result['AAA']['win_probability'] > result['BBB']['win_probability']
        assert result['AAA']['win_probability'] + result['BBB']['win_probability'] == 1.0

    def test_finished_game_uses_actual_score_not_projection(self, tmp_path):
        for week_num in range(1, 3):
            save_week(2026, week_num, archived_week(week_num, 2026), [['AAA', 'BBB']], True, data_dir=tmp_path)

        week_data = live_week(3)
        # Both games are final for week 3 in this schedule variant.
        finished_schedule = [dict(row, result='20-17') for row in SCHEDULE_ROWS if row['week'] == 3]
        week_data['teams'][0]['roster'][0]['score'] = 40.0
        week_data['teams'][1]['roster'][0]['score'] = 3.0

        result = calculate_week_projections(
            week_data, [['AAA', 'BBB']], 2026, 3, finished_schedule, data_dir=tmp_path
        )

        assert result['AAA']['projected_total'] == 40.0
        assert result['BBB']['projected_total'] == 3.0
        # Both starters are resolved, so the outcome is certain either way.
        assert result['AAA']['win_probability'] == 1.0
        assert result['BBB']['win_probability'] == 0.0

    def test_bye_week_player_contributes_zero(self, tmp_path):
        week_data = live_week(3)
        # No schedule rows at all for week 3: both teams are treated as on bye.
        result = calculate_week_projections(week_data, [['AAA', 'BBB']], 2026, 3, [], data_dir=tmp_path)

        assert week_data['teams'][0]['roster'][0]['projected_points'] == 0.0
        assert result['AAA']['projected_total'] == 0.0
        assert result['AAA']['win_probability'] == 0.5

    def test_no_history_falls_back_gracefully(self, tmp_path):
        week_data = live_week(1)
        result = calculate_week_projections(
            week_data, [['AAA', 'BBB']], 2026, 1, SCHEDULE_ROWS, data_dir=tmp_path
        )

        # With zero history, the position/player averages are both zero -
        # this must not raise (e.g. divide by zero) and should just project 0.
        assert result['AAA']['projected_total'] == 0.0
        assert result['BBB']['projected_total'] == 0.0
