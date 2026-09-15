"""Lineup efficiency and season-long team stats.

Both are ported from the sibling QPFL site's scripts/export_for_web.py; these
tests pin the one real behavioral difference (OPFL's roster entries never
carry a `taxi` flag) and the OPFL-specific inputs (schedule as a top-level
dict rather than a per-week matchups list, `week['final']` gating).
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / 'scripts'))

from export_for_web import calculate_lineup_efficiency, calculate_team_stats  # noqa: E402


def roster_player(position, score, starter):
    return {
        'name': f'{position}-{score}',
        'nfl_team': 'KC',
        'position': position,
        'score': score,
        'starter': starter,
    }


class TestLineupEfficiency:
    def test_optimal_lineup_leaves_nothing_on_the_table(self):
        team = {
            'roster': [
                roster_player('QB', 20, True),
                roster_player('QB', 10, False),
                roster_player('RB', 15, True),
                roster_player('RB', 5, False),
            ]
        }
        result = calculate_lineup_efficiency(team)
        assert result['actual_points'] == 35
        assert result['optimal_points'] == 35
        assert result['points_left_on_table'] == 0

    def test_starting_the_lower_scorer_leaves_points_on_the_table(self):
        team = {
            'roster': [
                roster_player('QB', 10, True),
                roster_player('QB', 25, False),
            ]
        }
        result = calculate_lineup_efficiency(team)
        assert result['actual_points'] == 10
        assert result['optimal_points'] == 25
        assert result['points_left_on_table'] == 15

    def test_optimal_lineup_respects_slot_counts_not_every_bench_player(self):
        """Two RB starters means the optimal lineup is the top 2 RB scores,
        not every RB on the roster."""
        team = {
            'roster': [
                roster_player('RB', 20, True),
                roster_player('RB', 15, True),
                roster_player('RB', 30, False),
                roster_player('RB', 5, False),
            ]
        }
        result = calculate_lineup_efficiency(team)
        assert result['actual_points'] == 35
        assert result['optimal_points'] == 50  # 30 + 20, the top two
        assert result['points_left_on_table'] == 15

    def test_no_starters_returns_none(self):
        team = {'roster': [roster_player('QB', 20, False)]}
        assert calculate_lineup_efficiency(team) is None

    def test_missing_roster_returns_none(self):
        assert calculate_lineup_efficiency({}) is None

    def test_non_numeric_scores_are_skipped_not_crashed_on(self):
        """A not-yet-scored player has score=None; must not blow up sorted()."""
        team = {
            'roster': [
                roster_player('QB', 20, True),
                {'name': 'unscored', 'position': 'QB', 'score': None, 'starter': False},
            ]
        }
        result = calculate_lineup_efficiency(team)
        assert result['optimal_points'] == 20


def week(number, final, scores, pairings):
    """Build a week in the shape opfl.week_archive stores."""
    return {
        'week': number,
        'final': final,
        'has_scores': True,
        'teams': [
            {
                'abbrev': abbrev,
                'name': abbrev,
                'owner': abbrev,
                'total_score': score,
                'roster': [
                    roster_player('QB', score, True),
                ],
            }
            for abbrev, score in scores.items()
        ],
    }, {str(number): pairings}


class TestCalculateTeamStats:
    def _two_week_season(self):
        w1, s1 = week(1, True, {'A': 50, 'B': 30, 'C': 40, 'D': 20}, [['A', 'B'], ['C', 'D']])
        w2, s2 = week(2, True, {'A': 25, 'B': 45, 'C': 35, 'D': 55}, [['A', 'B'], ['C', 'D']])
        weeks = [w1, w2]
        schedule = {**s1, **s2}
        standings = [
            {
                'abbrev': 'A',
                'name': 'A',
                'wins': 1,
                'losses': 1,
                'ties': 0,
                'points_for': 75,
                'points_against': 75,
            },
            {
                'abbrev': 'B',
                'name': 'B',
                'wins': 1,
                'losses': 1,
                'ties': 0,
                'points_for': 75,
                'points_against': 75,
            },
            {
                'abbrev': 'C',
                'name': 'C',
                'wins': 1,
                'losses': 1,
                'ties': 0,
                'points_for': 75,
                'points_against': 75,
            },
            {
                'abbrev': 'D',
                'name': 'D',
                'wins': 1,
                'losses': 1,
                'ties': 0,
                'points_for': 75,
                'points_against': 75,
            },
        ]
        return weeks, schedule, standings

    def test_empty_standings_returns_empty_stats(self):
        assert calculate_team_stats([], {}, []) == {}

    def test_ppg_and_totals_come_from_standings_not_recomputed(self):
        weeks, schedule, standings = self._two_week_season()
        stats = calculate_team_stats(weeks, schedule, standings)
        assert stats['A']['total_points_for'] == 75
        assert stats['A']['ppg'] == 37.5

    def test_best_and_worst_week_numbers_are_correct(self):
        weeks, schedule, standings = self._two_week_season()
        stats = calculate_team_stats(weeks, schedule, standings)
        # A scored 50 in week 1, 25 in week 2.
        assert stats['A']['best_week'] == 50
        assert stats['A']['best_week_num'] == 1
        assert stats['A']['worst_week'] == 25
        assert stats['A']['worst_week_num'] == 2

    def test_current_streak_tracks_the_most_recent_result(self):
        weeks, schedule, standings = self._two_week_season()
        stats = calculate_team_stats(weeks, schedule, standings)
        # D lost week 1 (20 vs 40), won week 2 (55 vs 35): a 1-game win streak.
        assert stats['D']['streak'] == {'type': 'W', 'count': 1}
        # A won week 1 (50 vs 30), lost week 2 (25 vs 45): a 1-game loss streak.
        assert stats['A']['streak'] == {'type': 'L', 'count': 1}

    def test_streak_count_extends_across_consecutive_matching_results(self):
        w1, s1 = week(1, True, {'A': 50, 'B': 30}, [['A', 'B']])
        w2, s2 = week(2, True, {'A': 60, 'B': 20}, [['A', 'B']])
        w3, s3 = week(3, True, {'A': 40, 'B': 10}, [['A', 'B']])
        standings = [
            {
                'abbrev': 'A',
                'name': 'A',
                'wins': 3,
                'losses': 0,
                'ties': 0,
                'points_for': 150,
                'points_against': 60,
            },
            {
                'abbrev': 'B',
                'name': 'B',
                'wins': 0,
                'losses': 3,
                'ties': 0,
                'points_for': 60,
                'points_against': 150,
            },
        ]
        stats = calculate_team_stats([w1, w2, w3], {**s1, **s2, **s3}, standings)
        assert stats['A']['streak'] == {'type': 'W', 'count': 3}

    def test_an_in_progress_week_is_excluded(self):
        w1, s1 = week(1, True, {'A': 50, 'B': 30}, [['A', 'B']])
        w2, s2 = week(2, False, {'A': 99, 'B': 1}, [['A', 'B']])
        standings = [
            {
                'abbrev': 'A',
                'name': 'A',
                'wins': 1,
                'losses': 0,
                'ties': 0,
                'points_for': 50,
                'points_against': 30,
            },
            {
                'abbrev': 'B',
                'name': 'B',
                'wins': 0,
                'losses': 1,
                'ties': 0,
                'points_for': 30,
                'points_against': 50,
            },
        ]
        stats = calculate_team_stats([w1, w2], {**s1, **s2}, standings)
        # If week 2 leaked in, A's best week would be 99, not 50.
        assert stats['A']['best_week'] == 50
        assert stats['A']['ppg'] == 50.0

    def test_lineup_efficiency_accumulates_across_weeks(self):
        weeks, schedule, standings = self._two_week_season()
        stats = calculate_team_stats(weeks, schedule, standings)
        # Every roster here has exactly one player who is also the starter,
        # so lineup_optimal_points always equals total_score - no points left.
        assert stats['A']['lineup_weeks'] == 2
        assert stats['A']['points_left_on_table'] == 0

    def test_no_offensive_line_position_appears_in_output(self):
        """OPFL has no OL slot; nothing in the stats pipeline should invent one."""
        weeks, schedule, standings = self._two_week_season()
        stats = calculate_team_stats(weeks, schedule, standings)
        assert 'OL' not in str(stats)

    def test_opr_and_adjusted_opr_are_present_and_relative_to_league_average(self):
        weeks, schedule, standings = self._two_week_season()
        stats = calculate_team_stats(weeks, schedule, standings)
        oprs = [s['opr'] for s in stats.values()]
        avg = sum(oprs) / len(oprs)
        for s in stats.values():
            assert s['adjusted_opr'] == round(s['opr'] / avg, 3)
