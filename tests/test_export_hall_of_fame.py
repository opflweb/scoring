"""Player/team records, fun stats, and head-to-head history.

Verified against the real 2025 archive elsewhere (manually cross-checked
KAM's championship-week 93 and the AND-over-STL 51-22 week 9 head-to-head
result). These tests exercise the aggregation logic with small synthetic
seasons so the edge cases (ties, missing scores, the completed-weeks-only
rule) are pinned independent of the real data.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / 'scripts'))

from export_hall_of_fame import (  # noqa: E402
    calculate_fun_stats,
    calculate_head_to_head,
    calculate_player_records,
    calculate_team_records,
    load_all_seasons,
    week_label,
)

from opfl.week_archive import save_week  # noqa: E402


def player(name, position, score, starter=True):
    return {'name': name, 'position': position, 'score': score, 'starter': starter}


def team(abbrev, total_score, roster):
    return {'abbrev': abbrev, 'name': abbrev, 'total_score': total_score, 'roster': roster}


def week(number, teams, final=True):
    return {'week': number, 'final': final, 'teams': teams}


def season(year, weeks, schedule):
    return {'season': year, 'weeks': weeks, 'schedule': schedule}


class TestWeekLabel:
    def test_regular_weeks_are_numbered(self):
        assert week_label(1) == 'Week 1'
        assert week_label(15) == 'Week 15'

    def test_playoff_weeks_get_names(self):
        assert week_label(16) == 'Semifinals'
        assert week_label(17) == 'Championship Week'


class TestPlayerRecords:
    def _season(self):
        w1 = week(
            1,
            [
                team('A', 50, [player('QB Best', 'QB', 30), player('K Worst', 'K', 0)]),
                team(
                    'B',
                    40,
                    [player('RB Mid', 'RB', 12), player('WR Bench', 'WR', 5, starter=False)],
                ),
            ],
        )
        return [season(2025, [w1], {'1': [['A', 'B']]})]

    def test_most_points_is_sorted_descending(self):
        records = calculate_player_records(self._season())
        assert records['most_points'][0] == 'QB Best (A) - 30.0 (Week 1, 2025)'

    def test_most_points_non_qb_excludes_quarterbacks(self):
        records = calculate_player_records(self._season())
        assert all('QB Best' not in r for r in records['most_points_non_qb'])
        assert 'RB Mid (B) - 12.0 (Week 1, 2025)' in records['most_points_non_qb']

    def test_bench_players_are_excluded(self):
        records = calculate_player_records(self._season())
        assert not any('WR Bench' in r for r in records['most_points_non_qb'])

    def test_least_points_offensive_only_covers_qb_rb_wr_te(self):
        records = calculate_player_records(self._season())
        assert not any('K Worst' in r for r in records['least_points_offensive'])

    def test_least_points_kicker_is_separate_from_offense(self):
        records = calculate_player_records(self._season())
        assert 'K K Worst (A) - 0.0 (Week 1, 2025)' in records['least_points_kicker']


class TestTeamRecords:
    def _season(self):
        w1 = week(1, [team('A', 80, []), team('B', 20, [])])
        return [season(2025, [w1], {'1': [['A', 'B']]})]

    def test_most_and_least_points(self):
        records = calculate_team_records(self._season())
        assert records['most_points'][0] == 'A (A) - 80.0 (Week 1, 2025)'
        assert records['least_points'][0] == 'B (B) - 20.0 (Week 1, 2025)'

    def test_largest_margin_names_the_winner_first(self):
        records = calculate_team_records(self._season())
        assert records['largest_margin'][0] == 'A (A) over B (B) - by 60.0 (Week 1, 2025)'

    def test_a_team_with_zero_score_is_excluded_not_a_real_margin(self):
        """A 0 score usually means a player-parsing gap, not a real blowout."""
        w1 = week(1, [team('A', 0, []), team('B', 40, [])])
        records = calculate_team_records([season(2025, [w1], {'1': [['A', 'B']]})])
        assert records['largest_margin'] == []


class TestFunStats:
    def test_highest_and_lowest_scoring_weeks(self):
        w1 = week(1, [team('A', 60, []), team('B', 60, [])])
        w2 = week(2, [team('A', 20, []), team('B', 20, [])])
        stats = calculate_fun_stats([season(2025, [w1, w2], {})])
        highest = next(s for s in stats if s['title'].startswith('Highest'))
        lowest = next(s for s in stats if s['title'].startswith('Lowest'))
        assert highest['records'][0] == '120.0 points (Week 1, 2025)'
        assert lowest['records'][0] == '40.0 points (Week 2, 2025)'

    def test_closest_games_sorted_by_smallest_margin(self):
        w1 = week(1, [team('A', 50, []), team('B', 49, [])])
        stats = calculate_fun_stats([season(2025, [w1], {'1': [['A', 'B']]})])
        closest = next(s for s in stats if s['title'] == 'Closest Games')
        assert 'margin 1.0' in closest['records'][0]


class TestHeadToHead:
    def test_wins_losses_and_points_accumulate_across_weeks(self):
        w1 = week(1, [team('A', 50, []), team('B', 30, [])])
        w2 = week(2, [team('A', 20, []), team('B', 40, [])])
        seasons = [season(2025, [w1, w2], {'1': [['A', 'B']], '2': [['A', 'B']]})]
        records = calculate_head_to_head(seasons)
        assert len(records) == 1
        rec = records[0]
        assert (rec['team1_wins'], rec['team2_wins'], rec['ties']) == (1, 1, 0)
        assert rec['leader'] is None

    def test_leader_is_the_team_with_more_wins(self):
        w1 = week(1, [team('A', 50, []), team('B', 30, [])])
        w2 = week(2, [team('A', 40, []), team('B', 10, [])])
        seasons = [season(2025, [w1, w2], {'1': [['A', 'B']], '2': [['A', 'B']]})]
        records = calculate_head_to_head(seasons)
        assert records[0]['leader'] == 'A'

    def test_ties_are_counted(self):
        w1 = week(1, [team('A', 30, []), team('B', 30, [])])
        records = calculate_head_to_head([season(2025, [w1], {'1': [['A', 'B']]})])
        assert records[0]['ties'] == 1

    def test_pairs_are_orientation_independent(self):
        """A-vs-B and B-vs-A must accumulate into the same record."""
        w1 = week(1, [team('A', 50, []), team('B', 30, [])])
        w2 = week(2, [team('B', 45, []), team('A', 20, [])])
        seasons = [season(2025, [w1, w2], {'1': [['A', 'B']], '2': [['B', 'A']]})]
        records = calculate_head_to_head(seasons)
        assert len(records) == 1
        assert records[0]['games'] == 2


class TestLoadAllSeasonsExcludesUnfinishedWeeks:
    def test_an_in_progress_week_cannot_set_a_record(self, tmp_path):
        """The same rule Phase 1 applied to standings: a partial score must not
        become a Hall of Fame entry."""
        final_week = week(1, [team('A', 50, [player('QB Real', 'QB', 30)])])
        live_week = week(2, [team('A', 999, [player('QB Huge', 'QB', 999)])], final=False)
        save_week(2025, 1, final_week, [], final=True, data_dir=tmp_path)
        save_week(2025, 2, live_week, [], final=False, data_dir=tmp_path)

        all_seasons = load_all_seasons([2025], data_dir=tmp_path)
        weeks = all_seasons[0]['weeks']

        assert [w['week'] for w in weeks] == [1]
        records = calculate_player_records(all_seasons)
        assert records['most_points'] == ['QB Real (A) - 30.0 (Week 1, 2025)']
