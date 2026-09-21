"""The week archive: how a one-week-at-a-time workbook becomes a season.

The Matchups tab only ever holds the current week, so each run archives what it
scored and the exporter reads the season back out of data/weeks/.
"""

import json

from opfl.week_archive import archive_path, load_all_weeks, load_week, save_week


def week(number, scores=None):
    """A minimal week payload in the shape export_matchup_week produces."""
    scores = scores or {'K/D': 44.0, 'J/M': 37.0}
    return {
        'week': number,
        'has_scores': True,
        'teams': [
            {
                'abbrev': abbrev,
                'name': abbrev,
                'owner': abbrev,
                'total_score': score,
                'score_rank': rank,
                'roster': [
                    {
                        'name': f'Player {abbrev}',
                        'nfl_team': 'KC',
                        'position': 'QB',
                        'score': score,
                        'starter': True,
                    }
                ],
            }
            for rank, (abbrev, score) in enumerate(scores.items(), 1)
        ],
    }


class TestSaveAndLoad:
    def test_round_trips_a_week(self, tmp_path):
        save_week(2026, 1, week(1), [['K/D', 'J/M']], final=True, data_dir=tmp_path)
        stored = load_week(2026, 1, data_dir=tmp_path)
        assert stored['week'] == 1
        assert stored['season'] == 2026
        assert stored['final'] is True
        assert stored['pairings'] == [['K/D', 'J/M']]
        assert len(stored['teams']) == 2

    def test_missing_week_is_none(self, tmp_path):
        assert load_week(2026, 9, data_dir=tmp_path) is None

    def test_writes_under_season_and_week(self, tmp_path):
        save_week(2026, 3, week(3), [], final=False, data_dir=tmp_path)
        assert archive_path(2026, 3, tmp_path).exists()
        assert archive_path(2026, 3, tmp_path).parent.name == '2026'

    def test_an_in_progress_week_is_rewritten_on_each_run(self, tmp_path):
        """Scores keep moving until the games end, so the archive must too."""
        save_week(2026, 1, week(1, {'K/D': 10.0}), [], final=False, data_dir=tmp_path)
        assert save_week(2026, 1, week(1, {'K/D': 44.0}), [], final=False, data_dir=tmp_path)
        assert load_week(2026, 1, data_dir=tmp_path)['teams'][0]['total_score'] == 44.0

    def test_a_final_week_is_frozen(self, tmp_path):
        """nflverse restates stats weeks later; a completed result must not
        drift underneath the standings."""
        save_week(2026, 1, week(1, {'K/D': 44.0}), [], final=True, data_dir=tmp_path)
        assert not save_week(2026, 1, week(1, {'K/D': 99.0}), [], final=True, data_dir=tmp_path)
        assert load_week(2026, 1, data_dir=tmp_path)['teams'][0]['total_score'] == 44.0

    def test_force_overrides_the_freeze(self, tmp_path):
        save_week(2026, 1, week(1, {'K/D': 44.0}), [], final=True, data_dir=tmp_path)
        assert save_week(
            2026, 1, week(1, {'K/D': 99.0}), [], final=True, data_dir=tmp_path, force=True
        )
        assert load_week(2026, 1, data_dir=tmp_path)['teams'][0]['total_score'] == 99.0


class TestLoadAllWeeks:
    def test_empty_season(self, tmp_path):
        assert load_all_weeks(2026, data_dir=tmp_path) == ([], {})

    def test_returns_weeks_and_schedule_in_data_json_shape(self, tmp_path):
        save_week(2026, 1, week(1), [['K/D', 'J/M']], final=True, data_dir=tmp_path)
        save_week(2026, 2, week(2), [['K/D', 'G/G']], final=False, data_dir=tmp_path)

        weeks, schedule = load_all_weeks(2026, data_dir=tmp_path)
        assert [w['week'] for w in weeks] == [1, 2]
        assert [w['final'] for w in weeks] == [True, False]
        assert schedule == {'1': [['K/D', 'J/M']], '2': [['K/D', 'G/G']]}

    def test_weeks_sort_numerically_not_lexically(self, tmp_path):
        """week_10 must not land between week_1 and week_2."""
        for n in (1, 2, 9, 10, 11, 17):
            save_week(2026, n, week(n), [], final=True, data_dir=tmp_path)
        weeks, _ = load_all_weeks(2026, data_dir=tmp_path)
        assert [w['week'] for w in weeks] == [1, 2, 9, 10, 11, 17]

    def test_a_week_without_pairings_is_omitted_from_the_schedule(self, tmp_path):
        """Backfilled W-sheets record lineups but not always who played whom."""
        save_week(2026, 1, week(1), [], final=True, data_dir=tmp_path)
        weeks, schedule = load_all_weeks(2026, data_dir=tmp_path)
        assert len(weeks) == 1
        assert schedule == {}

    def test_seasons_are_isolated(self, tmp_path):
        save_week(2025, 1, week(1), [], final=True, data_dir=tmp_path)
        save_week(2026, 1, week(1), [], final=True, data_dir=tmp_path)
        assert len(load_all_weeks(2025, data_dir=tmp_path)[0]) == 1
        assert len(load_all_weeks(2026, data_dir=tmp_path)[0]) == 1

    def test_ignores_unrelated_files(self, tmp_path):
        save_week(2026, 1, week(1), [], final=True, data_dir=tmp_path)
        (tmp_path / 'weeks' / '2026' / 'notes.txt').write_text('scratch')
        (tmp_path / 'weeks' / '2026' / 'week_draft.json').write_text(json.dumps({'week': 0}))
        weeks, _ = load_all_weeks(2026, data_dir=tmp_path)
        assert [w['week'] for w in weeks] == [1]
