"""resolve_matchups_week: catching a stale Matchups tab.

Found live: nflreadpy's calendar-based nfl.get_current_week() advanced to
week 2 a full day before the commissioner rolled the workbook's Matchups tab
over. An unguarded export scored week 1's players under
OPFLScorer(season, week_num=2) - looking up each player's week 2 NFL stats,
which mostly don't exist yet - and archived the result as
data/weeks/2026/week_2.json. That file was hand-deleted after the fact; this
test exists so the bug can't recur silently.
"""

import shutil
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / 'scripts'))

import export_for_web  # noqa: E402
from export_for_web import ROSTERS_SHEET, export_season, resolve_matchups_week  # noqa: E402

from opfl.week_archive import save_week  # noqa: E402

PROJECT_ROOT = Path(__file__).parent.parent
WORKBOOK = PROJECT_ROOT / 'OPFL Scoring 2026.xlsx'

pytestmark = pytest.mark.skipif(not WORKBOOK.exists(), reason='workbook not present')


def test_week_1_is_never_second_guessed():
    """No prior archive exists yet - nothing to compare against."""
    assert resolve_matchups_week('irrelevant.xlsx', 1, 2026) == 1


def _real_workbook_starters(tmp_path):
    """The actual starters currently on the real Matchups tab, in the shape
    resolve_matchups_week and the archive both use."""
    from opfl import build_matchup_week

    workbook = tmp_path / 'workbook.xlsx'
    shutil.copy(WORKBOOK, workbook)
    teams_by_code, _ = build_matchup_week(str(workbook), rosters_sheet=ROSTERS_SHEET)
    return {
        code: sorted(
            name for players in team.players.values() for name, _t, started in players if started
        )
        for code, team in teams_by_code.items()
    }


def test_a_stale_tab_falls_back_to_the_prior_week(tmp_path):
    """The exact scenario that happened: the requested week's lineups on the
    tab are identical to what we already archived for week N-1."""
    starters = _real_workbook_starters(tmp_path)
    archived_week = {
        'week': 1,
        'teams': [
            {'abbrev': code, 'roster': [{'name': n, 'starter': True} for n in names]}
            for code, names in starters.items()
        ],
    }
    save_week(2026, 1, archived_week, [], final=True, data_dir=tmp_path)

    resolved = resolve_matchups_week(
        str(WORKBOOK), requested_week=2, season=2026, data_dir=tmp_path
    )

    assert resolved == 1


def test_a_genuinely_different_lineup_is_accepted_as_the_new_week(tmp_path):
    """If the archived 'previous week' has different starters than the tab
    shows now, the tab really has moved on - trust the requested week."""
    fake_previous = {
        'week': 1,
        'teams': [
            {'abbrev': 'K/D', 'roster': [{'name': 'Nobody Real', 'starter': True}]},
        ],
    }
    save_week(2026, 1, fake_previous, [], final=True, data_dir=tmp_path)

    resolved = resolve_matchups_week(
        str(WORKBOOK), requested_week=2, season=2026, data_dir=tmp_path
    )

    assert resolved == 2


def test_no_prior_archive_trusts_the_requested_week(tmp_path):
    resolved = resolve_matchups_week(
        str(WORKBOOK), requested_week=2, season=2026, data_dir=tmp_path
    )
    assert resolved == 2


def test_export_season_checks_the_tab_even_with_an_explicit_week(monkeypatch):
    """The actual regression: the workflow always passes --week explicitly
    (computed from nflreadpy's calendar), which used to skip the stale-tab
    check entirely - it only ran when week_num was left as None. That let a
    week get archived with the previous week's lineups and pairings silently
    relabeled as the new one. export_season must run the check regardless of
    where week_num came from.
    """
    calls = []

    def fake_resolve(excel_path, requested_week, season, data_dir=None):
        calls.append(requested_week)
        return requested_week - 1

    monkeypatch.setattr(export_for_web, 'resolve_matchups_week', fake_resolve)
    monkeypatch.setattr(export_for_web, 'load_schedule_rows', lambda season: [])
    monkeypatch.setattr(export_for_web, 'get_current_nfl_week', lambda: 2)
    monkeypatch.setattr(
        export_for_web,
        'export_matchup_week',
        lambda excel_path, week_num, season: ({'week': week_num, 'teams': []}, []),
    )
    monkeypatch.setattr(export_for_web, 'week_games_are_final', lambda *a, **k: False)
    monkeypatch.setattr(export_for_web, 'save_week', lambda *a, **k: True)
    monkeypatch.setattr(export_for_web, 'load_all_weeks', lambda season: ([], {}))
    monkeypatch.setattr(export_for_web, 'load_season_schedule', lambda season: {})
    monkeypatch.setattr(export_for_web, 'build_standings', lambda weeks, schedule: [])
    monkeypatch.setattr(
        export_for_web, 'calculate_team_stats', lambda weeks, schedule, standings: {}
    )
    monkeypatch.setattr(export_for_web, 'generate_hall_of_fame', lambda: {})
    monkeypatch.setattr(export_for_web, 'build_game_times', lambda schedule_rows: {})
    monkeypatch.setattr(export_for_web, 'parse_taxi_squads', lambda *a, **k: {})
    monkeypatch.setattr(export_for_web, 'build_previous_seasons', lambda season: {})

    export_for_web.export_season('irrelevant.xlsx', week_num=2, season=2026)

    assert calls == [2]
