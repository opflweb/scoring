"""The workflow's human-readable scoring report."""

from argparse import Namespace
from types import SimpleNamespace

import autoscorer
from opfl.models import PlayerScore


def _player(name, points, starter):
    return PlayerScore(
        name=name,
        position='QB',
        team='BUF',
        total_points=points,
        found_in_stats=True,
        is_starter=starter,
    )


def test_matchup_report_lists_bench_players_but_only_totals_starters(monkeypatch, capsys):
    teams = {
        'AAA': SimpleNamespace(name='Alpha'),
        'BBB': SimpleNamespace(name='Beta'),
    }
    matchups = [{'teams': [{'code': 'AAA'}, {'code': 'BBB'}]}]
    starter = _player('Starter', 10.0, True)
    bench = _player('Bench', 25.0, False)

    class FakeScorer:
        def __init__(self, season, week):
            pass

        def score_fantasy_team(self, team, starters_only):
            assert starters_only is False
            return {'QB': [starter, bench]}

    monkeypatch.setattr(autoscorer, 'build_matchup_week', lambda *args, **kwargs: (teams, matchups))
    monkeypatch.setattr(autoscorer, 'OPFLScorer', FakeScorer)

    results = autoscorer.score_matchups(
        Namespace(
            excel='unused.xlsx',
            rosters_sheet='Rosters',
            matchups_sheet='Matchups',
            season=2026,
            week=2,
            quiet=False,
            update=False,
        )
    )

    output = capsys.readouterr().out
    assert 'QB  Starter (BUF): 10.0 pts ✓' in output
    assert 'QB  Bench (BUF): 25.0 pts ✓ [BENCH]' in output
    assert output.count('TOTAL: 10.0 points') == 2
    assert results['AAA']['total'] == 10.0
    assert results['BBB']['total'] == 10.0
