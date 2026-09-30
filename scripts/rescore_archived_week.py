#!/usr/bin/env python3
"""Rescore an archived week against current nflverse data, keeping its lineups.

`export_for_web.py --force-rescore` takes lineups from the Matchups tab, which
only ever holds the live week - so it cannot correct a past week. This rescores
every player in data/weeks/{season}/week_{N}.json using the lineups already
archived there, optionally correcting a starter first:

    python scripts/rescore_archived_week.py --week 2
    python scripts/rescore_archived_week.py --week 3 \\
        --start KEV "Deebo Samuel" --bench KEV "Zay Flowers"

Run export_for_web.py afterwards to rebuild web/data.json and the standings.
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from opfl.scorer import OPFLScorer
from opfl.week_archive import load_week, save_week


def set_starter(teams: list[dict], code: str, name: str, starter: bool) -> None:
    team = next((t for t in teams if t['abbrev'] == code), None)
    if team is None:
        raise SystemExit(f'No team {code!r} in the archived week')
    player = next((p for p in team['roster'] if p['name'] == name), None)
    if player is None:
        raise SystemExit(f'No player {name!r} on {code} in the archived week')
    player['starter'] = starter


def rescore(week: dict, season: int, week_num: int) -> None:
    scorer = OPFLScorer(season, week_num)
    for team in week['teams']:
        total = 0.0
        for player in team['roster']:
            ps = scorer.score_player(player['name'], player['nfl_team'], player['position'])
            player['score'] = round(ps.total_points, 1)
            player['found'] = ps.found_in_stats
            player['breakdown'] = ps.breakdown
            if player['starter']:
                total += ps.total_points
        team['total_score'] = round(total, 1)

    ranked = sorted(week['teams'], key=lambda t: t['total_score'], reverse=True)
    for rank, team in enumerate(ranked, 1):
        team['score_rank'] = rank


def main():
    parser = argparse.ArgumentParser(description=__doc__.split('\n\n')[0])
    parser.add_argument('--week', '-w', type=int, required=True)
    parser.add_argument('--season', '-y', type=int, default=2026)
    parser.add_argument('--start', nargs=2, action='append', default=[], metavar=('TEAM', 'PLAYER'))
    parser.add_argument('--bench', nargs=2, action='append', default=[], metavar=('TEAM', 'PLAYER'))
    args = parser.parse_args()

    week = load_week(args.season, args.week)
    if not week:
        raise SystemExit(f'Week {args.week} of {args.season} is not archived')

    for code, name in args.start:
        set_starter(week['teams'], code, name, True)
    for code, name in args.bench:
        set_starter(week['teams'], code, name, False)

    before = {t['abbrev']: t['total_score'] for t in week['teams']}
    rescore(week, args.season, args.week)

    for team in week['teams']:
        old, new = before[team['abbrev']], team['total_score']
        print(f'  {team["abbrev"]:4} {old:5.1f} -> {new:5.1f}' + ('' if old == new else '  *'))

    save_week(
        args.season,
        args.week,
        week,
        week['pairings'],
        week['final'],
        force=True,
    )
    return 0


if __name__ == '__main__':
    sys.exit(main())
