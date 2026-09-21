#!/usr/bin/env python3
"""Rebuild the week archive from a workbook's W-sheets.

The commissioner archives each finished week to a `W{n}` sheet inside the same
workbook. Those sheets are static values (not formulas) in the same roster-block
layout as the Rosters tab, with that week's starters starred — so a whole season
can be reconstructed from them.

By default this reads each player's points straight out of the sheet rather
than re-deriving them from nflreadpy stats. That is the only reliable option
for older seasons: their W-sheets never recorded a player's NFL team (just
"D. Adams", not "D. Adams (LV)"), so re-scoring has to guess at stat lookups
for any ambiguous surname and quietly undercounts the week. The sheet's own
point values don't have that problem - they're what the league actually used.
Pass --rescore to re-derive from nflreadpy instead, e.g. to repair a current
season week where a late stat correction hasn't been typed into the sheet.

    python scripts/backfill_weeks.py --excel "data/previous_seasons/OPFL Scoring 2022.xlsx" --season 2022
    python scripts/backfill_weeks.py --season 2025 --weeks 3 4 5 --force --rescore
"""

import argparse
import json
import re
import sys
from pathlib import Path

import nflreadpy as nfl
import openpyxl

sys.path.insert(0, str(Path(__file__).parent.parent))

from opfl import OPFLScorer, parse_roster_from_excel
from opfl.constants import CODE_TO_OWNER, resolve_team_code
from opfl.excel_parser import is_valid_player_name, parse_week_sheet_points
from opfl.week_archive import save_week
from opfl.week_status import week_games_are_final


def find_week_sheets(excel_path):
    """Return [(week_number, sheet_name)] for every W-sheet, in week order."""
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    sheets = []
    for name in wb.sheetnames:
        match = re.fullmatch(r'W(\d+)', name)
        if match:
            sheets.append((int(match.group(1)), name))
    wb.close()
    return sorted(sheets)


def _finish_week_data(teams_data):
    for rank, team in enumerate(
        sorted(teams_data, key=lambda t: t['total_score'], reverse=True), 1
    ):
        team['score_rank'] = rank
    return {
        'week': None,  # filled in by the caller
        'teams': teams_data,
        'has_scores': any(t['total_score'] > 0 for t in teams_data),
    }


def score_week_sheet_from_points(excel_path, sheet_name, week_num, season):
    """Build the web week shape from the sheet's own recorded points."""
    teams_data = []
    for team_name, roster in parse_week_sheet_points(excel_path, sheet_name):
        code = resolve_team_code(team_name)
        if not code:
            print(f'    WARNING: could not resolve team name {team_name!r}; skipping')
            continue

        total_score = round(sum(p['score'] for p in roster if p['starter']), 1)
        teams_data.append(
            {
                'name': CODE_TO_OWNER.get(code, team_name),
                'owner': CODE_TO_OWNER.get(code, team_name),
                'abbrev': code,
                'roster': roster,
                'total_score': total_score,
            }
        )

    week_data = _finish_week_data(teams_data)
    week_data['week'] = week_num
    return week_data


def score_week_sheet(excel_path, sheet_name, week_num, season):
    """Score one archived W-sheet by re-deriving points from nflreadpy stats."""
    teams = parse_roster_from_excel(excel_path, sheet_name)
    scorer = OPFLScorer(season, week_num)

    teams_data = []
    for team in teams:
        code = resolve_team_code(team.name)
        if not code:
            print(f'    WARNING: could not resolve team name {team.name!r}; skipping')
            continue

        roster = []
        total_score = 0.0
        scores = scorer.score_fantasy_team(team, starters_only=False)

        for position, player_scores in scores.items():
            for ps in player_scores:
                if not is_valid_player_name(ps.name, position):
                    continue
                roster.append(
                    {
                        'name': ps.name,
                        'nfl_team': ps.team,
                        'position': position,
                        'score': round(ps.total_points, 1),
                        'starter': ps.is_starter,
                    }
                )
                if ps.is_starter:
                    total_score += ps.total_points

        teams_data.append(
            {
                'name': CODE_TO_OWNER.get(code, team.name),
                'owner': CODE_TO_OWNER.get(code, team.name),
                'abbrev': code,
                'roster': roster,
                'total_score': round(total_score, 1),
            }
        )

    week_data = _finish_week_data(teams_data)
    week_data['week'] = week_num
    return week_data


def load_season_schedule(season):
    """Load that season's head-to-head pairings, if we have them.

    W-sheets record lineups but not who played whom, so pairings come from
    data/schedules/{season}.json. Without them a week still contributes points
    and leaders, just not W/L records.
    """
    path = Path(__file__).parent.parent / 'data' / 'schedules' / f'{season}.json'
    if not path.exists():
        print(f'  No data/schedules/{season}.json - weeks will have no matchups (points only)')
        return {}
    return json.loads(path.read_text()).get('weeks', {})


def main():
    parser = argparse.ArgumentParser(description='Rebuild the week archive from W-sheets')
    parser.add_argument('--excel', '-e', required=True, help='Workbook containing the W-sheets')
    parser.add_argument('--season', '-y', type=int, required=True, help='NFL season year')
    parser.add_argument(
        '--weeks', '-w', type=int, nargs='+', help='Only these weeks (default: all W-sheets)'
    )
    parser.add_argument(
        '--force', action='store_true', help='Overwrite weeks already archived as final'
    )
    parser.add_argument('--dry-run', action='store_true', help='Score but do not write')
    parser.add_argument(
        '--rescore',
        action='store_true',
        help='Re-derive points from nflreadpy stats instead of reading the sheet\'s own values',
    )
    args = parser.parse_args()

    project_dir = Path(__file__).parent.parent
    excel_path = Path(args.excel)
    if not excel_path.is_absolute():
        excel_path = project_dir / excel_path
    if not excel_path.exists():
        print(f'Error: {excel_path} not found')
        return 1

    sheets = find_week_sheets(excel_path)
    if args.weeks:
        wanted = set(args.weeks)
        sheets = [(w, name) for w, name in sheets if w in wanted]

    if not sheets:
        print(f'No W-sheets found in {excel_path.name}')
        return 1

    print(
        f'Found {len(sheets)} week sheet(s) in {excel_path.name}: '
        + ', '.join(n for _, n in sheets)
    )

    try:
        schedule_rows = list(nfl.load_schedules(seasons=args.season).iter_rows(named=True))
    except Exception as e:
        print(
            f'Warning: could not load the {args.season} schedule ({e}); weeks will be '
            'archived as not-final'
        )
        schedule_rows = []

    season_schedule = load_season_schedule(args.season)

    score_fn = score_week_sheet if args.rescore else score_week_sheet_from_points

    written = 0
    for week_num, sheet_name in sheets:
        print(f'\nScoring {sheet_name} (week {week_num})...')
        week_data = score_fn(str(excel_path), sheet_name, week_num, args.season)
        is_final = week_games_are_final(schedule_rows, week_num, args.season)
        week_data['final'] = is_final

        totals = ', '.join(
            f'{t["abbrev"]} {t["total_score"]:.0f}'
            for t in sorted(week_data['teams'], key=lambda t: -t['total_score'])
        )
        print(f'  {len(week_data["teams"])} teams ({"final" if is_final else "in progress"})')
        print(f'  {totals}')

        if args.dry_run:
            continue

        pairings = season_schedule.get(str(week_num), [])
        if save_week(args.season, week_num, week_data, pairings, is_final, force=args.force):
            written += 1
        else:
            print('  already archived as final; pass --force to rescore')

    print(f'\nArchived {written} week(s) to data/weeks/{args.season}/')
    return 0


if __name__ == '__main__':
    sys.exit(main())
