#!/usr/bin/env python3
"""Rebuild the week archive from a workbook's W-sheets.

The commissioner archives each finished week to a `W{n}` sheet inside the same
workbook. Those sheets are static values (not formulas) in the same roster-block
layout as the Rosters tab, with that week's starters starred — so a whole season
can be reconstructed from them.

Use this to seed a past season, or to repair the current one if a scoring run
was missed and the Matchups tab has already moved on.

    python scripts/backfill_weeks.py --excel "OPFL Scoring 2025 (5).xlsx" --season 2025
    python scripts/backfill_weeks.py --season 2025 --weeks 3 4 5 --force
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
from opfl.excel_parser import is_valid_player_name
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


def score_week_sheet(excel_path, sheet_name, week_num, season):
    """Score one archived W-sheet into the web week shape."""
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

    for rank, team in enumerate(
        sorted(teams_data, key=lambda t: t['total_score'], reverse=True), 1
    ):
        team['score_rank'] = rank

    return {
        'week': week_num,
        'teams': teams_data,
        'has_scores': any(t['total_score'] > 0 for t in teams_data),
    }


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

    written = 0
    for week_num, sheet_name in sheets:
        print(f'\nScoring {sheet_name} (week {week_num})...')
        week_data = score_week_sheet(str(excel_path), sheet_name, week_num, args.season)
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
