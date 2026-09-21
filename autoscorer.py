#!/usr/bin/env python3
"""
OPFL Autoscorer CLI

Automatically scores fantasy football lineups using nflreadpy for real-time NFL stats.

The 2026 workbook keeps official rosters on the "Rosters" tab and the week's
head-to-head lineups on the "Matchups" tab, so that is the default mode. Older
workbooks with per-week sheets (W1, W2, ...) are still supported via --sheet.

Usage:
    python autoscorer.py --week 1
    python autoscorer.py --excel "Griff OPFL Scoring 2025.xlsx" --sheet W12 --season 2025
"""

import argparse

from opfl import OPFLScorer, build_matchup_week, score_week, update_excel_scores
from opfl.config import get_current_season
from opfl.constants import CODE_TO_OWNER


def score_matchups(args):
    """Score a week from the Rosters + Matchups tabs."""
    teams_by_code, matchups = build_matchup_week(
        args.excel, rosters_sheet=args.rosters_sheet, matchups_sheet=args.matchups_sheet
    )
    scorer = OPFLScorer(args.season, args.week)

    results = {}
    for code, team in teams_by_code.items():
        scores = scorer.score_fantasy_team(team, starters_only=False)

        player_points = {}
        total = 0.0
        for player_scores in scores.values():
            for ps in player_scores:
                if not ps.is_starter:
                    continue
                player_points[ps.name] = round(ps.total_points, 1)
                total += ps.total_points

        results[code] = {'players': player_points, 'total': round(total, 1), 'scores': scores}

        if not args.quiet:
            print(f'\n{"=" * 60}')
            print(f'{CODE_TO_OWNER.get(code, code)}')
            print('=' * 60)
            for position in ['QB', 'RB', 'WR', 'TE', 'K', 'DF', 'HC']:
                for ps in scores.get(position, []):
                    status = '✓' if ps.found_in_stats else '✗'
                    bench = '' if ps.is_starter else ' [BENCH]'
                    matched = (
                        f' -> {ps.matched_name}'
                        if ps.matched_name and ps.matched_name != ps.name
                        else ''
                    )
                    print(
                        f'  {position:3s} {ps.name} ({ps.team}){matched}: '
                        f'{ps.total_points:.1f} pts {status}{bench}'
                    )
                    for key, val in ps.breakdown.items():
                        if key != 'floor_applied':
                            print(f'        {key}: {val}')
                        else:
                            print('        (floor applied - points capped at 0)')
                    for note in ps.data_notes:
                        print(f'        ⚠️  {note}')
            print(f'\n  TOTAL: {total:.1f} points')

    print('\n' + '=' * 60)
    print(f'WEEK {args.week} MATCHUPS')
    print('=' * 60)
    for matchup in matchups:
        code1, code2 = [side['code'] for side in matchup['teams']]
        name1 = CODE_TO_OWNER.get(code1, code1)
        name2 = CODE_TO_OWNER.get(code2, code2)
        print(
            f'  {name1:>14s} {results[code1]["total"]:6.1f}  -  {results[code2]["total"]:<6.1f} {name2}'
        )

    print('\n' + '=' * 60)
    print(f'WEEK {args.week} STANDINGS')
    print('=' * 60)
    ranked = sorted(results.items(), key=lambda kv: kv[1]['total'], reverse=True)
    for rank, (code, result) in enumerate(ranked, 1):
        print(f'  {rank:2d}. {CODE_TO_OWNER.get(code, code)}: {result["total"]:.1f} pts')

    if args.update:
        print(
            '\n--update is not supported for the Matchups tab: the 2026 workbook is\n'
            'formula-driven (Matchups <- Scoring <- Rosters) and openpyxl would strip\n'
            'its cached values. Run scripts/export_for_web.py to publish scores instead.'
        )

    return results


def score_sheet(args):
    """Score a legacy per-week sheet (W1, W2, ...)."""
    print(f'\n{"#" * 60}')
    print(f'# SCORING {args.sheet} (WEEK {args.week})')
    print(f'{"#" * 60}')

    teams, results = score_week(
        excel_path=args.excel,
        sheet_name=args.sheet,
        season=args.season,
        week=args.week,
        verbose=not args.quiet,
    )

    print('\n' + '=' * 60)
    print(f'WEEK {args.week} STANDINGS')
    print('=' * 60)
    for rank, (team_name, (total, _)) in enumerate(
        sorted(results.items(), key=lambda x: x[1][0], reverse=True), 1
    ):
        print(f'  {rank:2d}. {team_name}: {total:.1f} pts')

    if args.update:
        update_excel_scores(args.excel, args.sheet, teams, results)

    return results


def main():
    parser = argparse.ArgumentParser(description='OPFL Fantasy Football Autoscorer')
    parser.add_argument(
        '--excel',
        '-e',
        default=f'OPFL Scoring {get_current_season()}.xlsx',
        help='Path to the Excel file with rosters',
    )
    parser.add_argument(
        '--season',
        '-y',
        type=int,
        default=get_current_season(),
        help='NFL season year',
    )
    parser.add_argument(
        '--week',
        '-w',
        type=int,
        default=1,
        help='Week number to score',
    )
    parser.add_argument(
        '--sheet',
        '-s',
        default=None,
        help='Score a legacy per-week sheet (e.g. W12) instead of the Matchups tab',
    )
    parser.add_argument(
        '--rosters-sheet',
        default='Rosters',
        help='Name of the official rosters sheet',
    )
    parser.add_argument(
        '--matchups-sheet',
        default='Matchups',
        help='Name of the weekly matchups sheet',
    )
    parser.add_argument(
        '--update',
        '-u',
        action='store_true',
        help='Write scores back to the Excel file (legacy W-sheets only)',
    )
    parser.add_argument(
        '--quiet',
        '-q',
        action='store_true',
        help='Suppress per-player output',
    )

    args = parser.parse_args()

    if args.sheet:
        score_sheet(args)
    else:
        score_matchups(args)


if __name__ == '__main__':
    main()
