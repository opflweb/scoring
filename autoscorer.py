#!/usr/bin/env python3
"""
OPFL Autoscorer CLI

Automatically scores fantasy football lineups using nflreadpy for real-time NFL stats.

Usage:
    python autoscorer.py --excel "OPFL Scoring 2025.xlsx" --week 12 --update
    python autoscorer.py --all-weeks --update  # Score all weeks 1-14
"""

import argparse
import re

import openpyxl

from opfl import score_week, update_excel_scores


def get_available_weeks(excel_path):
    """Get list of week numbers that have sheets in the Excel file."""
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    weeks = []
    for sheet_name in wb.sheetnames:
        match = re.match(r'^W(\d+)$', sheet_name)
        if match:
            weeks.append(int(match.group(1)))
    wb.close()
    return sorted(weeks)


def score_single_week(args, week_num):
    """Score a single week and optionally update Excel."""
    sheet_name = f"W{week_num}"
    
    print(f"\n{'#'*60}")
    print(f"# SCORING WEEK {week_num}")
    print(f"{'#'*60}")
    
    teams, results = score_week(
        excel_path=args.excel,
        sheet_name=sheet_name,
        season=args.season,
        week=week_num,
        verbose=not args.quiet,
    )
    
    # Print summary
    print("\n" + "="*60)
    print(f"WEEK {week_num} STANDINGS")
    print("="*60)
    
    sorted_results = sorted(results.items(), key=lambda x: x[1][0], reverse=True)
    for rank, (team_name, (total, _)) in enumerate(sorted_results, 1):
        print(f"  {rank}. {team_name}: {total:.1f} pts")
    
    # Update Excel if requested
    if args.update:
        update_excel_scores(args.excel, sheet_name, teams, results)
    
    return teams, results


def main():
    parser = argparse.ArgumentParser(description="OPFL Fantasy Football Autoscorer")
    parser.add_argument(
        "--excel", "-e",
        default="Griff OPFL Scoring 2025.xlsx",
        help="Path to the Excel file with rosters",
    )
    parser.add_argument(
        "--season", "-y",
        type=int,
        default=2025,
        help="NFL season year",
    )
    parser.add_argument(
        "--week", "-w",
        type=int,
        default=None,
        help="Week number to score (e.g., 12). If not specified with --all-weeks, defaults to current week.",
    )
    parser.add_argument(
        "--all-weeks", "-a",
        action="store_true",
        help="Score all available weeks in the Excel file",
    )
    parser.add_argument(
        "--update", "-u",
        action="store_true",
        help="Update Excel file with scores",
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress detailed output",
    )
    
    args = parser.parse_args()
    
    if args.all_weeks:
        # Score all available weeks
        weeks = get_available_weeks(args.excel)
        print(f"Found {len(weeks)} weeks to score: {weeks}")
        
        for week_num in weeks:
            try:
                score_single_week(args, week_num)
            except Exception as e:
                print(f"Error scoring week {week_num}: {e}")
                continue
        
        print("\n" + "="*60)
        print("ALL WEEKS SCORED!")
        print("="*60)
    else:
        # Score single week
        week_num = args.week if args.week is not None else 14  # Default to week 14
        score_single_week(args, week_num)


if __name__ == "__main__":
    main()
