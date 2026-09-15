#!/usr/bin/env python3
"""Compute all-time and season records for the Hall of Fame page.

Ported conceptually from the sibling QPFL site's scripts/export_hall_of_fame.py
(2694 lines) - this is a from-scratch OPFL-native rewrite of just the generic
parts: player records, team records, league-wide fun stats, and head-to-head
history. QPFL's version is mostly Connor Bowl / franchise-lineage machinery
specific to that league (COMBINED_TEAM_OWNERS, FRANCHISE_LINEAGE,
TEAM_FINISH_PATTERNS) and a multi-era file-format branch for pre-2026 Excel
seasons - none of which OPFL has or needs, since every OPFL season is already
uniform in opfl.week_archive's shape.

Only completed weeks count (week['final']), so a week still in progress can't
set a record with a partial score - the same rule Phase 1 applied to
standings.
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from opfl.week_archive import load_all_weeks

DATA_DIR = Path(__file__).parent.parent / 'data'
WEEKS_DIR = DATA_DIR / 'weeks'

OFFENSIVE_POSITIONS = {'QB', 'RB', 'WR', 'TE'}


def discover_seasons() -> list[int]:
    """Every season with an archive directory under data/weeks/."""
    if not WEEKS_DIR.exists():
        return []
    return sorted(int(p.name) for p in WEEKS_DIR.iterdir() if p.is_dir() and p.name.isdigit())


def week_label(week_num: int, regular_season_weeks: int = 15) -> str:
    if week_num == regular_season_weeks + 1:
        return 'Semifinals'
    if week_num == regular_season_weeks + 2:
        return 'Championship Week'
    return f'Week {week_num}'


def load_all_seasons(seasons: list[int], data_dir: Path | str = DATA_DIR) -> list[dict]:
    """Load every completed week of every season into one flat structure."""
    all_seasons = []
    for season in seasons:
        weeks, schedule = load_all_weeks(season, data_dir=data_dir)
        completed = [w for w in weeks if w.get('final')]
        all_seasons.append({'season': season, 'weeks': completed, 'schedule': schedule})
    return all_seasons


def calculate_player_records(all_seasons: list[dict]) -> dict:
    """Top-5 lists of individual player performances. Starters only."""
    most_points = []
    most_points_non_qb = []
    least_points_offensive = []
    least_points_kicker = []

    for season_data in all_seasons:
        season = season_data['season']
        for week in season_data['weeks']:
            label = week_label(week['week'])
            for team in week['teams']:
                for player in team.get('roster', []):
                    if not player.get('starter'):
                        continue
                    score = player.get('score')
                    if not isinstance(score, (int, float)):
                        continue
                    position = player['position']
                    record = (score, player['name'], team['abbrev'], position, label, season)

                    most_points.append(record)
                    if position != 'QB':
                        most_points_non_qb.append(record)
                    if position in OFFENSIVE_POSITIONS:
                        least_points_offensive.append(record)
                    if position == 'K':
                        least_points_kicker.append(record)

    most_points.sort(key=lambda r: r[0], reverse=True)
    most_points_non_qb.sort(key=lambda r: r[0], reverse=True)
    least_points_offensive.sort(key=lambda r: r[0])
    least_points_kicker.sort(key=lambda r: r[0])

    def fmt(record, include_position=False):
        score, name, abbrev, position, label, season = record
        prefix = f'{position} ' if include_position else ''
        return f'{prefix}{name} ({abbrev}) - {score:.1f} ({label}, {season})'

    return {
        'most_points': [fmt(r) for r in most_points[:5]],
        'most_points_non_qb': [fmt(r) for r in most_points_non_qb[:5]],
        'least_points_offensive': [fmt(r, True) for r in least_points_offensive[:5]],
        'least_points_kicker': [fmt(r, True) for r in least_points_kicker[:5]],
    }


def calculate_team_records(all_seasons: list[dict]) -> dict:
    """Top-5 lists of team weekly performances and margins of victory."""
    most_points = []
    least_points = []
    margins = []

    for season_data in all_seasons:
        season = season_data['season']
        schedule = season_data['schedule']
        for week in season_data['weeks']:
            label = week_label(week['week'])
            teams_by_abbrev = {t['abbrev']: t for t in week['teams']}
            pairings = schedule.get(str(week['week']), [])

            for team in week['teams']:
                score = team.get('total_score')
                if isinstance(score, (int, float)) and score > 0:
                    most_points.append((score, team['name'], team['abbrev'], label, season))
                    least_points.append((score, team['name'], team['abbrev'], label, season))

            for abbrev1, abbrev2 in pairings:
                margin_record = _margin_record(teams_by_abbrev, abbrev1, abbrev2, label, season)
                if margin_record:
                    margins.append(margin_record)

    most_points.sort(key=lambda r: r[0], reverse=True)
    least_points.sort(key=lambda r: r[0])
    margins.sort(key=lambda r: r[0], reverse=True)

    def fmt_team(record):
        score, name, abbrev, label, season = record
        return f'{name} ({abbrev}) - {score:.1f} ({label}, {season})'

    def fmt_margin(record):
        margin, w_name, w_abbrev, l_name, l_abbrev, label, season = record
        return f'{w_name} ({w_abbrev}) over {l_name} ({l_abbrev}) - by {margin:.1f} ({label}, {season})'

    return {
        'most_points': [fmt_team(r) for r in most_points[:5]],
        'least_points': [fmt_team(r) for r in least_points[:5]],
        'largest_margin': [fmt_margin(r) for r in margins[:5]],
    }


def _margin_record(teams_by_abbrev, abbrev1, abbrev2, label, season):
    """(margin, winner_name, winner_abbrev, loser_name, loser_abbrev, label, season),
    or None if either side's score is missing/non-positive."""
    t1 = teams_by_abbrev.get(abbrev1)
    t2 = teams_by_abbrev.get(abbrev2)
    if not t1 or not t2:
        return None
    s1, s2 = t1.get('total_score'), t2.get('total_score')
    if not isinstance(s1, (int, float)) or not isinstance(s2, (int, float)):
        return None
    if s1 <= 0 or s2 <= 0:
        return None
    winner, w_score, loser, l_score = (t1, s1, t2, s2) if s1 > s2 else (t2, s2, t1, s1)
    return (
        w_score - l_score,
        winner['name'],
        winner['abbrev'],
        loser['name'],
        loser['abbrev'],
        label,
        season,
    )


def calculate_fun_stats(all_seasons: list[dict]) -> list[dict]:
    """League-wide superlatives: highest/lowest scoring weeks, closest games."""
    weekly_totals = []
    closest_games = []

    for season_data in all_seasons:
        season = season_data['season']
        schedule = season_data['schedule']
        for week in season_data['weeks']:
            label = week_label(week['week'])
            scores = [
                t['total_score']
                for t in week['teams']
                if isinstance(t.get('total_score'), (int, float))
            ]
            if scores and sum(scores) > 0:
                weekly_totals.append((sum(scores), label, season))

            teams_by_abbrev = {t['abbrev']: t for t in week['teams']}
            pairings = schedule.get(str(week['week']), [])
            for abbrev1, abbrev2 in pairings:
                margin_record = _margin_record(teams_by_abbrev, abbrev1, abbrev2, label, season)
                if margin_record:
                    margin, w_name, w_abbrev, l_name, l_abbrev, label2, season2 = margin_record
                    t1 = teams_by_abbrev[abbrev1]
                    closest_games.append(
                        (margin, w_name, w_abbrev, l_name, l_abbrev, label2, season2, t1)
                    )

    fun_stats = []

    weekly_totals.sort(key=lambda x: x[0], reverse=True)
    fun_stats.append(
        {
            'title': 'Highest Scoring Week (League Total)',
            'records': [f'{t[0]:.1f} points ({t[1]}, {t[2]})' for t in weekly_totals[:3]],
        }
    )

    weekly_totals.sort(key=lambda x: x[0])
    fun_stats.append(
        {
            'title': 'Lowest Scoring Week (League Total)',
            'records': [f'{t[0]:.1f} points ({t[1]}, {t[2]})' for t in weekly_totals[:3]],
        }
    )

    closest_games.sort(key=lambda g: g[0])
    fun_stats.append(
        {
            'title': 'Closest Games',
            'records': [
                f'{g[1]} ({g[2]}) over {g[3]} ({g[4]}) - margin {g[0]:.1f} ({g[5]}, {g[6]})'
                for g in closest_games[:5]
            ],
        }
    )

    return fun_stats


def calculate_head_to_head(all_seasons: list[dict]) -> list[dict]:
    """All-time head-to-head record for every pair of franchises that has played."""
    pairs = {}

    for season_data in all_seasons:
        schedule = season_data['schedule']
        for week in season_data['weeks']:
            teams_by_abbrev = {t['abbrev']: t for t in week['teams']}
            pairings = schedule.get(str(week['week']), [])
            for abbrev1, abbrev2 in pairings:
                t1 = teams_by_abbrev.get(abbrev1)
                t2 = teams_by_abbrev.get(abbrev2)
                if not t1 or not t2:
                    continue
                s1, s2 = t1.get('total_score'), t2.get('total_score')
                if not isinstance(s1, (int, float)) or not isinstance(s2, (int, float)):
                    continue

                key = tuple(sorted((abbrev1, abbrev2)))
                if key not in pairs:
                    pairs[key] = {
                        'team1': key[0],
                        'team2': key[1],
                        'team1_wins': 0,
                        'team2_wins': 0,
                        'ties': 0,
                        'team1_pf': 0.0,
                        'team2_pf': 0.0,
                        'games': 0,
                    }
                rec = pairs[key]
                score_for = {abbrev1: s1, abbrev2: s2}
                s_a, s_b = score_for[key[0]], score_for[key[1]]
                rec['team1_pf'] = round(rec['team1_pf'] + s_a, 1)
                rec['team2_pf'] = round(rec['team2_pf'] + s_b, 1)
                rec['games'] += 1
                if s_a > s_b:
                    rec['team1_wins'] += 1
                elif s_b > s_a:
                    rec['team2_wins'] += 1
                else:
                    rec['ties'] += 1

    records = []
    for rec in pairs.values():
        leader = None
        if rec['team1_wins'] > rec['team2_wins']:
            leader = rec['team1']
        elif rec['team2_wins'] > rec['team1_wins']:
            leader = rec['team2']
        records.append({**rec, 'leader': leader})

    records.sort(key=lambda r: r['games'], reverse=True)
    return records


def generate_hall_of_fame(seasons: list[int] | None = None) -> dict:
    seasons = seasons if seasons is not None else discover_seasons()
    all_seasons = load_all_seasons(seasons)

    completed_through = {}
    for season_data in all_seasons:
        weeks = [w['week'] for w in season_data['weeks']]
        completed_through[str(season_data['season'])] = max(weeks, default=0)

    return {
        'seasons': seasons,
        'completed_through': completed_through,
        'player_records': calculate_player_records(all_seasons),
        'team_records': calculate_team_records(all_seasons),
        'fun_stats': calculate_fun_stats(all_seasons),
        'head_to_head': calculate_head_to_head(all_seasons),
    }


def main():
    parser = argparse.ArgumentParser(description='Compute OPFL Hall of Fame records')
    parser.add_argument('--output', '-o', default=None, help='Write JSON here instead of stdout')
    args = parser.parse_args()

    hall_of_fame = generate_hall_of_fame()

    if args.output:
        Path(args.output).write_text(json.dumps(hall_of_fame, indent=2))
        print(f'Wrote {args.output}')
    else:
        print(json.dumps(hall_of_fame, indent=2))


if __name__ == '__main__':
    main()
