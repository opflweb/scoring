#!/usr/bin/env python3
"""Export OPFL Excel scores to JSON for web display."""

import argparse
import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import nflreadpy as nfl
import openpyxl

# Add project root to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from export_hall_of_fame import generate_hall_of_fame

from opfl import OPFLScorer, build_matchup_week, parse_taxi_squads
from opfl.config import get_config
from opfl.constants import ALL_TEAM_CODES, CODE_TO_OWNER, resolve_team_code
from opfl.projections import calculate_week_projections
from opfl.week_archive import load_all_weeks, load_week, save_week
from opfl.week_status import week_games_are_final

# Single source of truth for these is data/league_config.json - a season
# rollover or roster-shape change only needs to touch that file.
_config = get_config()

ALL_TEAMS = ALL_TEAM_CODES
TEAM_COLUMNS = [4, 7, 10, 13, 16, 19]
POSITIONS = _config.positions
TRADE_DEADLINE_WEEK = _config.trade_deadline_week
REGULAR_SEASON_WEEKS = _config.regular_season_weeks
PLAYOFF_WEEKS = _config.playoff_weeks

SEASON = _config.current_season
ROSTERS_SHEET = _config.rosters_sheet
MATCHUPS_SHEET = _config.matchups_sheet

# Draft pick defaults: 6 preseason rounds and 3 waiver rounds per team per season.
DEFAULT_PRESEASON_ROUNDS = list(range(1, 7))
DEFAULT_WAIVER_ROUNDS = list(range(1, 4))
DRAFT_PICK_SEASONS = ['2026', '2027', '2028']

# 2026 schedule team numbers, taken from the "Key" column on the Matchups tab.
# Weekly pairings are read from the workbook rather than hardcoded here.
TEAM_NUMBER_MAP = {
    1: 'AND',
    2: 'JOH',
    3: 'W/B',
    4: 'STL',
    5: 'G/G',
    6: 'J/M',
    7: 'KEV',
    8: 'KAM',
    9: 'K/D',
    10: 'E/J',
    11: 'ADA',
    12: 'D/J',
}


def parse_player_name(cell_value):
    if not cell_value:
        return '', ''
    cell_value = str(cell_value).strip()
    match = re.match(r'^(.+?)\s*\(([A-Za-z]{2,3})\)$', cell_value)
    if match:
        return match.group(1).strip(), match.group(2).upper()
    return cell_value, ''


def is_valid_player_name(name, position):
    """Check if a player name is valid (not just a number or empty)."""
    if not name:
        return False

    # Filter out numeric entries (scores, jersey numbers, etc.)
    # These can appear as integers, floats, or string representations thereof
    try:
        float(name)
        # If we get here, the name is purely numeric - not a valid player/coach name
        return False
    except (ValueError, TypeError):
        pass

    # Filter out phone numbers (patterns like "325-1289" or "325-1289 (J)")
    import re

    if re.match(r'^\d{3}-\d{4}', name):
        return False

    # For HC position, do additional validation
    if position == 'HC':
        # Head coach names should contain letters
        if not any(c.isalpha() for c in name):
            return False
        # Name should start with a letter (coach names don't start with numbers)
        if not name[0].isalpha():
            return False

    return True


def find_position_rows(ws, start_row, end_row):
    position_rows = {}
    current_position = None
    for row in range(start_row, min(end_row + 1, ws.max_row + 1)):
        cell_value = ws.cell(row=row, column=1).value
        if cell_value:
            cell_str = str(cell_value).strip().upper()
            if cell_str in POSITIONS:
                current_position = cell_str
                if current_position not in position_rows:
                    position_rows[current_position] = []
                position_rows[current_position].append(row)
        elif current_position:
            has_content = any(ws.cell(row=row, column=c).value for c in TEAM_COLUMNS)
            if has_content:
                position_rows[current_position].append(row)
    return position_rows


def extract_team_name(header_value):
    if not header_value:
        return '', ''
    match = re.match(r'^(.+?)\s*\((\d+)\)$', str(header_value).strip())
    if match:
        return match.group(1).strip(), match.group(2)
    return str(header_value).strip(), ''


def export_matchup_week(excel_path, week_num, season=SEASON):
    """
    Build a week of web data from the Rosters + Matchups tabs.

    The Rosters tab supplies each team's full roster; the Matchups tab supplies
    both the week's head-to-head pairings and which players actually started.

    Returns:
        (week_dict, pairings) where pairings is a list of [abbrev1, abbrev2].
    """
    teams_by_code, matchups = build_matchup_week(
        excel_path, rosters_sheet=ROSTERS_SHEET, matchups_sheet=MATCHUPS_SHEET
    )
    scorer = OPFLScorer(season, week_num)

    teams_data = []
    for code, team in teams_by_code.items():
        owner = CODE_TO_OWNER.get(code, team.name)
        roster = []
        total_score = 0.0

        # Score every player so the site can show bench points too.
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
                # Only starters count toward the team total.
                if ps.is_starter:
                    total_score += ps.total_points

        teams_data.append(
            {
                'name': owner,
                'owner': owner,
                'abbrev': code,
                'roster': roster,
                'total_score': round(total_score, 1),
            }
        )

    sorted_by_score = sorted(teams_data, key=lambda t: t['total_score'], reverse=True)
    for rank, team in enumerate(sorted_by_score, 1):
        team['score_rank'] = rank

    pairings = [[side['code'] for side in m['teams']] for m in matchups]

    print(f'  Week {week_num}: scored {len(teams_data)} teams across {len(pairings)} matchups')

    return {
        'week': week_num,
        'teams': teams_data,
        'has_scores': any(t['total_score'] > 0 for t in teams_data),
    }, pairings


def get_current_nfl_week():
    return nfl.get_current_week()


def load_schedule_rows(season=SEASON):
    """Fetch the season's NFL schedule once; both week-finality and kickoff
    times read from it, and it is the slowest call in the export."""
    try:
        return list(nfl.load_schedules(seasons=season).iter_rows(named=True))
    except Exception as e:
        print(f'Warning: could not load the {season} schedule: {e}')
        return []


def load_season_schedule(season=SEASON):
    """Load the full-season fixture list (week -> [[code, code], ...]).

    W-sheets and the Matchups tab only carry pairings for weeks that have
    already been scored, so future weeks fall back to data/schedules/{season}.json
    (the printed schedule, translated from team numbers to this year's teams).
    """
    path = Path(__file__).parent.parent / 'data' / 'schedules' / f'{season}.json'
    if not path.exists():
        return {}
    return json.loads(path.read_text()).get('weeks', {})


def load_projection_schedule_rows(season=SEASON):
    """Fetch this season's and last season's NFL schedules.

    Projections need the prior season's opponents too, since its archived
    weeks feed the history the model is built on.
    """
    try:
        return list(nfl.load_schedules(seasons=[season - 1, season]).iter_rows(named=True))
    except Exception as e:
        print(f'Warning: could not load schedule context for projections: {e}')
        return []


def build_game_times(schedule_rows):
    """Map week -> NFL team -> kickoff, for the frontend's lineup-lock display."""
    game_times = {}
    eastern = ZoneInfo('America/New_York')
    for row in schedule_rows:
        week = row.get('week')
        game_date = row.get('gameday', '')
        game_time = row.get('gametime', '')
        if not week or not game_date or not game_time:
            continue
        try:
            dt = datetime.strptime(f'{game_date} {game_time}', '%Y-%m-%d %H:%M').replace(
                tzinfo=eastern
            )
        except (ValueError, TypeError):
            continue
        kickoff_iso = dt.isoformat(timespec='seconds')
        slot = game_times.setdefault(week, {})
        for side in ('home_team', 'away_team'):
            if row.get(side):
                slot[row[side]] = kickoff_iso
    return game_times


def parse_draft_picks(excel_path):
    picks = {}
    for team in ALL_TEAMS:
        picks[team] = {}
        for season in DRAFT_PICK_SEASONS:
            picks[team][season] = {
                'preseason': [(r, team) for r in DEFAULT_PRESEASON_ROUNDS],
                'waiver': [(r, team) for r in DEFAULT_WAIVER_ROUNDS],
            }

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb['Future Traded Picks']
    except Exception as e:
        print(f'Warning: Could not load traded picks: {e}')
        return format_picks_for_output(picks)

    pattern = r"([A-Za-z/\.\s]+)\s+holds?\s+([A-Za-z/\.\s]+)'s\s+(\d{4})\s+#(\d+)\s+pick"

    for row in range(3, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=2).value
        if not cell_value:
            continue
        match = re.search(pattern, str(cell_value).strip(), re.IGNORECASE)
        if match:
            holder_name = match.group(1).strip()
            original_name = match.group(2).strip()
            season = match.group(3)
            round_num = int(match.group(4))

            holder = resolve_team_code(holder_name)
            original = resolve_team_code(original_name)

            if holder and original and season in DRAFT_PICK_SEASONS:
                draft_type = 'preseason' if round_num <= 6 else 'waiver'
                original_picks = picks[original][season][draft_type]
                for i, (r, owner) in enumerate(original_picks):
                    if r == round_num and owner == original:
                        original_picks.pop(i)
                        break
                picks[holder][season][draft_type].append((round_num, original))

    wb.close()
    return format_picks_for_output(picks)


def format_picks_for_output(picks):
    formatted = {}
    for team in ALL_TEAMS:
        formatted[team] = {}
        for season in picks[team]:
            formatted[team][season] = {}
            for draft_type in picks[team][season]:
                team_picks = sorted(picks[team][season][draft_type], key=lambda x: (x[0], x[1]))
                formatted[team][season][draft_type] = [
                    {'round': r, 'from': owner, 'own': owner == team} for r, owner in team_picks
                ]
    return formatted


def get_existing_banners(banners_dir):
    """Get list of banner images from directory, sorted by year descending."""
    if not os.path.exists(banners_dir):
        return []
    images = [f for f in os.listdir(banners_dir) if f.endswith('.png')]

    def get_year(filename):
        try:
            return int(filename.replace('.png', ''))
        except ValueError:
            return 0

    return sorted(images, key=get_year, reverse=True)


def resolve_matchups_week(excel_path, requested_week, season, data_dir=None):
    """Detect when the workbook hasn't caught up to nflreadpy's current week yet.

    The Matchups tab carries no explicit week marker, so when no --week is
    given we default to nfl.get_current_week() - a calendar-based guess, not
    a read of the workbook. That guess can outrun the commissioner: nflreadpy
    advances the moment the calendar crosses into the next week, which can be
    a day or more before the Matchups tab is actually updated. Scoring
    requested_week against a tab that still shows requested_week - 1's
    lineups would silently score every player under the wrong week's NFL
    stats.

    Compares the tab's current starters against the previously archived
    week; if they're identical, the workbook hasn't been rolled over and this
    is still that prior week.
    """
    if requested_week <= 1:
        return requested_week

    teams_by_code, _ = build_matchup_week(
        excel_path, rosters_sheet=ROSTERS_SHEET, matchups_sheet=MATCHUPS_SHEET
    )
    current_starters = {
        code: sorted(
            name
            for players in team.players.values()
            for name, _nfl_team, started in players
            if started
        )
        for code, team in teams_by_code.items()
    }

    load_week_kwargs = {'data_dir': data_dir} if data_dir is not None else {}
    previous = load_week(season, requested_week - 1, **load_week_kwargs)
    if not previous:
        return requested_week

    previous_starters = {
        t['abbrev']: sorted(p['name'] for p in t['roster'] if p['starter'])
        for t in previous['teams']
    }

    if current_starters == previous_starters:
        print(
            f"  Matchups tab still shows week {requested_week - 1}'s lineups "
            f'(nflreadpy says week {requested_week}) - scoring week {requested_week - 1} again'
        )
        return requested_week - 1
    return requested_week


def export_season(excel_path, week_num=None, season=SEASON, force_rescore=False):
    """Build the full data.json payload for the season."""
    schedule_rows = load_schedule_rows(season)
    current_nfl_week = get_current_nfl_week()

    # Always run the stale-Matchups-tab check, not just when week_num was left
    # to default: the workflow passes an explicit --week computed from
    # nflreadpy's calendar-based current week, which is exactly the guess
    # resolve_matchups_week exists to double-check against the workbook. An
    # explicit --week that skipped this check is how a week got archived
    # with the previous week's lineups and pairings silently relabeled.
    requested_week = week_num if week_num is not None else current_nfl_week
    week_num = resolve_matchups_week(excel_path, requested_week, season)

    print(f'Scoring week {week_num} from the {MATCHUPS_SHEET} tab...')
    week_data, pairings = export_matchup_week(excel_path, week_num, season)

    # A week only counts once every NFL game in it has a final result. Mid-week,
    # a matchup where one manager's Thursday starter has played and the other's
    # have not is an unfinished game, not a result.
    is_final = week_games_are_final(schedule_rows, week_num, season)
    week_data['final'] = is_final

    # Archive eagerly: the Matchups tab is overwritten each week, so a week not
    # captured before then is only recoverable from its W-sheet.
    written = save_week(season, week_num, week_data, pairings, is_final, force=force_rescore)
    state = 'final' if is_final else 'in progress'
    print(f'  Week {week_num} is {state}' + ('' if written else ' (already archived as final)'))

    # Only the live week needs projections - a final week's rosters already
    # have real scores, and projecting them would just be wasted work.
    if not is_final:
        try:
            projection_schedule_rows = load_projection_schedule_rows(season)
            week_data['projections'] = calculate_week_projections(
                week_data, pairings, season, week_num, projection_schedule_rows
            )
        except Exception as e:
            print(f'  Could not build week projections: {e}')

    # The archive is the season; the live week we just scored overrides its own
    # entry so an in-progress week still shows current scores.
    weeks, schedule = load_all_weeks(season)
    weeks = [w for w in weeks if w['week'] != week_num] + [week_data]
    weeks.sort(key=lambda w: w['week'])
    schedule[str(week_num)] = pairings

    # Weeks that haven't been played yet have no archived or live pairings -
    # show the printed schedule for those so the site doesn't say "not set".
    for future_week, future_pairings in load_season_schedule(season).items():
        schedule.setdefault(future_week, future_pairings)

    # Standings only count completed weeks.
    completed = [w for w in weeks if w.get('final')]
    standings = build_standings(completed, schedule)
    standings_through = max((w['week'] for w in completed), default=0)
    team_stats = calculate_team_stats(completed, schedule, standings)

    # Playoffs only exist once the regular season is in the books.
    playoffs = None
    if current_nfl_week > REGULAR_SEASON_WEEKS and len(standings) >= 4:
        playoffs = compute_playoff_data(weeks, standings, current_nfl_week)

    return {
        'updated_at': datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
        'season': season,
        'current_week': current_nfl_week,
        'regular_season_weeks': REGULAR_SEASON_WEEKS,
        'weeks': weeks,
        'schedule': schedule,
        'team_number_map': {str(k): v for k, v in TEAM_NUMBER_MAP.items()},
        'standings': standings,
        'standings_through_week': standings_through,
        'standings_in_progress': not is_final,
        'team_stats': team_stats,
        'playoffs': playoffs,
        'hall_of_fame': generate_hall_of_fame(),
        'game_times': build_game_times(schedule_rows),
        'trade_deadline_week': TRADE_DEADLINE_WEEK,
        'taxi_squads': parse_taxi_squads(excel_path, ROSTERS_SHEET),
        'previous_seasons': build_previous_seasons(season),
    }


def build_previous_seasons(current_season):
    """Season-level archives for years before the live one.

    Built entirely from data/weeks/ and data/schedules/ - no workbook needed,
    since those seasons are done and archived. A year with no archived weeks
    (no data/weeks/{year}/ directory, or none marked final) is left out rather
    than shown empty.
    """
    previous = {}
    weeks_root = Path(__file__).parent.parent / 'data' / 'weeks'
    if not weeks_root.exists():
        return previous

    for season_dir in sorted(weeks_root.iterdir()):
        if not season_dir.is_dir():
            continue
        try:
            season = int(season_dir.name)
        except ValueError:
            continue
        if season >= current_season:
            continue

        weeks, schedule = load_all_weeks(season)
        completed = [w for w in weeks if w.get('final')]
        if not completed:
            continue

        standings = build_standings(completed, schedule)
        previous[str(season)] = {
            'season': season,
            'weeks': completed,
            'schedule': schedule,
            'standings': standings,
        }

    return previous


def build_standings(weeks, schedule):
    """
    Build standings from scored weeks using the pairings read from the workbook.

    Rank points = 1 per win, 0.5 per tie, plus 0.5 for each top-half (top 6)
    scoring finish in a week.
    """
    standings = {}

    for week_data in weeks:
        if not week_data.get('has_scores'):
            continue

        week_num = week_data['week']
        pairings = schedule.get(str(week_num), [])

        team_scores = {}
        for team in week_data['teams']:
            abbrev = team['abbrev']
            team_scores[abbrev] = team['total_score']
            if abbrev not in standings:
                standings[abbrev] = {
                    'name': team['name'],
                    'owner': team['owner'],
                    'abbrev': abbrev,
                    'rank_points': 0.0,
                    'wins': 0,
                    'losses': 0,
                    'ties': 0,
                    'top_half': 0,
                    'points_for': 0.0,
                    'points_against': 0.0,
                }
            standings[abbrev]['points_for'] += team['total_score']

        # Head-to-head results
        for abbrev1, abbrev2 in pairings:
            if abbrev1 not in team_scores or abbrev2 not in team_scores:
                continue

            score1 = team_scores[abbrev1]
            score2 = team_scores[abbrev2]

            standings[abbrev1]['points_against'] += score2
            standings[abbrev2]['points_against'] += score1

            if score1 > score2:
                standings[abbrev1]['wins'] += 1
                standings[abbrev1]['rank_points'] += 1
                standings[abbrev2]['losses'] += 1
            elif score2 > score1:
                standings[abbrev2]['wins'] += 1
                standings[abbrev2]['rank_points'] += 1
                standings[abbrev1]['losses'] += 1
            else:
                standings[abbrev1]['ties'] += 1
                standings[abbrev1]['rank_points'] += 0.5
                standings[abbrev2]['ties'] += 1
                standings[abbrev2]['rank_points'] += 0.5

        # Top 6 scoring bonus (1 full rank point each), split evenly across
        # ties that straddle the cutoff.
        teams_by_score = sorted(week_data['teams'], key=lambda x: x['total_score'], reverse=True)
        current_rank = 1
        i = 0
        while i < len(teams_by_score):
            current_score = teams_by_score[i]['total_score']
            tied_teams = []
            while i < len(teams_by_score) and teams_by_score[i]['total_score'] == current_score:
                tied_teams.append(teams_by_score[i])
                i += 1

            tied_positions = list(range(current_rank, current_rank + len(tied_teams)))
            positions_in_top6 = [p for p in tied_positions if p <= 6]

            if positions_in_top6:
                points_per_team = (1.0 * len(positions_in_top6)) / len(tied_teams)
                for team in tied_teams:
                    standings[team['abbrev']]['rank_points'] += points_per_team
                    standings[team['abbrev']]['top_half'] += 1

            current_rank += len(tied_teams)

    return sorted(
        standings.values(),
        key=lambda x: (x['rank_points'], x['points_for']),
        reverse=True,
    )


def calculate_lineup_efficiency(team):
    """Compare a team's submitted starters with its best legal lineup at the
    same slot counts - "points left on the table" from a suboptimal start.

    Ported from the sibling QPFL site's scripts/export_for_web.py. Unlike that
    version this doesn't filter a `taxi` flag: OPFL's roster entries never
    include taxi-squad players, since parse_roster_from_excel stops at the
    `TS` block rather than reading into it.
    """
    roster = team.get('roster', [])
    if not isinstance(roster, list):
        return None

    slot_counts = defaultdict(int)
    players_by_position = defaultdict(list)
    actual_points = 0.0

    for player in roster:
        score = player.get('score')
        if not isinstance(score, (int, float)):
            continue
        position = player.get('position', '')
        players_by_position[position].append(score)
        if player.get('starter'):
            slot_counts[position] += 1
            actual_points += score

    if not slot_counts:
        return None

    optimal_points = 0.0
    for position, count in slot_counts.items():
        scores = sorted(players_by_position[position], reverse=True)
        optimal_points += sum(scores[:count])

    if optimal_points <= 0:
        return None

    return {
        'actual_points': round(actual_points, 1),
        'optimal_points': round(optimal_points, 1),
        'points_left_on_table': round(max(0.0, optimal_points - actual_points), 1),
    }


def _add_lineup_efficiency(stats, team):
    efficiency = calculate_lineup_efficiency(team)
    if not efficiency:
        return
    stats['lineup_actual_points'] += efficiency['actual_points']
    stats['lineup_optimal_points'] += efficiency['optimal_points']
    stats['points_left_on_table'] += efficiency['points_left_on_table']
    stats['lineup_weeks'] += 1


def calculate_team_stats(weeks, schedule, standings):
    """Per-team season stats: PPG, margins, weekly ranks, streaks, best/worst
    week, lineup efficiency, and the OPR power ranking.

    Ported from the sibling QPFL site's scripts/export_for_web.py. That
    version reads week.matchups[].team1/team2; OPFL's weeks are flatter
    (week['teams'] + the top-level `schedule` dict for pairings, matching how
    build_standings() already reads them above), so pairings are looked up
    the same way rather than duplicated onto every week.

    Args:
        weeks: Archived weeks (opfl.week_archive shape - 'week', 'teams', 'final')
        schedule: {week_str: [[abbrev1, abbrev2], ...]}
        standings: Output of build_standings() - the authoritative W/L/T/PF/PA

    Returns:
        Dict keyed by team abbrev.
    """
    import statistics

    team_stats = {}
    for standing in standings:
        abbrev = standing['abbrev']
        team_stats[abbrev] = {
            'abbrev': abbrev,
            'name': standing.get('name', abbrev),
            'points_for': [],
            'points_against': [],
            'margins': [],
            'weekly_ranks': [],
            'week_numbers': [],
            'wins': 0,
            'losses': 0,
            'ties': 0,
            'streak': {'type': None, 'count': 0},
            'current_streak': [],
            'lineup_actual_points': 0.0,
            'lineup_optimal_points': 0.0,
            'points_left_on_table': 0.0,
            'lineup_weeks': 0,
        }

    if not standings:
        return {}

    # Regular-season length from the standings W+L+T, so a partial season
    # (or a future change to REGULAR_SEASON_WEEKS) doesn't need a matching
    # code change here. Standings only ever include completed weeks, so this
    # is 0 before anything finishes - fall back to the configured length.
    s0 = standings[0]
    reg_season_weeks = (s0.get('wins') or 0) + (s0.get('losses') or 0) + (s0.get('ties') or 0)
    if reg_season_weeks == 0:
        reg_season_weeks = REGULAR_SEASON_WEEKS

    regular_weeks = [
        w for w in weeks if w.get('final') and (w.get('week') or 0) <= reg_season_weeks
    ]

    for week_data in regular_weeks:
        week_num = week_data['week']
        pairings = schedule.get(str(week_num), [])
        teams_by_abbrev = {t['abbrev']: t for t in week_data['teams']}

        weekly_scores = sorted(
            ((t['abbrev'], t['total_score']) for t in week_data['teams']),
            key=lambda x: x[1],
            reverse=True,
        )
        rank_map = {abbrev: rank + 1 for rank, (abbrev, _) in enumerate(weekly_scores)}

        for abbrev1, abbrev2 in pairings:
            t1 = teams_by_abbrev.get(abbrev1)
            t2 = teams_by_abbrev.get(abbrev2)
            if not t1 or not t2:
                continue
            s1, s2 = t1['total_score'], t2['total_score']

            for abbrev, team, score, opp_score in ((abbrev1, t1, s1, s2), (abbrev2, t2, s2, s1)):
                if abbrev not in team_stats:
                    continue
                stats = team_stats[abbrev]
                stats['points_for'].append(score)
                stats['points_against'].append(opp_score)
                stats['week_numbers'].append(week_num)
                margin = score - opp_score
                stats['margins'].append(margin)
                if abbrev in rank_map:
                    stats['weekly_ranks'].append(rank_map[abbrev])

                if margin > 0:
                    stats['wins'] += 1
                    stats['current_streak'].append('W')
                elif margin < 0:
                    stats['losses'] += 1
                    stats['current_streak'].append('L')
                else:
                    stats['ties'] += 1
                    stats['current_streak'].append('T')
                _add_lineup_efficiency(stats, team)

    # Standings are the authoritative W/L/T/PF/PA (they may reflect manual
    # corrections); override the values accumulated from raw matchups.
    standings_by_abbrev = {s['abbrev']: s for s in standings}

    for abbrev, stats in team_stats.items():
        pf = stats['points_for']
        pa = stats['points_against']
        margins = stats['margins']
        ranks = stats['weekly_ranks']
        week_numbers = stats['week_numbers']

        games_played = len(pf)
        if games_played == 0:
            continue

        std = standings_by_abbrev.get(abbrev, {})
        stats['wins'] = std.get('wins', stats['wins'])
        stats['losses'] = std.get('losses', stats['losses'])
        stats['ties'] = std.get('ties', stats['ties'])
        stats['total_points_for'] = std.get('points_for', sum(pf))
        stats['total_points_against'] = std.get('points_against', sum(pa))
        stats['point_differential'] = round(
            stats['total_points_for'] - stats['total_points_against'], 1
        )

        std_games = stats['wins'] + stats['losses'] + stats['ties']
        denom = std_games if std_games > 0 else games_played
        stats['ppg'] = round(stats['total_points_for'] / denom, 2)
        stats['ppg_against'] = round(stats['total_points_against'] / denom, 2)
        stats['avg_margin'] = round(sum(margins) / games_played, 2)
        stats['avg_rank'] = round(sum(ranks) / len(ranks), 2) if ranks else 0

        stats['std_dev'] = round(statistics.stdev(pf), 2) if len(pf) > 1 else 0

        best_idx = pf.index(max(pf))
        worst_idx = pf.index(min(pf))
        stats['best_week'] = pf[best_idx]
        stats['worst_week'] = pf[worst_idx]
        stats['best_week_num'] = week_numbers[best_idx]
        stats['worst_week_num'] = week_numbers[worst_idx]

        stats['largest_win'] = round(max(margins), 1) if margins else 0
        stats['largest_loss'] = round(min(margins), 1) if margins else 0

        total_games = stats['wins'] + stats['losses'] + stats['ties']
        stats['win_pct'] = (
            round((stats['wins'] + 0.5 * stats['ties']) / total_games, 3) if total_games > 0 else 0
        )

        streak = stats['current_streak']
        if streak:
            last_result = streak[-1]
            streak_count = 0
            for result in reversed(streak):
                if result == last_result:
                    streak_count += 1
                else:
                    break
            stats['streak'] = {'type': last_result, 'count': streak_count}

        stats['games_above_500'] = stats['wins'] - stats['losses']

        stats['record'] = f'{stats["wins"]}-{stats["losses"]}'
        if stats['ties'] > 0:
            stats['record'] += f'-{stats["ties"]}'

        # OPR ("Oberon Power Ranking"): a single-number blend of scoring
        # volume, consistency (best + worst week), and win rate.
        avg_points = stats['ppg']
        high_score = stats['best_week']
        low_score = stats['worst_week']
        win_pct_100 = stats['win_pct'] * 100
        stats['opr'] = round(
            (5 * avg_points + 2 * (high_score + low_score) + 3 * win_pct_100) / 10, 2
        )

        optimal_points = stats['lineup_optimal_points']
        if optimal_points > 0:
            stats['owner_success_rate'] = round(
                stats['lineup_actual_points'] / optimal_points * 100, 1
            )
            stats['points_left_on_table_pct'] = round(
                stats['points_left_on_table'] / optimal_points * 100, 1
            )
        else:
            stats['owner_success_rate'] = None
            stats['points_left_on_table_pct'] = None

        del stats['points_for']
        del stats['points_against']
        del stats['margins']
        del stats['weekly_ranks']
        del stats['week_numbers']
        del stats['current_streak']

    opr_values = [s['opr'] for s in team_stats.values() if 'opr' in s]
    league_avg_opr = sum(opr_values) / len(opr_values) if opr_values else 1

    for stats in team_stats.values():
        if 'opr' in stats:
            stats['adjusted_opr'] = (
                round(stats['opr'] / league_avg_opr, 3) if league_avg_opr > 0 else 0
            )
            stats['league_avg_opr'] = round(league_avg_opr, 2)

    return team_stats


def load_pending_trades():
    path = Path(__file__).parent.parent / 'data' / 'pending_trades.json'
    if path.exists():
        with open(path) as f:
            return json.load(f).get('trades', [])
    return []


def compute_playoff_data(weeks_data, standings, current_nfl_week):
    """
    Compute playoff bracket and Jamboree data.

    Playoffs:
    - Week 16: Semifinals (1 vs 4, 2 vs 3)
    - Week 17: Oakland Bowl (winners), 3rd place game (losers)

    Jamboree:
    - All non-playoff teams (seeds 5-12)
    - Winner is whoever scores most total points over weeks 16-17
    """
    # Get final regular season standings (after week 15)
    if not standings or len(standings) < 4:
        return None

    # Playoff teams are top 4 seeds
    playoff_teams = [s['abbrev'] for s in standings[:4]]
    jamboree_teams = [s['abbrev'] for s in standings[4:]]

    # Create week lookup
    week_data_map = {w['week']: w for w in weeks_data}

    # Initialize playoff structure
    playoffs = {
        'playoff_teams': playoff_teams,
        'jamboree_teams': jamboree_teams,
        'seeds': {standings[i]['abbrev']: i + 1 for i in range(len(standings))},
        'week_16': {
            'semifinal_1': {  # 1 vs 4
                'higher_seed': standings[0]['abbrev'],
                'lower_seed': standings[3]['abbrev'],
                'higher_score': 0,
                'lower_score': 0,
                'winner': None,
                'loser': None,
            },
            'semifinal_2': {  # 2 vs 3
                'higher_seed': standings[1]['abbrev'],
                'lower_seed': standings[2]['abbrev'],
                'higher_score': 0,
                'lower_score': 0,
                'winner': None,
                'loser': None,
            },
        },
        'week_17': {
            'championship': {  # Oakland Bowl
                'team1': None,
                'team2': None,
                'score1': 0,
                'score2': 0,
                'winner': None,
                'loser': None,
            },
            'third_place': {
                'team1': None,
                'team2': None,
                'score1': 0,
                'score2': 0,
                'winner': None,
                'loser': None,
            },
        },
        'jamboree': {
            'standings': [],  # List of {abbrev, name, week_16_score, week_17_score, total}
            'winner': None,
        },
    }

    # Get week 16 scores
    week_16_data = week_data_map.get(16)
    if week_16_data and week_16_data.get('teams'):
        team_scores_16 = {t['abbrev']: t['total_score'] for t in week_16_data['teams']}

        # Update semifinal 1 (1 vs 4)
        sf1 = playoffs['week_16']['semifinal_1']
        sf1['higher_score'] = team_scores_16.get(sf1['higher_seed'], 0)
        sf1['lower_score'] = team_scores_16.get(sf1['lower_seed'], 0)
        if sf1['higher_score'] > 0 or sf1['lower_score'] > 0:
            if sf1['higher_score'] >= sf1['lower_score']:
                sf1['winner'] = sf1['higher_seed']
                sf1['loser'] = sf1['lower_seed']
            else:
                sf1['winner'] = sf1['lower_seed']
                sf1['loser'] = sf1['higher_seed']

        # Update semifinal 2 (2 vs 3)
        sf2 = playoffs['week_16']['semifinal_2']
        sf2['higher_score'] = team_scores_16.get(sf2['higher_seed'], 0)
        sf2['lower_score'] = team_scores_16.get(sf2['lower_seed'], 0)
        if sf2['higher_score'] > 0 or sf2['lower_score'] > 0:
            if sf2['higher_score'] >= sf2['lower_score']:
                sf2['winner'] = sf2['higher_seed']
                sf2['loser'] = sf2['lower_seed']
            else:
                sf2['winner'] = sf2['lower_seed']
                sf2['loser'] = sf2['higher_seed']

        # Set up week 17 matchups based on week 16 results
        if sf1['winner'] and sf2['winner']:
            playoffs['week_17']['championship']['team1'] = sf1['winner']
            playoffs['week_17']['championship']['team2'] = sf2['winner']
            playoffs['week_17']['third_place']['team1'] = sf1['loser']
            playoffs['week_17']['third_place']['team2'] = sf2['loser']

    # Get week 17 scores
    week_17_data = week_data_map.get(17)
    if week_17_data and week_17_data.get('teams'):
        team_scores_17 = {t['abbrev']: t['total_score'] for t in week_17_data['teams']}

        # Update championship (Oakland Bowl)
        champ = playoffs['week_17']['championship']
        if champ['team1'] and champ['team2']:
            champ['score1'] = team_scores_17.get(champ['team1'], 0)
            champ['score2'] = team_scores_17.get(champ['team2'], 0)
            if champ['score1'] > 0 or champ['score2'] > 0:
                if champ['score1'] >= champ['score2']:
                    champ['winner'] = champ['team1']
                    champ['loser'] = champ['team2']
                else:
                    champ['winner'] = champ['team2']
                    champ['loser'] = champ['team1']

        # Update third place game
        third = playoffs['week_17']['third_place']
        if third['team1'] and third['team2']:
            third['score1'] = team_scores_17.get(third['team1'], 0)
            third['score2'] = team_scores_17.get(third['team2'], 0)
            if third['score1'] > 0 or third['score2'] > 0:
                if third['score1'] >= third['score2']:
                    third['winner'] = third['team1']
                    third['loser'] = third['team2']
                else:
                    third['winner'] = third['team2']
                    third['loser'] = third['team1']

    # Compute Jamboree standings (combined scores for weeks 16-17)
    jamboree_standings = []
    for abbrev in jamboree_teams:
        team_info = next((s for s in standings if s['abbrev'] == abbrev), None)
        if not team_info:
            continue

        week_16_score = 0
        week_17_score = 0

        if week_16_data and week_16_data.get('teams'):
            team_16 = next((t for t in week_16_data['teams'] if t['abbrev'] == abbrev), None)
            if team_16:
                week_16_score = team_16.get('total_score', 0)

        if week_17_data and week_17_data.get('teams'):
            team_17 = next((t for t in week_17_data['teams'] if t['abbrev'] == abbrev), None)
            if team_17:
                week_17_score = team_17.get('total_score', 0)

        jamboree_standings.append(
            {
                'abbrev': abbrev,
                'name': team_info['name'],
                'week_16_score': round(week_16_score, 1),
                'week_17_score': round(week_17_score, 1),
                'total': round(week_16_score + week_17_score, 1),
            }
        )

    # Sort by total score descending
    jamboree_standings.sort(key=lambda x: x['total'], reverse=True)
    playoffs['jamboree']['standings'] = jamboree_standings

    # Determine Jamboree winner if week 17 is complete
    if jamboree_standings and current_nfl_week > 17:
        playoffs['jamboree']['winner'] = jamboree_standings[0]['abbrev']

    return playoffs


def main():
    parser = argparse.ArgumentParser(description='Export OPFL scores to web JSON')
    parser.add_argument('--excel', '-e', default=None, help='Path to the OPFL workbook')
    parser.add_argument(
        '--week',
        '-w',
        type=int,
        default=None,
        help='Week the Matchups tab represents (defaults to the current NFL week)',
    )
    parser.add_argument('--season', '-y', type=int, default=SEASON, help='NFL season year')
    parser.add_argument(
        '--force-rescore',
        action='store_true',
        help='Rescore and overwrite a week already archived as final',
    )
    args = parser.parse_args()

    project_dir = Path(__file__).parent.parent

    if args.excel:
        excel_path = Path(args.excel)
        if not excel_path.is_absolute():
            excel_path = project_dir / excel_path
        if not excel_path.exists():
            print(f'Error: {excel_path} not found!')
            return
    else:
        excel_path = None
        for fname in [
            'OPFL Scoring 2026.xlsx',
            'Griff OPFL Scoring 2025.xlsx',
            'OPFL Scoring 2025.xlsx',
        ]:
            if (project_dir / fname).exists():
                excel_path = project_dir / fname
                break

        if not excel_path:
            print('Error: No OPFL Excel file found!')
            return

    output_path = project_dir / 'web' / 'data.json'
    output_path.parent.mkdir(parents=True, exist_ok=True)

    print(f'Exporting {excel_path} to {output_path}...')

    data = export_season(
        str(excel_path),
        week_num=args.week,
        season=args.season,
        force_rescore=args.force_rescore,
    )
    data['pending_trades'] = load_pending_trades()

    draft_picks_path = project_dir / 'OPFL Draft & Future Traded Picks.xlsx'
    if draft_picks_path.exists():
        print('Parsing draft picks...')
        data['draft_picks'] = parse_draft_picks(str(draft_picks_path))

    banners_dir = project_dir / 'web' / 'images' / 'banners'
    if banners_dir.exists():
        print('Loading banners...')
        data['banners'] = get_existing_banners(str(banners_dir))

    with open(output_path, 'w') as f:
        json.dump(data, f, indent=2)

    print(f'Exported {len(data["weeks"])} week(s)')
    print(f'Standings: {len(data["standings"])} teams')
    if 'banners' in data:
        print(f'Banners: {len(data["banners"])} images')
    print(f'Updated at: {data["updated_at"]}')


if __name__ == '__main__':
    main()
