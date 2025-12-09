#!/usr/bin/env python3
"""Export OPFL Excel scores to JSON for web display."""

import json
import re
import os
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import nflreadpy as nfl
import openpyxl

OWNER_TO_CODE = {
    'KIRK/DAVID': 'K/D',
    'STEVE L.': 'STL',
    'JOHN': 'JOH',
    'KEVIN': 'KEV',
    'DANNY/JOEY': 'D/J',
    'GREG/GRIFFIN': 'G/G',
    'ERIC/JEFF': 'E/J',
    'ANDREW': 'AND',
    'WES/BILL': 'W/B',
    'KEMP/A/M': 'KAM',
    'JARRETT/MATT': 'J/M',
    'ADAM': 'ADA',
}

CODE_TO_OWNER = {v: k for k, v in OWNER_TO_CODE.items()}
ALL_TEAMS = list(OWNER_TO_CODE.values())
TEAM_COLUMNS = [4, 7, 10, 13, 16, 19]
POSITIONS = ['QB', 'RB', 'WR', 'TE', 'K', 'DF', 'HC']
TRADE_DEADLINE_WEEK = 12


def parse_player_name(cell_value):
    if not cell_value:
        return "", ""
    cell_value = str(cell_value).strip()
    match = re.match(r'^(.+?)\s*\(([A-Za-z]{2,3})\)$', cell_value)
    if match:
        return match.group(1).strip(), match.group(2).upper()
    return cell_value, ""


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
        return "", ""
    match = re.match(r'^(.+?)\s*\((\d+)\)$', str(header_value).strip())
    if match:
        return match.group(1).strip(), match.group(2)
    return str(header_value).strip(), ""


def export_week(ws, week_num):
    teams_data = []
    blocks = [(1, 1, 38), (39, 39, 80)]
    
    for header_row, start_row, end_row in blocks:
        position_rows = find_position_rows(ws, start_row, end_row)
        
        for col in TEAM_COLUMNS:
            header = ws.cell(row=header_row, column=col).value
            if not header:
                continue
            
            team_name, standing = extract_team_name(header)
            abbrev = OWNER_TO_CODE.get(team_name, team_name[:3].upper())
            star_col = col - 1
            roster = []
            total_score = 0.0
            
            for position, rows in position_rows.items():
                for row in rows:
                    player_cell = ws.cell(row=row, column=col)
                    score_cell = ws.cell(row=row, column=col - 2)
                    star_cell = ws.cell(row=row, column=star_col)
                    
                    if player_cell.value:
                        player_name, nfl_team = parse_player_name(str(player_cell.value))
                        is_starter = star_cell.value == '*'
                        score = float(score_cell.value) if score_cell.value else 0.0
                        
                        if player_name:
                            roster.append({
                                'name': player_name,
                                'nfl_team': nfl_team,
                                'position': position,
                                'score': score,
                                'starter': is_starter,
                            })
                            if is_starter:
                                total_score += score
            
            if roster:
                teams_data.append({
                    'name': team_name,
                    'owner': team_name,
                    'abbrev': abbrev,
                    'roster': roster,
                    'total_score': round(total_score, 1),
                })
    
    sorted_by_score = sorted(teams_data, key=lambda t: t['total_score'], reverse=True)
    for rank, team in enumerate(sorted_by_score, 1):
        team['score_rank'] = rank
    
    return {
        'week': week_num,
        'teams': teams_data,
        'has_scores': any(t['total_score'] > 0 for t in teams_data),
    }


def get_current_nfl_week():
    return nfl.get_current_week()


def get_game_times(season=2025):
    try:
        schedule = nfl.load_schedules(seasons=season)
        game_times = {}
        for week in range(1, 19):
            week_games = schedule.filter(schedule['week'] == week)
            if week_games.height == 0:
                continue
            game_times[week] = {}
            for row in week_games.iter_rows(named=True):
                game_date = row.get('gameday', '')
                game_time = row.get('gametime', '')
                if game_date and game_time:
                    try:
                        dt = datetime.strptime(f"{game_date} {game_time}", "%Y-%m-%d %H:%M")
                        kickoff_iso = dt.strftime("%Y-%m-%dT%H:%M:00-05:00")
                        if row.get('home_team'):
                            game_times[week][row['home_team']] = kickoff_iso
                        if row.get('away_team'):
                            game_times[week][row['away_team']] = kickoff_iso
                    except (ValueError, TypeError):
                        pass
        return game_times
    except Exception as e:
        print(f"Warning: Could not load game times: {e}")
        return {}


def parse_draft_picks(excel_path):
    DEFAULT_PRESEASON = list(range(1, 7))
    DEFAULT_WAIVER = list(range(1, 4))
    SEASONS = ['2026', '2027', '2028']
    
    picks = {}
    for team in ALL_TEAMS:
        picks[team] = {}
        for season in SEASONS:
            picks[team][season] = {
                'preseason': [(r, team) for r in DEFAULT_PRESEASON],
                'waiver': [(r, team) for r in DEFAULT_WAIVER],
            }
    
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb['Future Traded Picks']
    except Exception as e:
        print(f"Warning: Could not load traded picks: {e}")
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
            
            holder = OWNER_TO_CODE.get(holder_name)
            original = OWNER_TO_CODE.get(original_name)
            
            if not holder or not original:
                for name, code in OWNER_TO_CODE.items():
                    if holder_name.upper() in name.upper() or name.upper() in holder_name.upper():
                        holder = code
                    if original_name.upper() in name.upper() or name.upper() in original_name.upper():
                        original = code
            
            if holder and original and season in SEASONS:
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
                    {'round': r, 'from': owner, 'own': owner == team}
                    for r, owner in team_picks
                ]
    return formatted


def extract_banners(docx_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    images = []
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            for name in z.namelist():
                if name.startswith('word/media/'):
                    img_name = name.split('/')[-1]
                    with open(os.path.join(output_dir, img_name), 'wb') as f:
                        f.write(z.read(name))
                    images.append(img_name)
    except Exception as e:
        print(f"Warning: Could not extract banners: {e}")
    return sorted(images)


def export_all_weeks(excel_path):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    weeks = []
    standings = {}
    
    week_sheets = []
    for sheet_name in wb.sheetnames:
        match = re.match(r'^W(\d+)$', sheet_name)
        if match:
            week_sheets.append((int(match.group(1)), sheet_name))
    week_sheets.sort(key=lambda x: x[0])
    
    for week_num, sheet_name in week_sheets:
        weeks.append(export_week(wb[sheet_name], week_num))
    
    current_nfl_week = get_current_nfl_week()
    print(f"Current NFL week: {current_nfl_week}, standings include weeks 1-{current_nfl_week - 1}")
    
    for week_data in weeks:
        if not week_data.get('has_scores', False) or week_data['week'] >= current_nfl_week:
            continue
        
        for team in week_data['teams']:
            abbrev = team['abbrev']
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
                points_per_team = (0.5 * len(positions_in_top6)) / len(tied_teams)
                for team in tied_teams:
                    standings[team['abbrev']]['rank_points'] += points_per_team
                    standings[team['abbrev']]['top_half'] += 1
            
            current_rank += len(tied_teams)
    
    wb.close()
    
    return {
        'updated_at': datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
        'season': 2025,
        'current_week': get_current_nfl_week(),
        'weeks': weeks,
        'standings': sorted(standings.values(), key=lambda x: (x['rank_points'], x['points_for']), reverse=True),
        'game_times': get_game_times(2025),
        'trade_deadline_week': TRADE_DEADLINE_WEEK,
        'taxi_squads': {team: [] for team in ALL_TEAMS},
    }


def load_pending_trades():
    path = Path(__file__).parent.parent / 'data' / 'pending_trades.json'
    if path.exists():
        with open(path) as f:
            return json.load(f).get('trades', [])
    return []


def main():
    project_dir = Path(__file__).parent.parent
    
    excel_path = None
    for fname in ['Griff OPFL Scoring 2025.xlsx', 'OPFL Scoring 2025.xlsx']:
        if (project_dir / fname).exists():
            excel_path = project_dir / fname
            break
    
    if not excel_path:
        print('Error: No OPFL Excel file found!')
        return
    
    output_path = project_dir / 'web' / 'data.json'
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    print(f'Exporting {excel_path} to {output_path}...')
    
    data = export_all_weeks(str(excel_path))
    data['pending_trades'] = load_pending_trades()
    
    draft_picks_path = project_dir / 'OPFL Draft & Future Traded Picks.xlsx'
    if draft_picks_path.exists():
        print('Parsing draft picks...')
        data['draft_picks'] = parse_draft_picks(str(draft_picks_path))
    
    banner_docx = project_dir / 'OPFL Banner Room.docx'
    if banner_docx.exists():
        print('Extracting banners...')
        data['banners'] = extract_banners(str(banner_docx), str(project_dir / 'web' / 'images' / 'banners'))
    
    with open(output_path, 'w') as f:
        json.dump(data, f, indent=2)
    
    print(f'Exported {len(data["weeks"])} weeks')
    print(f'Standings: {len(data["standings"])} teams')
    if 'banners' in data:
        print(f'Banners: {len(data["banners"])} images')
    print(f'Updated at: {data["updated_at"]}')


if __name__ == '__main__':
    main()
