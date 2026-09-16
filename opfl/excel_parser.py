"""Excel roster parsing utilities for OPFL format."""

import re

import openpyxl

from .constants import (
    DEFENSE_NAME_TO_ABBREV,
    POSITION_LABELS,
    TEAM_ABBREV_NORMALIZE,
    resolve_team_code,
)
from .models import FantasyTeam


def parse_player_name(cell_value: str) -> tuple[str, str]:
    """
    Parse player name from Excel format "Player Name (TEAM)" to (name, team_abbrev).

    Examples:
        "Patrick Mahomes II (KC)" -> ("Patrick Mahomes II", "KC")
        "Caleb Williams (Chi)" -> ("Caleb Williams", "CHI")
        "Baltimore" -> ("Baltimore", "BAL")  # Defense
    """
    if not cell_value:
        return '', ''

    cell_value = str(cell_value).strip()

    # Check if it's a defense (just team name)
    if cell_value in DEFENSE_NAME_TO_ABBREV:
        return cell_value, DEFENSE_NAME_TO_ABBREV[cell_value]

    # Try to match "Player Name (TEAM)" pattern
    match = re.match(r'^(.+?)\s*\(([A-Za-z]{2,3})\)$', cell_value)
    if match:
        name = match.group(1).strip()
        team = match.group(2).upper()
        # Normalize team abbreviation
        team = TEAM_ABBREV_NORMALIZE.get(team, team)
        return name, team

    return cell_value, ''


def is_valid_player_name(name: str, position: str = '') -> bool:
    """Reject roster cells that hold filler rather than a player/coach name.

    The workbook pads roster blocks with zeros and keeps phone numbers in the
    same columns, so those have to be screened out before scoring.
    """
    if not name:
        return False

    name = str(name).strip()
    if not name:
        return False

    # Padding zeros and stray numbers.
    try:
        float(name)
        return False
    except (ValueError, TypeError):
        pass

    # Phone numbers, e.g. "407-5858 cell" or "325-1289 (J)".
    if re.match(r'^\(?\d{3}\)?[-\s]?\d{3}[-\s]?\d{4}', name) or re.match(r'^\d{3}-\d{4}', name):
        return False

    if not any(c.isalpha() for c in name):
        return False

    # Coach names always start with a letter; roster cells sometimes don't.
    return not (position == 'HC' and not name[0].isalpha())


# Valid team columns in OPFL Excel format (columns D, G, J, M, P, S)
# Skip columns V (22) and Y (25) which are duplicate starter-only views
VALID_TEAM_COLUMNS = [4, 7, 10, 13, 16, 19]


def find_team_columns(ws, header_row: int = 1) -> list[tuple[int, str, int]]:
    """
    Find fantasy team columns in the worksheet.

    Only checks the known valid columns (4, 7, 10, 13, 16, 19) to avoid
    picking up duplicate columns like V (22) and Y (25).

    Returns:
        List of (column_index, team_name, header_row) tuples
    """
    teams = []

    # Only check the valid team columns
    for col in VALID_TEAM_COLUMNS:
        if col <= ws.max_column:
            cell_value = ws.cell(row=header_row, column=col).value
            if cell_value:
                cell_str = str(cell_value).strip()
                # Team names in OPFL look like "KIRK/DAVID (11)" or "STEVE L. (43)"
                if re.match(r'^[A-Z\s/\.]+\s*\(\d+\)$', cell_str, re.IGNORECASE):
                    teams.append((col, cell_str, header_row))

    return teams


def find_position_rows(ws, start_row: int = 1, end_row: int = 100) -> dict:
    """
    Find the row ranges for each position in the worksheet within a specific range.

    Args:
        ws: Worksheet
        start_row: First row to search
        end_row: Last row to search

    Returns:
        Dict mapping position -> list of row numbers containing players
    """
    position_rows = {}
    current_position = None

    for row in range(start_row, min(end_row + 1, ws.max_row + 1)):
        cell_value = ws.cell(row=row, column=1).value
        if cell_value and str(cell_value).strip():
            cell_str = str(cell_value).strip().upper()
            if cell_str in POSITION_LABELS:
                current_position = cell_str
                if current_position not in position_rows:
                    position_rows[current_position] = []
                # This row contains position players
                position_rows[current_position].append(row)
            else:
                # A non-position label (PH = phone numbers, TS = taxi squad)
                # ends the roster section; anything after it is not a lineup.
                current_position = None
        elif current_position:
            # Empty column A but might still have players for current position
            # Check if there's actual content in the row
            has_content = any(
                ws.cell(row=row, column=c).value for c in range(2, min(25, ws.max_column + 1))
            )
            if has_content:
                position_rows[current_position].append(row)

    return position_rows


def parse_roster_from_excel(filepath: str, sheet_name: str = 'W1') -> list[FantasyTeam]:
    """
    Parse fantasy team rosters from OPFL Excel file.

    The OPFL Excel format:
        - Teams are arranged horizontally, each occupying 3 columns: Points | Star | Player
        - Two blocks of 6 teams: rows 1-38 and rows 39+
        - Positions are in column A (QB, RB, WR, TE, K, DF, HC)
        - A star (*) in the star column indicates a starter
        - Player names are in format "Name (Team)" or just "TeamName" for defenses

    Args:
        filepath: Path to the Excel file
        sheet_name: Name of the sheet to read (e.g., "W1", "W12")

    Returns:
        List of FantasyTeam objects
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[sheet_name]

    teams = []
    seen_team_names = set()  # Track teams we've already added

    # Find teams in first block (row 1)
    team_columns_block1 = find_team_columns(ws, header_row=1)

    # If no teams found via pattern, try the valid columns directly
    if not team_columns_block1:
        for col in VALID_TEAM_COLUMNS:
            if col <= ws.max_column:
                header = ws.cell(row=1, column=col).value
                if header:
                    team_columns_block1.append((col, str(header).strip(), 1))

    # Find teams in second block (row 39)
    team_columns_block2 = find_team_columns(ws, header_row=39)

    # If no teams found via pattern, try the valid columns directly
    if not team_columns_block2:
        for col in VALID_TEAM_COLUMNS:
            if col <= ws.max_column:
                header = ws.cell(row=39, column=col).value
                if header and str(header).strip():
                    header_str = str(header).strip()
                    if re.match(r'^[A-Z\s/\.]+\s*\(\d+\)$', header_str, re.IGNORECASE):
                        team_columns_block2.append((col, header_str, 39))

    # Pre-compute position rows for each block
    position_rows_block1 = find_position_rows(ws, start_row=1, end_row=38)
    position_rows_block2 = find_position_rows(ws, start_row=39, end_row=80)

    # Parse teams from block 1
    for player_col, team_name_raw, _header_row in team_columns_block1:
        match = re.match(r'^(.+?)\s*\(\d+\)$', team_name_raw)
        team_name = match.group(1).strip() if match else team_name_raw

        if team_name in seen_team_names:
            continue
        seen_team_names.add(team_name)

        star_col = player_col - 1

        team = FantasyTeam(
            name=team_name,
            owner=team_name,
            abbreviation='',
            column_index=player_col,
            players={},
        )

        # Use position rows from block 1
        for position, rows in position_rows_block1.items():
            team.players[position] = []

            for row in rows:
                player_cell = ws.cell(row=row, column=player_col)
                star_cell = ws.cell(row=row, column=star_col)

                if player_cell.value:
                    player_name, nfl_team = parse_player_name(str(player_cell.value))

                    if is_valid_player_name(player_name, position):
                        is_started = star_cell.value == '*'
                        team.players[position].append((player_name, nfl_team, is_started))

        teams.append(team)

    # Parse teams from block 2
    for player_col, team_name_raw, _header_row in team_columns_block2:
        match = re.match(r'^(.+?)\s*\(\d+\)$', team_name_raw)
        team_name = match.group(1).strip() if match else team_name_raw

        if team_name in seen_team_names:
            continue
        seen_team_names.add(team_name)

        star_col = player_col - 1

        team = FantasyTeam(
            name=team_name,
            owner=team_name,
            abbreviation='',
            column_index=player_col,
            players={},
        )

        # Use position rows from block 2
        for position, rows in position_rows_block2.items():
            team.players[position] = []

            for row in rows:
                player_cell = ws.cell(row=row, column=player_col)
                star_cell = ws.cell(row=row, column=star_col)

                if player_cell.value:
                    player_name, nfl_team = parse_player_name(str(player_cell.value))

                    if is_valid_player_name(player_name, position):
                        is_started = star_cell.value == '*'
                        team.players[position].append((player_name, nfl_team, is_started))

        teams.append(team)

    wb.close()
    return teams


def parse_roster_from_rosters_sheet(filepath: str) -> list[FantasyTeam]:
    """
    Parse fantasy team rosters from the 'Rosters' sheet in OPFL Excel file.

    Args:
        filepath: Path to the Excel file

    Returns:
        List of FantasyTeam objects
    """
    return parse_roster_from_excel(filepath, sheet_name='Rosters')


def parse_week_sheet_points(filepath: str, sheet_name: str = 'W1') -> list[tuple[str, list[dict]]]:
    """Read a W-sheet's own recorded points, rather than re-deriving them.

    Each player's points are static values sitting one column to the left of
    the star column - the commissioner (or an older season's scoring engine)
    already computed them at the time, so for archived seasons this is more
    trustworthy than re-scoring from nflreadpy stats, which requires an NFL
    team per player that older-format sheets never recorded.

    Returns:
        [(team_name, [{'name', 'nfl_team', 'position', 'score', 'starter'}, ...]), ...]
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[sheet_name]

    results: list[tuple[str, list[dict]]] = []
    seen_team_names: set[str] = set()

    for header_row, end_row in ((1, 38), (39, 80)):
        team_columns = find_team_columns(ws, header_row=header_row)
        if not team_columns:
            for col in VALID_TEAM_COLUMNS:
                if col <= ws.max_column:
                    header = ws.cell(row=header_row, column=col).value
                    if header and str(header).strip():
                        team_columns.append((col, str(header).strip(), header_row))

        position_rows = find_position_rows(ws, start_row=header_row, end_row=end_row)

        for player_col, team_name_raw, _header_row in team_columns:
            match = re.match(r'^(.+?)\s*\(\d+\)$', team_name_raw)
            team_name = match.group(1).strip() if match else team_name_raw
            if team_name in seen_team_names:
                continue
            seen_team_names.add(team_name)

            star_col = player_col - 1
            points_col = player_col - 2

            roster = []
            for position, rows in position_rows.items():
                for row in rows:
                    player_cell = ws.cell(row=row, column=player_col)
                    if not player_cell.value:
                        continue
                    name, nfl_team = parse_player_name(str(player_cell.value))
                    if not is_valid_player_name(name, position):
                        continue

                    points_value = ws.cell(row=row, column=points_col).value
                    score = float(points_value) if isinstance(points_value, (int, float)) else 0.0
                    starter = ws.cell(row=row, column=star_col).value == '*'

                    roster.append(
                        {
                            'name': name,
                            'nfl_team': nfl_team,
                            'position': position,
                            'score': round(score, 1),
                            'starter': starter,
                        }
                    )

            results.append((team_name, roster))

    wb.close()
    return results


# The Matchups tab lists one lineup per block, always in this slot order.
MATCHUP_LINEUP_POSITIONS = ['QB', 'RB', 'RB', 'WR', 'WR', 'TE', 'K', 'DF', 'HC']

# Each matchup block puts the two opponents in these player-name columns.
# The cell immediately to the right of each holds that player's points.
MATCHUP_COLUMNS = [1, 4]


def parse_matchups_sheet(filepath: str, sheet_name: str = 'Matchups') -> list[dict]:
    """
    Parse the weekly matchups tab of the OPFL workbook.

    Layout: stacked blocks of two opponents side by side. Each block starts with
    a row holding the two owner names (columns A and D), followed by the nine
    starting slots in QB/RB/RB/WR/WR/TE/K/DF/HC order.

    Returns:
        List of matchup dicts:
        {'teams': [{'code', 'raw_name', 'lineup': [(position, name, nfl_team)]}, ...]}
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[sheet_name]

    matchups = []
    row = 1
    while row <= ws.max_row:
        left = ws.cell(row=row, column=MATCHUP_COLUMNS[0]).value
        right = ws.cell(row=row, column=MATCHUP_COLUMNS[1]).value

        left_code = resolve_team_code(left) if left else ''
        right_code = resolve_team_code(right) if right else ''

        # A header row is two resolvable owner names sitting side by side.
        if not (left_code and right_code):
            row += 1
            continue

        sides = []
        for col, code, raw in zip(
            MATCHUP_COLUMNS, (left_code, right_code), (left, right), strict=True
        ):
            lineup = []
            for offset, position in enumerate(MATCHUP_LINEUP_POSITIONS, start=1):
                cell = ws.cell(row=row + offset, column=col).value
                if not cell:
                    continue
                name, nfl_team = parse_player_name(str(cell))
                if is_valid_player_name(name, position):
                    lineup.append((position, name, nfl_team))
            sides.append({'code': code, 'raw_name': str(raw).strip(), 'lineup': lineup})

        matchups.append({'teams': sides})
        row += len(MATCHUP_LINEUP_POSITIONS) + 1

    wb.close()
    return matchups


def parse_taxi_squads(filepath: str, sheet_name: str = 'Rosters') -> dict:
    """
    Parse the taxi squad ("TS") block under each roster.

    Taxi entries are written as "RB Aaron Jones" - position prefix, then name,
    with no NFL team.

    Returns:
        Dict mapping team code -> list of {'position', 'name'}.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[sheet_name]

    taxi = {}
    for header_row in (1, 39):
        # Find the TS label that belongs to this roster block.
        ts_row = None
        for row in range(header_row, min(header_row + 40, ws.max_row + 1)):
            value = ws.cell(row=row, column=1).value
            if value and str(value).strip().upper() == 'TS':
                ts_row = row
                break
        if ts_row is None:
            continue

        for col in VALID_TEAM_COLUMNS:
            code = resolve_team_code(ws.cell(row=header_row, column=col).value)
            if not code:
                continue

            players = []
            for row in range(ts_row, ts_row + 6):
                cell = ws.cell(row=row, column=col).value
                if not cell:
                    continue
                text = str(cell).strip()
                match = re.match(r'^(QB|RB|WR|TE|K|DF|HC)\s+(.+)$', text, re.IGNORECASE)
                # Only "POS Name" rows are taxi entries; this also skips the
                # next block's owner header, which sits a few rows below.
                if match and is_valid_player_name(match.group(2)):
                    players.append(
                        {'position': match.group(1).upper(), 'name': match.group(2).strip()}
                    )
            taxi[code] = players

    wb.close()
    return taxi


# NOTE: the 2026 workbook is formula-driven end to end - the Matchups tab pulls
# names and totals from the Scoring tab, which in turn pulls each team's starred
# players off the Rosters tab. openpyxl cannot evaluate formulas, so saving the
# workbook would strip every cached value and leave the file unreadable to this
# pipeline until Excel reopened and recalculated it. Scores are therefore
# published to web/data.json only; nothing is written back to the workbook.


def update_excel_scores(
    excel_path: str,
    sheet_name: str,
    teams: list[FantasyTeam],
    results: dict,
):
    """
    Update the Excel file with calculated scores.

    Args:
        excel_path: Path to the Excel file
        sheet_name: Sheet to update
        teams: List of FantasyTeam objects
        results: Dict mapping team name to (total_score, position_scores)
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[sheet_name]

    for team in teams:
        if team.name not in results:
            continue

        total, scores = results[team.name]
        player_col = team.column_index
        points_col = player_col - 2

        for position, player_list in team.players.items():
            if position not in scores:
                continue

            for player_name, _nfl_team, _is_started in player_list:
                player_score = None
                for ps in scores[position]:
                    if ps.name == player_name:
                        player_score = ps
                        break

                if player_score is None:
                    continue

                # Find the row for this player
                # Search appropriate range based on team's column location
                for row in range(1, ws.max_row + 1):
                    cell = ws.cell(row=row, column=player_col)
                    if cell.value:
                        parsed_name, _ = parse_player_name(str(cell.value))
                        if parsed_name == player_name:
                            score_cell = ws.cell(row=row, column=points_col)
                            score_cell.value = player_score.total_points
                            break

    wb.save(excel_path)
    print(f'\nScores saved to {excel_path}')
