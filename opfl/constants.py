"""Constants and mappings for OPFL autoscorer."""

# Team name to abbreviation mapping
TEAM_ABBREV_MAP = {
    'Arizona Cardinals': 'ARI',
    'Atlanta Falcons': 'ATL',
    'Baltimore Ravens': 'BAL',
    'Buffalo Bills': 'BUF',
    'Carolina Panthers': 'CAR',
    'Chicago Bears': 'CHI',
    'Cincinnati Bengals': 'CIN',
    'Cleveland Browns': 'CLE',
    'Dallas Cowboys': 'DAL',
    'Denver Broncos': 'DEN',
    'Detroit Lions': 'DET',
    'Green Bay Packers': 'GB',
    'Houston Texans': 'HOU',
    'Indianapolis Colts': 'IND',
    'Jacksonville Jaguars': 'JAX',
    'Kansas City Chiefs': 'KC',
    'Las Vegas Raiders': 'LV',
    'Los Angeles Chargers': 'LAC',
    'Los Angeles Rams': 'LA',
    'Miami Dolphins': 'MIA',
    'Minnesota Vikings': 'MIN',
    'New England Patriots': 'NE',
    'New Orleans Saints': 'NO',
    'New York Giants': 'NYG',
    'New York Jets': 'NYJ',
    'Philadelphia Eagles': 'PHI',
    'Pittsburgh Steelers': 'PIT',
    'San Francisco 49ers': 'SF',
    'Seattle Seahawks': 'SEA',
    'Tampa Bay Buccaneers': 'TB',
    'Tennessee Titans': 'TEN',
    'Washington Commanders': 'WAS',
}

# Reverse mapping
ABBREV_TO_TEAM = {v: k for k, v in TEAM_ABBREV_MAP.items()}

# Team abbreviation normalization (Excel format -> nflreadpy format)
# OPFL uses various abbreviations, we need to map them to nflreadpy format
TEAM_ABBREV_NORMALIZE = {
    'LAR': 'LA',  # Los Angeles Rams
    'JAC': 'JAX',  # Jacksonville Jaguars
    'ARZ': 'ARI',  # Arizona Cardinals (OPFL sometimes uses ARZ)
    'Arz': 'ARI',  # Arizona Cardinals (OPFL sometimes uses Arz)
}

# Defense team name to abbreviation mapping
DEFENSE_NAME_TO_ABBREV = {
    'Arizona': 'ARI',
    'Atlanta': 'ATL',
    'Baltimore': 'BAL',
    'Buffalo': 'BUF',
    'Carolina': 'CAR',
    'Chicago': 'CHI',
    'Cincinnati': 'CIN',
    'Cleveland': 'CLE',
    'Dallas': 'DAL',
    'Denver': 'DEN',
    'Detroit': 'DET',
    'Green Bay': 'GB',
    'Houston': 'HOU',
    'Indianapolis': 'IND',
    'Jacksonville': 'JAX',
    'Kansas City': 'KC',
    'Las Vegas': 'LV',
    'LA Chargers': 'LAC',
    'LA Rams': 'LA',
    'Los Angeles Chargers': 'LAC',
    'Los Angeles Rams': 'LA',
    'Miami': 'MIA',
    'Minnesota': 'MIN',
    'New England': 'NE',
    'New Orleans': 'NO',
    'NY Giants': 'NYG',
    'NY Jets': 'NYJ',
    'New York Giants': 'NYG',
    'New York Jets': 'NYJ',
    'Philadelphia': 'PHI',
    'Pittsburgh': 'PIT',
    'San Francisco': 'SF',
    'Seattle': 'SEA',
    'Tampa Bay': 'TB',
    'Tennessee': 'TEN',
    'Washington': 'WAS',
}

# --- OPFL franchise identity -------------------------------------------------
# Canonical owner name -> 3-char team code used by the website and data files.
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
ALL_TEAM_CODES = list(OWNER_TO_CODE.values())

# The workbook spells the same franchise a dozen different ways depending on the
# tab ("LAM" on Rosters, "Steve L" on Matchups, "SL" in the schedule key).
# Everything here is matched case-insensitively after stripping.
OWNER_ALIASES = {
    'KIRK/DAVID': 'K/D',
    'KIRK/D': 'K/D',
    'K/D': 'K/D',
    'KIRK': 'K/D',
    'STEVE L.': 'STL',
    'STEVE L': 'STL',
    'SL': 'STL',
    'LAM': 'STL',
    'STL': 'STL',
    'JOHN': 'JOH',
    'JOH': 'JOH',
    'KEVIN': 'KEV',
    'KEV': 'KEV',
    'KL': 'KEV',
    'DANNY/JOEY': 'D/J',
    'D/J': 'D/J',
    'DANNY/J': 'D/J',
    'GREG/GRIFFIN': 'G/G',
    'GREG/GRIFF': 'G/G',
    'GREG/G': 'G/G',
    'G/G': 'G/G',
    'ERIC/JEFF': 'E/J',
    'ERIC/J': 'E/J',
    'E/J': 'E/J',
    'ANDREW': 'AND',
    'AND': 'AND',
    'WES/BILL': 'W/B',
    'WES/B': 'W/B',
    'BILL/WES': 'W/B',
    'W/B': 'W/B',
    'KEMP/A/M': 'KAM',
    'K/A/M': 'KAM',
    'KAM': 'KAM',
    'JARRETT/MATT': 'J/M',
    'JARRETT/M': 'J/M',
    'J/M': 'J/M',
    'ADAM': 'ADA',
    'ADA': 'ADA',
}


def resolve_team_code(raw_name: str) -> str:
    """Map any workbook spelling of a franchise to its canonical team code.

    Returns "" when the name cannot be resolved so callers can flag it rather
    than silently inventing a team.
    """
    if not raw_name:
        return ''

    name = str(raw_name).strip()
    # Roster headers carry a trailing seed like "KIRK/DAVID (1)".
    if '(' in name:
        name = name.split('(')[0].strip()
    key = name.upper()

    if key in OWNER_ALIASES:
        return OWNER_ALIASES[key]

    # Fall back to a containment match so a new spelling of an existing
    # franchise ("Greg/Griffin & co") still resolves.
    for alias, code in OWNER_ALIASES.items():
        if key.startswith(alias) or alias.startswith(key):
            return code

    return ''


# OPFL Excel layout: positions in Column A, pattern repeats for each team
# Each team occupies 3 columns: Points | Star (*) | Player Name
# Teams are at columns: 4, 7, 10, 13, 16, 19, ... (player name column)
# The row structure is different - positions are identified by the label in column A

# Position identifiers used in OPFL Excel
POSITION_LABELS = ['QB', 'RB', 'WR', 'TE', 'K', 'DF', 'HC']

# Roster/starter slot counts per position live in data/league_config.json
# (opfl.config.get_roster_slots / get_starter_slots) rather than here.
