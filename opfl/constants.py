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
    'LAR': 'LA',   # Los Angeles Rams
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

# OPFL Excel layout: positions in Column A, pattern repeats for each team
# Each team occupies 3 columns: Points | Star (*) | Player Name
# Teams are at columns: 4, 7, 10, 13, 16, 19, ... (player name column)
# The row structure is different - positions are identified by the label in column A

# Position identifiers used in OPFL Excel
POSITION_LABELS = ['QB', 'RB', 'WR', 'TE', 'K', 'DF', 'HC']

# Expected player slots per position
PLAYERS_PER_POSITION = {
    'QB': 3,
    'RB': 4,
    'WR': 4,
    'TE': 3,
    'K': 2,
    'DF': 2,  # Defense
    'HC': 2,  # Head Coach
}
