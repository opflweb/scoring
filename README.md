# OPFL Autoscorer

Automated fantasy football scoring for the OPFL using real-time NFL stats from [nflreadpy](https://github.com/nflverse/nflreadpy).

## Installation

```bash
pip install nflreadpy polars openpyxl thefuzz
```

Or using the project dependencies:

```bash
pip install -e .
```

## Usage

### Basic Usage

Score the current week with detailed output:

```bash
python autoscorer.py --excel "OPFL Scoring 2025.xlsx" --sheet "W12" --season 2025 --week 12
```

### Command Line Options

| Option | Short | Default | Description |
|--------|-------|---------|-------------|
| `--excel` | `-e` | `OPFL Scoring 2025.xlsx` | Path to the Excel file with rosters |
| `--sheet` | `-s` | `W12` | Sheet name to score (W1, W2, etc.) |
| `--season` | `-y` | `2025` | NFL season year |
| `--week` | `-w` | `12` | Week number to score |
| `--update` | `-u` | - | Update Excel file with calculated scores |
| `--quiet` | `-q` | - | Suppress detailed output, show only standings |

### Examples

```bash
# Score Week 12 with full breakdown
python autoscorer.py --sheet W12 --week 12

# Score a different week
python autoscorer.py --sheet W10 --week 10

# Quick standings only
python autoscorer.py --quiet

# Score and save results back to Excel
python autoscorer.py --update
```

### Validating Scores

Compare calculated scores against manually entered scores:

```bash
# Validate a single week
python validate_scores.py --sheet W12 --week 12

# Validate all weeks
python validate_scores.py --all

# Show only summary
python validate_scores.py --all --summary
```

## OPFL Scoring Rules

### Quarterback (QB)
- **Passing yards**: 200 yards = 2 pts, +1 pt per 50 yards thereafter (below 200 = 0 pts)
- **Rushing yards**: 75 yards = 2 pts, +1 pt per 25 yards thereafter
- **Receiving yards**: 75 yards = 2 pts, +1 pt per 25 yards thereafter
- **Touchdowns**: 6 points each (passing, rushing, receiving)
- **2-point conversions**: 2 points each
- **Interceptions**: -1 pt each
- **Interceptions returned for TD (pick-6)**: -3 pts total (replaces the -1)
- **Fumbles lost**: -1 pt each
- **Fumbles returned for TD (fumble-6)**: -3 pts total (replaces the -1)
- **Points cannot be less than zero**

### Running Back / Wide Receiver (RB/WR)
**Individual Category Scoring:**
- **Rushing yards**: 75 yards = 2 pts, +1 pt per 25 yards thereafter
- **Receiving yards**: 75 yards = 2 pts, +1 pt per 25 yards thereafter

**Alternate Combined Bonus:**
- 100 combined rush/rec yards = 2 pts, +1 pt per 25 yards thereafter
- Player gets whichever method yields more points

**Other:**
- **Touchdowns**: 6 points each
- **2-point conversions**: 2 points each
- **Fumbles lost / INTs thrown**: -1 pt each
- **Turnovers returned for TD**: -3 pts total (replaces the -1)
- **Points cannot be less than zero**

### Tight End (TE)
**Individual Category Scoring:**
- **Rushing yards**: 50 yards = 2 pts, +1 pt per 25 yards thereafter
- **Receiving yards**: 50 yards = 2 pts, +1 pt per 25 yards thereafter

**Alternate Combined Bonus:**
- 75 combined rush/rec yards = 2 pts, +1 pt per 25 yards thereafter
- Player gets whichever method yields more points

**Other:**
- **Touchdowns**: 6 points each
- **2-point conversions**: 2 points each
- **Fumbles lost / INTs thrown**: -1 pt each
- **Turnovers returned for TD**: -3 pts total (replaces the -1)
- **Points cannot be less than zero**

### Kicker (K)
- **PAT made**: 1 pt each
- **PAT missed/blocked**: -1 pt each
- **FG 1-29 yards**: 1 pt each
- **FG 30-39 yards**: 2 pts each
- **FG 40-49 yards**: 3 pts each
- **FG 50+ yards**: 4 pts each
- **FG missed/blocked**: -2 pts each

### Defense (DF)
| Points Allowed | Points |
|----------------|--------|
| 0 (Shutout) | +8 |
| 2-9 | +6 |
| 10-13 | +4 |
| 14-17 | +2 |
| 18-27 | 0 |
| 28-31 | -2 |
| 32-35 | -4 |
| 36+ | -6 |

- **Interception**: 2 pts each
- **Fumble recovery**: 2 pts each
- **Safety**: 2 pts each
- **Blocked punt or FG**: 2 pts each
- **Blocked PAT**: 1 pt each
- **Defensive TD**: 4 pts each (INTs, fumble recoveries, blocked kicks count; punt/kick returns don't)
- **Sack**: 1 pt each

### Head Coach (HC)
Scoring based on ESPN spread:
| Result | Points |
|--------|--------|
| Home favorite win | 4 pts |
| Road favorite win | 5 pts |
| Home underdog win | 6 pts |
| Road underdog win | 7 pts |
| Loss | 0 pts |

## Fuzzy Matching

The OPFL Excel file often contains misspelled player names or incorrect team codes. The autoscorer uses fuzzy string matching (via `thefuzz` library) to find the correct player:

- First attempts exact match
- Then tries partial name matching
- Falls back to fuzzy matching with a configurable threshold
- Searches across all teams if player not found on specified team

When a fuzzy match is made, the output shows both the original name and the matched name:

```
RB TreVeyon Henderson (NE) -> Tre'Veon Henderson: 6.0 pts ✓
```

## Excel File Format

The autoscorer reads rosters from the OPFL Excel file format:

- **Teams** are arranged horizontally in columns
- **Each team** occupies 3 columns: Points | Star (*) | Player Name
- **Position labels** (QB, RB, WR, TE, K, DF, HC) are in column A
- **Started players** are indicated by a `*` in the star column
- **Player format**: `Player Name (Team)` (e.g., "Patrick Mahomes (KC)")
- **Defense format**: Just team name (e.g., "Baltimore", "Denver")
- **Weekly sheets** are named W1, W2, W3, etc.

## Output

The autoscorer displays:
- Individual player scores with breakdowns
- ✓ indicates player found in stats
- ✗ indicates player not found (bye week, game not played, or name mismatch)
- Fuzzy match indicator when name didn't match exactly
- Final standings ranked by total points

## Notes

- Games that haven't been played yet will show players as "not found"
- Team abbreviation differences (LAR→LA, JAC→JAX, ARZ→ARI) are handled automatically
- Stats are pulled from nflverse data, updated after games complete
- All scores have a floor of 0 points (cannot go negative)
