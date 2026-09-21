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

The 2026 workbook (`OPFL Scoring 2026.xlsx`) keeps the **official rosters** on the
`Rosters` tab and each week's **head-to-head lineups and pairings** on the
`Matchups` tab. That is the default mode:

```bash
python autoscorer.py --week 1
```

### Command Line Options

| Option | Short | Default | Description |
|--------|-------|---------|-------------|
| `--excel` | `-e` | `OPFL Scoring 2026.xlsx` | Path to the Excel workbook |
| `--week` | `-w` | `1` | Week number to score |
| `--season` | `-y` | `2026` | NFL season year |
| `--sheet` | `-s` | - | Score a legacy per-week sheet (W1, W2, ...) instead |
| `--rosters-sheet` | - | `Rosters` | Name of the official rosters sheet |
| `--matchups-sheet` | - | `Matchups` | Name of the weekly matchups sheet |
| `--update` | `-u` | - | Write scores back to Excel (legacy W-sheets only) |
| `--quiet` | `-q` | - | Suppress per-player output |

### Examples

```bash
# Score week 1 with a full per-player breakdown
python autoscorer.py --week 1

# Just the matchup results and standings
python autoscorer.py --week 1 --quiet

# Score a week from the old 2025 workbook format
python autoscorer.py --excel "Griff OPFL Scoring 2025.xlsx" --sheet W12 --week 12 --season 2025
```

### Publishing to the website

`scripts/export_for_web.py` scores the week and writes `web/data.json`, which is
what the site reads:

```bash
python scripts/export_for_web.py                       # current NFL week
python scripts/export_for_web.py --week 1 --season 2026
```

The export publishes each team's full roster (starters and bench points), the
week's matchups as read from the workbook, standings, taxi squads, and draft picks.

> **The 2026 workbook is formula-driven and is never written to.** The `Matchups`
> tab pulls names and totals from the `Scoring` tab, which pulls each team's
> starred players from `Rosters`. openpyxl cannot evaluate formulas, so saving the
> file would strip every cached value. Scores are published to `web/data.json` only.

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

### 2026 workbook

**`Rosters` tab** - the official rosters, two blocks of six teams (headers on
rows 1 and 39):

- Each team occupies 3 columns: Points | Star (`*`) | Player Name
- Position labels (QB, RB, WR, TE, K, DF, HC) are in column A
- A `*` in the star column marks a starter
- Player format: `Player Name (Team)`, e.g. `Patrick Mahomes (KC)`
- Defense format: just the team name, e.g. `Baltimore`, `Denver`
- A `PH` block (phone numbers) and `TS` block (taxi squad, written as
  `RB Aaron Jones`) follow each roster and are not part of the lineup

**`Matchups` tab** - the week's six head-to-head games, stacked vertically:

- Each block starts with the two owner names in columns A and D
- The nine following rows are the starting lineup in
  QB / RB / RB / WR / WR / TE / K / DF / HC order
- The `Key` column maps schedule team numbers to owners

Owner names are spelled inconsistently across tabs (`LAM`, `Steve L`, `SL` are
all the same franchise); `opfl/constants.py:resolve_team_code` normalizes them.

### Legacy 2025 workbook

Per-week sheets named `W1`, `W2`, ... using the same roster block layout. Score
these with `--sheet W12`.

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
