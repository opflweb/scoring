# OPFL Autoscorer and Archive

The Oakland Perennial Football League site is a read-only, GitHub Pages–first archive. Seasons from 1988–2024 contain summary standings and finishes. The 2025 season contains full weekly lineups, player scores, matchups, playoffs, and the two-week Jamboree.

Excel is an import and troubleshooting source only. After a season or lineup is imported, `data/` JSON is authoritative for scoring, integrity checks, and public exports.

## Data model

```text
data/
  league.json                     season-aware rules and public season list
  teams.json                      twelve-team registry
  schedules/2025.txt              canonical import schedule
  archive/
    seasons/1988.json … 2024.json summary archives
    shared/                       history, drafts, trades, picks, rules, banners
  seasons/2025/
    metadata.json
    rosters.json
    schedule.json
    standings.json
    playoffs.json
    lineups/week_N.json
    weeks/week_N.json
web/
  index.html
  styles.css
  app.js
  data/                            generated split public tree
  data.json                        one-release current-season compatibility snapshot
```

Public season metadata declares capability flags. Summary seasons intentionally disable weekly matchups, schedules, rosters, player stats, and advanced standings instead of returning misleading empty values.

## Setup and checks

Use the frozen lockfile:

```bash
uv sync --frozen --all-extras
```

Run the full local verification suite:

```bash
.venv/bin/ruff check opfl scripts tests
.venv/bin/ruff format --check opfl scripts tests
.venv/bin/mypy opfl scripts
.venv/bin/pytest --cov --cov-report=term-missing
.venv/bin/python scripts/validate_data.py
node --check web/app.js
```

Serve the static site locally from `web/`:

```bash
python -m http.server 8000 --directory web
```

## Safe commissioner imports

Both import commands are dry-run by default. `--apply` is required to write files. If a target JSON file already exists, `--replace-existing` is also required.

Import a complete season:

```bash
.venv/bin/python scripts/import_season.py \
  --season 2026 \
  --excel "OPFL Scoring 2026.xlsx" \
  --schedule data/schedules/2026.txt
```

After reviewing the validation summary:

```bash
.venv/bin/python scripts/import_season.py \
  --season 2026 \
  --excel "OPFL Scoring 2026.xlsx" \
  --schedule data/schedules/2026.txt \
  --apply
```

Import one lineup week:

```bash
.venv/bin/python scripts/import_lineups.py \
  --season 2026 --week 1 \
  --excel "OPFL Scoring 2026.xlsx" --sheet W1
```

Schedules must use exactly six lines per week, with every team appearing once:

```text
Week 1: KAM versus D/J
Week 1: STL versus K/D
```

The importer requires every regular-season week from 1 through 15. A 2026 season is not added to `public_seasons` until its roster and schedule pass validation.

## JSON scoring

Score a week from JSON rosters and lineups:

```bash
.venv/bin/python scripts/score_json_week.py --season 2026 --week 1
```

The command scores starters and bench players for statistics, but only the nine configured starters count toward the team total. It supports player score overrides and team-level manual corrections in the week lineup JSON. Saving and overwriting use the same explicit safety flags:

```bash
.venv/bin/python scripts/score_json_week.py \
  --season 2026 --week 1 --apply --replace-existing
```

The lineup schema is configuration-driven. `FLEX` already accepts RB/WR/TE when its count is enabled, but its count remains zero for 2026.

## Standings rules

Rank Points are:

```text
wins + Top-6 points + 0.5 × matchup ties
```

Each of the six weekly top-half places is worth one point. When a scoring tie crosses sixth place, the remaining points are split evenly among all teams in the tie. Advanced standings also publish PF, average PF, PA, differential, all-play expected record normalized across eleven opponents, luck, remaining schedule strength, and weekly rank history.

Standings ties are ordered by Rank Points, matchup wins, Top-6 points, and then points for.

Four teams reach the playoffs. The other eight teams compete in the two-week Jamboree during Weeks 16 and 17.

## Public export

Regenerate the static tree and compatibility snapshot:

```bash
.venv/bin/python scripts/export_public.py
.venv/bin/python scripts/validate_data.py
```

The export is built in a staging directory and replaces `web/data/` only after source validation succeeds. Vercel is configured for static files only; the incomplete write endpoints under `api/` are not deployed.

## Automation

CI runs Ruff, Mypy, pytest with coverage, schema checks, cross-file integrity, and JavaScript syntax validation. The scoring workflow reads the active public season from `data/league.json`, scores only from JSON, validates before and after export, commits generated data when it changes, and deploys `web/` to GitHub Pages.

The 2025 Excel scorer remains available only for archive reconciliation and troubleshooting. It is not part of automated scoring or public export input.
