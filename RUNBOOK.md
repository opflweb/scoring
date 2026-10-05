# Weekly Runbook

What to do each week to get OPFL scores onto the site. Once the workbook is
pushed, the GitHub Action (`.github/workflows/score.yml`) does the scoring,
publishing, and emails on its own schedule. Your job is to keep the workbook
correct.

## 1. Update the workbook (`OPFL Scoring 2026.xlsx`)

Do this before Thursday night's kickoff, and again whenever lineups change.

### Rosters tab

- **Transactions:** replace dropped players with the added ones. Use the format
  `Player Name (TEAM)`, for example `Jalen Coker (Car)`. Write a defense as just
  the team name, for example `Detroit`. Taxi squad entries in the `TS` block use
  the format `RB Aaron Jones`.
- **Lineups:** clear last week's `*` marks and star exactly **9 starters per
  team** in the star column (the column just left of the player names):
  QB, RB, RB, WR, WR, TE, K, DF, HC.
  - A team with no stars scores **0**, because every player counts as bench.

### Matchups tab: nothing to do

The scorer doesn't read this tab for regular-season weeks. Pairings come from
`data/schedules/2026.json`, so the tab can show any week. It's only used for
weeks the schedule file doesn't cover (the playoffs, weeks 16 and 17). Set it
correctly for those weeks.

### Last week's archive (commissioner bookkeeping)

- Copy last week's final lineups and points into a static `W{n}` sheet.
- Update the **Standings** tab.

The scorer doesn't read either one, but `scripts/backfill_weeks.py` relies on
the `W{n}` sheets to rebuild a season.

> Don't save the workbook from a script. It's formula-driven and openpyxl
> strips the cached values (see `tests/test_no_workbook_writes.py`). Edit it in
> Excel or Sheets only.

## 2. Check it locally

```bash
uv sync                                   # first time / after dependency changes
uv run python scripts/nfl_week.py --season 2026     # which week the Action will score
uv run python autoscorer.py --week 4 --quiet        # matchups + standings only
uv run python autoscorer.py --week 4 > week4.txt    # full per-player breakdown
```

Check for these:

- [ ] **No `WARNING: ... starters on the Rosters tab` lines.** These flag a team
      with missing or extra stars, for example `RB 1/2`.
- [ ] **Fuzzy matches look right.** `Name (TEAM) -> Other Name` lines are fine
      for spelling fixes (`JK Dobbins -> J.K. Dobbins`). The scorer now rejects
      matches to a different player (`Josh Jacobs -> Josh Jobe`). If one still
      gets through, fix the spelling or team in the workbook.
- [ ] A `✗` on a starter only means "no stats yet". That's expected before their
      game is played or before nflverse publishes it (about 30 min to a few hours
      after the game).

## 3. Commit and push the workbook

```bash
git add "OPFL Scoring 2026.xlsx"
git commit -m "Week 4 lineups"
git pull --rebase    # the Action commits score updates to main too
git push
```

The Action only sees what's on `main`. Local edits do nothing until you push them.

## 4. Let the Action run (or trigger it)

It runs automatically during the season (all times ET):

| When | Run | Email? |
|------|-----|--------|
| Daily 5:30 AM | catch-all | no |
| Fri 1:00 AM | after TNF | yes |
| Sun 5:30 PM | early games | no |
| Sun 7:35 PM | late games | no |
| Mon 1:00 AM | after SNF | yes |
| Tue 1:00 AM | after MNF | yes |

To score right away after pushing:

```bash
gh workflow run score.yml            # auto-detects the week
gh workflow run score.yml -f week=4  # or force a specific week
gh run watch                         # follow it
```

Each run scores the week, writes `web/data.json` and `data/weeks/`, commits as
`github-actions[bot]`, and deploys the site. A week is archived as **final** once
every team that played has stats in nflverse.

## 5. Verify

- Check the site at https://opflweb.github.io/scoring/ (hard-refresh if it looks
  stale).
- The score breakdown email reaches griffin.ansel@gmail.com after TNF, SNF, and MNF.
- If a run fails, the commissioner gets an alert email with the run link. Use
  `gh run view --log-failed` to see what broke, then re-run with
  `gh run rerun <id>`.

## Fixing a past week

The Matchups tab only holds the live week, so correct finished weeks from their
archive:

```bash
# Rescore against the latest nflverse stats (stat corrections, late data)
uv run python scripts/rescore_archived_week.py --week 3

# Fix a lineup mistake
uv run python scripts/rescore_archived_week.py --week 3 \
    --start KEV "Deebo Samuel" --bench KEV "Zay Flowers"

# Rebuild web/data.json and standings, then commit and push
uv run python scripts/export_for_web.py --week 4 --season 2026
git add web/data.json data/ && git commit -m "Rescore week 3" && git push
```

## Weekly checklist

- [ ] Rosters: transactions entered
- [ ] Rosters: old stars cleared, 9 new stars per team (all 12 teams)
- [ ] Playoffs only: Matchups tab set to the week's games
- [ ] Last week: `W{n}` sheet + Standings updated
- [ ] `autoscorer.py --week N --quiet` looks right
- [ ] Workbook committed and pushed
- [ ] Site updated after the next run (or `gh workflow run score.yml`)
