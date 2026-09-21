"""Per-week score archive.

The workbook's Matchups tab only ever holds the *current* week — the
commissioner overwrites it each week and archives the finished one to a `W{n}`
sheet. That means a single export run can only ever see one week, which is not
enough for season totals, standings, leaders or records.

So every run writes what it scored to `data/weeks/{season}/week_{n}.json` and
the exporter reads the whole season back out. The archive is a derived cache,
not a source of truth: it can be regenerated from the workbook's `W{n}` sheets
at any time via `scripts/backfill_weeks.py`.

Weeks are archived eagerly rather than only once complete, because if the
Matchups tab were overwritten for week N+1 before a run captured week N and
before its `W{n}` sheet existed, that week's lineups would be unrecoverable. A
week is rewritten on every run until its NFL games are all final, after which
it is frozen unless explicitly forced.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / 'data'


def season_dir(season: int, data_dir: Path | str = DATA_DIR) -> Path:
    return Path(data_dir) / 'weeks' / str(season)


def archive_path(season: int, week: int, data_dir: Path | str = DATA_DIR) -> Path:
    return season_dir(season, data_dir) / f'week_{week}.json'


def load_week(season: int, week: int, data_dir: Path | str = DATA_DIR) -> dict | None:
    """Read one archived week, or None if it has not been archived."""
    path = archive_path(season, week, data_dir)
    if not path.exists():
        return None
    return json.loads(path.read_text())


def save_week(
    season: int,
    week: int,
    week_data: dict,
    pairings: list[list[str]],
    final: bool,
    data_dir: Path | str = DATA_DIR,
    force: bool = False,
) -> bool:
    """Archive a scored week. Returns True if the file was written.

    A week already archived as final is left alone — nflverse occasionally
    restates stats weeks later, and a completed week's official result should
    not drift underneath the standings. Pass `force` to rescore it anyway.
    """
    existing = load_week(season, week, data_dir)
    if existing and existing.get('final') and not force:
        return False

    path = archive_path(season, week, data_dir)
    path.parent.mkdir(parents=True, exist_ok=True)

    payload = {
        'season': season,
        'week': week,
        'final': final,
        'scored_at': datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
        'has_scores': week_data.get('has_scores', False),
        'pairings': pairings,
        'teams': week_data['teams'],
    }
    path.write_text(json.dumps(payload, indent=2))
    return True


def load_all_weeks(
    season: int, data_dir: Path | str = DATA_DIR
) -> tuple[list[dict], dict[str, list[list[str]]]]:
    """Read the whole archived season.

    Returns:
        (weeks, schedule) in the shape web/data.json expects — weeks sorted by
        week number, schedule keyed by week number as a string.
    """
    directory = season_dir(season, data_dir)
    if not directory.exists():
        return [], {}

    weeks, schedule = [], {}
    for path in sorted(directory.glob('week_*.json'), key=_week_number):
        # A stray or half-written file in here must not take down the whole
        # export - skip it loudly and carry on with the weeks that do parse.
        try:
            archived = json.loads(path.read_text())
            entry = {
                'week': archived['week'],
                'teams': archived['teams'],
                'has_scores': archived.get('has_scores', False),
                'final': archived.get('final', False),
            }
        except (json.JSONDecodeError, KeyError, TypeError) as e:
            print(f'Warning: skipping unreadable week archive {path.name}: {e}')
            continue

        weeks.append(entry)
        if archived.get('pairings'):
            schedule[str(archived['week'])] = archived['pairings']

    return weeks, schedule


def _week_number(path: Path) -> int:
    """Sort key so week_10 lands after week_9 rather than after week_1."""
    try:
        return int(path.stem.removeprefix('week_'))
    except ValueError:
        return 0
