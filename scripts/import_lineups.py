#!/usr/bin/env python3
"""Import one OPFL lineup week from its required Excel workbook."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from opfl.importing import import_weekly_workbook, load_teams, weekly_workbook_path
from opfl.integrity import read_json


def build_lineups(
    season: int, week: int, excel_path: Path | None, data_root: Path
) -> dict[str, Any]:
    teams = load_teams(data_root / "teams.json")
    expected_path = weekly_workbook_path(data_root, season, week)
    source = excel_path or expected_path
    if source.resolve() != expected_path.resolve():
        raise ValueError(f"Season {season} week {week} must use {expected_path}")
    if not source.exists():
        raise FileNotFoundError(f"Missing weekly workbook: {source}")
    _rosters, lineups = import_weekly_workbook(source, teams, season, week)
    schedule_path = data_root / "seasons" / str(season) / "schedule.json"
    if schedule_path.exists() and week <= 15:
        schedule = read_json(schedule_path)["weeks"]
        expected = next(item["matchups"] for item in schedule if int(item["week"]) == week)
        actual_pairs = {frozenset((item["away"], item["home"])) for item in lineups["pairings"]}
        expected_pairs = {frozenset((item["away"], item["home"])) for item in expected}
        if actual_pairs != expected_pairs:
            raise ValueError(f"Workbook matchups do not match the Week {week} schedule")
    return lineups


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--season", required=True, type=int)
    result.add_argument("--week", required=True, type=int)
    result.add_argument(
        "--excel",
        type=Path,
        help="Must equal data/seasons/SEASON/workbooks/week_WEEK.xlsx",
    )
    result.add_argument("--data-root", default="data", type=Path)
    result.add_argument("--apply", action="store_true")
    result.add_argument("--replace-existing", action="store_true")
    return result


def main() -> int:
    args = parser().parse_args()
    payload = build_lineups(args.season, args.week, args.excel, args.data_root)
    print(f"Validated 12 lineups for season {args.season}, week {args.week}.")
    if not args.apply:
        print("Dry run: no files written. Pass --apply to import.")
        return 0
    path = args.data_root / "seasons" / str(args.season) / "lineups" / f"week_{args.week}.json"
    if path.exists() and not args.replace_existing:
        raise FileExistsError(f"Refusing to overwrite {path}; pass --replace-existing")
    if path.exists():
        existing = read_json(path)
        for field in ("player_score_overrides", "manual_corrections"):
            if existing.get(field):
                payload[field] = existing[field]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    workbook_dir = path.parent.parent / "workbooks"
    available_weeks = [
        int(candidate.stem.split("_")[1]) for candidate in workbook_dir.glob("week_*.xlsx")
    ]
    if args.week == max(available_weeks, default=args.week):
        roster_path = path.parent.parent / "rosters.json"
        roster_path.write_text(
            json.dumps(payload["roster_snapshot"], indent=2) + "\n", encoding="utf-8"
        )
        print(f"Updated current roster snapshot at {roster_path}")
    print(f"Wrote {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
