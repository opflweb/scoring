#!/usr/bin/env python3
"""Import one OPFL lineup week from Excel. Dry-run by default."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from opfl.importing import extract_sheet, load_teams


def build_lineups(
    season: int, week: int, excel_path: Path, sheet_name: str, data_root: Path
) -> dict[str, object]:
    teams = load_teams(data_root / "teams.json")
    workbook = openpyxl.load_workbook(excel_path, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Workbook has no sheet named {sheet_name}")
        extracted = extract_sheet(workbook[sheet_name], teams, season, week)
    finally:
        workbook.close()
    roster_path = data_root / "seasons" / str(season) / "rosters.json"
    with roster_path.open(encoding="utf-8") as handle:
        rosters = json.load(handle)["teams"]
    lineups: dict[str, list[str]] = {}
    for team in extracted:
        abbreviation = team["abbrev"]
        roster_ids = {player["player_id"] for player in rosters[abbreviation]["players"]}
        starter_ids = [player["player_id"] for player in team["roster"] if player["starter"]]
        missing = set(starter_ids) - roster_ids
        if missing:
            raise ValueError(
                f"{abbreviation} lineup contains players absent from authoritative rosters: "
                f"{', '.join(sorted(missing))}"
            )
        lineups[abbreviation] = starter_ids
    return {"season": season, "week": week, "lineups": lineups, "manual_corrections": {}}


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--season", required=True, type=int)
    result.add_argument("--week", required=True, type=int)
    result.add_argument("--excel", required=True, type=Path)
    result.add_argument("--sheet", required=True)
    result.add_argument("--data-root", default="data", type=Path)
    result.add_argument("--apply", action="store_true")
    result.add_argument("--replace-existing", action="store_true")
    return result


def main() -> int:
    args = parser().parse_args()
    payload = build_lineups(args.season, args.week, args.excel, args.sheet, args.data_root)
    print(f"Validated 12 lineups for season {args.season}, week {args.week}.")
    if not args.apply:
        print("Dry run: no files written. Pass --apply to import.")
        return 0
    path = args.data_root / "seasons" / str(args.season) / "lineups" / f"week_{args.week}.json"
    if path.exists() and not args.replace_existing:
        raise FileExistsError(f"Refusing to overwrite {path}; pass --replace-existing")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
