#!/usr/bin/env python3
"""Import an OPFL season workbook into authoritative JSON. Dry-run by default."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from opfl.importing import attach_regular_matchups, import_workbook, load_teams
from opfl.league import season_config
from opfl.schedule import parse_schedule_file
from opfl.standings import calculate_playoffs, calculate_standings


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def build_season(
    season: int, excel_path: Path, schedule_path: Path, data_root: Path
) -> dict[str, Any]:
    teams = load_teams(data_root / "teams.json")
    config = season_config(season, data_root.parent)
    schedule = parse_schedule_file(schedule_path, (team["abbrev"] for team in teams))
    if len(schedule) != int(config["regular_season_weeks"]):
        raise ValueError(f"Season {season} requires 15 complete regular-season weeks")
    rosters, weeks, lineups = import_workbook(excel_path, teams, season)
    attach_regular_matchups(weeks, schedule)
    standings = calculate_standings(teams, schedule, weeks.values(), 15)
    playoffs = calculate_playoffs(standings, weeks.values()) if {16, 17}.issubset(weeks) else None
    if playoffs:
        weeks[16]["matchups"] = playoffs["semifinals"]
        weeks[17]["matchups"] = [
            game for game in (playoffs["championship"], playoffs["third_place"]) if game
        ]
    metadata = {
        "year": season,
        "detail_level": "full",
        "status": "final" if season == 2025 else "preseason",
        "weeks_available": sorted(weeks),
        "regular_season_weeks": config["regular_season_weeks"],
        "playoff_weeks": config["playoff_weeks"],
        "capabilities": {
            "standings": True,
            "advanced_standings": True,
            "matchups": True,
            "schedules": True,
            "rosters": True,
            "player_stats": True,
            "transactions": True,
            "drafts": True,
            "history": True,
        },
        "configuration": config,
        "source": {
            "type": "excel_import",
            "workbook": excel_path.name,
            "schedule": schedule_path.name,
            "json_authoritative_after_import": True,
        },
    }
    return {
        "metadata": metadata,
        "rosters": rosters,
        "schedule": {"season": season, "weeks": schedule},
        "weeks": weeks,
        "lineups": lineups,
        "standings": {"season": season, "standings": standings},
        "playoffs": {"season": season, **playoffs} if playoffs else None,
    }


def apply_season(payload: dict[str, Any], data_root: Path, replace_existing: bool = False) -> Path:
    season = int(payload["metadata"]["year"])
    season_dir = data_root / "seasons" / str(season)
    managed_targets = [
        season_dir / "metadata.json",
        season_dir / "rosters.json",
        season_dir / "schedule.json",
        season_dir / "standings.json",
        season_dir / "playoffs.json",
    ]
    managed_targets.extend(season_dir / "weeks" / f"week_{week}.json" for week in payload["weeks"])
    managed_targets.extend(
        season_dir / "lineups" / f"week_{week}.json" for week in payload["lineups"]
    )
    existing = [path for path in managed_targets if path.exists()]
    if existing and not replace_existing:
        raise FileExistsError(
            f"Refusing to overwrite {len(existing)} season files; pass --replace-existing"
        )
    write_json(season_dir / "metadata.json", payload["metadata"])
    write_json(season_dir / "rosters.json", payload["rosters"])
    write_json(season_dir / "schedule.json", payload["schedule"])
    write_json(season_dir / "standings.json", payload["standings"])
    if payload["playoffs"]:
        write_json(season_dir / "playoffs.json", payload["playoffs"])
    for week, value in payload["weeks"].items():
        write_json(season_dir / "weeks" / f"week_{week}.json", value)
    for week, value in payload["lineups"].items():
        write_json(season_dir / "lineups" / f"week_{week}.json", value)
    return season_dir


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--season", required=True, type=int)
    result.add_argument("--excel", required=True, type=Path)
    result.add_argument("--schedule", required=True, type=Path)
    result.add_argument("--data-root", default="data", type=Path)
    result.add_argument("--apply", action="store_true")
    result.add_argument("--replace-existing", action="store_true")
    return result


def main() -> int:
    args = parser().parse_args()
    payload = build_season(args.season, args.excel, args.schedule, args.data_root)
    print(
        f"Validated season {args.season}: 12 teams, {len(payload['weeks'])} week sheets, "
        "15 complete schedule weeks."
    )
    if not args.apply:
        print("Dry run: no files written. Pass --apply to import.")
        return 0
    path = apply_season(payload, args.data_root, args.replace_existing)
    print(f"Wrote authoritative season JSON to {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
