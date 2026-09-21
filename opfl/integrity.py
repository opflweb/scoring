"""Cross-file integrity checks for authoritative and public OPFL JSON."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .league import POSITIONS, season_config, starter_count
from .schedule import validate_schedule


class IntegrityError(ValueError):
    """Raised when related OPFL JSON files disagree."""


def read_json(path: str | Path) -> Any:
    try:
        with Path(path).open(encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        raise IntegrityError(f"Unable to read JSON file {path}") from exc


def validate_source_season(season_dir: str | Path, root: Path | None = None) -> list[str]:
    directory = Path(season_dir)
    metadata = read_json(directory / "metadata.json")
    season = int(metadata["year"])
    config = season_config(season, root)
    teams_document = read_json((root or directory.parents[2]) / "data" / "teams.json")
    teams = teams_document["teams"]
    abbreviations = {team["abbrev"] for team in teams}
    if len(teams) != 12 or len(abbreviations) != 12:
        raise IntegrityError("Team registry must contain twelve unique teams")
    schedule_document = read_json(directory / "schedule.json")
    schedule = schedule_document["weeks"]
    validate_schedule(schedule, abbreviations, int(config["regular_season_weeks"]))
    rosters = read_json(directory / "rosters.json")
    if set(rosters["teams"]) != abbreviations:
        raise IntegrityError("Rosters must contain each registered team exactly once")
    player_owners: dict[str, str] = {}
    limits = config["roster_limits"]
    for abbreviation, roster in rosters["teams"].items():
        counts = {position: 0 for position in POSITIONS}
        for player in roster["players"]:
            position = player.get("position")
            if position not in POSITIONS:
                raise IntegrityError(f"{abbreviation} has invalid position {position!r}")
            counts[position] += 1
            player_id = player["player_id"]
            if player_id in player_owners:
                raise IntegrityError(
                    f"Player {player_id} appears on {player_owners[player_id]} and {abbreviation}"
                )
            player_owners[player_id] = abbreviation
        for position, expected in limits.items():
            if counts[position] != int(expected):
                raise IntegrityError(
                    f"{abbreviation} has {counts[position]} {position} players; expected {expected}"
                )
        if len(roster.get("taxi", [])) > int(config["taxi_limit"]):
            raise IntegrityError(f"{abbreviation} exceeds the taxi limit")

    weeks_available = [int(week) for week in metadata.get("weeks_available", [])]
    weekly_excel = metadata.get("source", {}).get("type") == "weekly_excel"
    for week in weeks_available:
        week_path = directory / "weeks" / f"week_{week}.json"
        lineup_path = directory / "lineups" / f"week_{week}.json"
        week_data = read_json(week_path)
        lineup_data = read_json(lineup_path)
        if weekly_excel:
            expected_source = f"workbooks/week_{week}.xlsx"
            if lineup_data.get("source_workbook") != expected_source:
                raise IntegrityError(
                    f"Week {week} must identify {expected_source} as its source workbook"
                )
            if not (directory / expected_source).is_file():
                raise IntegrityError(f"Week {week} is missing source workbook {expected_source}")
        week_teams = week_data.get("teams", [])
        if len(week_teams) != 12 or {item["abbrev"] for item in week_teams} != abbreviations:
            raise IntegrityError(f"Week {week} must contain all twelve teams")
        if set(lineup_data.get("lineups", {})) != abbreviations:
            raise IntegrityError(f"Week {week} lineups must contain all twelve teams")
        for team in week_teams:
            roster_by_id = {player["player_id"]: player for player in team["roster"]}
            starters = lineup_data["lineups"][team["abbrev"]]
            if len(starters) != starter_count(config) or len(starters) != len(set(starters)):
                raise IntegrityError(
                    f"Week {week} {team['abbrev']} must contain {starter_count(config)} unique starters"
                )
            if not set(starters).issubset(roster_by_id):
                raise IntegrityError(f"Week {week} {team['abbrev']} starter is not on its roster")
            for player in roster_by_id.values():
                if player["position"] not in POSITIONS:
                    raise IntegrityError(
                        f"Week {week} {team['abbrev']} has invalid position {player['position']}"
                    )
                if float(player["score"]) < 0:
                    raise IntegrityError(
                        f"Week {week} {team['abbrev']} contains a score below the OPFL floor"
                    )
            starter_total = sum(float(roster_by_id[player]["score"]) for player in starters)
            expected_total = starter_total + float(team.get("manual_correction", 0))
            if abs(expected_total - float(team["total_score"])) > 0.001:
                raise IntegrityError(f"Week {week} {team['abbrev']} total does not reconcile")
    actual_week_files = sorted(
        int(path.stem.split("_")[1]) for path in (directory / "weeks").glob("week_*.json")
    )
    if actual_week_files != weeks_available:
        raise IntegrityError("metadata weeks_available does not match week files")
    return [
        f"season {season}: 12 teams",
        f"season {season}: {len(weeks_available)} detailed weeks",
        f"season {season}: schedule and cross-file integrity valid",
    ]


def validate_archives(data_root: str | Path) -> list[str]:
    root = Path(data_root)
    season_files = sorted((root / "archive" / "seasons").glob("*.json"))
    years = [int(path.stem) for path in season_files]
    if years != list(range(1988, 2025)):
        raise IntegrityError("Summary archives must contain every season from 1988 through 2024")
    drafts = read_json(root / "archive" / "shared" / "drafts.json")["drafts"]
    trades = read_json(root / "archive" / "shared" / "transactions.json")["trades"]
    banners = read_json(root / "archive" / "shared" / "banners.json")["banners"]
    if len(drafts) != 12:
        raise IntegrityError(f"Expected 12 draft boards, found {len(drafts)}")
    if len(trades) != 33:
        raise IntegrityError(f"Expected 33 archived trades, found {len(trades)}")
    if len(banners) != 37:
        raise IntegrityError(f"Expected 37 banners, found {len(banners)}")
    return ["archives: 37 summary seasons", "archives: 12 drafts, 33 trades, 37 banners"]


def validate_public_tree(web_data_root: str | Path) -> list[str]:
    root = Path(web_data_root)
    index = read_json(root / "index.json")
    entries = index.get("seasons", [])
    years = [int(entry["year"]) for entry in entries]
    current_season = int(index["current_season"])
    if years != list(range(1988, current_season + 1)):
        raise IntegrityError(
            f"Public index must contain every season from 1988 through {current_season}"
        )
    for entry in entries:
        year = int(entry["year"])
        metadata = read_json(root / "seasons" / str(year) / "metadata.json")
        if metadata["detail_level"] != entry["detail_level"]:
            raise IntegrityError(f"Season {year} detail level disagrees with the public index")
        if metadata["weeks_available"] != entry["weeks_available"]:
            raise IntegrityError(f"Season {year} week availability disagrees with the public index")
        capabilities = metadata.get("capabilities", {})
        if year < 2025 and any(
            capabilities.get(capability)
            for capability in ("matchups", "player_stats", "rosters", "advanced_standings")
        ):
            raise IntegrityError(f"Summary season {year} exposes an unavailable capability")
    return [
        f"public tree: {len(entries)} indexed seasons",
        "public tree: summary capabilities are explicit",
    ]
