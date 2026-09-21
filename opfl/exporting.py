"""Build the split, static public data tree from authoritative OPFL JSON."""

from __future__ import annotations

import json
import os
import shutil
import statistics
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .integrity import read_json, validate_archives, validate_source_season


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def export_public_data(
    data_root: Path, output_root: Path, compatibility_path: Path | None = None
) -> dict[str, Any]:
    validate_archives(data_root)
    league = read_json(data_root / "league.json")
    public_years = [int(year) for year in league["public_seasons"]]
    current = int(league["current_public_season"])
    output_root.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=".opfl-data-build-", dir=output_root.parent))
    try:
        entries = []
        for year in public_years:
            destination = staging / "seasons" / str(year)
            if year < 2025:
                summary = read_json(data_root / "archive" / "seasons" / f"{year}.json")
                metadata = _summary_metadata(year)
                write_json(destination / "metadata.json", metadata)
                write_json(
                    destination / "standings.json",
                    {
                        "season": year,
                        "detail_level": "summary",
                        "standings": summary["standings"],
                        "finish": summary["finish"],
                    },
                )
            else:
                source = data_root / "seasons" / str(year)
                validate_source_season(source, data_root.parent)
                metadata = read_json(source / "metadata.json")
                for name in (
                    "metadata.json",
                    "standings.json",
                    "rosters.json",
                    "schedule.json",
                    "playoffs.json",
                ):
                    source_path = source / name
                    if source_path.exists():
                        _copy(source_path, destination / name)
                for week in metadata["weeks_available"]:
                    _copy(
                        source / "weeks" / f"week_{week}.json",
                        destination / "weeks" / f"week_{week}.json",
                    )
                standings = read_json(source / "standings.json")["standings"]
                write_json(destination / "team-stats.json", _team_stats(year, standings))
                write_json(destination / "player-stats.json", _player_stats(source, metadata))
            entries.append(
                {
                    "year": year,
                    "detail_level": metadata["detail_level"],
                    "weeks_available": metadata["weeks_available"],
                    "capabilities": metadata["capabilities"],
                }
            )
        for source_name, public_name in (
            ("history.json", "history.json"),
            ("drafts.json", "drafts.json"),
            ("transactions.json", "transactions.json"),
            ("future-picks.json", "future-picks.json"),
            ("rules.json", "rules.json"),
            ("banners.json", "banners.json"),
        ):
            _copy(data_root / "archive" / "shared" / source_name, staging / "shared" / public_name)
        generated_at = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
        index = {
            "league": league["league_name"],
            "current_season": current,
            "generated_at": generated_at,
            "seasons": entries,
            "shared": {
                "history": "shared/history.json",
                "drafts": "shared/drafts.json",
                "transactions": "shared/transactions.json",
                "future_picks": "shared/future-picks.json",
                "rules": "shared/rules.json",
                "banners": "shared/banners.json",
            },
        }
        write_json(staging / "index.json", index)
        _replace_generated_tree(staging, output_root)
        if compatibility_path:
            write_json(
                compatibility_path,
                _compatibility_snapshot(data_root, output_root, current, generated_at),
            )
        return index
    except Exception:
        if staging.exists():
            shutil.rmtree(staging)
        raise


def _summary_metadata(year: int) -> dict[str, Any]:
    return {
        "year": year,
        "detail_level": "summary",
        "status": "final",
        "weeks_available": [],
        "capabilities": {
            "standings": True,
            "advanced_standings": False,
            "matchups": False,
            "schedules": False,
            "rosters": False,
            "player_stats": False,
            "transactions": True,
            "drafts": year >= 2020,
            "history": True,
        },
    }


def _team_stats(year: int, standings: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "season": year,
        "teams": [
            {
                key: row.get(key)
                for key in (
                    "rank",
                    "abbrev",
                    "name",
                    "wins",
                    "losses",
                    "ties",
                    "points_for",
                    "points_against",
                    "point_differential",
                    "average_points_for",
                    "high_score",
                    "low_score",
                    "median_score",
                )
            }
            for row in standings
        ],
    }


def _player_stats(source: Path, metadata: dict[str, Any]) -> dict[str, Any]:
    players: dict[str, dict[str, Any]] = {}
    for week in metadata["weeks_available"]:
        week_data = read_json(source / "weeks" / f"week_{week}.json")
        for team in week_data["teams"]:
            for player in team["roster"]:
                if player.get("placeholder"):
                    continue
                item = players.setdefault(
                    player["player_id"],
                    {
                        "player_id": player["player_id"],
                        "name": player["name"],
                        "position": player["position"],
                        "nfl_team": player.get("nfl_team", ""),
                        "fantasy_teams": set(),
                        "scores": [],
                        "starts": 0,
                    },
                )
                item["fantasy_teams"].add(team["abbrev"])
                item["scores"].append(float(player["score"]))
                item["starts"] += int(player["starter"])
    output = []
    for item in players.values():
        scores = item.pop("scores")
        item["fantasy_teams"] = sorted(item["fantasy_teams"])
        item["games"] = len(scores)
        item["points"] = round(sum(scores), 2)
        item["ppg"] = round(sum(scores) / len(scores), 2) if scores else 0.0
        item["high"] = max(scores) if scores else 0.0
        item["median"] = round(statistics.median(scores), 2) if scores else 0.0
        output.append(item)
    output.sort(key=lambda item: (-item["points"], item["name"]))
    return {"season": metadata["year"], "players": output}


def _compatibility_snapshot(
    data_root: Path, public_root: Path, season: int, generated_at: str
) -> dict[str, Any]:
    season_root = public_root / "seasons" / str(season)
    metadata = read_json(season_root / "metadata.json")
    standings = read_json(season_root / "standings.json")["standings"]
    rosters = read_json(season_root / "rosters.json")["teams"]
    weeks = [
        read_json(season_root / "weeks" / f"week_{week}.json")
        for week in metadata["weeks_available"]
    ]
    shared = public_root / "shared"
    return {
        "updated_at": generated_at,
        "season": season,
        "current_week": max(metadata["weeks_available"], default=0),
        "regular_season_weeks": metadata["regular_season_weeks"],
        "weeks": weeks,
        "standings": [{**row, "top_half": row["top_six"]} for row in standings],
        "playoffs": (
            read_json(season_root / "playoffs.json")
            if (season_root / "playoffs.json").exists()
            else None
        ),
        "trade_deadline_week": 12,
        "taxi_squads": {team: roster.get("taxi", []) for team, roster in rosters.items()},
        "pending_trades": [],
        "draft_picks": read_json(shared / "future-picks.json"),
        "banners": [
            item["image"].split("/")[-1] for item in read_json(shared / "banners.json")["banners"]
        ],
    }


def _copy(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)


def _replace_generated_tree(staging: Path, output: Path) -> None:
    backup = output.parent / f".{output.name}-previous"
    if backup.exists():
        shutil.rmtree(backup)
    if output.exists():
        os.replace(output, backup)
    os.replace(staging, output)
    if backup.exists():
        shutil.rmtree(backup)
