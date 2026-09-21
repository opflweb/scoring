"""Season-aware OPFL league configuration."""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any

POSITIONS = ("QB", "RB", "WR", "TE", "K", "DF", "HC")
TEAM_COUNT = 12


class ConfigurationError(ValueError):
    """Raised when league configuration is missing or inconsistent."""


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def league_config_path(root: Path | None = None) -> Path:
    return (root or project_root()) / "data" / "league.json"


def load_league_config(root: Path | None = None) -> dict[str, Any]:
    path = league_config_path(root)
    try:
        with path.open(encoding="utf-8") as handle:
            config: dict[str, Any] = json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        raise ConfigurationError(f"Unable to load league configuration: {path}") from exc
    validate_league_config(config)
    return config


def season_config(season: int, root: Path | None = None) -> dict[str, Any]:
    league = load_league_config(root)
    result = deepcopy(league["defaults"])
    override = league.get("season_overrides", {}).get(str(season), {})
    _deep_update(result, override)
    result["season"] = season
    result["league_name"] = league["league_name"]
    result["team_count"] = league["team_count"]
    return result


def active_lineup_slots(config: dict[str, Any]) -> list[dict[str, Any]]:
    return [slot for slot in config["lineup_slots"] if int(slot.get("count", 0)) > 0]


def starter_count(config: dict[str, Any]) -> int:
    return sum(int(slot["count"]) for slot in active_lineup_slots(config))


def eligible_slot_count(config: dict[str, Any], position: str) -> int:
    return sum(
        int(slot["count"])
        for slot in active_lineup_slots(config)
        if position in slot.get("eligible_positions", [])
    )


def validate_league_config(config: dict[str, Any]) -> None:
    if config.get("team_count") != TEAM_COUNT:
        raise ConfigurationError("OPFL must have exactly 12 teams")
    defaults = config.get("defaults")
    if not isinstance(defaults, dict):
        raise ConfigurationError("League configuration requires defaults")
    if defaults.get("regular_season_weeks") != 15:
        raise ConfigurationError("OPFL regular season must contain 15 weeks")
    if defaults.get("playoff_seeds") != 4:
        raise ConfigurationError("OPFL playoffs must contain four seeds")
    jamboree = defaults.get("jamboree", {})
    if jamboree.get("teams") != 8 or jamboree.get("weeks") != 2:
        raise ConfigurationError("OPFL Jamboree must contain eight teams over two weeks")
    slots = defaults.get("lineup_slots", [])
    if not isinstance(slots, list):
        raise ConfigurationError("lineup_slots must be a list")
    for slot in slots:
        eligible = slot.get("eligible_positions", [])
        if not eligible or any(position not in POSITIONS for position in eligible):
            raise ConfigurationError(f"Invalid eligible positions for {slot.get('name', 'slot')}")
    merged = deepcopy(defaults)
    _deep_update(merged, config.get("season_overrides", {}).get("2026", {}))
    if starter_count(merged) != 9:
        raise ConfigurationError("The 2026 lineup must contain nine starters")
    if merged.get("roster_limits", {}).get("WR") != 5:
        raise ConfigurationError("The 2026 roster must contain five WR slots")


def _deep_update(target: dict[str, Any], changes: dict[str, Any]) -> None:
    for key, value in changes.items():
        if isinstance(value, dict) and isinstance(target.get(key), dict):
            _deep_update(target[key], value)
        else:
            target[key] = deepcopy(value)
