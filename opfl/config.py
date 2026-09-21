"""League configuration.

A plain dataclass loaded from data/league_config.json and cached after first
load, so a season rollover only requires editing one file instead of hunting
down the same constants scattered across the exporter, the CI workflow, and
opfl/constants.py.

QPFL's equivalent (qpfl/config.py) validates this file with a pydantic model;
that is disproportionate for one small file here, so this just reads the JSON
and fails loudly if a key is missing.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

CONFIG_PATH = Path(__file__).parent.parent / 'data' / 'league_config.json'


@dataclass(frozen=True)
class LeagueConfig:
    current_season: int
    trade_deadline_week: int
    regular_season_weeks: int
    playoff_weeks: list[int]
    roster_slots: dict[str, int]
    starter_slots: dict[str, int]
    playoff_structure: dict[str, list[int]]
    rosters_sheet: str = 'Rosters'
    matchups_sheet: str = 'Matchups'
    positions: list[str] = field(init=False)

    def __post_init__(self):
        # Position order follows the roster_slots key order rather than being
        # declared a third time.
        object.__setattr__(self, 'positions', list(self.roster_slots))


@lru_cache(maxsize=1)
def get_config() -> LeagueConfig:
    """Load and cache data/league_config.json."""
    raw = json.loads(CONFIG_PATH.read_text())
    return LeagueConfig(
        current_season=raw['current_season'],
        trade_deadline_week=raw['trade_deadline_week'],
        regular_season_weeks=raw['regular_season_weeks'],
        playoff_weeks=raw['playoff_weeks'],
        roster_slots=raw['roster_slots'],
        starter_slots=raw['starter_slots'],
        playoff_structure=raw['playoff_structure'],
        rosters_sheet=raw.get('rosters_sheet', 'Rosters'),
        matchups_sheet=raw.get('matchups_sheet', 'Matchups'),
    )


def clear_config_cache() -> None:
    """Drop the cached config, e.g. between tests that write different files."""
    get_config.cache_clear()


def get_current_season() -> int:
    return get_config().current_season


def get_trade_deadline_week() -> int:
    return get_config().trade_deadline_week


def get_regular_season_weeks() -> int:
    return get_config().regular_season_weeks


def get_playoff_weeks() -> list[int]:
    return get_config().playoff_weeks


def get_roster_slots() -> dict[str, int]:
    return get_config().roster_slots


def get_starter_slots() -> dict[str, int]:
    return get_config().starter_slots


def get_playoff_structure() -> dict[str, list[int]]:
    return get_config().playoff_structure


def get_positions() -> list[str]:
    return get_config().positions
