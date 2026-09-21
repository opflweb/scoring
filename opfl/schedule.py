"""Parsing and validation for canonical OPFL schedules."""

from __future__ import annotations

import re
from collections.abc import Iterable
from pathlib import Path
from typing import Any

SCHEDULE_LINE = re.compile(
    r"^Week\s+(?P<week>\d+):\s+(?P<away>\S+)\s+versus\s+(?P<home>\S+)\s*$",
    re.IGNORECASE,
)


class ScheduleError(ValueError):
    """Raised when a schedule does not satisfy the OPFL contract."""


def parse_schedule_text(text: str, team_abbrevs: Iterable[str]) -> list[dict[str, Any]]:
    allowed = set(team_abbrevs)
    by_week: dict[int, list[dict[str, str]]] = {}
    for line_number, raw_line in enumerate(text.splitlines(), 1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        match = SCHEDULE_LINE.fullmatch(line)
        if not match:
            raise ScheduleError(
                f"Line {line_number} must use 'Week N: ABBR versus ABBR': {raw_line!r}"
            )
        week = int(match.group("week"))
        away, home = match.group("away"), match.group("home")
        unknown = {away, home} - allowed
        if unknown:
            raise ScheduleError(
                f"Week {week} contains unknown team(s): {', '.join(sorted(unknown))}"
            )
        if away == home:
            raise ScheduleError(f"Week {week} cannot match {away} against itself")
        by_week.setdefault(week, []).append({"away": away, "home": home})
    schedule = [{"week": week, "matchups": by_week[week]} for week in sorted(by_week)]
    validate_schedule(schedule, allowed)
    return schedule


def parse_schedule_file(path: str | Path, team_abbrevs: Iterable[str]) -> list[dict[str, Any]]:
    return parse_schedule_text(Path(path).read_text(encoding="utf-8"), team_abbrevs)


def validate_schedule(
    schedule: list[dict[str, Any]], team_abbrevs: Iterable[str], expected_weeks: int | None = None
) -> None:
    teams = set(team_abbrevs)
    weeks = [int(week["week"]) for week in schedule]
    if len(weeks) != len(set(weeks)):
        raise ScheduleError("Schedule contains duplicate weeks")
    if expected_weeks is not None and weeks != list(range(1, expected_weeks + 1)):
        raise ScheduleError(f"Schedule must contain every week from 1 through {expected_weeks}")
    for week_data in schedule:
        week = int(week_data["week"])
        matchups = week_data.get("matchups", [])
        if len(matchups) != 6:
            raise ScheduleError(f"Week {week} must contain exactly six matchups")
        appearances = [team for matchup in matchups for team in (matchup["away"], matchup["home"])]
        if len(appearances) != len(set(appearances)):
            raise ScheduleError(f"Week {week} contains a duplicate team")
        missing = teams - set(appearances)
        extra = set(appearances) - teams
        if missing or extra:
            parts = []
            if missing:
                parts.append(f"missing {', '.join(sorted(missing))}")
            if extra:
                parts.append(f"unknown {', '.join(sorted(extra))}")
            raise ScheduleError(f"Week {week} is invalid: {'; '.join(parts)}")


def schedule_lookup(schedule: list[dict[str, Any]]) -> dict[int, list[dict[str, str]]]:
    return {int(week["week"]): list(week["matchups"]) for week in schedule}
