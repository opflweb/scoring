"""NFL week-completion helpers used by scoring automation."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import Any


def week_games_are_final(
    schedule_rows: Iterable[Mapping[str, Any]], week: int, season: int | None = None
) -> bool:
    """Whether every NFL game in `week` has a final result.

    Standings can only count a week once it is over. Mid-week, a matchup where
    one manager's Thursday starter has played and the other's have not is not a
    1-0 record, it is an unfinished game. Returns False when the week has no
    games in `schedule_rows` at all, since "no evidence" is not "finished".

    Pass `season` whenever the rows span more than one - projection history
    carries the prior season too, and its games are all final, which would
    otherwise vote for a week the current season has not played yet.
    """
    games = [
        row
        for row in schedule_rows
        if row.get('game_type') == 'REG'
        and row.get('week') == week
        and (season is None or row.get('season') == season)
    ]
    return bool(games) and all(game.get('result') not in (None, '') for game in games)


def teams_missing_stats(
    schedule_rows: Iterable[Mapping[str, Any]],
    week: int,
    season: int | None,
    stat_teams: Iterable[str],
) -> set[str]:
    """NFL teams that played in `week` but have no rows in the stats feed yet.

    The schedule's final score lands well before nflverse publishes that
    game's player and team stats - Week 2 of 2026 was archived as final at
    03:33 UTC, two hours before the NYG@LA Monday night stats existed, so
    every Rams player was frozen at 0 and never rescored. A week is only
    safe to archive as final once this is empty.
    """
    published = set(stat_teams)
    playing = {
        team
        for row in schedule_rows
        if row.get('game_type') == 'REG'
        and row.get('week') == week
        and (season is None or row.get('season') == season)
        for team in (row.get('home_team'), row.get('away_team'))
        if team
    }
    return playing - published


def latest_completed_week(schedule_rows: Iterable[Mapping[str, Any]], max_week: int = 17) -> int:
    """Return the latest fantasy week whose NFL games all have final results."""
    games_by_week: dict[int, list[Mapping[str, Any]]] = {}
    for row in schedule_rows:
        if row.get('game_type') != 'REG':
            continue
        week = row.get('week')
        if not isinstance(week, int) or not 1 <= week <= max_week:
            continue
        games_by_week.setdefault(week, []).append(row)

    completed = [
        week
        for week, games in games_by_week.items()
        if games and all(game.get('result') not in (None, '') for game in games)
    ]
    return max(completed, default=0)
