"""Weekly matchup point projections.

A lightweight cousin of QPFL's projection model, adapted to what OPFL actually
has on hand: two archived seasons of scored weeks (`data/weeks/{season}/`)
rather than QPFL's multi-season history, and no live injury/depth-chart feed.
So this skips availability, market lines, and the home/away split entirely and
keeps only the two adjustments that matter with a thin sample: shrinking a
player's own history toward his position average, and scaling for the
strength of his week's opponent.

For each starter still to play, the projection is a shrunk blend of his own
scoring history and his position's average, adjusted for how many points his
opponent's defense has historically allowed at that position. A starter whose
game has gone final contributes his real score instead - the team total
converges to the final result as the week resolves, exactly like the roster
it is scored from.
"""

from __future__ import annotations

import math
from collections import defaultdict
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from pathlib import Path
from statistics import pstdev
from typing import Any

from .constants import TEAM_ABBREV_NORMALIZE
from .week_archive import DATA_DIR, load_all_weeks

# How many seasons of archived weeks feed a projection, counting the current one.
HISTORY_SEASONS = 2
# A prior-season observation counts for less than a current-season one - the
# roster and the league's scoring environment have both moved on since.
PRIOR_SEASON_WEIGHT = 0.5
# How strongly a thin-history player is pulled toward his position's average,
# expressed as a number of average games' worth of weight on that average.
PLAYER_SHRINKAGE_GAMES = 6
# Samples (weighted) needed before an opponent's defensive strength is trusted
# at full weight; fewer than that fades the adjustment toward neutral.
OPPONENT_FULL_WEIGHT_SAMPLES = 20
MIN_OPPONENT_MULTIPLIER = 0.85
MAX_OPPONENT_MULTIPLIER = 1.15


@dataclass(frozen=True)
class GameContext:
    opponent: str | None
    final: bool


@dataclass
class TeamProjection:
    projected_total: float
    variance: float
    starters_remaining: int
    win_probability: float | None = None


@dataclass(frozen=True)
class _Observation:
    season: int
    player: tuple[str, str]
    position: str
    opponent: str | None
    score: float


def normalize_team(team: str | None) -> str:
    value = str(team or '').strip().upper()
    return TEAM_ABBREV_NORMALIZE.get(value, value)


def normalize_player_name(name: str) -> str:
    return ' '.join(str(name).strip().casefold().split())


def player_key(name: str, position: str) -> tuple[str, str]:
    return normalize_player_name(name), position


def build_schedule_lookup(
    schedule_rows: Iterable[Mapping[str, Any]],
) -> dict[tuple[int, int, str], GameContext]:
    lookup: dict[tuple[int, int, str], GameContext] = {}
    for row in schedule_rows:
        if row.get('game_type') not in (None, 'REG'):
            continue
        season = row.get('season')
        week = row.get('week')
        if not isinstance(season, int) or not isinstance(week, int):
            continue
        home = normalize_team(row.get('home_team'))
        away = normalize_team(row.get('away_team'))
        if not home or not away:
            continue
        final = row.get('result') not in (None, '')
        lookup[(season, week, home)] = GameContext(away, final)
        lookup[(season, week, away)] = GameContext(home, final)
    return lookup


def _load_history(
    season: int,
    target_week: int,
    schedule_lookup: Mapping[tuple[int, int, str], GameContext],
    data_dir: Path | str = DATA_DIR,
) -> list[_Observation]:
    """Every archived starter score from the last `HISTORY_SEASONS` seasons.

    Only final weeks count, and only starters - bench appearances are not a
    consistent signal of what a player is expected to score when started.
    """
    observations: list[_Observation] = []
    for hist_season in range(season - (HISTORY_SEASONS - 1), season + 1):
        weeks, _schedule = load_all_weeks(hist_season, data_dir)
        for week_entry in weeks:
            if not week_entry.get('final'):
                continue
            if hist_season == season and week_entry['week'] >= target_week:
                continue
            for team in week_entry.get('teams', []) or []:
                for player in team.get('roster', []) or []:
                    if not player.get('starter'):
                        continue
                    score = player.get('score')
                    position = player.get('position')
                    name = player.get('name')
                    nfl_team = normalize_team(player.get('nfl_team'))
                    if not isinstance(score, (int, float)) or not position or not nfl_team or not name:
                        continue
                    game = schedule_lookup.get((hist_season, week_entry['week'], nfl_team))
                    observations.append(
                        _Observation(
                            season=hist_season,
                            player=player_key(name, position),
                            position=position,
                            opponent=game.opponent if game else None,
                            score=float(score),
                        )
                    )
    return observations


def _season_weight(obs_season: int, target_season: int) -> float:
    return 1.0 if obs_season == target_season else PRIOR_SEASON_WEIGHT


def _weighted_mean(values_weights: list[tuple[float, float]], default: float = 0.0) -> float:
    total_weight = sum(weight for _, weight in values_weights)
    if total_weight <= 0:
        return default
    return sum(value * weight for value, weight in values_weights) / total_weight


def calculate_week_projections(
    week_data: dict[str, Any],
    pairings: Iterable[Iterable[str]],
    season: int,
    week: int,
    schedule_rows: Iterable[Mapping[str, Any]],
    data_dir: Path | str = DATA_DIR,
) -> dict[str, dict[str, Any]]:
    """Project every starter still to play, and roll those up into matchups.

    Mutates `week_data`'s rosters in place, stamping a `projected_points` on
    each starter. Returns `{abbrev: {'projected_total', 'win_probability'}}`
    for every team with a scored roster, which `pairings` turns into
    head-to-head win probabilities.
    """
    schedule_lookup = build_schedule_lookup(schedule_rows)
    observations = _load_history(season, week, schedule_lookup, data_dir)

    player_obs: dict[tuple[str, str], list[tuple[float, float]]] = defaultdict(list)
    position_obs: dict[str, list[tuple[float, float]]] = defaultdict(list)
    defense_obs: dict[tuple[str, str], list[tuple[float, float]]] = defaultdict(list)

    for obs in observations:
        weight = _season_weight(obs.season, season)
        player_obs[obs.player].append((obs.score, weight))
        position_obs[obs.position].append((obs.score, weight))
        if obs.opponent:
            defense_obs[(obs.opponent, obs.position)].append((obs.score, weight))

    position_mean = {pos: _weighted_mean(vals) for pos, vals in position_obs.items()}
    position_stdev = {
        pos: pstdev([value for value, _ in vals]) if len(vals) >= 2 else 0.0
        for pos, vals in position_obs.items()
    }

    def opponent_multiplier(opponent: str, position: str) -> float:
        league_mean = position_mean.get(position)
        if not league_mean:
            return 1.0
        vals = defense_obs.get((opponent, position), [])
        allowed = _weighted_mean(vals, default=league_mean)
        bounded = min(MAX_OPPONENT_MULTIPLIER, max(MIN_OPPONENT_MULTIPLIER, allowed / league_mean))
        reliability = min(1.0, sum(weight for _, weight in vals) / OPPONENT_FULL_WEIGHT_SAMPLES)
        return 1.0 + (bounded - 1.0) * reliability

    def player_baseline(key: tuple[str, str], position: str) -> tuple[float, float]:
        pos_mean = position_mean.get(position, 0.0)
        vals = player_obs.get(key, [])
        total_weight = sum(weight for _, weight in vals)
        weighted_sum = sum(value * weight for value, weight in vals)
        baseline = (weighted_sum + PLAYER_SHRINKAGE_GAMES * pos_mean) / (
            total_weight + PLAYER_SHRINKAGE_GAMES
        )
        raw_scores = [value for value, _ in vals]
        stdev = pstdev(raw_scores) if len(raw_scores) >= 2 else position_stdev.get(position, 0.0)
        return baseline, stdev

    team_projections: dict[str, TeamProjection] = {}

    for team in week_data.get('teams', []) or []:
        abbrev = team.get('abbrev')
        if not abbrev:
            continue
        effective_total = 0.0
        variance = 0.0
        starters_remaining = 0

        for player in team.get('roster', []) or []:
            if not player.get('starter'):
                continue
            position = player.get('position')
            nfl_team = normalize_team(player.get('nfl_team'))
            game = schedule_lookup.get((season, week, nfl_team))
            on_bye = game is None
            baseline, stdev = player_baseline(player_key(player.get('name', ''), position), position)

            if on_bye:
                projected = 0.0
            else:
                multiplier = opponent_multiplier(game.opponent, position) if game.opponent else 1.0
                # Additive form of `baseline * multiplier` that leaves the
                # sign alone, so a below-average (even negative, for coaches)
                # baseline scales the same direction the multiplier intends.
                projected = baseline + abs(baseline) * (multiplier - 1)
                stdev *= multiplier

            player['projected_points'] = round(projected, 1)

            if game and game.final:
                effective_total += player.get('score', 0.0) or 0.0
            elif on_bye:
                continue
            else:
                effective_total += projected
                variance += stdev**2
                starters_remaining += 1

        team_projections[abbrev] = TeamProjection(
            projected_total=round(effective_total, 1),
            variance=variance,
            starters_remaining=starters_remaining,
        )

    for pairing in pairings:
        team1, team2 = list(pairing)[:2]
        projection1 = team_projections.get(team1)
        projection2 = team_projections.get(team2)
        if not projection1 or not projection2:
            continue
        diff = projection1.projected_total - projection2.projected_total
        if projection1.starters_remaining == 0 and projection2.starters_remaining == 0:
            probability1 = 1.0 if diff > 0 else 0.0 if diff < 0 else 0.5
        else:
            total_variance = projection1.variance + projection2.variance
            if total_variance <= 0:
                probability1 = 0.99 if diff > 0 else 0.01 if diff < 0 else 0.5
            else:
                z_score = diff / math.sqrt(total_variance)
                probability1 = min(0.99, max(0.01, 0.5 * (1 + math.erf(z_score / math.sqrt(2)))))
        projection1.win_probability = probability1
        projection2.win_probability = 1 - probability1

    return {
        abbrev: {
            'projected_total': projection.projected_total,
            'win_probability': projection.win_probability,
        }
        for abbrev, projection in team_projections.items()
    }
