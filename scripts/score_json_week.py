#!/usr/bin/env python3
"""Score one OPFL week from authoritative JSON rosters and lineups."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from opfl.importing import attach_regular_matchups
from opfl.models import FantasyTeam
from opfl.scorer import OPFLScorer
from opfl.standings import calculate_playoffs, calculate_standings


def read_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def score_json_week(
    season: int,
    week: int,
    data_root: Path,
    scorer: OPFLScorer | None = None,
) -> dict[str, Any]:
    season_dir = data_root / "seasons" / str(season)
    lineups = read_json(season_dir / "lineups" / f"week_{week}.json")
    rosters = lineups.get("roster_snapshot") or read_json(season_dir / "rosters.json")
    rosters = rosters["teams"]
    schedule = read_json(season_dir / "schedule.json")["weeks"]
    team_registry = read_json(data_root / "teams.json")["teams"]
    scoring_engine = scorer or OPFLScorer(season, week)
    player_overrides = lineups.get("player_score_overrides", {})
    team_corrections = lineups.get("manual_corrections", {})
    teams_output = []
    for registry_team in team_registry:
        abbreviation = registry_team["abbrev"]
        roster = rosters[abbreviation]["players"]
        starter_ids = set(lineups["lineups"][abbreviation])
        roster_ids = {player["player_id"] for player in roster}
        if len(starter_ids) != 9 or not starter_ids.issubset(roster_ids):
            raise ValueError(f"{abbreviation} must have nine valid starters before scoring")
        players_by_position: dict[str, list[tuple[str, str, bool]]] = {}
        for player in roster:
            players_by_position.setdefault(player["position"], []).append(
                (player["name"], player.get("nfl_team", ""), player["player_id"] in starter_ids)
            )
        fantasy_team = FantasyTeam(
            name=registry_team["name"],
            owner=registry_team["owner"],
            abbreviation=abbreviation,
            column_index=0,
            players=players_by_position,
        )
        score_groups = scoring_engine.score_fantasy_team(fantasy_team, starters_only=False)
        score_queues: dict[tuple[str, str, str], list[Any]] = {}
        for position, scores in score_groups.items():
            for score in scores:
                score_queues.setdefault((position, score.name, score.team), []).append(score)
        roster_output = []
        starter_total = 0.0
        for player in roster:
            score = score_queues[
                (player["position"], player["name"], player.get("nfl_team", ""))
            ].pop(0)
            points = float(score.total_points)
            breakdown = dict(score.breakdown)
            if player["player_id"] in player_overrides:
                points = max(0.0, float(player_overrides[player["player_id"]]))
                breakdown = {"manual_score_override": points}
            is_starter = player["player_id"] in starter_ids
            if is_starter:
                starter_total += points
            roster_output.append(
                {
                    **player,
                    "starter": is_starter,
                    "score": round(points, 2),
                    "breakdown": breakdown,
                    "found_in_stats": score.found_in_stats,
                    "matched_name": score.matched_name,
                    "data_notes": score.data_notes,
                }
            )
        correction = float(team_corrections.get(abbreviation, 0))
        teams_output.append(
            {
                "abbrev": abbreviation,
                "name": registry_team["name"],
                "owner": registry_team["owner"],
                "roster": roster_output,
                "starter_count": len(starter_ids),
                "starter_total": round(starter_total, 2),
                "manual_correction": correction,
                "total_score": round(starter_total + correction, 2),
            }
        )
    ranked = sorted(teams_output, key=lambda item: (-item["total_score"], item["abbrev"]))
    ranks = {item["abbrev"]: rank for rank, item in enumerate(ranked, 1)}
    for team in teams_output:
        team["score_rank"] = ranks[team["abbrev"]]
    result = {
        "season": season,
        "week": week,
        "phase": "regular" if week <= 15 else "playoffs",
        "status": "scored",
        "score_source": "json_scorer",
        "teams": teams_output,
        "matchups": [],
    }
    if week <= 15:
        attach_regular_matchups({week: result}, schedule)
    else:
        existing_weeks = []
        for path in sorted((season_dir / "weeks").glob("week_*.json")):
            existing_week = read_json(path)
            if int(existing_week["week"]) != week:
                existing_weeks.append(existing_week)
        all_weeks = [*existing_weeks, result]
        standings = calculate_standings(team_registry, schedule, all_weeks, 15)
        playoffs = calculate_playoffs(standings, all_weeks)
        if week == 16:
            result["matchups"] = playoffs["semifinals"]
        elif week == 17:
            result["matchups"] = [
                game for game in (playoffs["championship"], playoffs["third_place"]) if game
            ]
    return result


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--season", required=True, type=int)
    result.add_argument("--week", required=True, type=int)
    result.add_argument("--data-root", default="data", type=Path)
    result.add_argument("--apply", action="store_true")
    result.add_argument("--replace-existing", action="store_true")
    return result


def main() -> int:
    args = parser().parse_args()
    payload = score_json_week(args.season, args.week, args.data_root)
    print(
        f"Scored season {args.season}, week {args.week}: "
        f"{len(payload['teams'])} teams and {len(payload['matchups'])} matchups."
    )
    if not args.apply:
        print("Dry run: no files written. Pass --apply to save JSON scores.")
        return 0
    season_dir = args.data_root / "seasons" / str(args.season)
    path = season_dir / "weeks" / f"week_{args.week}.json"
    if path.exists() and not args.replace_existing:
        raise FileExistsError(f"Refusing to overwrite {path}; pass --replace-existing")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    metadata_path = season_dir / "metadata.json"
    metadata = read_json(metadata_path)
    metadata["weeks_available"] = sorted({*metadata.get("weeks_available", []), args.week})
    metadata["status"] = "in_progress"
    metadata_path.write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")
    week_documents = []
    for week_path in sorted((season_dir / "weeks").glob("week_*.json")):
        week_documents.append(payload if week_path == path else read_json(week_path))
    team_registry = read_json(args.data_root / "teams.json")["teams"]
    schedule = read_json(season_dir / "schedule.json")["weeks"]
    standings = calculate_standings(team_registry, schedule, week_documents, 15)
    standings_path = season_dir / "standings.json"
    standings_path.write_text(
        json.dumps({"season": args.season, "standings": standings}, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Wrote {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
