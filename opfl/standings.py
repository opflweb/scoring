"""OPFL standings, all-play, schedule strength, and playoff calculations."""

from __future__ import annotations

import statistics
from collections.abc import Iterable
from typing import Any


def allocate_top_six(scores: dict[str, float], cutoff: int = 6) -> dict[str, float]:
    """Allocate one Top-6 point per occupied cutoff place, splitting cutoff ties."""
    credits = {team: 0.0 for team in scores}
    ordered = sorted(scores.items(), key=lambda item: (-item[1], item[0]))
    place = 1
    index = 0
    while index < len(ordered):
        score = ordered[index][1]
        tied: list[str] = []
        while index < len(ordered) and ordered[index][1] == score:
            tied.append(ordered[index][0])
            index += 1
        available = max(0, min(cutoff, place + len(tied) - 1) - place + 1)
        credit = available / len(tied) if available else 0.0
        for team in tied:
            credits[team] = credit
        place += len(tied)
    return credits


def calculate_standings(
    teams: list[dict[str, Any]],
    schedule: list[dict[str, Any]],
    weeks: Iterable[dict[str, Any]],
    regular_season_weeks: int = 15,
) -> list[dict[str, Any]]:
    schedule_by_week = {int(item["week"]): item["matchups"] for item in schedule}
    week_list = sorted(
        (week for week in weeks if int(week["week"]) <= regular_season_weeks),
        key=lambda item: int(item["week"]),
    )
    table: dict[str, dict[str, Any]] = {}
    for team in teams:
        abbreviation = team["abbrev"]
        table[abbreviation] = {
            "abbrev": abbreviation,
            "name": team["name"],
            "owner": team.get("owner", team["name"]),
            "wins": 0,
            "losses": 0,
            "ties": 0,
            "top_six": 0.0,
            "rank_points": 0.0,
            "points_for": 0.0,
            "points_against": 0.0,
            "all_play_wins": 0.0,
            "games_played": 0,
            "weekly_rank_history": [],
            "weekly_scores": [],
            "opponents_played": [],
        }

    for week_data in week_list:
        week = int(week_data["week"])
        score_map = {item["abbrev"]: float(item["total_score"]) for item in week_data["teams"]}
        if set(score_map) != set(table):
            raise ValueError(f"Week {week} must contain a score for every team")
        top_six = allocate_top_six(score_map)
        weekly_order = sorted(score_map, key=lambda team: (-score_map[team], team))
        for rank, abbreviation in enumerate(weekly_order, 1):
            row = table[abbreviation]
            row["points_for"] += score_map[abbreviation]
            row["top_six"] += top_six[abbreviation]
            row["rank_points"] += top_six[abbreviation]
            row["games_played"] += 1
            row["weekly_scores"].append(score_map[abbreviation])
            row["weekly_rank_history"].append({"week": week, "rank": rank})
            all_play = 0.0
            for opponent, opponent_score in score_map.items():
                if opponent == abbreviation:
                    continue
                if score_map[abbreviation] > opponent_score:
                    all_play += 1
                elif score_map[abbreviation] == opponent_score:
                    all_play += 0.5
            row["all_play_wins"] += all_play
        for matchup in schedule_by_week.get(week, []):
            away, home = matchup["away"], matchup["home"]
            away_score, home_score = score_map[away], score_map[home]
            table[away]["points_against"] += home_score
            table[home]["points_against"] += away_score
            table[away]["opponents_played"].append(home)
            table[home]["opponents_played"].append(away)
            if away_score > home_score:
                table[away]["wins"] += 1
                table[away]["rank_points"] += 1
                table[home]["losses"] += 1
            elif home_score > away_score:
                table[home]["wins"] += 1
                table[home]["rank_points"] += 1
                table[away]["losses"] += 1
            else:
                for abbreviation in (away, home):
                    table[abbreviation]["ties"] += 1
                    table[abbreviation]["rank_points"] += 0.5

    completed_weeks = {int(week["week"]) for week in week_list}
    for abbreviation, row in table.items():
        games = row["games_played"]
        scores = row.pop("weekly_scores")
        row["points_for"] = round(row["points_for"], 2)
        row["points_against"] = round(row["points_against"], 2)
        row["average_points_for"] = round(row["points_for"] / games, 2) if games else 0.0
        row["point_differential"] = round(row["points_for"] - row["points_against"], 2)
        row["high_score"] = max(scores) if scores else 0.0
        row["low_score"] = min(scores) if scores else 0.0
        row["median_score"] = round(statistics.median(scores), 2) if scores else 0.0
        row["expected_wins"] = round(row["all_play_wins"] / 11, 2)
        row["expected_losses"] = round(games - row["expected_wins"], 2)
        actual_wins = row["wins"] + 0.5 * row["ties"]
        row["luck"] = round(actual_wins - row["expected_wins"], 2)
        remaining_opponents: list[str] = []
        for item in schedule:
            if int(item["week"]) in completed_weeks:
                continue
            for matchup in item["matchups"]:
                if matchup["away"] == abbreviation:
                    remaining_opponents.append(matchup["home"])
                elif matchup["home"] == abbreviation:
                    remaining_opponents.append(matchup["away"])
        opponent_strengths = [
            table[opponent]["rank_points"] for opponent in remaining_opponents if opponent in table
        ]
        row["remaining_sos"] = (
            round(sum(opponent_strengths) / len(opponent_strengths), 2)
            if opponent_strengths
            else None
        )
        row["top_six"] = round(row["top_six"], 3)
        row["rank_points"] = round(row["wins"] + row["top_six"] + 0.5 * row["ties"], 3)
        row.pop("opponents_played")

    ordered = sorted(
        table.values(),
        key=lambda row: (
            -row["rank_points"],
            -row["wins"],
            -row["top_six"],
            -row["points_for"],
            row["abbrev"],
        ),
    )
    for rank, row in enumerate(ordered, 1):
        row["rank"] = rank
    return ordered


def calculate_playoffs(
    standings: list[dict[str, Any]], weeks: Iterable[dict[str, Any]]
) -> dict[str, Any]:
    if len(standings) != 12:
        raise ValueError("Playoff output requires twelve final standings rows")
    seeds = {row["abbrev"]: index + 1 for index, row in enumerate(standings)}
    scores_by_week = {
        int(week["week"]): {team["abbrev"]: team["total_score"] for team in week["teams"]}
        for week in weeks
    }
    playoff_teams = [row["abbrev"] for row in standings[:4]]
    jamboree_teams = [row["abbrev"] for row in standings[4:]]

    def game(team1: str, team2: str, week: int) -> dict[str, Any]:
        score1 = scores_by_week.get(week, {}).get(team1)
        score2 = scores_by_week.get(week, {}).get(team2)
        winner = loser = None
        if score1 is not None and score2 is not None:
            if score1 > score2 or (score1 == score2 and seeds[team1] < seeds[team2]):
                winner, loser = team1, team2
            else:
                winner, loser = team2, team1
        return {
            "week": week,
            "team1": team1,
            "team2": team2,
            "score1": score1,
            "score2": score2,
            "winner": winner,
            "loser": loser,
        }

    semifinal1 = game(playoff_teams[0], playoff_teams[3], 16)
    semifinal2 = game(playoff_teams[1], playoff_teams[2], 16)
    championship = third_place = None
    if semifinal1["winner"] and semifinal2["winner"]:
        championship = game(semifinal1["winner"], semifinal2["winner"], 17)
        third_place = game(semifinal1["loser"], semifinal2["loser"], 17)
    jamboree = []
    for abbreviation in jamboree_teams:
        week16 = scores_by_week.get(16, {}).get(abbreviation)
        week17 = scores_by_week.get(17, {}).get(abbreviation)
        total = None if week16 is None or week17 is None else week16 + week17
        jamboree.append(
            {"abbrev": abbreviation, "week_16": week16, "week_17": week17, "total": total}
        )
    jamboree.sort(
        key=lambda row: (row["total"] is None, -(row["total"] or 0), seeds[row["abbrev"]])
    )
    return {
        "seeds": seeds,
        "playoff_teams": playoff_teams,
        "semifinals": [semifinal1, semifinal2],
        "championship": championship,
        "third_place": third_place,
        "jamboree": {"teams": jamboree_teams, "weeks": [16, 17], "standings": jamboree},
    }
