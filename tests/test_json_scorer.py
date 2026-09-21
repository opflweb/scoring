import json
import shutil
from pathlib import Path

from opfl.models import PlayerScore
from scripts.score_json_week import score_json_week

ROOT = Path(__file__).resolve().parent.parent


class FakeScorer:
    def score_fantasy_team(self, team, starters_only=False):
        assert starters_only is False
        result = {}
        for position, players in team.players.items():
            result[position] = []
            for name, nfl_team, started in players:
                result[position].append(
                    PlayerScore(
                        name=name,
                        position=position,
                        team=nfl_team,
                        total_points=1,
                        breakdown={"fixture": 1},
                        found_in_stats=True,
                        matched_name=name,
                        is_starter=started,
                    )
                )
        return result


def test_json_scorer_scores_bench_but_counts_starters_and_corrections(tmp_path: Path) -> None:
    data_root = tmp_path / "data"
    season_root = data_root / "seasons" / "2025"
    (season_root / "lineups").mkdir(parents=True)
    shutil.copy(ROOT / "data" / "teams.json", data_root / "teams.json")
    shutil.copy(ROOT / "data" / "seasons" / "2025" / "rosters.json", season_root / "rosters.json")
    shutil.copy(ROOT / "data" / "seasons" / "2025" / "schedule.json", season_root / "schedule.json")
    lineup = json.loads(
        (ROOT / "data" / "seasons" / "2025" / "lineups" / "week_1.json").read_text()
    )
    rosters = json.loads((season_root / "rosters.json").read_text())["teams"]
    required = {"QB": 1, "RB": 2, "WR": 2, "TE": 1, "K": 1, "DF": 1, "HC": 1}
    for abbreviation, roster in rosters.items():
        starters = []
        for position, count in required.items():
            starters.extend(
                [
                    player["player_id"]
                    for player in roster["players"]
                    if player["position"] == position
                ][:count]
            )
        lineup["lineups"][abbreviation] = starters
    first_starter = lineup["lineups"]["K/D"][0]
    lineup["player_score_overrides"] = {first_starter: 5}
    lineup["manual_corrections"] = {"K/D": 2}
    (season_root / "lineups" / "week_1.json").write_text(json.dumps(lineup), encoding="utf-8")
    payload = score_json_week(2025, 1, data_root, scorer=FakeScorer())
    kirk = next(team for team in payload["teams"] if team["abbrev"] == "K/D")
    assert len(kirk["roster"]) == 20
    assert all(player["score"] >= 1 for player in kirk["roster"])
    assert kirk["starter_total"] == 13
    assert kirk["manual_correction"] == 2
    assert kirk["total_score"] == 15
    assert len(payload["matchups"]) == 6
