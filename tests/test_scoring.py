from opfl.models import FantasyTeam, PlayerScore
from opfl.scorer import OPFLScorer
from opfl.scoring import (
    score_defense,
    score_head_coach,
    score_kicker,
    score_qb,
    score_rb_wr,
    score_te,
)


def test_qb_scoring_and_turnover_return_penalties() -> None:
    points, breakdown = score_qb(
        {"passing_yards": 200, "passing_tds": 1, "passing_interceptions": 2},
        {"pick_sixes": 1},
    )
    assert points == 4
    assert breakdown["passing_yards"] == 2
    assert breakdown["interceptions"] == -1
    assert breakdown["pick_sixes"] == -3


def test_qb_score_floor() -> None:
    points, breakdown = score_qb({"passing_interceptions": 1}, {"pick_sixes": 1})
    assert points == 0
    assert breakdown["floor_applied"] is True


def test_rb_wr_uses_best_yardage_method() -> None:
    combined_points, combined = score_rb_wr(
        {"rushing_yards": 60, "receiving_yards": 40, "rushing_tds": 1}
    )
    individual_points, individual = score_rb_wr({"rushing_yards": 75, "receiving_yards": 75})
    assert combined_points == 8
    assert combined["combined_rush_rec_yards"] == 2
    assert individual_points == 4
    assert individual["rushing_yards"] == 2
    assert individual["receiving_yards"] == 2


def test_tight_end_combined_bonus_and_floor() -> None:
    points, breakdown = score_te({"rushing_yards": 40, "receiving_yards": 35})
    assert points == 2
    assert breakdown["combined_rush_rec_yards"] == 2
    floor, floor_breakdown = score_te({"passing_interceptions": 1}, {"pick_sixes": 1})
    assert floor == 0
    assert floor_breakdown["floor_applied"] is True


def test_kicker_distance_misses_and_floor() -> None:
    points, breakdown = score_kicker(
        {
            "pat_made": 2,
            "pat_missed": 1,
            "fg_made_20_29": 1,
            "fg_made_30_39": 1,
            "fg_made_40_49": 1,
            "fg_made_50_59": 1,
            "fg_missed": 1,
        }
    )
    assert points == 9
    assert breakdown["fg_50+"] == 4
    floor, floor_breakdown = score_kicker({"fg_missed": 1})
    assert floor == 0
    assert floor_breakdown["floor_applied"] is True


def test_defense_all_categories_and_floor() -> None:
    points, breakdown = score_defense(
        {
            "def_interceptions": 1,
            "fumble_recovery_opp": 1,
            "def_safeties": 1,
            "def_tds": 1,
        },
        {"fg_blocked": 1, "pat_blocked": 1, "_blocked_punts": 1},
        {"points_allowed": 0},
        pbp_sacks=3,
    )
    assert points == 26
    assert breakdown == {
        "points_allowed": 8,
        "interceptions": 2,
        "fumble_recoveries": 2,
        "sacks": 3,
        "safeties": 2,
        "blocked_kicks": 4,
        "blocked_pats": 1,
        "defensive_tds": 4,
    }
    floor, floor_breakdown = score_defense({}, {}, {"points_allowed": 40})
    assert floor == 0
    assert floor_breakdown["floor_applied"] is True


def test_head_coach_spread_matrix() -> None:
    win = {"team_score": 20, "opponent_score": 10}
    assert score_head_coach({**win, "is_home": True}, {"spread": 3})[0] == 4
    assert score_head_coach({**win, "is_home": True}, {"spread": -3})[0] == 5
    assert score_head_coach({**win, "is_home": False}, {"spread": 3})[0] == 6
    assert score_head_coach({**win, "is_home": False}, {"spread": -3})[0] == 7
    assert (
        score_head_coach({"team_score": 10, "opponent_score": 20, "is_home": True}, {"spread": 3})[
            0
        ]
        == 0
    )


def test_only_starters_count_toward_team_total() -> None:
    scores = {
        "QB": [
            PlayerScore("Starter", "QB", "BUF", total_points=10, is_starter=True),
            PlayerScore("Bench", "QB", "KC", total_points=40, is_starter=False),
        ]
    }
    assert OPFLScorer.calculate_team_total(scores) == 10
    assert OPFLScorer.calculate_team_total(scores, starters_only=False) == 50


class FakeNFLData:
    def find_player(self, name, team, position):
        if name == "Missing":
            return None
        stats = {"player_display_name": name, "player_id": name}
        if position == "QB":
            stats["passing_tds"] = 1
        elif position in {"RB", "WR"}:
            stats["rushing_tds"] = 1
        elif position == "TE":
            stats["receiving_tds"] = 1
        elif position == "K":
            stats["pat_made"] = 1
        return stats

    def get_turnovers_returned_for_td(self, player_id):
        return {}

    def get_team_stats(self, team):
        return {"def_interceptions": 1}

    def get_opponent_stats(self, team):
        return {}

    def get_game_info(self, team):
        return {
            "team_score": 20,
            "opponent_score": 10,
            "points_allowed": 10,
            "is_home": True,
        }

    def get_defensive_sacks(self, team):
        return {"aggregated": 1, "pbp": 1, "value": 1, "discrepancy": False}

    def get_blocked_punts(self, team):
        return 0

    def get_blocked_kick_tds(self, team):
        return 0

    def get_spread_info(self, team):
        return {"spread": 3}

    def find_coach(self, name, team):
        return None


def test_scoring_engine_adapts_every_opfl_position() -> None:
    scorer = OPFLScorer.__new__(OPFLScorer)
    scorer.season = 2025
    scorer.week = 1
    scorer.data = FakeNFLData()
    expected = {"QB": 6, "RB": 6, "WR": 6, "TE": 6, "K": 1, "DF": 7, "HC": 4}
    for position, points in expected.items():
        result = scorer.score_player(position, "BUF", position)
        assert result.found_in_stats is True
        assert result.total_points == points
    assert scorer.score_player("Missing", "BUF", "QB").found_in_stats is False

    team = FantasyTeam(
        "Fixture",
        "Fixture",
        "FIX",
        0,
        players={"QB": [("QB", "BUF", True), ("Missing", "BUF", False)]},
    )
    scores = scorer.score_fantasy_team(team)
    assert len(scores["QB"]) == 2
    assert scorer.score_fantasy_team(team, starters_only=True)["QB"][0].is_starter
