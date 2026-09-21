import pytest

from opfl.schedule import ScheduleError, parse_schedule_text, validate_schedule
from opfl.standings import allocate_top_six, calculate_playoffs, calculate_standings

TEAM_CODES = [f"T{index:02d}" for index in range(12)]


def one_week_schedule() -> list[dict[str, object]]:
    pairs = [(0, 2), (1, 3), (4, 5), (6, 7), (8, 9), (10, 11)]
    return [
        {
            "week": 1,
            "matchups": [
                {"away": TEAM_CODES[away], "home": TEAM_CODES[home]} for away, home in pairs
            ],
        }
    ]


def test_schedule_canonical_format() -> None:
    text = "\n".join(
        f"Week 1: {game['away']} versus {game['home']}"
        for game in one_week_schedule()[0]["matchups"]
    )
    schedule = parse_schedule_text(text, TEAM_CODES)
    assert len(schedule) == 1
    assert len(schedule[0]["matchups"]) == 6


@pytest.mark.parametrize(
    "text",
    [
        "Week 1 - T00 vs T01",
        "Week 1: T00 versus UNKNOWN",
    ],
)
def test_schedule_rejects_noncanonical_or_unknown_teams(text: str) -> None:
    with pytest.raises(ScheduleError):
        parse_schedule_text(text, TEAM_CODES)


def test_schedule_rejects_duplicate_and_missing_teams() -> None:
    schedule = one_week_schedule()
    schedule[0]["matchups"][5] = {"away": "T00", "home": "T11"}
    with pytest.raises(ScheduleError, match="duplicate"):
        validate_schedule(schedule, TEAM_CODES)


def test_schedule_rejects_incomplete_season() -> None:
    with pytest.raises(ScheduleError, match="every week"):
        validate_schedule(one_week_schedule(), TEAM_CODES, expected_weeks=15)


def test_top_six_full_value_and_cutoff_ties() -> None:
    two_way = {f"T{index}": float(12 - index) for index in range(12)}
    two_way["T5"] = two_way["T6"] = 7
    credits = allocate_top_six(two_way)
    assert credits["T0"] == 1
    assert credits["T5"] == credits["T6"] == 0.5
    assert sum(credits.values()) == 6

    three_way = {f"T{index}": float(12 - index) for index in range(12)}
    three_way["T4"] = three_way["T5"] = three_way["T6"] = 8
    credits = allocate_top_six(three_way)
    assert credits["T4"] == pytest.approx(2 / 3)
    assert sum(credits.values()) == pytest.approx(6)


def test_rank_points_expected_record_matchup_tie_and_tiebreaks() -> None:
    teams = [{"abbrev": code, "name": code, "owner": code} for code in TEAM_CODES]
    scores = {
        "T00": 1,
        "T01": 100,
        "T02": 0,
        "T03": 101,
        "T04": 50,
        "T05": 49,
        "T06": 48,
        "T07": 47,
        "T08": 46,
        "T09": 45,
        "T10": 44,
        "T11": 44,
    }
    week = {
        "week": 1,
        "teams": [{"abbrev": code, "total_score": score} for code, score in scores.items()],
    }
    standings = calculate_standings(teams, one_week_schedule(), [week], 15)
    rows = {row["abbrev"]: row for row in standings}
    assert rows["T00"]["rank_points"] == 1
    assert rows["T01"]["rank_points"] == 1
    assert rows["T00"]["wins"] == 1
    assert standings.index(rows["T00"]) < standings.index(rows["T01"])
    assert rows["T03"]["expected_wins"] == 1
    assert rows["T10"]["ties"] == rows["T11"]["ties"] == 1
    assert rows["T10"]["rank_points"] == pytest.approx(
        rows["T10"]["wins"] + rows["T10"]["top_six"] + 0.5
    )


def test_playoff_and_jamboree_output() -> None:
    standings = [{"abbrev": code} for code in TEAM_CODES]
    week16 = {
        "week": 16,
        "teams": [
            {"abbrev": code, "total_score": 100 - index} for index, code in enumerate(TEAM_CODES)
        ],
    }
    week17 = {
        "week": 17,
        "teams": [
            {"abbrev": code, "total_score": 50 + index} for index, code in enumerate(TEAM_CODES)
        ],
    }
    playoffs = calculate_playoffs(standings, [week16, week17])
    assert len(playoffs["playoff_teams"]) == 4
    assert len(playoffs["jamboree"]["teams"]) == 8
    assert playoffs["jamboree"]["weeks"] == [16, 17]
