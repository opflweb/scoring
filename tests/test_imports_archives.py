import json
from pathlib import Path

import pytest

from opfl.archives import extract_drafts_trades_picks, extract_history
from opfl.importing import import_workbook, load_teams
from opfl.league import season_config, starter_count
from scripts.import_archives import build_archives
from scripts.import_season import apply_season

ROOT = Path(__file__).resolve().parent.parent


def test_season_configuration_is_opfl_specific() -> None:
    config = season_config(2026, ROOT)
    assert config["team_count"] == 12
    assert config["roster_limits"]["WR"] == 5
    assert config["taxi_limit"] == 3
    assert config["regular_season_weeks"] == 15
    assert config["playoff_seeds"] == 4
    assert config["jamboree"] == {"teams": 8, "weeks": 2}
    assert starter_count(config) == 9
    flex = next(slot for slot in config["lineup_slots"] if slot["name"] == "FLEX")
    assert flex == {"name": "FLEX", "count": 0, "eligible_positions": ["RB", "WR", "TE"]}


def test_final_2025_workbook_import_counts_and_reconciliation() -> None:
    teams = load_teams(ROOT / "data" / "teams.json")
    rosters, weeks, lineups = import_workbook(ROOT / "OPFL Scoring 2025 (5).xlsx", teams, 2025)
    assert len(rosters["teams"]) == 12
    assert sorted(weeks) == list(range(1, 18))
    assert len(lineups) == 17
    for week in weeks.values():
        assert len(week["teams"]) == 12
        for team in week["teams"]:
            assert team["starter_count"] == 9
            assert team["total_score"] == pytest.approx(
                team["starter_total"] + team["manual_correction"]
            )
    assert any(player.get("placeholder") for player in weeks[8]["teams"][3]["roster"])


def test_archive_source_counts() -> None:
    seasons, history = extract_history(ROOT / "web" / "docs" / "Playoffs & W-L Records - 2024.xlsx")
    drafts, trades, future_picks = extract_drafts_trades_picks(
        ROOT / "OPFL Draft & Future Traded Picks.xlsx"
    )
    assert len(seasons) == 37
    assert seasons[0]["year"] == 1988
    assert seasons[-1]["year"] == 2024
    assert len(history["seasons"]) == 37
    assert len(drafts) == 12
    assert len(trades) == 33
    assert future_picks["notes"]


def test_full_archive_build_includes_existing_banners() -> None:
    payload = build_archives(
        ROOT / "web" / "docs" / "Playoffs & W-L Records - 2024.xlsx",
        ROOT / "OPFL Draft & Future Traded Picks.xlsx",
        ROOT / "web" / "images" / "banners",
    )
    assert len(payload["banners"]["banners"]) == 37
    assert payload["rules"]["title"] == "OPFL Scoring Rules"


def test_apply_refuses_accidental_overwrite(tmp_path: Path) -> None:
    season_dir = tmp_path / "seasons" / "2026"
    season_dir.mkdir(parents=True)
    (season_dir / "metadata.json").write_text("{}", encoding="utf-8")
    payload = {"metadata": {"year": 2026}, "weeks": {}, "lineups": {}}
    with pytest.raises(FileExistsError, match="replace-existing"):
        apply_season(payload, tmp_path)


def test_generated_source_matches_expected_2025_results() -> None:
    standings = json.loads((ROOT / "data" / "seasons" / "2025" / "standings.json").read_text())[
        "standings"
    ]
    playoffs = json.loads((ROOT / "data" / "seasons" / "2025" / "playoffs.json").read_text())
    assert [(row["abbrev"], row["wins"], row["losses"], row["ties"]) for row in standings[:4]] == [
        ("KAM", 10, 4, 1),
        ("G/G", 10, 5, 0),
        ("STL", 9, 6, 0),
        ("K/D", 8, 7, 0),
    ]
    assert playoffs["championship"]["winner"] == "K/D"
    assert playoffs["jamboree"]["standings"][0] == {
        "abbrev": "W/B",
        "week_16": 75.0,
        "week_17": 43.0,
        "total": 118.0,
    }
