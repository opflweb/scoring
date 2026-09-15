"""League configuration loading and shape."""

import json

import pytest

from opfl.config import (
    CONFIG_PATH,
    clear_config_cache,
    get_config,
    get_current_season,
    get_playoff_structure,
    get_playoff_weeks,
    get_positions,
    get_regular_season_weeks,
    get_roster_slots,
    get_starter_slots,
    get_trade_deadline_week,
)


@pytest.fixture(autouse=True)
def _clear_cache():
    """get_config() is lru_cached; tests that mutate the file need a clean slate."""
    clear_config_cache()
    yield
    clear_config_cache()


def test_loads_the_real_config_file():
    config = get_config()
    assert config.current_season == 2026
    assert config.regular_season_weeks == 15
    assert config.playoff_weeks == [16, 17]


def test_roster_and_starter_slots_match_the_real_workbook():
    """QB3/RB4/WR5/TE3/K2/DF2/HC2 = 21 roster, 1/2/2/1/1/1/1 = 9 starters."""
    roster = get_roster_slots()
    starters = get_starter_slots()
    assert roster == {'QB': 3, 'RB': 4, 'WR': 5, 'TE': 3, 'K': 2, 'DF': 2, 'HC': 2}
    assert starters == {'QB': 1, 'RB': 2, 'WR': 2, 'TE': 1, 'K': 1, 'DF': 1, 'HC': 1}
    assert sum(roster.values()) == 21
    assert sum(starters.values()) == 9


def test_no_offensive_line_position():
    """OPFL has no OL slot, unlike the sibling QPFL league this config format
    was ported from."""
    assert 'OL' not in get_roster_slots()
    assert 'OL' not in get_starter_slots()


def test_positions_follow_roster_slots_order():
    assert get_positions() == list(get_roster_slots())


def test_accessors_agree_with_get_config():
    config = get_config()
    assert get_current_season() == config.current_season
    assert get_trade_deadline_week() == config.trade_deadline_week
    assert get_regular_season_weeks() == config.regular_season_weeks
    assert get_playoff_weeks() == config.playoff_weeks
    assert get_playoff_structure() == config.playoff_structure


def test_config_is_cached_across_calls():
    """Two calls without clearing the cache return the same object."""
    assert get_config() is get_config()


def test_cache_clear_picks_up_a_changed_file(tmp_path, monkeypatch):
    custom = tmp_path / 'league_config.json'
    custom.write_text(json.dumps(json.loads(CONFIG_PATH.read_text()) | {'trade_deadline_week': 9}))
    monkeypatch.setattr('opfl.config.CONFIG_PATH', custom)

    clear_config_cache()
    assert get_trade_deadline_week() == 9


def test_playoff_structure_covers_all_twelve_teams_with_no_overlap():
    structure = get_playoff_structure()
    championship = set(structure['championship_seeds'])
    jamboree = set(structure['jamboree_seeds'])
    assert championship & jamboree == set()
    assert championship | jamboree == set(range(1, 13))
