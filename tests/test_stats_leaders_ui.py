"""Points leaders section: markup and script wiring.

No browser or JS runtime here - these read index.html/app.js/styles.css as
text, matching the sibling QPFL repo's *_ui.py pattern. CI separately runs
`node --check web/app.js` to catch syntax errors these can't.
"""

from pathlib import Path

INDEX_HTML = (Path(__file__).parent.parent / 'web' / 'index.html').read_text()
APP_JS = (Path(__file__).parent.parent / 'web' / 'app.js').read_text()
STYLES_CSS = (Path(__file__).parent.parent / 'web' / 'styles.css').read_text()


def test_nav_has_a_leaders_button():
    assert 'data-section="stats"' in INDEX_HTML


def test_stats_section_exists_with_its_containers():
    assert '<section id="stats"' in INDEX_HTML
    assert 'id="stats-position-selector"' in INDEX_HTML
    assert 'id="stats-leaders-container"' in INDEX_HTML


def test_leaders_button_sits_between_standings_and_teams():
    """Matches the nav order: Matchups, Schedule, Standings, Leaders, Teams..."""
    standings_pos = INDEX_HTML.index('data-section="standings"')
    leaders_pos = INDEX_HTML.index('data-section="stats"')
    teams_pos = INDEX_HTML.index('data-section="teams"')
    assert standings_pos < leaders_pos < teams_pos


def test_get_stats_leaders_is_defined():
    assert 'function getStatsLeaders()' in APP_JS


def test_render_stats_leaders_is_wired_into_init():
    assert 'function renderStatsLeaders()' in APP_JS
    init_fn = APP_JS.split('function initApp()')[1].split('\n        function ')[0]
    assert 'renderStatsLeaders();' in init_fn


def test_leaders_read_the_flat_opfl_week_shape():
    """OPFL's week.teams[] is flatter than QPFL's week.matchups[].team1/team2 -
    the ported aggregation loop must use the OPFL shape, not the QPFL one."""
    leaders_fn = APP_JS.split('function getStatsLeaders()')[1].split('\nfunction ')[0]
    assert 'team.roster' in leaders_fn
    assert 'week.teams' in leaders_fn
    assert 'matchups' not in leaders_fn


def test_leaders_skip_in_progress_weeks():
    leaders_fn = APP_JS.split('function getStatsLeaders()')[1].split('\nfunction ')[0]
    assert 'week.final' in leaders_fn


def test_no_offensive_line_or_dst_naming():
    """OPFL has no OL position and calls the defense slot DF, not D/ST."""
    leaders_fn = APP_JS.split('function getStatsLeaders()')[1].split('\nfunction ')[0]
    render_fn = APP_JS.split('function renderStatsLeaders()')[1].split('\nfunction ')[0]
    assert 'OL' not in leaders_fn
    assert "'D/ST'" not in render_fn


def test_stats_positions_match_opfl_roster_positions():
    from opfl.config import get_positions

    assert "const STATS_POSITIONS = ['QB', 'RB', 'WR', 'TE', 'K', 'DF', 'HC'];" in APP_JS
    assert get_positions() == ['QB', 'RB', 'WR', 'TE', 'K', 'DF', 'HC']


def test_leaders_css_targets_the_real_container_ids():
    assert '#stats-leaders-container' in STYLES_CSS
    assert '.stats-pos-btn' in STYLES_CSS
    assert '.stats-leader-row' in STYLES_CSS
