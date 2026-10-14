"""Player profile modal, position ranks, and Greg/Griffin-first ordering."""

import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).parent.parent
APP_JS = (ROOT / 'web' / 'app.js').read_text()
INDEX_HTML = (ROOT / 'web' / 'index.html').read_text()
STYLES_CSS = (ROOT / 'web' / 'styles.css').read_text()


def _function_body(name):
    return APP_JS.split(f'function {name}(')[1].split('\n        function ')[0]


def _run_node(script):
    result = subprocess.run(['node', '-e', script], check=True, capture_output=True, text=True)
    return json.loads(result.stdout)


def test_modal_markup_exists():
    assert 'id="player-modal-overlay"' in INDEX_HTML
    assert 'id="player-modal-weeks"' in INDEX_HTML
    assert 'class="player-modal-close"' in INDEX_HTML
    assert '.modal-overlay.active {' in STYLES_CSS
    assert '.player-profile-trigger {' in STYLES_CSS


def test_player_names_open_the_modal_everywhere_they_render():
    for fn in ('renderRosterList', 'renderTeamPage', 'renderStatsLeaders'):
        assert 'playerProfileButton(' in _function_body(fn), fn
    assert "e.target.closest('.player-profile-trigger')" in APP_JS
    assert 'initPlayerModal();' in _function_body('initApp')


def test_my_team_matches_teams_json():
    teams = json.loads((ROOT / 'data' / 'teams.json').read_text())
    greg_griffin = [t for t in teams['teams'] if t['owner_key'] == 'greg_griffin']
    assert greg_griffin and greg_griffin[0]['abbrev'] == 'G/G'
    assert "const MY_TEAM = 'G/G';" in APP_JS
    assert '.is-my-team {' in STYLES_CSS


def test_ranked_lists_highlight_but_do_not_reorder():
    for fn in ('renderStandings', 'renderStatsLeaders', 'renderPlayoffOdds'):
        body = _function_body(fn)
        assert 'myTeamClass(' in body, fn
        assert 'floatMyTeam(' not in body, fn
    assert 'floatMyTeam(' in _function_body('renderTeamSelector')
    assert 'myTeamFirstPairs(schedule)' in _function_body('renderMatchups')


def test_my_team_floats_first_and_takes_the_left_side():
    helpers = APP_JS[APP_JS.index('const MY_TEAM') : APP_JS.index('function escapeHtml')]
    rendered = _run_node(
        f"""
{helpers}
const pairs = myTeamFirstPairs([['K/D', 'STL'], ['JOH', 'G/G'], ['KEV', 'D/J']]);
const teams = floatMyTeam([{{a: 'K/D'}}, {{a: 'G/G'}}, {{a: 'STL'}}], t => t.a).map(t => t.a);
const untouched = floatMyTeam([{{a: 'K/D'}}, {{a: 'STL'}}], t => t.a).map(t => t.a);
process.stdout.write(JSON.stringify({{pairs, teams, untouched}}));
"""
    )
    assert rendered['pairs'] == [['G/G', 'JOH'], ['K/D', 'STL'], ['KEV', 'D/J']]
    assert rendered['teams'] == ['G/G', 'K/D', 'STL']
    assert rendered['untouched'] == ['K/D', 'STL']


def test_position_rank_counts_bench_and_skips_live_week():
    metrics_js = APP_JS[
        APP_JS.index('const _playerMetricsCache') : APP_JS.index('function positionRankTitle')
    ]
    rb = lambda name, score, starter=True: {  # noqa: E731
        'name': name,
        'position': 'RB',
        'score': score,
        'starter': starter,
    }
    data = {
        'weeks': [
            {'week': 1, 'final': True, 'teams': [
                {'abbrev': 'G/G', 'roster': [rb('Ace', 10), rb('Bench Guy', 30, False)]},
                {'abbrev': 'K/D', 'roster': [rb('Cee', 20)]},
            ]},
            {'week': 2, 'final': True, 'teams': [
                {'abbrev': 'G/G', 'roster': [rb('Ace', 10), rb('Bench Guy', 0, False)]},
                {'abbrev': 'K/D', 'roster': [rb('Cee', 0)]},
            ]},
            {'week': 3, 'final': False, 'teams': [
                {'abbrev': 'K/D', 'roster': [rb('Cee', 50)]},
            ]},
        ],
    }
    rendered = _run_node(
        f"""
let data = {json.dumps(data)};
{metrics_js}
const m = getPlayerSeasonMetrics();
const pick = n => [m['RB|' + n].position_rank, m['RB|' + n].total_points, m['RB|' + n].pool_size];
process.stdout.write(JSON.stringify({{bench: pick('Bench Guy'), ace: pick('Ace'), cee: pick('Cee')}}));
"""
    )
    assert rendered['bench'] == [1, 30, 3]
    # Ace and Cee tie on 20 (Cee's live week 3 doesn't count); name breaks it.
    assert rendered['ace'] == [2, 20, 3]
    assert rendered['cee'] == [3, 20, 3]
