"""Matchup-card insights (optimal lineup, top 6, head-to-head) and Team Stats."""

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT / 'scripts'))

from export_for_web import calculate_lineup_efficiency  # noqa: E402

APP_JS = (ROOT / 'web' / 'app.js').read_text()
INDEX_HTML = (ROOT / 'web' / 'index.html').read_text()
DATA = json.loads((ROOT / 'web' / 'data.json').read_text())

HELPERS = APP_JS[APP_JS.index('const NFL_TEAM_ALIASES') : APP_JS.index('// Historical championship data')]
MY_TEAM_HELPERS = APP_JS[APP_JS.index('const DEFAULT_TEAM_NUMBER_MAP') : APP_JS.index('const NFL_TEAM_ALIASES')]
INSIGHTS = APP_JS[
    APP_JS.index('// Optimal lineup and bench mistakes') : APP_JS.index('const BREAKDOWN_LABELS')
]
TOP_SIX = APP_JS[APP_JS.index('const TOP_SCORERS_CUTOFF') : APP_JS.index('function renderMatchups')]


def _run_node(data, body):
    script = f"""
let data = {json.dumps(data)};
{MY_TEAM_HELPERS}
{HELPERS}
{TOP_SIX}
{INSIGHTS}
{body}
"""
    result = subprocess.run(['node', '-e', script], check=True, capture_output=True, text=True)
    return json.loads(result.stdout)


def _player(name, position, score, starter):
    return {'name': name, 'position': position, 'nfl_team': 'XX', 'score': score, 'starter': starter}


def test_matchup_cards_and_leaders_tab_render_the_new_pieces():
    body = APP_JS.split('function renderMatchups(')[1].split('\n        function ')[0]
    for call in ('renderTopSixBadge(', 'renderScoreRankLine(', 'renderTeamOptimal(', 'renderOptimalSummary(', 'renderH2HBadge('):
        assert call in body, call
    assert 'id="team-stats-container"' in INDEX_HTML
    assert 'renderTeamStats();' in APP_JS.split('function initApp(')[1].split('\n        function ')[0]


def test_optimal_lineup_agrees_with_exporter_on_every_final_week():
    teams = [t for w in DATA['weeks'] if w.get('final') for t in w['teams']]
    assert teams
    rendered = _run_node(
        {'weeks': []},
        f"""
const teams = {json.dumps(teams)};
process.stdout.write(JSON.stringify(teams.map(t => {{
    const opt = computeOptimalLineup(t.roster);
    return opt && {{optimal: opt.optimalTotal, left: opt.leftOnBench}};
}})));
""",
    )
    for team, js in zip(teams, rendered):
        py = calculate_lineup_efficiency(team)
        assert js['optimal'] == py['optimal_points'], team['abbrev']
        assert round(js['left'], 1) == py['points_left_on_table'], team['abbrev']


def test_bench_mistakes_pair_best_bench_with_worst_starter():
    roster = [
        _player('Starter A', 'RB', 10, True),
        _player('Starter B', 'RB', 2, True),
        _player('Bench C', 'RB', 15, False),
        _player('Bench D', 'RB', 1, False),
        _player('QB', 'QB', 20, True),
        _player('Tied Bench QB', 'QB', 20, False),
    ]
    rendered = _run_node(
        {},
        f"""
const opt = computeOptimalLineup({json.dumps(roster)});
process.stdout.write(JSON.stringify({{
    optimal: opt.optimalTotal,
    left: opt.leftOnBench,
    mistakes: opt.mistakes.map(m => [m.benched?.name, m.started.name, m.margin]),
}}));
""",
    )
    assert rendered['optimal'] == 45
    assert rendered['left'] == 13
    # A tie at QB is not a mistake.
    assert rendered['mistakes'] == [['Bench C', 'Starter B', 13]]


def test_top_six_secured_only_when_nobody_can_pass():
    def team(abbrev, score, starter_team):
        return {'abbrev': abbrev, 'total_score': score, 'roster': [_player(abbrev, 'QB', score, True) | {'nfl_team': starter_team}]}

    # 12 teams; DONE plays in a final game, LIVE in one still going.
    teams = [team(f'T{i}', 100 - i * 5, 'DONE') for i in range(12)]
    teams[5] = team('T5', 75, 'DONE')
    teams[6] = team('T6', 70, 'LIVE')
    week = {'week': 4, 'final': False, 'teams': teams}
    data = {
        'game_opponents': {'4': {'DONE': {'opponent': 'X', 'final': True}, 'LIVE': {'opponent': 'Y', 'final': False}}},
    }
    rendered = _run_node(
        data,
        f"""
const ranks = computeTopSixRanks({json.dumps(week)});
const pick = a => {{ const r = ranks.get(a); return [r.rank, r.topSix, r.secured]; }};
const line = renderScoreRankLine({json.dumps(teams[6])}, {json.dumps(week)}, ranks);
process.stdout.write(JSON.stringify({{t0: pick('T0'), t5: pick('T5'), t6: pick('T6'), line}}));
""",
    )
    assert rendered['t0'] == [1, True, True]
    # T5 is 6th and done, but T6 is still playing and could pass it.
    assert rendered['t5'] == [6, True, False]
    assert rendered['t6'] == [7, False, False]
    assert '7th of 12 in points · 1 starter left' in rendered['line']


def test_h2h_is_from_first_teams_perspective():
    data = {
        'season': 2026,
        'schedule': {'1': [['G/G', 'STL']], '2': [['STL', 'G/G']]},
        'weeks': [
            {'week': 1, 'final': True, 'teams': [{'abbrev': 'G/G', 'total_score': 50}, {'abbrev': 'STL', 'total_score': 40}]},
            {'week': 2, 'final': False, 'teams': [{'abbrev': 'G/G', 'total_score': 1}, {'abbrev': 'STL', 'total_score': 60}]},
        ],
        'hall_of_fame': {
            'seasons': [2022, 2026],
            'head_to_head': [{'team1': 'STL', 'team2': 'G/G', 'team1_wins': 3, 'team2_wins': 5, 'ties': 1}],
        },
    }
    rendered = _run_node(data, "process.stdout.write(JSON.stringify(renderH2HBadge('G/G', 'STL')));")
    assert 'Since 2022: 5–3–1' in rendered
    # Week 2 isn't final, so only week 1 counts this season.
    assert '2026: 1–0' in rendered
