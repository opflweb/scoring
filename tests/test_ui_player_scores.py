"""Player score display behavior for games that have not kicked off."""

import json
import os
import subprocess
from pathlib import Path

APP_JS = (Path(__file__).parent.parent / 'web' / 'app.js').read_text()
STYLES_CSS = (Path(__file__).parent.parent / 'web' / 'styles.css').read_text()


def test_player_score_display_uses_kickoff_times():
    assert 'function playerGameHasStarted(player, weekNum)' in APP_JS
    assert 'data?.game_times' in APP_JS
    assert 'data?.game_opponents' in APP_JS
    assert 'kickoffTime <= Date.now()' in APP_JS


def test_future_players_render_as_a_dash_without_hiding_real_zeroes():
    score_fn = APP_JS.split('function playerScoreText(player, weekNum)')[1].split(
        '\n        function '
    )[0]
    assert "return '-';" in score_fn
    assert 'return player.score ?? 0;' in score_fn
    assert 'player.score || 0' not in score_fn


def test_current_matchup_rosters_pass_the_selected_week_to_score_display():
    assert 'renderRosterList(t1.roster, currentWeek)' in APP_JS
    assert 'renderRosterList(t2.roster, currentWeek)' in APP_JS


def test_matchup_rosters_show_opponent_kickoff_and_projection():
    assert 'function getPlayerGameDetails(player, weekNum)' in APP_JS
    assert '`@${opponent.opponent}`' in APP_JS
    assert '`vs ${opponent.opponent}`' in APP_JS
    assert 'class="player-game-time ${game.colorClass}"' in APP_JS
    assert 'Proj ${p.projected_points.toFixed(1)}' in APP_JS


def test_scored_players_can_expand_their_breakdown():
    assert 'function renderBreakdown(breakdown)' in APP_JS
    assert '<details class="score-breakdown">' in APP_JS
    assert 'BREAKDOWN_LABELS' in APP_JS
    assert '.breakdown-content {' in STYLES_CSS
    assert '.player-projection {' in STYLES_CSS


def test_game_context_and_breakdown_render_like_qpfl():
    helpers = APP_JS[
        APP_JS.index('const NFL_TEAM_ALIASES') : APP_JS.index(
            '// Historical championship data'
        )
    ]
    breakdown_helpers = APP_JS[
        APP_JS.index('const BREAKDOWN_LABELS') : APP_JS.index('function toggleRoster')
    ]
    script = f"""
let data = {{
    game_times: {{'2': {{LA: '2026-09-21T20:15:00-04:00'}}}},
    game_opponents: {{'2': {{
        LA: {{opponent: 'NYG', is_home: true, final: false}},
        NYG: {{opponent: 'LA', is_home: false, final: false}}
    }}}}
}};
{helpers}
{breakdown_helpers}
Date.now = () => Date.parse('2026-09-21T17:00:00-04:00');
const player = {{nfl_team: 'LA', score: 0, projected_points: 8.4}};
const future = getPlayerGameDetails(player, 2);
const futureScore = playerScoreText(player, 2);
data.game_opponents['2'].LA.final = true;
const finalGame = getPlayerGameDetails(player, 2);
const breakdown = renderBreakdown({{touchdowns: 6, interceptions: -1}});
process.stdout.write(JSON.stringify({{future, futureScore, finalGame, breakdown}}));
"""
    env = dict(os.environ, TZ='America/New_York')
    result = subprocess.run(
        ['node', '-e', script], check=True, capture_output=True, text=True, env=env
    )
    rendered = json.loads(result.stdout)

    assert rendered['future']['matchup'] == 'vs NYG'
    assert rendered['future']['gameTime'] == 'Mon 8:15p'
    assert rendered['futureScore'] == '-'
    assert rendered['finalGame']['gameTime'] == 'Final'
    assert 'TD' in rendered['breakdown']
    assert '+6' in rendered['breakdown']
    assert 'INT' in rendered['breakdown']
    assert '-1' in rendered['breakdown']
