"""Player score display behavior for games that have not kicked off."""

from pathlib import Path

APP_JS = (Path(__file__).parent.parent / 'web' / 'app.js').read_text()


def test_player_score_display_uses_kickoff_times():
    assert 'function playerGameHasStarted(player, weekNum)' in APP_JS
    assert 'data?.game_times' in APP_JS
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
