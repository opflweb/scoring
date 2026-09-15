"""Week completion: a week only counts once every NFL game in it is over.

This is what stops a Sunday-evening scoring run from posting W/L records while
the Monday night game has not kicked off — the bug that shipped Week 1 of 2026
into the standings with DEN@KC unplayed.
"""

from opfl.week_status import latest_completed_week, week_games_are_final


def game(week, result=None, game_type='REG', season=2026):
    return {'week': week, 'result': result, 'game_type': game_type, 'season': season}


class TestWeekGamesAreFinal:
    def test_all_games_final(self):
        rows = [game(1, result=7), game(1, result=-3)]
        assert week_games_are_final(rows, 1, 2026)

    def test_one_unplayed_game_holds_the_whole_week_open(self):
        """The real Week 1 2026 case: fifteen games done, DEN@KC still to play."""
        rows = [game(1, result=7)] * 15 + [game(1, result=None)]
        assert not week_games_are_final(rows, 1, 2026)

    def test_a_zero_margin_still_counts_as_a_result(self):
        """A tie has result 0, which is falsy — it must not read as unplayed."""
        assert week_games_are_final([game(1, result=0)], 1, 2026)

    def test_empty_string_result_counts_as_unplayed(self):
        assert not week_games_are_final([game(1, result='')], 1, 2026)

    def test_a_week_with_no_games_is_not_final(self):
        """No evidence is not the same as finished."""
        assert not week_games_are_final([game(1, result=7)], 5, 2026)
        assert not week_games_are_final([], 1, 2026)

    def test_playoff_games_are_ignored(self):
        """Only REG games bear on a fantasy week."""
        rows = [game(1, result=7), game(1, result=None, game_type='POST')]
        assert week_games_are_final(rows, 1, 2026)

    def test_other_seasons_do_not_vote(self):
        """Prior-season rows are all final and would otherwise carry the week."""
        rows = [game(1, result=7, season=2025), game(1, result=None, season=2026)]
        assert not week_games_are_final(rows, 1, 2026)
        assert week_games_are_final(rows, 1, 2025)


class TestLatestCompletedWeek:
    def test_returns_the_last_fully_final_week(self):
        rows = [game(1, result=7), game(2, result=3), game(3, result=None)]
        assert latest_completed_week(rows) == 2

    def test_zero_when_nothing_has_finished(self):
        assert latest_completed_week([game(1, result=None)]) == 0
        assert latest_completed_week([]) == 0

    def test_weeks_beyond_the_cap_are_ignored(self):
        rows = [game(1, result=7), game(18, result=7)]
        assert latest_completed_week(rows, max_week=17) == 1

    def test_a_gap_does_not_advance_past_an_unfinished_week(self):
        """Week 2 unplayed but week 3 done should not report 3 as 'completed
        through'. It reports the max completed week, so callers that need a
        contiguous run must check each week themselves."""
        rows = [game(1, result=7), game(2, result=None), game(3, result=7)]
        assert latest_completed_week(rows) == 3
        assert not week_games_are_final(rows, 2, 2026)
