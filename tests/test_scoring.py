"""OPFL position scoring rules.

Each test asserts the total *and* the breakdown dict, so a rule change that
happens to preserve the total still shows up as a failure. Rules are as
documented in README.md under "OPFL Scoring Rules".
"""

from opfl.scoring import (
    score_defense,
    score_head_coach,
    score_kicker,
    score_qb,
    score_rb_wr,
    score_te,
)


class TestQuarterback:
    def test_under_200_passing_yards_scores_nothing(self):
        """Below the 200-yard threshold a QB gets no passing points at all."""
        points, breakdown = score_qb({'passing_yards': 199})
        assert points == 0
        assert 'passing_yards' not in breakdown

    def test_200_passing_yards_is_two_points(self):
        points, breakdown = score_qb({'passing_yards': 200})
        assert points == 2
        assert breakdown['passing_yards'] == 2

    def test_passing_yards_add_one_per_50_thereafter(self):
        """249 rounds down to the 200 band; 250 crosses into the next one."""
        assert score_qb({'passing_yards': 249})[0] == 2
        assert score_qb({'passing_yards': 250})[0] == 3
        assert score_qb({'passing_yards': 400})[0] == 6

    def test_rushing_yards_use_the_75_yard_band(self):
        points, breakdown = score_qb({'rushing_yards': 100})
        assert points == 3
        assert breakdown['rushing_yards'] == 3

    def test_touchdowns_are_six_each_across_all_types(self):
        points, breakdown = score_qb({'passing_tds': 2, 'rushing_tds': 1, 'fumble_recovery_tds': 1})
        assert points == 24
        assert breakdown['touchdowns'] == 24

    def test_two_point_conversions_are_two_each(self):
        points, breakdown = score_qb({'passing_2pt_conversions': 1, 'rushing_2pt_conversions': 1})
        assert points == 4
        assert breakdown['two_point_conversions'] == 4

    def test_interception_costs_one_point(self):
        points, breakdown = score_qb({'passing_tds': 1, 'passing_interceptions': 1})
        assert points == 5
        assert breakdown['interceptions'] == -1

    def test_pick_six_costs_three_and_replaces_the_regular_interception(self):
        """A pick-6 is -3 total, not -1 for the INT plus -3 for the return."""
        points, breakdown = score_qb(
            {'passing_tds': 2, 'passing_interceptions': 1},
            turnover_tds={'pick_sixes': 1},
        )
        assert points == 9
        assert breakdown['pick_sixes'] == -3
        assert 'interceptions' not in breakdown

    def test_pick_six_alongside_a_regular_interception(self):
        points, breakdown = score_qb(
            {'passing_tds': 2, 'passing_interceptions': 2},
            turnover_tds={'pick_sixes': 1},
        )
        assert points == 8
        assert breakdown['interceptions'] == -1
        assert breakdown['pick_sixes'] == -3

    def test_fumble_lost_costs_one_point(self):
        points, breakdown = score_qb({'passing_tds': 1, 'rushing_fumbles_lost': 1})
        assert points == 5
        assert breakdown['fumbles_lost'] == -1

    def test_score_is_floored_at_zero(self):
        points, breakdown = score_qb({'passing_interceptions': 3})
        assert points == 0
        assert breakdown['floor_applied'] is True


class TestRunningBackWideReceiver:
    def test_under_75_yards_in_each_category_scores_nothing(self):
        points, breakdown = score_rb_wr({'rushing_yards': 74, 'receiving_yards': 20})
        assert points == 0
        assert breakdown == {}

    def test_individual_category_starts_at_75_yards(self):
        points, breakdown = score_rb_wr({'rushing_yards': 75})
        assert points == 2
        assert breakdown['rushing_yards'] == 2

    def test_combined_bonus_wins_when_it_beats_the_individual_total(self):
        """45 rush + 80 rec: individually only the 80 rec clears 75, for 2 pts.

        Combined is 125, giving 2 + (125-100)//25 = 3, so the combined path wins
        and replaces the individual line items entirely.
        """
        points, breakdown = score_rb_wr({'rushing_yards': 45, 'receiving_yards': 80})
        assert points == 3
        assert breakdown['combined_rush_rec_yards'] == 3
        assert 'receiving_yards' not in breakdown

    def test_combined_bonus_still_wins_when_both_categories_clear_75(self):
        """95 rush + 95 rec = 2 + 2 = 4 individually, vs 2 + (190-100)//25 = 5 combined."""
        points, breakdown = score_rb_wr({'rushing_yards': 95, 'receiving_yards': 95})
        assert points == 5
        assert breakdown['combined_rush_rec_yards'] == 5

    def test_real_week_1_2026_saquon_barkley(self):
        """83 rush, -3 rec: clears the 75 individual band, misses 100 combined."""
        points, breakdown = score_rb_wr({'rushing_yards': 83, 'receiving_yards': -3})
        assert points == 2
        assert breakdown == {'rushing_yards': 2}

    def test_touchdown_plus_yards(self):
        points, breakdown = score_rb_wr({'rushing_yards': 80, 'rushing_tds': 2})
        assert points == 14
        assert breakdown['touchdowns'] == 12


class TestTightEnd:
    def test_individual_band_starts_at_50_yards(self):
        """TEs get a lower threshold than RB/WR: 50 instead of 75."""
        points, breakdown = score_te({'receiving_yards': 50})
        assert points == 2
        assert breakdown['receiving_yards'] == 2
        assert score_te({'receiving_yards': 49})[0] == 0

    def test_combined_bonus_threshold_is_75_for_tight_ends(self):
        points, breakdown = score_te({'rushing_yards': 40, 'receiving_yards': 45})
        assert points == 2
        assert breakdown['combined_rush_rec_yards'] == 2


class TestKicker:
    def test_pat_made_is_one_point_each(self):
        points, breakdown = score_kicker({'pat_made': 3})
        assert points == 3
        assert breakdown['pat_made'] == 3

    def test_field_goals_score_by_distance_band(self):
        points, breakdown = score_kicker(
            {
                'fg_made_20_29': 1,  # 1 pt
                'fg_made_30_39': 1,  # 2 pts
                'fg_made_40_49': 1,  # 3 pts
                'fg_made_50_59': 1,  # 4 pts
            }
        )
        assert points == 10
        assert breakdown['fg_1_29'] == 1
        assert breakdown['fg_30_39'] == 2
        assert breakdown['fg_40_49'] == 3
        assert breakdown['fg_50+'] == 4

    def test_60_plus_yard_field_goals_count_in_the_50_plus_band(self):
        points, breakdown = score_kicker({'fg_made_60_': 1})
        assert points == 4
        assert breakdown['fg_50+'] == 4

    def test_missed_field_goal_costs_two_and_missed_pat_costs_one(self):
        points, breakdown = score_kicker({'pat_made': 4, 'fg_missed': 1, 'pat_missed': 1})
        assert points == 1
        assert breakdown['fg_missed'] == -2
        assert breakdown['pat_missed'] == -1

    def test_kicker_score_is_floored_at_zero(self):
        points, breakdown = score_kicker({'fg_missed': 2})
        assert points == 0
        assert breakdown['floor_applied'] is True


class TestDefense:
    def test_shutout_is_eight_points(self):
        points, breakdown = score_defense({}, {}, {'points_allowed': 0})
        assert points == 8
        assert breakdown['points_allowed'] == 8

    def test_points_allowed_bands(self):
        bands = {0: 8, 9: 6, 13: 4, 17: 2, 27: 0, 31: -2, 35: -4, 36: -6, 50: -6}
        for allowed, expected in bands.items():
            points, breakdown = score_defense({}, {}, {'points_allowed': allowed})
            assert breakdown['points_allowed'] == expected, f'{allowed} allowed'

    def test_turnovers_and_sacks(self):
        points, breakdown = score_defense(
            {'def_interceptions': 2, 'fumble_recovery_opp': 1, 'def_sacks': 3},
            {},
            {'points_allowed': 20},
        )
        assert points == 9
        assert breakdown['interceptions'] == 4
        assert breakdown['fumble_recoveries'] == 2
        assert breakdown['sacks'] == 3

    def test_pbp_sack_count_overrides_the_aggregated_one(self):
        """Aggregated team stats and play-by-play disagree; PBP wins."""
        _, breakdown = score_defense({'def_sacks': 3}, {}, {'points_allowed': 20}, pbp_sacks=5)
        assert breakdown['sacks'] == 5

    def test_defensive_touchdown_is_four_points(self):
        points, breakdown = score_defense({'def_tds': 1}, {}, {'points_allowed': 20})
        assert points == 4
        assert breakdown['defensive_tds'] == 4

    def test_blocked_kicks_and_pats(self):
        points, breakdown = score_defense(
            {}, {'fg_blocked': 1, '_blocked_punts': 1, 'pat_blocked': 1}, {'points_allowed': 20}
        )
        assert points == 5
        assert breakdown['blocked_kicks'] == 4
        assert breakdown['blocked_pats'] == 1

    def test_defense_score_is_floored_at_zero(self):
        points, breakdown = score_defense({}, {}, {'points_allowed': 45})
        assert points == 0
        assert breakdown['floor_applied'] is True


class TestHeadCoach:
    def test_loss_is_zero_points(self):
        points, breakdown = score_head_coach({'team_score': 10, 'opponent_score': 20})
        assert points == 0
        assert breakdown['loss'] == 0

    def test_tie_is_zero_points(self):
        points, breakdown = score_head_coach({'team_score': 20, 'opponent_score': 20})
        assert points == 0
        assert breakdown['loss'] == 0

    def test_home_favorite_win_is_four(self):
        """nflverse spread_line is from the home team's perspective: positive = home favored."""
        points, breakdown = score_head_coach(
            {'team_score': 24, 'opponent_score': 20, 'is_home': True}, {'spread': 3.0}
        )
        assert points == 4
        assert breakdown['home_favorite_win'] == 4

    def test_home_underdog_win_is_six(self):
        points, breakdown = score_head_coach(
            {'team_score': 24, 'opponent_score': 20, 'is_home': True}, {'spread': -3.0}
        )
        assert points == 6
        assert breakdown['home_underdog_win'] == 6

    def test_road_favorite_win_is_five(self):
        points, breakdown = score_head_coach(
            {'team_score': 24, 'opponent_score': 20, 'is_home': False}, {'spread': 3.0}
        )
        assert points == 5
        assert breakdown['road_favorite_win'] == 5

    def test_road_underdog_win_is_seven(self):
        points, breakdown = score_head_coach(
            {'team_score': 24, 'opponent_score': 20, 'is_home': False}, {'spread': -3.0}
        )
        assert points == 7
        assert breakdown['road_underdog_win'] == 7

    def test_missing_spread_is_treated_as_favorite(self):
        """No line available: assume favorite, which is the more common case for a winner."""
        points, breakdown = score_head_coach(
            {'team_score': 24, 'opponent_score': 20, 'is_home': True}
        )
        assert points == 4
        assert breakdown['home_favorite_win'] == 4
