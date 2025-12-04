"""Scoring functions for each position type based on OPFL rules."""

import math
from typing import Dict, Tuple


def score_qb(stats: dict, turnover_tds: dict = None) -> Tuple[float, Dict[str, float]]:
    """
    Score a quarterback using OPFL rules.
    
    Scoring:
        - Passing yards: 200 yards = 2 pts, +1 pt per 50 yards thereafter (below 200 = 0)
        - Rushing yards: 75 yards = 2 pts, +1 pt per 25 yards thereafter (RB/WR rules)
        - Receiving yards: 75 yards = 2 pts, +1 pt per 25 yards thereafter
        - Touchdowns (pass/rush/rec): 6 pts each
        - 2pt conversions: 2 pts each
        - Interceptions: -1 pt each
        - Interceptions returned for TD: -3 pts additional each
        - Fumbles lost: -1 pt each
        - Fumbles returned for TD: -3 pts additional each
        - Points cannot be less than zero
    """
    points = 0.0
    breakdown = {}
    turnover_tds = turnover_tds or {}
    
    # Passing yards
    passing_yards = stats.get('passing_yards', 0) or 0
    if passing_yards >= 200:
        passing_pts = 2 + max(0, (passing_yards - 200) // 50)
        breakdown['passing_yards'] = passing_pts
        points += passing_pts
    
    # Rushing yards (use RB/WR rules - 75 yard threshold)
    rushing_yards = stats.get('rushing_yards', 0) or 0
    if rushing_yards >= 75:
        rushing_pts = 2 + max(0, (rushing_yards - 75) // 25)
        breakdown['rushing_yards'] = rushing_pts
        points += rushing_pts
    
    # Receiving yards (use RB/WR rules - 75 yard threshold)
    receiving_yards = stats.get('receiving_yards', 0) or 0
    if receiving_yards >= 75:
        receiving_pts = 2 + max(0, (receiving_yards - 75) // 25)
        breakdown['receiving_yards'] = receiving_pts
        points += receiving_pts
    
    # Touchdowns (6 points each)
    total_tds = (
        (stats.get('passing_tds', 0) or 0) +
        (stats.get('rushing_tds', 0) or 0) +
        (stats.get('receiving_tds', 0) or 0)
    )
    if total_tds:
        td_pts = 6 * total_tds
        breakdown['touchdowns'] = td_pts
        points += td_pts
    
    # Two point conversions (2 points each)
    two_pt = (
        (stats.get('passing_2pt_conversions', 0) or 0) +
        (stats.get('rushing_2pt_conversions', 0) or 0) +
        (stats.get('receiving_2pt_conversions', 0) or 0)
    )
    if two_pt:
        two_pt_pts = 2 * two_pt
        breakdown['two_point_conversions'] = two_pt_pts
        points += two_pt_pts
    
    # Turnovers: -1 pt each, BUT -3 pts total if returned for TD (not stacking)
    pick_sixes = turnover_tds.get('pick_sixes', 0)
    fumble_sixes = turnover_tds.get('fumble_sixes', 0)
    
    # Interceptions: -1 pt each (regular), -3 pts each (pick-6)
    interceptions = stats.get('passing_interceptions', 0) or 0
    regular_ints = max(0, interceptions - pick_sixes)
    if regular_ints:
        int_pts = -1 * regular_ints
        breakdown['interceptions'] = int_pts
        points += int_pts
    if pick_sixes:
        pick_six_pts = -3 * pick_sixes
        breakdown['pick_sixes'] = pick_six_pts
        points += pick_six_pts
    
    # Fumbles lost: -1 pt each (regular), -3 pts each (fumble-6)
    fumbles_lost = (
        (stats.get('sack_fumbles_lost', 0) or 0) +
        (stats.get('rushing_fumbles_lost', 0) or 0) +
        (stats.get('receiving_fumbles_lost', 0) or 0)
    )
    regular_fumbles = max(0, fumbles_lost - fumble_sixes)
    if regular_fumbles:
        fumble_pts = -1 * regular_fumbles
        breakdown['fumbles_lost'] = fumble_pts
        points += fumble_pts
    if fumble_sixes:
        fumble_six_pts = -3 * fumble_sixes
        breakdown['fumble_sixes'] = fumble_six_pts
        points += fumble_six_pts
    
    # Points cannot be less than zero
    if points < 0:
        points = 0
        breakdown['floor_applied'] = True
    
    return points, breakdown


def score_rb_wr(stats: dict, turnover_tds: dict = None) -> Tuple[float, Dict[str, float]]:
    """
    Score a running back or wide receiver using OPFL rules.
    
    Scoring:
        Individual category (rushing OR receiving):
            - 75 yards = 2 pts, +1 pt per 25 yards thereafter (below 75 = 0)
        
        Alternate combined bonus:
            - 100 combined rush/rec yards = 2 pts, +1 pt per 25 yards thereafter
            - Player gets whichever method yields more points
        
        Other:
            - Touchdowns (pass/rush/rec): 6 pts each
            - 2pt conversions: 2 pts each
            - Fumbles lost / INTs thrown: -1 pt each
            - Turnovers returned for TD: -3 pts additional each
            - Points cannot be less than zero
    """
    points = 0.0
    breakdown = {}
    turnover_tds = turnover_tds or {}
    
    rushing_yards = stats.get('rushing_yards', 0) or 0
    receiving_yards = stats.get('receiving_yards', 0) or 0
    combined_yards = rushing_yards + receiving_yards
    
    # Calculate individual category points
    individual_rushing_pts = 0
    if rushing_yards >= 75:
        individual_rushing_pts = 2 + max(0, (rushing_yards - 75) // 25)
    
    individual_receiving_pts = 0
    if receiving_yards >= 75:
        individual_receiving_pts = 2 + max(0, (receiving_yards - 75) // 25)
    
    individual_total = individual_rushing_pts + individual_receiving_pts
    
    # Calculate combined bonus (RB/WR: 100 yards threshold)
    combined_pts = 0
    if combined_yards >= 100:
        combined_pts = 2 + max(0, (combined_yards - 100) // 25)
    
    # Use whichever method yields more points
    if combined_pts > individual_total:
        breakdown['combined_rush_rec_yards'] = combined_pts
        points += combined_pts
    else:
        if individual_rushing_pts:
            breakdown['rushing_yards'] = individual_rushing_pts
            points += individual_rushing_pts
        if individual_receiving_pts:
            breakdown['receiving_yards'] = individual_receiving_pts
            points += individual_receiving_pts
    
    # Passing yards (if applicable, e.g., trick plays)
    passing_yards = stats.get('passing_yards', 0) or 0
    if passing_yards >= 200:
        passing_pts = 2 + max(0, (passing_yards - 200) // 50)
        breakdown['passing_yards'] = passing_pts
        points += passing_pts
    
    # Touchdowns (6 points each)
    total_tds = (
        (stats.get('passing_tds', 0) or 0) +
        (stats.get('rushing_tds', 0) or 0) +
        (stats.get('receiving_tds', 0) or 0)
    )
    if total_tds:
        td_pts = 6 * total_tds
        breakdown['touchdowns'] = td_pts
        points += td_pts
    
    # Two point conversions (2 points each)
    two_pt = (
        (stats.get('passing_2pt_conversions', 0) or 0) +
        (stats.get('rushing_2pt_conversions', 0) or 0) +
        (stats.get('receiving_2pt_conversions', 0) or 0)
    )
    if two_pt:
        two_pt_pts = 2 * two_pt
        breakdown['two_point_conversions'] = two_pt_pts
        points += two_pt_pts
    
    # Turnovers: -1 pt each, BUT -3 pts total if returned for TD (not stacking)
    pick_sixes = turnover_tds.get('pick_sixes', 0)
    fumble_sixes = turnover_tds.get('fumble_sixes', 0)
    
    # Interceptions thrown: -1 pt each (regular), -3 pts each (pick-6)
    interceptions = stats.get('passing_interceptions', 0) or 0
    regular_ints = max(0, interceptions - pick_sixes)
    if regular_ints:
        int_pts = -1 * regular_ints
        breakdown['interceptions'] = int_pts
        points += int_pts
    if pick_sixes:
        pick_six_pts = -3 * pick_sixes
        breakdown['pick_sixes'] = pick_six_pts
        points += pick_six_pts
    
    # Fumbles lost: -1 pt each (regular), -3 pts each (fumble-6)
    fumbles_lost = (
        (stats.get('sack_fumbles_lost', 0) or 0) +
        (stats.get('rushing_fumbles_lost', 0) or 0) +
        (stats.get('receiving_fumbles_lost', 0) or 0)
    )
    regular_fumbles = max(0, fumbles_lost - fumble_sixes)
    if regular_fumbles:
        fumble_pts = -1 * regular_fumbles
        breakdown['fumbles_lost'] = fumble_pts
        points += fumble_pts
    if fumble_sixes:
        fumble_six_pts = -3 * fumble_sixes
        breakdown['fumble_sixes'] = fumble_six_pts
        points += fumble_six_pts
    
    # Points cannot be less than zero
    if points < 0:
        points = 0
        breakdown['floor_applied'] = True
    
    return points, breakdown


def score_te(stats: dict, turnover_tds: dict = None) -> Tuple[float, Dict[str, float]]:
    """
    Score a tight end using OPFL rules.
    
    Scoring:
        Individual category (rushing OR receiving):
            - 50 yards = 2 pts, +1 pt per 25 yards thereafter (below 50 = 0)
        
        Alternate combined bonus:
            - 75 combined rush/rec yards = 2 pts, +1 pt per 25 yards thereafter
            - Player gets whichever method yields more points
        
        Other:
            - Touchdowns (pass/rush/rec): 6 pts each
            - 2pt conversions: 2 pts each
            - Fumbles lost / INTs thrown: -1 pt each
            - Turnovers returned for TD: -3 pts additional each
            - Points cannot be less than zero
    """
    points = 0.0
    breakdown = {}
    turnover_tds = turnover_tds or {}
    
    rushing_yards = stats.get('rushing_yards', 0) or 0
    receiving_yards = stats.get('receiving_yards', 0) or 0
    combined_yards = rushing_yards + receiving_yards
    
    # Calculate individual category points (TE: 50 yard threshold)
    individual_rushing_pts = 0
    if rushing_yards >= 50:
        individual_rushing_pts = 2 + max(0, (rushing_yards - 50) // 25)
    
    individual_receiving_pts = 0
    if receiving_yards >= 50:
        individual_receiving_pts = 2 + max(0, (receiving_yards - 50) // 25)
    
    individual_total = individual_rushing_pts + individual_receiving_pts
    
    # Calculate combined bonus (TE: 75 yards threshold)
    combined_pts = 0
    if combined_yards >= 75:
        combined_pts = 2 + max(0, (combined_yards - 75) // 25)
    
    # Use whichever method yields more points
    if combined_pts > individual_total:
        breakdown['combined_rush_rec_yards'] = combined_pts
        points += combined_pts
    else:
        if individual_rushing_pts:
            breakdown['rushing_yards'] = individual_rushing_pts
            points += individual_rushing_pts
        if individual_receiving_pts:
            breakdown['receiving_yards'] = individual_receiving_pts
            points += individual_receiving_pts
    
    # Passing yards (if applicable, e.g., trick plays)
    passing_yards = stats.get('passing_yards', 0) or 0
    if passing_yards >= 200:
        passing_pts = 2 + max(0, (passing_yards - 200) // 50)
        breakdown['passing_yards'] = passing_pts
        points += passing_pts
    
    # Touchdowns (6 points each)
    total_tds = (
        (stats.get('passing_tds', 0) or 0) +
        (stats.get('rushing_tds', 0) or 0) +
        (stats.get('receiving_tds', 0) or 0)
    )
    if total_tds:
        td_pts = 6 * total_tds
        breakdown['touchdowns'] = td_pts
        points += td_pts
    
    # Two point conversions (2 points each)
    two_pt = (
        (stats.get('passing_2pt_conversions', 0) or 0) +
        (stats.get('rushing_2pt_conversions', 0) or 0) +
        (stats.get('receiving_2pt_conversions', 0) or 0)
    )
    if two_pt:
        two_pt_pts = 2 * two_pt
        breakdown['two_point_conversions'] = two_pt_pts
        points += two_pt_pts
    
    # Turnovers: -1 pt each, BUT -3 pts total if returned for TD (not stacking)
    pick_sixes = turnover_tds.get('pick_sixes', 0)
    fumble_sixes = turnover_tds.get('fumble_sixes', 0)
    
    # Interceptions thrown: -1 pt each (regular), -3 pts each (pick-6)
    interceptions = stats.get('passing_interceptions', 0) or 0
    regular_ints = max(0, interceptions - pick_sixes)
    if regular_ints:
        int_pts = -1 * regular_ints
        breakdown['interceptions'] = int_pts
        points += int_pts
    if pick_sixes:
        pick_six_pts = -3 * pick_sixes
        breakdown['pick_sixes'] = pick_six_pts
        points += pick_six_pts
    
    # Fumbles lost: -1 pt each (regular), -3 pts each (fumble-6)
    fumbles_lost = (
        (stats.get('sack_fumbles_lost', 0) or 0) +
        (stats.get('rushing_fumbles_lost', 0) or 0) +
        (stats.get('receiving_fumbles_lost', 0) or 0)
    )
    regular_fumbles = max(0, fumbles_lost - fumble_sixes)
    if regular_fumbles:
        fumble_pts = -1 * regular_fumbles
        breakdown['fumbles_lost'] = fumble_pts
        points += fumble_pts
    if fumble_sixes:
        fumble_six_pts = -3 * fumble_sixes
        breakdown['fumble_sixes'] = fumble_six_pts
        points += fumble_six_pts
    
    # Points cannot be less than zero
    if points < 0:
        points = 0
        breakdown['floor_applied'] = True
    
    return points, breakdown


def score_kicker(stats: dict) -> Tuple[float, Dict[str, float]]:
    """
    Score a kicker using OPFL rules.
    
    Scoring:
        - PAT made: 1 pt each
        - Missed/blocked PAT: -1 pt each
        - FG 1-29 yards: 1 pt each
        - FG 30-39 yards: 2 pts each
        - FG 40-49 yards: 3 pts each
        - FG 50+ yards: 4 pts each
        - Missed/blocked FG: -2 pts each
    """
    points = 0.0
    breakdown = {}
    
    # PATs made (1 point each)
    pat_made = stats.get('pat_made', 0) or 0
    if pat_made:
        breakdown['pat_made'] = pat_made
    points += pat_made
    
    # Missed/blocked PATs (-1 point each)
    pat_missed = stats.get('pat_missed', 0) or 0
    pat_blocked = stats.get('pat_blocked', 0) or 0
    total_pat_missed = pat_missed + pat_blocked
    if total_pat_missed:
        breakdown['pat_missed'] = -1 * total_pat_missed
        points -= total_pat_missed
    
    # Field Goals by distance
    fg_0_19 = stats.get('fg_made_0_19', 0) or 0
    fg_20_29 = stats.get('fg_made_20_29', 0) or 0
    fg_1_29 = fg_0_19 + fg_20_29
    if fg_1_29:
        breakdown['fg_1_29'] = fg_1_29
    points += fg_1_29
    
    fg_30_39 = stats.get('fg_made_30_39', 0) or 0
    if fg_30_39:
        breakdown['fg_30_39'] = 2 * fg_30_39
    points += 2 * fg_30_39
    
    fg_40_49 = stats.get('fg_made_40_49', 0) or 0
    if fg_40_49:
        breakdown['fg_40_49'] = 3 * fg_40_49
    points += 3 * fg_40_49
    
    # 50+ yards (includes 50-59 and 60+)
    fg_50_59 = stats.get('fg_made_50_59', 0) or 0
    fg_60_plus = stats.get('fg_made_60_', 0) or 0
    fg_50_plus = fg_50_59 + fg_60_plus
    if fg_50_plus:
        breakdown['fg_50+'] = 4 * fg_50_plus
        points += 4 * fg_50_plus
    
    # Missed/blocked FGs (-2 points each)
    fg_missed = stats.get('fg_missed', 0) or 0
    fg_blocked = stats.get('fg_blocked', 0) or 0
    total_fg_missed = fg_missed + fg_blocked
    if total_fg_missed:
        breakdown['fg_missed'] = -2 * total_fg_missed
        points -= 2 * total_fg_missed
    
    # Points cannot be less than zero
    if points < 0:
        points = 0
        breakdown['floor_applied'] = True
    
    return points, breakdown


def score_defense(
    team_stats: dict,
    opponent_stats: dict,
    game_info: dict,
    pbp_sacks: int = None,
) -> Tuple[float, Dict[str, float]]:
    """
    Score a defense using OPFL rules.
    
    Scoring:
        - Shutout (0 points): 8 pts
        - 2-9 points allowed: 6 pts
        - 10-13 points allowed: 4 pts
        - 14-17 points allowed: 2 pts
        - 18-27 points allowed: 0 pts
        - 28-31 points allowed: -2 pts
        - 32-35 points allowed: -4 pts
        - 36+ points allowed: -6 pts
        
        - Interception: 2 pts each
        - Fumble recovery: 2 pts each
        - Safety: 2 pts each
        - Blocked punt or FG: 2 pts each
        - Blocked PAT: 1 pt each
        - Defensive/ST TD: 4 pts each (INTs, fumble recoveries, blocked kicks count; punt/kick returns don't)
        - Sack: 1 pt each
    """
    points = 0.0
    breakdown = {}
    
    # Points allowed
    points_allowed = game_info.get('points_allowed', 0) or 0
    
    if points_allowed == 0:
        pa_pts = 8  # Shutout
    elif points_allowed <= 9:
        pa_pts = 6  # 2-9 points (assumes 1 point is rare/impossible)
    elif points_allowed <= 13:
        pa_pts = 4  # 10-13 points
    elif points_allowed <= 17:
        pa_pts = 2  # 14-17 points
    elif points_allowed <= 27:
        pa_pts = 0  # 18-27 points
    elif points_allowed <= 31:
        pa_pts = -2  # 28-31 points
    elif points_allowed <= 35:
        pa_pts = -4  # 32-35 points
    else:
        pa_pts = -6  # 36+ points
    
    breakdown['points_allowed'] = pa_pts
    points += pa_pts
    
    # Interceptions (2 pts each)
    interceptions = team_stats.get('def_interceptions', 0) or 0
    if interceptions:
        breakdown['interceptions'] = 2 * interceptions
    points += 2 * interceptions
    
    # Fumble recoveries (2 pts each)
    our_recoveries = team_stats.get('fumble_recovery_opp', 0) or 0
    opponent_fumbles_lost = (
        (opponent_stats.get('sack_fumbles_lost', 0) or 0) +
        (opponent_stats.get('rushing_fumbles_lost', 0) or 0) +
        (opponent_stats.get('receiving_fumbles_lost', 0) or 0)
    )
    fumble_recoveries = max(our_recoveries, opponent_fumbles_lost)
    if fumble_recoveries:
        breakdown['fumble_recoveries'] = 2 * fumble_recoveries
    points += 2 * fumble_recoveries
    
    # Sacks (1 pt each)
    sacks = pbp_sacks if pbp_sacks is not None else (team_stats.get('def_sacks', 0) or 0)
    if sacks:
        breakdown['sacks'] = int(sacks)
    points += int(sacks)
    
    # Safeties (2 pts each)
    safeties = team_stats.get('def_safeties', 0) or 0
    if safeties:
        breakdown['safeties'] = 2 * safeties
    points += 2 * safeties
    
    # Blocked punts/FGs (2 pts each)
    # blocked_fg comes from opponent team stats
    # blocked_punts should be passed in from PBP data
    blocked_fg = opponent_stats.get('fg_blocked', 0) or 0
    blocked_punts = opponent_stats.get('_blocked_punts', 0) or 0  # Custom field from PBP
    total_blocked_kicks = blocked_fg + blocked_punts
    if total_blocked_kicks:
        breakdown['blocked_kicks'] = 2 * total_blocked_kicks
    points += 2 * total_blocked_kicks
    
    # Blocked PATs (1 pt each)
    blocked_pat = opponent_stats.get('pat_blocked', 0) or 0
    if blocked_pat:
        breakdown['blocked_pats'] = blocked_pat
    points += blocked_pat
    
    # Defensive TDs (4 pts each)
    # Per rules: TDs count on INTs, fumble recoveries, blocked punts/kicks
    # but NOT on punt/kickoff returns
    def_tds = team_stats.get('def_tds', 0) or 0
    fumble_recovery_tds = team_stats.get('fumble_recovery_tds', 0) or 0
    # Blocked kick TDs come from PBP (they're categorized as special_teams_tds in nflverse)
    blocked_kick_tds = opponent_stats.get('_blocked_kick_tds', 0) or 0
    # Note: special_teams_tds includes kick/punt return TDs which don't count per OPFL rules
    # So we count def_tds, fumble_recovery_tds, and blocked_kick_tds separately
    total_def_tds = def_tds + fumble_recovery_tds + blocked_kick_tds
    if total_def_tds:
        breakdown['defensive_tds'] = 4 * total_def_tds
        points += 4 * total_def_tds
    
    # Points cannot be less than zero
    if points < 0:
        points = 0
        breakdown['floor_applied'] = True
    
    return points, breakdown


def score_head_coach(game_info: dict, spread_info: dict = None) -> Tuple[float, Dict[str, float]]:
    """
    Score a head coach using OPFL rules.
    
    Scoring (based on ESPN spread):
        - Home favorite win: 4 pts
        - Road favorite win: 5 pts
        - Home underdog win: 6 pts
        - Road underdog win: 7 pts
        - Loss: 0 pts
    
    Args:
        game_info: Dict with team_score, opponent_score, is_home
        spread_info: Dict with 'spread' (from nflverse spread_line)
    """
    points = 0.0
    breakdown = {}
    
    team_score = game_info.get('team_score', 0) or 0
    opponent_score = game_info.get('opponent_score', 0) or 0
    is_home = game_info.get('is_home', True)
    
    # Determine if win
    if team_score > opponent_score:
        # Win - determine points based on home/away and favorite/underdog
        spread = 0
        if spread_info:
            spread = spread_info.get('spread', 0) or 0
        
        # nflverse spread_line convention (from home team's perspective):
        # - Positive spread = home team is FAVORITE (giving points)
        # - Negative spread = home team is UNDERDOG (getting points)
        # For away team, we negate the spread, so:
        # - Negative spread (for away) = away team is FAVORITE
        # - Positive spread (for away) = away team is UNDERDOG
        # If no spread available (0), assume favorite (more common for winning teams)
        is_favorite = spread > 0 or spread == 0
        is_underdog = spread < 0
        
        if is_home:
            if is_underdog:
                points = 6  # Home underdog win
                breakdown['home_underdog_win'] = 6
            else:
                points = 4  # Home favorite win
                breakdown['home_favorite_win'] = 4
        else:
            if is_underdog:
                points = 7  # Road underdog win
                breakdown['road_underdog_win'] = 7
            else:
                points = 5  # Road favorite win
                breakdown['road_favorite_win'] = 5
    else:
        # Loss or tie = 0 points
        breakdown['loss'] = 0
    
    return points, breakdown
