"""Main scoring engine that ties everything together for OPFL."""

from typing import Dict, List, Tuple

from .models import PlayerScore, FantasyTeam
from .data_fetcher import NFLDataFetcher
from .scoring import (
    score_qb,
    score_rb_wr,
    score_te,
    score_kicker,
    score_defense,
    score_head_coach,
)


class OPFLScorer:
    """Main scoring engine for OPFL fantasy football."""
    
    def __init__(self, season: int, week: int):
        self.season = season
        self.week = week
        self.data = NFLDataFetcher(season, week)
    
    def score_player(self, name: str, team: str, position: str) -> PlayerScore:
        """Score a single player using OPFL rules."""
        result = PlayerScore(name=name, position=position, team=team)
        
        if position == 'QB':
            stats = self.data.find_player(name, team, position)
            if stats:
                result.found_in_stats = True
                result.matched_name = stats.get('player_display_name', name)
                player_id = stats.get('player_id')
                turnover_tds = {}
                if player_id:
                    turnover_tds = self.data.get_turnovers_returned_for_td(player_id)
                result.total_points, result.breakdown = score_qb(stats, turnover_tds)
        
        elif position in ('RB', 'WR'):
            stats = self.data.find_player(name, team, position)
            if stats:
                result.found_in_stats = True
                result.matched_name = stats.get('player_display_name', name)
                player_id = stats.get('player_id')
                turnover_tds = {}
                if player_id:
                    turnover_tds = self.data.get_turnovers_returned_for_td(player_id)
                result.total_points, result.breakdown = score_rb_wr(stats, turnover_tds)
        
        elif position == 'TE':
            stats = self.data.find_player(name, team, position)
            if stats:
                result.found_in_stats = True
                result.matched_name = stats.get('player_display_name', name)
                player_id = stats.get('player_id')
                turnover_tds = {}
                if player_id:
                    turnover_tds = self.data.get_turnovers_returned_for_td(player_id)
                result.total_points, result.breakdown = score_te(stats, turnover_tds)
        
        elif position == 'K':
            stats = self.data.find_player(name, team, position)
            if stats:
                result.found_in_stats = True
                result.matched_name = stats.get('player_display_name', name)
                result.total_points, result.breakdown = score_kicker(stats)
        
        elif position == 'DF':  # Defense (OPFL uses DF, not D/ST)
            team_stats = self.data.get_team_stats(team)
            opponent_stats = self.data.get_opponent_stats(team) or {}
            game_info = self.data.get_game_info(team)
            
            if team_stats and game_info:
                result.found_in_stats = True
                # Get sack counts from both sources
                sack_info = self.data.get_defensive_sacks(team)
                if sack_info['discrepancy']:
                    result.data_notes.append(
                        f"Sack discrepancy: aggregated={sack_info['aggregated']}, PBP={sack_info['pbp']} (using PBP)"
                    )
                # Get blocked punts from PBP
                blocked_punts = self.data.get_blocked_punts(team)
                opponent_stats['_blocked_punts'] = blocked_punts
                
                # Get blocked kick TDs from PBP (special teams TDs on blocked punts/FGs)
                blocked_kick_tds = self.data.get_blocked_kick_tds(team)
                opponent_stats['_blocked_kick_tds'] = blocked_kick_tds
                
                result.total_points, result.breakdown = score_defense(
                    team_stats, opponent_stats, game_info, sack_info['value']
                )
        
        elif position == 'HC':
            # For HC, we need to find the coach's team
            # The name might be just the coach name, or coach name (team)
            game_info = self.data.get_game_info(team) if team else None
            spread_info = None
            
            if not game_info:
                # Try to find the coach by name
                coach_info = self.data.find_coach(name, team)
                if coach_info:
                    game_info = {
                        'team_score': coach_info.get('team_score', 0),
                        'opponent_score': coach_info.get('opponent_score', 0),
                        'is_home': coach_info.get('is_home', True),
                    }
                    spread_info = {
                        'spread': coach_info.get('spread', 0),
                    }
            else:
                spread_info = self.data.get_spread_info(team)
            
            if game_info:
                result.found_in_stats = True
                result.total_points, result.breakdown = score_head_coach(game_info, spread_info)
        
        return result
    
    def score_fantasy_team(self, team: FantasyTeam, starters_only: bool = False) -> Dict[str, List[PlayerScore]]:
        """Score players on a fantasy team.
        
        Args:
            team: The fantasy team to score
            starters_only: If True, only score started players. If False, score all players.
        """
        results = {}
        
        for position, players in team.players.items():
            results[position] = []
            
            for player_name, nfl_team, is_started in players:
                if starters_only and not is_started:
                    continue
                score = self.score_player(player_name, nfl_team, position)
                score.is_starter = is_started
                results[position].append(score)
        
        return results
    
    @staticmethod
    def calculate_team_total(scores: Dict[str, List[PlayerScore]], starters_only: bool = True) -> float:
        """Calculate total score for a fantasy team.
        
        Args:
            scores: Dict mapping position to list of PlayerScore objects
            starters_only: If True, only count starters. If False, count all players.
        """
        total = 0.0
        for position_scores in scores.values():
            for score in position_scores:
                if starters_only and not score.is_starter:
                    continue
                total += score.total_points
        return total


def score_week(
    excel_path: str,
    sheet_name: str,
    season: int,
    week: int,
    verbose: bool = True,
) -> Tuple[List[FantasyTeam], Dict[str, Tuple[float, Dict[str, List[PlayerScore]]]]]:
    """
    Score all fantasy teams for a given week.
    
    Returns:
        Tuple of (teams, results) where results maps team name to (total_score, position_scores)
    """
    from .excel_parser import parse_roster_from_excel
    
    teams = parse_roster_from_excel(excel_path, sheet_name)
    
    if verbose:
        print(f"\nFound {len(teams)} fantasy teams")
        for team in teams:
            started_count = sum(
                1 for players in team.players.values()
                for _, _, is_started in players if is_started
            )
            print(f"  - {team.name}: {started_count} started players")
    
    scorer = OPFLScorer(season, week)
    results = {}
    
    for team in teams:
        if verbose:
            print(f"\n{'='*60}")
            print(f"Scoring: {team.name}")
            print('='*60)
        
        scores = scorer.score_fantasy_team(team)
        total = scorer.calculate_team_total(scores)
        
        if verbose:
            for position, player_scores in scores.items():
                for ps in player_scores:
                    status = "✓" if ps.found_in_stats else "✗"
                    starter_marker = "*" if ps.is_starter else " "
                    matched = f" -> {ps.matched_name}" if ps.matched_name and ps.matched_name != ps.name else ""
                    print(f"  {starter_marker} {position} {ps.name} ({ps.team}){matched}: {ps.total_points:.1f} pts {status}")
                    if ps.breakdown:
                        for key, val in ps.breakdown.items():
                            if key != 'floor_applied':
                                print(f"        {key}: {val}")
                            else:
                                print(f"        (floor applied - points capped at 0)")
                    if ps.data_notes:
                        for note in ps.data_notes:
                            print(f"        ⚠️  {note}")
            
            print(f"\n  TOTAL: {total:.1f} points")
        
        results[team.name] = (total, scores)
    
    return teams, results
