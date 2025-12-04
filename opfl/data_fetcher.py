"""NFL data fetching using nflreadpy with fuzzy matching support."""

import re
from typing import Optional, List

import polars as pl

try:
    import nflreadpy as nfl
except ImportError:
    raise ImportError("Please install nflreadpy: pip install nflreadpy")

try:
    from thefuzz import fuzz, process
    FUZZY_AVAILABLE = True
except ImportError:
    FUZZY_AVAILABLE = False

from .constants import TEAM_ABBREV_NORMALIZE

# Offensive line positions
OL_POSITIONS = {'T', 'G', 'C', 'OT', 'OG', 'OL', 'LT', 'RT', 'LG', 'RG'}


def normalize_name(name: str) -> str:
    """
    Normalize a player name for matching.
    
    Removes suffixes, converts to lowercase, and removes extra whitespace.
    """
    # Remove common suffixes
    name = re.sub(r'\s+(Sr\.?|Jr\.?|II|III|IV|V)$', '', name.strip(), flags=re.IGNORECASE)
    # Remove periods and extra whitespace
    name = re.sub(r'\.', '', name)
    name = re.sub(r'\s+', ' ', name)
    return name.lower().strip()


def fuzzy_match_name(
    query: str,
    candidates: List[str],
    threshold: int = 80,
) -> Optional[str]:
    """
    Find the best fuzzy match for a name among candidates.
    
    Args:
        query: The name to search for
        candidates: List of candidate names to match against
        threshold: Minimum score (0-100) to consider a match
        
    Returns:
        Best matching name or None if no match above threshold
    """
    if not FUZZY_AVAILABLE or not candidates:
        return None
    
    # Use token_sort_ratio for better matching with name order variations
    result = process.extractOne(
        query, 
        candidates,
        scorer=fuzz.token_sort_ratio,
    )
    
    if result and result[1] >= threshold:
        return result[0]
    
    return None


class NFLDataFetcher:
    """Fetches and caches NFL stats from nflreadpy with fuzzy matching support."""
    
    def __init__(self, season: int, week: int):
        self.season = season
        self.week = week
        self._player_stats: Optional[pl.DataFrame] = None
        self._team_stats: Optional[pl.DataFrame] = None
        self._schedules: Optional[pl.DataFrame] = None
        self._pbp: Optional[pl.DataFrame] = None
        self._players_db: Optional[pl.DataFrame] = None
        self._player_name_cache: dict = {}  # Cache for player name lookups
    
    @property
    def player_stats(self) -> pl.DataFrame:
        """Lazy load player stats."""
        if self._player_stats is None:
            print(f"Loading player stats for {self.season} week {self.week}...")
            stats = nfl.load_player_stats(seasons=self.season, summary_level='week')
            self._player_stats = stats.filter(pl.col('week') == self.week)
        return self._player_stats
    
    @property
    def team_stats(self) -> pl.DataFrame:
        """Lazy load team stats."""
        if self._team_stats is None:
            print(f"Loading team stats for {self.season} week {self.week}...")
            stats = nfl.load_team_stats(seasons=self.season, summary_level='week')
            self._team_stats = stats.filter(pl.col('week') == self.week)
        return self._team_stats
    
    @property
    def schedules(self) -> pl.DataFrame:
        """Lazy load schedules."""
        if self._schedules is None:
            print(f"Loading schedules for {self.season}...")
            schedules = nfl.load_schedules(seasons=self.season)
            self._schedules = schedules.filter(pl.col('week') == self.week)
        return self._schedules
    
    @property
    def pbp(self) -> pl.DataFrame:
        """Lazy load play-by-play data."""
        if self._pbp is None:
            print(f"Loading play-by-play for {self.season} week {self.week}...")
            pbp = nfl.load_pbp(seasons=self.season)
            self._pbp = pbp.filter(pl.col('week') == self.week)
        return self._pbp
    
    @property
    def players_db(self) -> pl.DataFrame:
        """Lazy load players database."""
        if self._players_db is None:
            self._players_db = nfl.load_players()
        return self._players_db
    
    def _normalize_team(self, team: str) -> str:
        """Normalize team abbreviation to nflreadpy format."""
        if not team:
            return team
        return TEAM_ABBREV_NORMALIZE.get(team.upper(), team.upper())
    
    def _get_all_player_names(self, team: str = None) -> List[str]:
        """Get all player display names, optionally filtered by team."""
        stats = self.player_stats
        
        if team:
            normalized_team = self._normalize_team(team)
            stats = stats.filter(pl.col('team') == normalized_team)
        
        if 'player_display_name' in stats.columns:
            names = stats['player_display_name'].drop_nulls().to_list()
            return [str(n) for n in names]
        return []
    
    def find_player(
        self,
        name: str,
        team: str,
        position: str,
        use_fuzzy: bool = True,
        fuzzy_threshold: int = 75,
    ) -> Optional[dict]:
        """
        Find a player in the stats by name matching with fuzzy search support.
        
        Args:
            name: Player name from Excel (e.g., "Patrick Mahomes II")
            team: Team abbreviation (e.g., "KC")
            position: Position (e.g., "QB")
            use_fuzzy: Whether to use fuzzy matching if exact match fails
            fuzzy_threshold: Minimum fuzzy match score (0-100)
            
        Returns:
            Dict of player stats or None if not found
        """
        stats = self.player_stats
        
        # Clean up name
        clean_name = normalize_name(name)
        normalized_team = self._normalize_team(team)
        
        # Check cache first
        cache_key = (clean_name, normalized_team)
        if cache_key in self._player_name_cache:
            cached_name = self._player_name_cache[cache_key]
            if cached_name is None:
                return None
            # Look up by cached name
            matches = stats.filter(
                pl.col('player_display_name') == cached_name
            )
            if matches.height > 0:
                return matches.row(0, named=True)
        
        # Filter by team first if provided (but also search all teams for fuzzy)
        if normalized_team:
            team_stats = stats.filter(pl.col('team') == normalized_team)
        else:
            team_stats = stats
        
        # Try exact match on display name (case insensitive, remove periods for comparison)
        # This handles cases like "A.J. Brown" vs "AJ Brown"
        matches = team_stats.filter(
            pl.col('player_display_name').str.to_lowercase().str.replace_all(r'\.', '') == clean_name
        )
        if matches.height > 0:
            result = matches.row(0, named=True)
            self._player_name_cache[cache_key] = result.get('player_display_name')
            return result
        
        # Try contains match (also remove periods)
        matches = team_stats.filter(
            pl.col('player_display_name').str.to_lowercase().str.replace_all(r'\.', '').str.contains(clean_name)
        )
        if matches.height > 0:
            result = matches.row(0, named=True)
            self._player_name_cache[cache_key] = result.get('player_display_name')
            return result
        
        # Try matching just last name
        name_parts = name.split()
        if len(name_parts) >= 2:
            last_name = name_parts[-1].lower()
            matches = team_stats.filter(
                pl.col('player_display_name').str.to_lowercase().str.replace_all(r'\.', '').str.contains(last_name)
            )
            if matches.height == 1:
                result = matches.row(0, named=True)
                self._player_name_cache[cache_key] = result.get('player_display_name')
                return result
        
        # Use fuzzy matching if enabled and available
        if use_fuzzy and FUZZY_AVAILABLE:
            # First try with team filter
            candidates = self._get_all_player_names(team=normalized_team if normalized_team else None)
            
            if candidates:
                best_match = fuzzy_match_name(name, candidates, threshold=fuzzy_threshold)
                
                if best_match:
                    matches = stats.filter(
                        pl.col('player_display_name') == best_match
                    )
                    if matches.height > 0:
                        result = matches.row(0, named=True)
                        self._player_name_cache[cache_key] = result.get('player_display_name')
                        return result
            
            # If team filter didn't work, try all players (OPFL often has wrong team)
            if normalized_team:
                all_candidates = self._get_all_player_names(team=None)
                if all_candidates:
                    best_match = fuzzy_match_name(name, all_candidates, threshold=fuzzy_threshold)
                    
                    if best_match:
                        matches = stats.filter(
                            pl.col('player_display_name') == best_match
                        )
                        if matches.height > 0:
                            result = matches.row(0, named=True)
                            self._player_name_cache[cache_key] = result.get('player_display_name')
                            # Note: We found the player but on a different team
                            return result
        
        # No match found
        self._player_name_cache[cache_key] = None
        return None
    
    def get_team_stats(self, team: str) -> Optional[dict]:
        """Get team stats for D/ST scoring."""
        normalized_team = self._normalize_team(team)
        team_data = self.team_stats.filter(pl.col('team') == normalized_team)
        
        if team_data.height > 0:
            return team_data.row(0, named=True)
        return None
    
    def get_opponent_stats(self, team: str) -> Optional[dict]:
        """Get opponent's team stats (for D/ST scoring)."""
        game = self.get_game_info(team)
        if not game:
            return None
        
        opponent = game.get('opponent')
        if not opponent:
            return None
        
        return self.get_team_stats(opponent)
    
    def get_game_info(self, team: str) -> Optional[dict]:
        """Get game information for a team."""
        normalized_team = self._normalize_team(team)
        schedules = self.schedules
        
        # Check if home team
        home_game = schedules.filter(pl.col('home_team') == normalized_team)
        if home_game.height > 0:
            row = home_game.row(0, named=True)
            if row.get('home_score') is None:
                return None  # Game hasn't been played yet
            return {
                'team_score': row.get('home_score', 0),
                'opponent_score': row.get('away_score', 0),
                'points_allowed': row.get('away_score', 0),
                'opponent': row.get('away_team'),
                'coach': row.get('home_coach'),
                'is_home': True,
                'spread': row.get('spread_line'),  # Home team spread
            }
        
        # Check if away team
        away_game = schedules.filter(pl.col('away_team') == normalized_team)
        if away_game.height > 0:
            row = away_game.row(0, named=True)
            if row.get('away_score') is None:
                return None  # Game hasn't been played yet
            # For away team, spread is the negative of home spread
            home_spread = row.get('spread_line')
            away_spread = -home_spread if home_spread is not None else None
            return {
                'team_score': row.get('away_score', 0),
                'opponent_score': row.get('home_score', 0),
                'points_allowed': row.get('home_score', 0),
                'opponent': row.get('home_team'),
                'coach': row.get('away_coach'),
                'is_home': False,
                'spread': away_spread,
            }
        
        return None
    
    def get_spread_info(self, team: str) -> Optional[dict]:
        """Get betting spread information for head coach scoring.
        
        nflverse spread_line convention (from home team's perspective):
        - Positive = home team is FAVORITE (giving points)
        - Negative = home team is UNDERDOG (getting points)
        
        For away team, we negate this, so:
        - Negative = away team is FAVORITE
        - Positive = away team is UNDERDOG
        """
        game = self.get_game_info(team)
        if not game:
            return None
        
        spread = game.get('spread')
        if spread is None:
            return None
        
        return {
            'spread': spread,
            'is_favorite': spread > 0,
            'is_underdog': spread < 0,
        }
    
    def get_turnovers_returned_for_td(self, player_id: str) -> dict:
        """
        Get count of turnovers returned for TDs by this player.
        
        Returns dict with:
            - pick_sixes: number of interceptions returned for TD
            - fumble_sixes: number of fumbles returned for TD
        """
        pbp = self.pbp
        
        # Pick sixes (interceptions returned for TD where this player threw the INT)
        pick_sixes = pbp.filter(
            (pl.col('interception') == 1) & 
            (pl.col('return_touchdown') == 1) &
            (pl.col('passer_player_id') == player_id)
        ).height
        
        # Fumble sixes (fumbles returned for TD where this player fumbled)
        fumble_sixes = pbp.filter(
            (pl.col('fumble_lost') == 1) & 
            (pl.col('return_touchdown') == 1) &
            (pl.col('fumbled_1_player_id') == player_id)
        ).height
        
        return {
            'pick_sixes': pick_sixes,
            'fumble_sixes': fumble_sixes,
        }
    
    def get_extra_fumbles_lost(self, player_id: str, player_stats: dict) -> int:
        """
        Get fumbles lost from PBP that aren't in player stats.
        
        This catches fumbles on laterals and other plays that don't get
        attributed to the player in the standard stats.
        
        Args:
            player_id: Player's NFL ID
            player_stats: Player's stats dict (to compare against)
            
        Returns:
            Number of additional fumbles lost not in player stats
        """
        pbp = self.pbp
        
        # Count fumbles lost where this player fumbled (from PBP)
        pbp_fumbles = pbp.filter(
            (pl.col('fumble_lost') == 1) &
            (pl.col('fumbled_1_player_id') == player_id)
        ).height
        
        # Count fumbles in player stats
        stats_fumbles = (
            (player_stats.get('sack_fumbles_lost', 0) or 0) +
            (player_stats.get('rushing_fumbles_lost', 0) or 0) +
            (player_stats.get('receiving_fumbles_lost', 0) or 0)
        )
        
        # Extra fumbles = PBP fumbles not in stats
        extra = max(0, pbp_fumbles - stats_fumbles)
        return extra
    
    def get_defensive_sacks(self, team: str) -> dict:
        """
        Get sack count from both aggregated stats and play-by-play.
        
        The team_stats 'def_sacks' column can undercount sacks, so we
        also count directly from PBP for accuracy.
        
        Args:
            team: Team abbreviation (e.g., 'KC')
            
        Returns:
            Dict with 'aggregated', 'pbp', 'value' (the one to use), and 'discrepancy' flag
        """
        normalized_team = self._normalize_team(team)
        
        # Get aggregated stats sacks
        team_data = self.team_stats.filter(pl.col('team') == normalized_team)
        agg_sacks = int(team_data['def_sacks'][0]) if team_data.height > 0 else 0
        
        # Count from PBP
        pbp = self.pbp
        pbp_sacks = pbp.filter(
            (pl.col('defteam') == normalized_team) &
            (pl.col('sack') == 1)
        ).height
        
        # Use PBP if different (more accurate)
        discrepancy = agg_sacks != pbp_sacks
        
        return {
            'aggregated': agg_sacks,
            'pbp': pbp_sacks,
            'value': pbp_sacks if discrepancy else agg_sacks,
            'discrepancy': discrepancy,
        }
    
    def get_blocked_punts(self, team: str) -> int:
        """
        Get count of blocked punts by a defense from play-by-play.
        
        Args:
            team: Team abbreviation (e.g., 'HOU')
            
        Returns:
            Number of punts blocked by this team's defense
        """
        normalized_team = self._normalize_team(team)
        pbp = self.pbp
        
        # Count plays where punt was blocked and this team was on defense
        blocked_punts = pbp.filter(
            (pl.col('defteam') == normalized_team) &
            (pl.col('punt_blocked') == 1)
        ).height
        
        return blocked_punts
    
    def get_blocked_kick_tds(self, team: str) -> int:
        """
        Get count of TDs scored on blocked punts/FGs by a defense from play-by-play.
        
        These are categorized as special_teams_tds in nflverse, but per OPFL rules
        they should count as defensive TDs (4 pts each), while regular punt/kick
        return TDs do not count.
        
        Args:
            team: Team abbreviation (e.g., 'PHI')
            
        Returns:
            Number of blocked punt/FG TDs by this team's defense
        """
        normalized_team = self._normalize_team(team)
        pbp = self.pbp
        
        # Count blocked punt TDs where this team scored
        blocked_punt_tds = pbp.filter(
            (pl.col('punt_blocked') == 1) &
            (pl.col('touchdown') == 1) &
            (pl.col('td_team') == normalized_team)
        ).height
        
        # Count blocked FG TDs where this team scored (rare but possible)
        # Note: nflverse uses 'field_goal_attempt' and we check if it was blocked
        # A blocked FG returned for TD would have the scoring team as td_team
        blocked_fg_tds = pbp.filter(
            (pl.col('field_goal_attempt') == 1) &
            (pl.col('touchdown') == 1) &
            (pl.col('td_team') == normalized_team) &
            (pl.col('defteam') == normalized_team)  # Defense scored, so FG was blocked
        ).height
        
        return blocked_punt_tds + blocked_fg_tds
    
    def find_coach(self, coach_name: str, team: str = None) -> Optional[dict]:
        """
        Find a coach's team from the schedule data.
        
        Args:
            coach_name: Coach name from Excel
            team: Optional team abbreviation hint
            
        Returns:
            Game info dict if found, None otherwise
        """
        schedules = self.schedules
        clean_name = coach_name.lower().strip()
        
        # Check home coaches
        for row in schedules.iter_rows(named=True):
            home_coach = row.get('home_coach', '') or ''
            if clean_name in home_coach.lower():
                return {
                    'team': row.get('home_team'),
                    'team_score': row.get('home_score', 0),
                    'opponent_score': row.get('away_score', 0),
                    'is_home': True,
                    'spread': row.get('spread_line'),
                }
            
            away_coach = row.get('away_coach', '') or ''
            if clean_name in away_coach.lower():
                home_spread = row.get('spread_line')
                return {
                    'team': row.get('away_team'),
                    'team_score': row.get('away_score', 0),
                    'opponent_score': row.get('home_score', 0),
                    'is_home': False,
                    'spread': -home_spread if home_spread is not None else None,
                }
        
        return None
