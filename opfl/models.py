"""Data models for OPFL autoscorer."""

from dataclasses import dataclass, field


@dataclass
class PlayerScore:
    """Container for a player's score breakdown."""

    name: str
    position: str
    team: str
    total_points: float = 0.0
    breakdown: dict[str, float] = field(default_factory=dict)
    found_in_stats: bool = False
    data_notes: list[str] = field(default_factory=list)  # Flags for data discrepancies
    matched_name: str = ''  # The name that was actually matched in stats (for fuzzy matches)
    is_starter: bool = True  # Whether this player is a starter


@dataclass
class FantasyTeam:
    """Container for a fantasy team's roster."""

    name: str
    owner: str
    abbreviation: str
    column_index: int  # 1-based column index in Excel
    players: dict[str, list[tuple[str, str, bool]]] = field(default_factory=dict)
    # players[position] = [(player_name, nfl_team, is_started), ...]
