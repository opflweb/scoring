"""Every hardcoded season literal must agree with data/league_config.json.

The season used to be duplicated across scripts/export_for_web.py, autoscorer.py,
and .github/workflows/score.yml. Nothing enforced that they matched, so a season
rollover that missed one of them would silently score the wrong year. This test
turns that drift into a red CI run instead of a live bug next August, following
the pattern from the sibling QPFL repo's tests/test_config_consistency.py.
"""

import re
from pathlib import Path

from opfl.config import get_current_season

PROJECT_ROOT = Path(__file__).parent.parent


def test_workflow_season_matches_league_config():
    workflow = (PROJECT_ROOT / '.github' / 'workflows' / 'score.yml').read_text()
    match = re.search(r"SEASON:\s*'(\d{4})'", workflow)
    assert match, 'score.yml has no SEASON environment variable'
    assert int(match.group(1)) == get_current_season()


def test_workflow_workbook_matches_the_current_season():
    workflow = (PROJECT_ROOT / '.github' / 'workflows' / 'score.yml').read_text()
    match = re.search(r"WORKBOOK:\s*'OPFL Scoring (\d{4})\.xlsx'", workflow)
    assert match, 'score.yml has no WORKBOOK environment variable'
    assert int(match.group(1)) == get_current_season()


def test_autoscorer_default_season_matches_league_config():
    """autoscorer.py's --season default is derived from get_current_season() at
    import time; this pins that it stays derived rather than reverting to a
    hardcoded literal."""
    source = (PROJECT_ROOT / 'autoscorer.py').read_text()
    assert 'get_current_season()' in source
    assert re.search(r"default=\d{4},\s*\n\s*help='NFL season year'", source) is None
