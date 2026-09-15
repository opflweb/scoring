#!/usr/bin/env python3
"""Print an NFL week number, for the scoring workflow to consume.

Replaces an inline Python heredoc in .github/workflows/score.yml — logic in a
YAML string cannot be tested or run locally.

    python scripts/nfl_week.py --season 2026            # current week
    python scripts/nfl_week.py --season 2026 --latest-completed
"""

import argparse
import sys
from pathlib import Path

import nflreadpy as nfl

sys.path.insert(0, str(Path(__file__).parent.parent))

from opfl.week_status import latest_completed_week


def main():
    parser = argparse.ArgumentParser(description='Print an NFL week number')
    parser.add_argument('--season', '-y', type=int, required=True)
    parser.add_argument(
        '--latest-completed',
        action='store_true',
        help='Print the last week whose NFL games are all final (0 if none)',
    )
    parser.add_argument('--max-week', type=int, default=17)
    args = parser.parse_args()

    if args.latest_completed:
        rows = list(nfl.load_schedules(seasons=args.season).iter_rows(named=True))
        print(latest_completed_week(rows, max_week=args.max_week))
        return 0

    print(min(nfl.get_current_week(), args.max_week))
    return 0


if __name__ == '__main__':
    sys.exit(main())
