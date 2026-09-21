#!/usr/bin/env python3
"""Generate the split GitHub Pages data tree from authoritative JSON."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from opfl.exporting import export_public_data


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--data-root", default="data", type=Path)
    result.add_argument("--output", default="web/data", type=Path)
    result.add_argument("--compatibility", default="web/data.json", type=Path)
    return result


def main() -> int:
    args = parser().parse_args()
    index = export_public_data(args.data_root, args.output, args.compatibility)
    print(
        f"Exported {len(index['seasons'])} public seasons; current season "
        f"is {index['current_season']}."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
