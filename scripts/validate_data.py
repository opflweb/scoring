#!/usr/bin/env python3
"""Validate JSON schemas and OPFL cross-file integrity."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from jsonschema import Draft202012Validator

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from opfl.integrity import (
    read_json,
    validate_archives,
    validate_public_tree,
    validate_source_season,
)


def validate_schema(instance_path: Path, schema_path: Path) -> None:
    validator = Draft202012Validator(read_json(schema_path))
    errors = sorted(
        validator.iter_errors(read_json(instance_path)), key=lambda error: list(error.path)
    )
    if errors:
        details = "; ".join(
            f"{'/'.join(map(str, error.path)) or '<root>'}: {error.message}" for error in errors
        )
        raise ValueError(f"Schema validation failed for {instance_path}: {details}")


def run_validation(
    data_root: Path,
    web_root: Path,
    schemas: Path,
    *,
    source_only: bool = False,
    public_only: bool = False,
) -> list[str]:
    results = []
    if not public_only:
        validate_schema(data_root / "league.json", schemas / "league.schema.json")
        results.append("schema: league.json")
        results.extend(validate_archives(data_root))
        for season_dir in sorted((data_root / "seasons").glob("*")):
            if season_dir.is_dir():
                validate_schema(season_dir / "metadata.json", schemas / "metadata.schema.json")
                for week_path in sorted((season_dir / "weeks").glob("week_*.json")):
                    validate_schema(week_path, schemas / "week.schema.json")
                results.extend(validate_source_season(season_dir, data_root.parent))
    if not source_only:
        validate_schema(web_root / "index.json", schemas / "public-index.schema.json")
        results.extend(validate_public_tree(web_root))
    return results


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--data-root", default="data", type=Path)
    result.add_argument("--web-root", default="web/data", type=Path)
    result.add_argument("--schemas", default="schemas", type=Path)
    scope = result.add_mutually_exclusive_group()
    scope.add_argument("--source-only", action="store_true")
    scope.add_argument("--public-only", action="store_true")
    return result


def main() -> int:
    args = parser().parse_args()
    for result in run_validation(
        args.data_root,
        args.web_root,
        args.schemas,
        source_only=args.source_only,
        public_only=args.public_only,
    ):
        print(f"✓ {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
