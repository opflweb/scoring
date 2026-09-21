#!/usr/bin/env python3
"""Import OPFL history and shared archives. Dry-run unless --apply is supplied."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from opfl.archives import extract_drafts_trades_picks, extract_history, rules_archive


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def build_archives(history_path: Path, activity_path: Path, banners_dir: Path) -> dict[str, Any]:
    seasons, history = extract_history(history_path)
    drafts, trades, future_picks = extract_drafts_trades_picks(activity_path)
    banners = []
    finish_by_year = {item["year"]: item for item in history["seasons"]}
    for path in sorted(banners_dir.glob("*.png"), reverse=True):
        if not path.stem.isdigit():
            continue
        year = int(path.stem)
        banners.append(
            {
                "year": year,
                "image": f"images/banners/{path.name}",
                "champion": finish_by_year.get(year, {}).get("first", ""),
            }
        )
    return {
        "seasons": seasons,
        "history": history,
        "drafts": {"drafts": drafts},
        "transactions": {"trades": trades},
        "future_picks": future_picks,
        "rules": rules_archive(),
        "banners": {"banners": banners},
    }


def apply_archives(payload: dict[str, Any], data_root: Path, replace_existing: bool) -> None:
    archive_root = data_root / "archive"
    targets = [archive_root / "seasons" / f"{item['year']}.json" for item in payload["seasons"]]
    targets.extend(
        archive_root / "shared" / name
        for name in (
            "history.json",
            "drafts.json",
            "transactions.json",
            "future-picks.json",
            "rules.json",
            "banners.json",
        )
    )
    existing = [path for path in targets if path.exists()]
    if existing and not replace_existing:
        raise FileExistsError(
            f"Refusing to overwrite {len(existing)} archive files; pass --replace-existing"
        )
    for season in payload["seasons"]:
        write_json(archive_root / "seasons" / f"{season['year']}.json", season)
    shared = archive_root / "shared"
    write_json(shared / "history.json", payload["history"])
    write_json(shared / "drafts.json", payload["drafts"])
    write_json(shared / "transactions.json", payload["transactions"])
    write_json(shared / "future-picks.json", payload["future_picks"])
    write_json(shared / "rules.json", payload["rules"])
    write_json(shared / "banners.json", payload["banners"])


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument(
        "--history", default="web/docs/Playoffs & W-L Records - 2024.xlsx", type=Path
    )
    result.add_argument("--activity", default="OPFL Draft & Future Traded Picks.xlsx", type=Path)
    result.add_argument("--banners", default="web/images/banners", type=Path)
    result.add_argument("--data-root", default="data", type=Path)
    result.add_argument("--apply", action="store_true")
    result.add_argument("--replace-existing", action="store_true")
    return result


def main() -> int:
    args = parser().parse_args()
    payload = build_archives(args.history, args.activity, args.banners)
    print(
        f"Validated {len(payload['seasons'])} summary seasons, "
        f"{len(payload['drafts']['drafts'])} drafts, "
        f"{len(payload['transactions']['trades'])} trades, and "
        f"{len(payload['banners']['banners'])} banners."
    )
    if not args.apply:
        print("Dry run: no files written. Pass --apply to import.")
        return 0
    apply_archives(payload, args.data_root, args.replace_existing)
    print(f"Wrote authoritative archives under {args.data_root / 'archive'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
