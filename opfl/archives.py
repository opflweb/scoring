"""Extraction of OPFL summary history, drafts, trades, rules, and banners."""

from __future__ import annotations

import re
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl


def extract_history(path: str | Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    workbook = openpyxl.load_workbook(path, data_only=True)
    try:
        records = workbook["W-L Records"]
        playoffs = workbook["Playoffs"]
        owners = [
            (column, str(records.cell(1, column).value).strip())
            for column in range(3, 103, 5)
            if records.cell(1, column).value
        ]
        finishes: dict[int, dict[str, Any]] = {}
        for row in range(4, 41):
            year_value = playoffs.cell(row, 1).value
            if not isinstance(year_value, int):
                continue
            finishes[year_value] = {
                "year": year_value,
                "first": _text(playoffs.cell(row, 2).value),
                "second": _text(playoffs.cell(row, 3).value),
                "third": _text(playoffs.cell(row, 4).value),
                "fourth": _text(playoffs.cell(row, 5).value),
                "championship_score": _text(playoffs.cell(row, 6).value),
                "toilet_bowl_score": _text(playoffs.cell(row, 7).value),
                "jamboree": _text(playoffs.cell(row, 8).value),
                "comments": " ".join(
                    value
                    for value in (
                        _text(playoffs.cell(row, 9).value),
                        _text(playoffs.cell(row, 10).value),
                    )
                    if value
                ),
                "team_count": int(playoffs.cell(row, 11).value),
            }
        seasons: list[dict[str, Any]] = []
        for row in range(3, 40):
            year_value = records.cell(row, 1).value
            if not isinstance(year_value, int):
                continue
            standings: list[dict[str, Any]] = []
            for column, owner in owners:
                wins = records.cell(row, column).value
                losses = records.cell(row, column + 1).value
                ties = records.cell(row, column + 2).value
                top_six = records.cell(row, column + 3).value
                points_for = records.cell(row, column + 4).value
                if wins is None and points_for is None:
                    continue
                wins_value = _number(wins)
                ties_value = _number(ties)
                top_six_value = _optional_number(top_six)
                standings.append(
                    {
                        "name": owner,
                        "owner": owner,
                        "wins": wins_value,
                        "losses": _number(losses),
                        "ties": ties_value,
                        "top_six": top_six_value,
                        "rank_points": (
                            round(wins_value + ties_value * 0.5 + top_six_value, 3)
                            if top_six_value is not None
                            else None
                        ),
                        "points_for": _number(points_for),
                    }
                )
            if any(item["rank_points"] is not None for item in standings):
                standings.sort(
                    key=lambda item: (
                        -(item["rank_points"] or 0),
                        -item["wins"],
                        -(item["top_six"] or 0),
                        -item["points_for"],
                        item["name"],
                    )
                )
            else:
                standings.sort(
                    key=lambda item: (
                        -(item["wins"] + 0.5 * item["ties"]),
                        -item["points_for"],
                        item["name"],
                    )
                )
            for rank, item in enumerate(standings, 1):
                item["rank"] = rank
            seasons.append(
                {
                    "year": year_value,
                    "detail_level": "summary",
                    "standings": standings,
                    "finish": finishes[year_value],
                }
            )
        owner_summaries = _owner_summaries(seasons)
        return seasons, {"seasons": [item["finish"] for item in seasons], "owners": owner_summaries}
    finally:
        workbook.close()


def extract_drafts_trades_picks(path: str | Path) -> tuple[list[Any], list[Any], dict[str, Any]]:
    workbook = openpyxl.load_workbook(path, data_only=True)
    try:
        drafts = [
            _extract_draft_board(workbook[sheet], sheet)
            for sheet in workbook.sheetnames
            if re.fullmatch(r"20(?:20|21|22|23|24|25) (?:Preseason|Waiver)", sheet)
        ]
        drafts.sort(key=lambda item: (item["year"], item["type"]))
        trades = []
        sheet = workbook["Trade History"]
        for row in range(5, 38):
            value = sheet.cell(row, 1).value
            if not isinstance(value, (date, datetime)):
                continue
            trades.append(
                {
                    "date": value.date().isoformat()
                    if isinstance(value, datetime)
                    else value.isoformat(),
                    "teams": [_text(sheet.cell(row, 3).value), _text(sheet.cell(row, 5).value)],
                    "team_a": _text(sheet.cell(row, 3).value),
                    "team_a_receives": _lines(sheet.cell(row, 4).value),
                    "team_b": _text(sheet.cell(row, 5).value),
                    "team_b_receives": _lines(sheet.cell(row, 6).value),
                    "notes": _text(sheet.cell(row, 7).value),
                }
            )
        future_notes = [
            _text(workbook["Future Traded Picks"].cell(row, 2).value)
            for row in range(3, 75)
            if workbook["Future Traded Picks"].cell(row, 2).value
        ]
        return (
            drafts,
            trades,
            {"notes": future_notes, "holdings": _parse_pick_holdings(future_notes)},
        )
    finally:
        workbook.close()


def rules_archive() -> dict[str, Any]:
    return {
        "title": "OPFL Scoring Rules",
        "sections": [
            {
                "title": "General",
                "rules": [
                    "Touchdowns and touchdown passes are 6 points; two-point conversions are 2.",
                    "Individual player scores cannot be less than zero.",
                    "Special-teams touchdowns count for the individual player.",
                ],
            },
            {
                "title": "Passing",
                "rules": [
                    "200 yards earns 2 points, plus 1 for each additional 50 yards.",
                    "Interceptions and fumbles lost are -1; a turnover returned for a touchdown is -3 total.",
                ],
            },
            {
                "title": "Rushing and receiving",
                "rules": [
                    "RB/WR: 75 yards in either category earns 2 points, plus 1 per additional 25 yards.",
                    "RB/WR alternate: 100 combined yards earns 2 points, plus 1 per additional 25 yards.",
                    "TE: 50 yards in either category earns 2 points, plus 1 per additional 25 yards.",
                    "TE alternate: 75 combined yards earns 2 points, plus 1 per additional 25 yards.",
                    "The combined bonus is an alternative; use whichever yardage method scores more.",
                ],
            },
            {
                "title": "Kicker",
                "rules": [
                    "PAT 1; field goals 1–29 yards 1, 30–39 yards 2, 40–49 yards 3, 50+ yards 4.",
                    "Missed or blocked field goals are -2; missed or blocked PATs are -1.",
                ],
            },
            {
                "title": "Defense",
                "rules": [
                    "Interceptions, fumble recoveries, safeties, blocked punts, and blocked field goals are 2; sacks and blocked PATs are 1; eligible defensive touchdowns are 4.",
                    "Points allowed: 0=8, 2–9=6, 10–13=4, 14–17=2, 18–27=0, 28–31=-2, 32–35=-4, 36+=-6.",
                    "Punt- and kickoff-return touchdowns do not count for DF scoring.",
                ],
            },
            {
                "title": "Head coach",
                "rules": [
                    "Using the ESPN spread: home favorite win 4, home underdog win 5, road favorite win 6, road underdog win 7; loss or tie 0."
                ],
            },
            {
                "title": "League format",
                "rules": [
                    "Rank Points = matchup wins + Top-6 points + 0.5 × matchup ties.",
                    "A tie crossing the weekly sixth-place cutoff splits the available Top-6 points evenly.",
                    "Four teams reach the playoffs; the other eight play a two-week Jamboree in Weeks 16–17.",
                ],
            },
        ],
    }


def _extract_draft_board(sheet: Any, name: str) -> dict[str, Any]:
    year = int(name[:4])
    board_type = name.split()[1].lower()
    picks: list[dict[str, Any]] = []
    for row in range(1, min(sheet.max_row, 40) + 1):
        for column in (1, 5, 9):
            overall = sheet.cell(row, column).value
            owner = sheet.cell(row, column + 1).value
            if isinstance(overall, (int, float)) and owner:
                picks.append(
                    {
                        "overall": int(overall),
                        "round": (int(overall) - 1) // 12 + 1,
                        "slot": (int(overall) - 1) % 12 + 1,
                        "owner": _text(owner),
                        "selection": _text(sheet.cell(row, column + 2).value),
                        "drop": _text(sheet.cell(row, column + 3).value),
                        "taxi": False,
                    }
                )
    header_row = next(
        (
            row
            for row in range(1, 8)
            if any("Taxi Round" in _text(sheet.cell(row, column).value) for column in range(1, 23))
        ),
        None,
    )
    if header_row:
        for column in range(1, 23):
            header = _text(sheet.cell(header_row, column).value)
            match = re.match(r"Taxi Round\s+(\d+)", header, re.IGNORECASE)
            if not match:
                continue
            taxi_round = int(match.group(1))
            for slot, row in enumerate(range(header_row + 1, header_row + 13), 1):
                owner = _text(sheet.cell(row, column).value)
                if not owner:
                    continue
                picks.append(
                    {
                        "overall": None,
                        "round": taxi_round,
                        "slot": slot,
                        "owner": owner,
                        "selection": _text(sheet.cell(row, column + 1).value),
                        "drop": "",
                        "taxi": True,
                    }
                )
    return {
        "year": year,
        "type": board_type,
        "title": _text(sheet.cell(1, 1).value),
        "date": _text(sheet.cell(2, 1).value),
        "picks": picks,
    }


def _parse_pick_holdings(notes: list[str]) -> list[dict[str, Any]]:
    holdings = []
    pattern = re.compile(
        r"^(?P<holder>.+?) holds? (?P<original>.+?)'s (?P<year>20\d{2}) #(?P<round>\d+) pick$",
        re.IGNORECASE,
    )
    for note in notes:
        match = pattern.match(note)
        if match:
            holdings.append(
                {
                    "holder": match.group("holder"),
                    "original_owner": match.group("original"),
                    "year": int(match.group("year")),
                    "round": int(match.group("round")),
                    "note": note,
                }
            )
    return holdings


def _owner_summaries(seasons: list[dict[str, Any]]) -> list[dict[str, Any]]:
    summary: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "seasons": 0,
            "wins": 0.0,
            "losses": 0.0,
            "ties": 0.0,
            "points_for": 0.0,
            "titles": 0.0,
        }
    )
    for season in seasons:
        for row in season["standings"]:
            owner = row["owner"]
            summary[owner]["seasons"] += 1
            for field in ("wins", "losses", "ties", "points_for"):
                summary[owner][field] += row[field]
        champions = [part.strip() for part in season["finish"]["first"].split(" & ")]
        for champion in champions:
            summary[champion]["titles"] += 1 / len(champions)
    return [
        {"owner": owner, **values}
        for owner, values in sorted(
            summary.items(), key=lambda item: (-item[1]["titles"], -item[1]["wins"], item[0])
        )
    ]


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _lines(value: object) -> list[str]:
    return [line.strip() for line in _text(value).splitlines() if line.strip()]


def _number(value: object) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float, str)):
        return float(value)
    raise TypeError(f"Expected a number, got {type(value).__name__}")


def _optional_number(value: object) -> float | None:
    return None if value in (None, "") else _number(value)
