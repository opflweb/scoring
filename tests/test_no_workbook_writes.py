"""The workbook is read-only. Forever.

`OPFL Scoring 2026.xlsx` is formula-driven end to end: the Matchups tab pulls
names and totals from the Scoring tab, which pulls each team's starred players
off the Rosters tab. openpyxl cannot evaluate formulas, so saving the workbook
strips every cached value and leaves the file unreadable to this pipeline until
someone opens it in Excel and lets it recalculate.

That is easy to reintroduce by accident — `wb.save()` is one line — so it is
pinned here rather than left as tribal knowledge.
"""

import hashlib
import re
import shutil
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
WORKBOOK = PROJECT_ROOT / 'OPFL Scoring 2026.xlsx'

# The one sanctioned writer: a legacy helper for the retired per-week W-sheet
# format, which is never reached by the current pipeline.
SANCTIONED_SAVE_SITES = {'opfl/excel_parser.py'}


def _digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


@pytest.mark.skipif(not WORKBOOK.exists(), reason='workbook not present')
def test_parsing_the_workbook_never_modifies_it(tmp_path):
    """Every read path leaves the file byte-identical."""
    from opfl.excel_parser import (
        parse_matchups_sheet,
        parse_roster_from_excel,
        parse_taxi_squads,
    )
    from opfl.scorer import build_matchup_week

    workbook = tmp_path / 'workbook.xlsx'
    shutil.copy(WORKBOOK, workbook)
    before = _digest(workbook)

    parse_roster_from_excel(str(workbook), 'Rosters')
    parse_matchups_sheet(str(workbook), 'Matchups')
    parse_taxi_squads(str(workbook), 'Rosters')
    build_matchup_week(str(workbook))

    assert _digest(workbook) == before, (
        'Parsing modified the workbook. openpyxl saves strip cached formula '
        'values — nothing in the read path may call wb.save().'
    )


@pytest.mark.skipif(not WORKBOOK.exists(), reason='workbook not present')
def test_cached_formula_values_survive_parsing(tmp_path):
    """The symptom to catch: a saved workbook reads back as all-None."""
    import openpyxl

    workbook = tmp_path / 'workbook.xlsx'
    shutil.copy(WORKBOOK, workbook)

    from opfl.scorer import build_matchup_week

    build_matchup_week(str(workbook))

    wb = openpyxl.load_workbook(workbook, data_only=True)
    matchups = wb['Matchups']
    # A1 is a formula; data_only gives us Excel's cached result for it.
    assert matchups.cell(row=1, column=1).value, (
        'Cached formula values were lost — the Matchups tab now reads as empty.'
    )
    wb.close()


def test_no_unsanctioned_workbook_saves_in_the_codebase():
    """Grep-level guard so a new `wb.save()` fails review rather than production."""
    offenders = []
    for path in list(PROJECT_ROOT.glob('opfl/*.py')) + list(PROJECT_ROOT.glob('scripts/*.py')):
        rel = path.relative_to(PROJECT_ROOT).as_posix()
        if rel in SANCTIONED_SAVE_SITES:
            continue
        if re.search(r'\.save\(', path.read_text()):
            offenders.append(rel)

    assert not offenders, (
        f'Workbook save detected in {offenders}. The OPFL workbook is formula-driven '
        'and must never be written by openpyxl; publish to web/data.json instead.'
    )
