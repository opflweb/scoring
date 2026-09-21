import json
import subprocess
from pathlib import Path

from opfl.integrity import validate_archives, validate_public_tree, validate_source_season

ROOT = Path(__file__).resolve().parent.parent


def test_authoritative_and_public_integrity() -> None:
    assert validate_source_season(ROOT / "data" / "seasons" / "2025", ROOT)
    assert validate_source_season(ROOT / "data" / "seasons" / "2026", ROOT)
    assert validate_archives(ROOT / "data")
    assert validate_public_tree(ROOT / "web" / "data")


def test_public_capabilities_and_split_contract() -> None:
    index = json.loads((ROOT / "web" / "data" / "index.json").read_text())
    assert len(index["seasons"]) == 39
    assert index["current_season"] == 2026
    summary = next(season for season in index["seasons"] if season["year"] == 1988)
    full = next(season for season in index["seasons"] if season["year"] == 2025)
    current = next(season for season in index["seasons"] if season["year"] == 2026)
    assert summary["detail_level"] == "summary"
    assert summary["weeks_available"] == []
    assert summary["capabilities"]["matchups"] is False
    assert summary["capabilities"]["player_stats"] is False
    assert full["weeks_available"] == list(range(1, 18))
    assert current["detail_level"] == "full"
    assert current["weeks_available"] == [1, 2]
    assert (ROOT / "web" / "data" / "seasons" / "2025" / "weeks" / "week_17.json").exists()
    assert (ROOT / "web" / "data" / "seasons" / "2026" / "weeks" / "week_2.json").exists()


def test_frontend_contract_and_static_only_hosting() -> None:
    html = (ROOT / "web" / "index.html").read_text()
    javascript = (ROOT / "web" / "app.js").read_text()
    vercel = json.loads((ROOT / "vercel.json").read_text())
    assert 'href="styles.css"' in html
    assert 'src="app.js"' in html
    assert "<style>" not in html
    assert "hashchange" in javascript
    assert "escapeHtml" in javascript
    assert "aria-disabled" in javascript
    assert "Summary-season limitation" in javascript
    assert "roster-search" in javascript
    assert "menu-toggle" in html
    assert all("api/" not in json.dumps(build) for build in vercel["builds"])


def test_javascript_syntax() -> None:
    result = subprocess.run(
        ["node", "--check", str(ROOT / "web" / "app.js")],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr
