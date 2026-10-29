"""Integrity of the submission page (site/index.html).

The page's numbers are typed by hand on purpose (like docs/answers.md) and do not update with the
data, so these tests only check that what the page points at exists.
"""
import re

from dm_analyst import config

SITE = config.ROOT / "site"
HTML = (SITE / "index.html").read_text(encoding="utf-8")
REPO = "https://github.com/OfekBP/data-analyst-data-moments"


def test_every_github_file_link_points_at_a_file_in_the_repo():
    paths = set(re.findall(re.escape(REPO) + r"/blob/main/([^\"#]+)", HTML))
    assert paths, "the page should link to the code"
    missing = [p for p in paths if not (config.ROOT / p).exists()]
    assert not missing, f"links to files that do not exist: {missing}"


def test_local_assets_exist():
    for src in set(re.findall(r'(?:src|href)="(assets/[^"]+)"', HTML)):
        assert (SITE / src).exists(), src


def test_deployed_downloads_are_produced_by_the_vercel_build():
    """calculations/* is copied into site/ at deploy time; make sure every source exists."""
    names = set(re.findall(r'href="calculations/tables/([^"]+)"', HTML))
    assert names
    for n in names:
        assert (config.OUTPUTS / "tables" / n).exists(), n
    assert (config.ROOT / "excel" / "DataMoments_Analyst_Calculations.xlsx").exists()
    assert "calculations/DataMoments_Analyst_Calculations.xlsx" in HTML


def test_every_question_in_the_brief_is_on_the_page():
    for q in ("ביצועי הסניפים", "ממוצע וחציון", "שינוי לאורך זמן", "חקירה פתוחה", "המלצות ומדדי KPI"):
        assert q in HTML or q.split()[0] in HTML
    for required in ("איכות הנתונים", "שימוש ב-AI", "חישובים תומכים", "github.com/OfekBP/data-analyst-data-moments"):
        assert required in HTML, required
