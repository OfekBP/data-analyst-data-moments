# Data Moments - analyst assignment

Analysis of a five-store pet-supply chain's 2025 sales (41,685 receipt lines): store comparison, receipt
distribution, category change Q1 vs. Q4, an open investigation, and two KPI recommendations.
The written answers (Hebrew) are in [`docs/answers.md`](docs/answers.md).

Every headline number is computed three independent ways and checked by tests:

| Layer | Where | Role |
|---|---|---|
| Python (pandas) | `dm_analyst/` | main analysis, writes `outputs/tables/*.csv` and `outputs/results.json` |
| SQL (SQLite) | `sql/*.sql`, `dm_analyst/sql_check.py` | independent re-computation of Q1-Q4 |
| Excel, live formulas | `excel/` | follow-the-numbers workbook for reviewers; `Checks` sheet compares to Python |

## Layout
```
data/raw/            analyst_exam.xlsx - the exam file (synthetic data, published with permission)
dm_analyst/          load.py (clean), quality.py (data checks), questions.py (Q1-Q5), run_analysis.py, sql_check.py
sql/                 one query per question
excel/               build_workbook.py, recalc.ps1, the workbook
tests/               pinned numbers + SQL == pandas + Excel == pandas
outputs/             generated tables and results.json
docs/answers.md      the deliverable
```

## Run
```bash
pip install -r requirements.txt
python -m dm_analyst.run_analysis        # tables + results.json + printed summary
python -m dm_analyst.sql_check           # SQL results
python -m pytest                         # proof: 13 tests
python -m excel.build_workbook           # rebuild the workbook (Windows + Excel: then run excel/recalc.ps1)
```

## Method notes
- A receipt is `StoreCode + BillNum`; a receipt line adds `LineNum`. `LineSales` is already net of the line discount.
- Only exact duplicate lines are removed (59). Returns, bulk lines and unidentified customers are kept.
- Parameters (bulk threshold = 20 units, focus store = Haifa) are in `dm_analyst/config.py`.
- The Excel workbook needs desktop Excel (or LibreOffice) to recalculate; the `Data` sheet relies on rows
  being sorted by store and receipt, which holds for this file.
