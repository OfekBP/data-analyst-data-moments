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
docs/answers.md      the written answers (the two-page version)
site/                the interactive submission page (deployed on Vercel)
PRODUCT.md           product/brand context for the page
```

## Run
```bash
pip install -r requirements.txt
python -m dm_analyst.run_analysis        # tables + results.json + printed summary
python -m dm_analyst.sql_check           # SQL results
python -m pytest                         # proof: invariants, adaptivity, reference figures
python -m excel.build_workbook           # rebuild the workbook (Windows + Excel: then run excel/recalc.ps1)
```

## The submission page (`site/`)
`site/index.html` is the official, interactive submission: one static page (Hebrew, RTL) in Data Moments' visual
identity with the five answers as an accordion, an executive summary, the data-quality description, the AI note and
the supporting calculations. Its numbers are typed by hand, like `docs/answers.md`, and do not update with the data.
`tests/test_site.py` only checks that every link it shows points at something that exists.

Deploy on Vercel: import the repository, leave the framework as "Other" and keep the defaults. `vercel.json`
sets the output directory to `site` and its build command copies the Excel workbook and `outputs/tables/` into
`site/calculations/`, so the page's download links work (those copies are git-ignored). To preview locally:
```bash
mkdir -p site/calculations && cp excel/DataMoments_Analyst_Calculations.xlsx site/calculations/ && cp -r outputs/tables site/calculations/tables
python -m http.server 5173 --directory site     # then open http://localhost:5173
```

## How the code is organised
- The five questions of the brief are independent, so `questions.py` has one function per question and each takes
  **only the cleaned data** (`q1_stores(df)`, `q2_receipt_distribution(df)`, ...). No question reads another's result;
  shared building blocks (`store_table`, `category_gap`, `attach_rates`) are helpers that run on the data.
- Nothing names a store or a category. The store to improve is the one with the largest *profit upside*
  (sales x (margin of the rest of the chain - own margin)); the categories behind its gap, the priciest category
  and the biggest movers between Q1 and Q4 are all derived from the numbers. Change the figures and the
  conclusions follow; change the table layout (column names) and it breaks.
- `python -m dm_analyst.run_analysis` prints the reasoning behind each answer and, at the end, **what changed since
  the previous run**. `docs/answers.md` is hand-written text and is not regenerated: after a change, re-read it.
- Tests: INVARIANTS (SQL == pandas == Excel, parts add up), ADAPTIVITY (make another store weak, check every
  question re-points to it) and REFERENCE figures of the delivered data (they fail on purpose when the data changes,
  with a message saying so).

## Method notes
- A receipt is `StoreCode + BillNum`; a receipt line adds `LineNum`. `LineSales` is already net of the line discount.
- Only exact duplicate lines are removed (59). Returns, bulk lines and unidentified customers are kept.
- Parameters are in `dm_analyst/config.py`: bulk threshold (20 units; `quality.py` shows it sits in a real gap in the
  quantity distribution) and an optional store override.
- The Excel workbook needs desktop Excel to recalculate (`excel/recalc.ps1`); its `Data` sheet relies on rows
  being sorted by store and receipt, which holds for this file. Its store/category choices are formulas.
