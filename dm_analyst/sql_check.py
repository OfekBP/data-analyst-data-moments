"""Re-run the main calculations in SQL (SQLite) on the same cleaned data.

The SQL files in sql/ are the independent second implementation: the tests compare
their output with the pandas results.

    python -m dm_analyst.sql_check
"""
import sqlite3

import pandas as pd

from . import config
from .load import load_clean


def build_database(df: pd.DataFrame) -> sqlite3.Connection:
    con = sqlite3.connect(":memory:")
    cols = ["BillDate", "Bill", "StoreName", "Category", "LineSales", "LineCost"]
    out = df[cols].copy()
    out["BillDate"] = out["BillDate"].dt.strftime("%Y-%m-%d")
    out.to_sql("sales", con, index=False)
    return con


def run(name: str, con: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql_query((config.SQL_DIR / f"{name}.sql").read_text(encoding="utf-8"), con)


def run_all(df: pd.DataFrame | None = None) -> dict[str, pd.DataFrame]:
    con = build_database(load_clean() if df is None else df)
    names = ["q1_stores", "q2_receipt_mean_median", "q3_category_q1_vs_q4", "q4_attach_rates"]
    return {n: run(n, con) for n in names}


if __name__ == "__main__":
    for name, table in run_all().items():
        print(f"\n-- {name}\n{table.to_string(index=False)}")
