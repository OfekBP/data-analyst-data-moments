"""Proof that the numbers in the answers are right.

Three layers:
  1. pinned headline numbers (so a silent change in the data or the code fails loudly);
  2. SQL results must equal the pandas results;
  3. the Excel workbook (live formulas, recalculated) must equal the pandas results.
"""
import math

import pandas as pd
import pytest

from dm_analyst import config, run_analysis, sql_check
from dm_analyst.load import load_raw


@pytest.fixture(scope="module")
def res():
    return run_analysis.compute_all()


@pytest.fixture(scope="module")
def sql(res):
    return sql_check.run_all(res["df"])


# ------------------------------------------------------------------ data quality
def test_cleaning_removes_only_exact_duplicate_lines(res):
    q = res["quality"]
    assert q["duplicate_lines_removed"] == 59
    assert q["rows_clean"] == q["rows_raw"] - 59
    assert q["store_ranking_unchanged_by_cleaning"]


def test_delivered_columns_are_arithmetically_consistent(res):
    q = res["quality"]
    assert q["line_sales_formula_violations"] == 0
    assert q["line_cost_formula_violations"] == 0
    assert q["unit_price_varies_within_item"] == 0


def test_bill_key_requires_store_code():
    raw = load_raw()
    assert raw.groupby("BillNum")["StoreCode"].nunique().max() == 1  # holds in this file, but the key still uses both


# ------------------------------------------------------------------ headline numbers
def test_chain_totals(res):
    s = res["q1"]["stores"]
    assert s.Sales.sum() == pytest.approx(3_520_099.61, abs=0.01)
    assert s.Profit.sum() == pytest.approx(1_157_552.43, abs=0.01)
    assert s.Bills.sum() == 15_588


def test_q1_focus_store_is_a_mix_problem_not_a_price_problem(res):
    q1, q4 = res["q1"], res["q4"]
    s = q1["stores"]
    assert s.Sales.idxmax() == config.FOCUS_STORE
    assert s.Margin.idxmin() == config.FOCUS_STORE
    assert q1["focus_margin_actual"] == pytest.approx(0.2792, abs=1e-3)
    assert q1["focus_margin_at_rest_mix"] == pytest.approx(0.3425, abs=1e-3)
    cat = q1["categories"]  # per-category margins are the same in the focus store and elsewhere
    assert (cat.FocusMargin - cat.RestMargin).abs().max() < 0.01
    assert q4["list_price_varies_across_stores_items"] == 0


def test_q2_mean_median(res):
    q2 = res["q2"]
    assert q2["mean"] == pytest.approx(225.82, abs=0.01)
    assert q2["median"] == pytest.approx(186.545, abs=0.001)
    assert q2["mean_ex_bulk"] == pytest.approx(214.74, abs=0.01)
    assert q2["median_ex_bulk"] == pytest.approx(185.73, abs=0.01)
    assert 0.2 < q2["gap_explained_by_bulk"] < 0.3  # bulk lines explain only about a quarter of the gap


def test_q3_mix_shift(res):
    c = res["q3"]["categories"]
    assert c.loc["Dry Food", "SalesChangePct"] == pytest.approx(-0.276, abs=1e-3)
    assert c.loc["Wet Food", "SalesChangePct"] == pytest.approx(1.330, abs=1e-3)
    assert res["q3"]["total_sales_change_pct"] == pytest.approx(0.012, abs=1e-3)
    assert res["q3"]["total_profit_change_pct"] == pytest.approx(0.125, abs=1e-3)
    assert (c.Margin_Q1 - c.Margin_Q4).abs().max() < 0.01  # category margins are stable


def test_q4_same_customers_buy_differently_in_focus_store(res):
    q4 = res["q4"]
    mix = q4["common_customer_category_mix"]["Dry Food"]
    assert mix[config.FOCUS_STORE] == pytest.approx(0.794, abs=1e-3)
    assert mix["Rest of chain"] == pytest.approx(0.497, abs=1e-3)
    assert q4["common_customers"] == 1294
    a = q4["attach_rates"]
    assert a.loc[config.FOCUS_STORE, "Wet"] == pytest.approx(0.391, abs=1e-3)


# ------------------------------------------------------------------ SQL == pandas
def test_sql_q1_matches_pandas(res, sql):
    py = res["q1"]["stores"].sort_index()
    q = sql["q1_stores"].set_index("StoreName").sort_index()
    for col in ["Sales", "Profit", "Bills"]:
        assert q[col].to_numpy() == pytest.approx(py[col].to_numpy(), abs=0.01)
    assert q["Margin"].to_numpy() == pytest.approx(py["Margin"].to_numpy(), abs=1e-5)
    assert q["AvgBill"].to_numpy() == pytest.approx(py["AvgBill"].to_numpy(), abs=1e-3)


def test_sql_q2_matches_pandas(res, sql):
    row = sql["q2_receipt_mean_median"].iloc[0]
    assert row.Bills == res["q2"]["bills"]
    assert row.Mean == pytest.approx(res["q2"]["mean"], abs=1e-3)
    assert row.Median == pytest.approx(res["q2"]["median"], abs=1e-3)


def test_sql_q3_matches_pandas(res, sql):
    py = res["q3"]["categories"].sort_index()
    q = sql["q3_category_q1_vs_q4"].set_index("Category").sort_index()
    for sql_col, py_col in [("Sales_Q1", "Sales_Q1"), ("Sales_Q4", "Sales_Q4"),
                            ("Profit_Q1", "Profit_Q1"), ("Profit_Q4", "Profit_Q4")]:
        assert q[sql_col].to_numpy() == pytest.approx(py[py_col].to_numpy(), abs=0.01)
    assert q["SalesChangePct"].to_numpy() == pytest.approx(py["SalesChangePct"].to_numpy(), abs=1e-5)


def test_sql_q4_matches_pandas(res, sql):
    py = res["q4"]["attach_rates"].sort_index()
    q = sql["q4_attach_rates"].set_index("StoreName").sort_index()
    for sql_col, py_col in [("Dry", "Dry"), ("Wet", "Wet"), ("Snacks", "Snacks")]:
        assert q[sql_col].to_numpy() == pytest.approx(py[py_col].to_numpy(), abs=1e-5)


# ------------------------------------------------------------------ Excel == pandas
def _excel_value(wb, sheet, cell):
    return wb[sheet][cell].value


@pytest.mark.skipif(not config.EXCEL_PATH.exists(), reason="workbook not built")
def test_excel_formulas_match_pandas(res):
    from openpyxl import load_workbook
    wb = load_workbook(config.EXCEL_PATH, data_only=True)
    if _excel_value(wb, "Q1_Stores", "G5") is None:
        pytest.skip("workbook has no cached values - open it in Excel and save")
    s = res["q1"]["stores"]
    assert _excel_value(wb, "Q1_Stores", "G5") == pytest.approx(s.Sales.sum(), abs=0.01)
    assert _excel_value(wb, "Q1_Stores", "C7") == pytest.approx(s.loc["Haifa", "Margin"], abs=1e-6)
    assert _excel_value(wb, "Q2_MeanMedian", "B6") == pytest.approx(res["q2"]["mean"], abs=1e-6)
    assert _excel_value(wb, "Q2_MeanMedian", "B7") == pytest.approx(res["q2"]["median"], abs=1e-6)
    assert _excel_value(wb, "Q3_Categories", "C6") == pytest.approx(res["q3"]["categories"].loc["Wet Food", "Sales_Q4"], abs=0.01)
    assert _excel_value(wb, "Q4_Haifa_Mix", "C6") == pytest.approx(res["q4"]["attach_rates"].loc["Haifa", "Wet"], abs=1e-6)
