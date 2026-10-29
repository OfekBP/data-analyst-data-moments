"""Run the whole analysis, write tables + results.json, and print what it found and what changed.

    python -m dm_analyst.run_analysis

The five questions are independent: each one is computed from the cleaned data alone.
"""
import json

import numpy as np
import pandas as pd

from . import config
from .load import clean, load_raw
from .quality import quality_report
from .questions import (q1_stores, q2_receipt_distribution, q3_category_change,
                        q4_open_investigation, q5_kpi_baselines)

RESULTS_JSON = config.OUTPUTS / "results.json"


def compute_all() -> dict:
    raw = load_raw()
    df = clean(raw)
    return {
        "quality": quality_report(raw),
        "q1": q1_stores(df),
        "q2": q2_receipt_distribution(df),
        "q3": q3_category_change(df),
        "q4": q4_open_investigation(df),
        "q5": q5_kpi_baselines(df),
        "df": df,
    }


# ---------------------------------------------------------------- flat view of the results
def flatten(r: dict) -> dict:
    """Every number in the results as {'section.name[.row.col]': value}, for json and change detection."""
    flat = {}

    def add(prefix, v):
        if isinstance(v, pd.DataFrame):
            for idx, row in v.iterrows():
                for col, x in row.items():
                    add(f"{prefix}.{idx}.{col}", x)
        elif isinstance(v, dict):
            for k, x in v.items():
                add(f"{prefix}.{k}", x)
        elif isinstance(v, pd.Series):
            for idx, x in v.items():
                add(f"{prefix}.{idx}", x)
        elif isinstance(v, (list, tuple)):
            flat[prefix] = ", ".join(map(str, v))
        elif isinstance(v, (bool, np.bool_)):
            flat[prefix] = bool(v)
        elif isinstance(v, (int, float, np.integer, np.floating)):
            flat[prefix] = round(float(v), 6)
        elif isinstance(v, str):
            flat[prefix] = v

    for section in ["quality", "q1", "q2", "q3", "q4"]:
        for key, v in r[section].items():
            add(f"{section}.{key}", v)
    add("q5", r["q5"].set_index("KPI")["Baseline"])
    return flat


def changes_since_last_run(new: dict, old: dict) -> list[str]:
    out = []
    for k in sorted(set(new) | set(old)):
        a, b = old.get(k), new.get(k)
        same = (a == b) or (isinstance(a, float) and isinstance(b, float) and abs(a - b) <= 1e-6 * max(1.0, abs(a)))
        if not same:
            out.append(f"  {k}: {a} -> {b}")
    return out


# ---------------------------------------------------------------- output
def write_outputs(r: dict) -> None:
    config.TABLES.mkdir(parents=True, exist_ok=True)
    r["q1"]["stores"].round(4).to_csv(config.TABLES / "q1_stores.csv")
    r["q1"]["categories"].round(4).to_csv(config.TABLES / "q1_category_mix_gap.csv")
    r["q2"]["avg_by_lines_in_bill"].round(2).to_csv(config.TABLES / "q2_avg_receipt_by_lines.csv")
    r["q2"]["gap_steps"].round(4).to_csv(config.TABLES / "q2_gap_explanations.csv")
    r["q3"]["categories"].round(4).to_csv(config.TABLES / "q3_category_q1_vs_q4.csv")
    r["q4"]["attach_rates"].round(4).to_csv(config.TABLES / "q4_attach_rates.csv")
    r["q4"]["rest_attach_rates"].round(4).to_csv(config.TABLES / "q4_rest_of_chain_attach_rates.csv", header=["RestOfChain"])
    r["q4"]["common_customer_category_mix"].round(4).to_csv(config.TABLES / "q4_common_customers_mix.csv")
    for cat, table in r["q4"]["share_by_quarter_of_chain_movers"].items():
        table.round(4).to_csv(config.TABLES / f"q4_share_of_{cat.lower().replace(' ', '_')}_by_quarter.csv")
    r["q5"].round(4).to_csv(config.TABLES / "q5_kpis.csv", index=False)


def explain(r: dict) -> None:
    """Print the reasoning chain so the reader can see WHY each conclusion came out as it did."""
    pd.set_option("display.width", 200)
    pd.set_option("display.max_columns", 30)
    q1, q2, q3, q4 = r["q1"], r["q2"], r["q3"], r["q4"]

    print("=" * 78, "\nDATA QUALITY\n" + "=" * 78)
    for k, v in r["quality"].items():
        print(f"  {k}: {v}")

    print("\n" + "=" * 78, "\nQ1  COMPARE THE STORES, CHOOSE ONE TO IMPROVE, AND WHY\n" + "=" * 78)
    cols = ["Sales", "Profit", "Margin", "Bills", "AvgBill", "SalesChangeQ4vsQ1", "ProfitChangeQ4vsQ1", "ProfitUpside"]
    print(q1["stores"][cols].round(3).to_string())
    print(f"\n  -> store to improve: {q1['focus']} ({q1['focus_reason']})")
    print("\n  Why is its margin different? Category mix vs the rest of the chain "
          "(Contribution = profit gained if its mix matched the rest):")
    print(q1["categories"][["FocusMix", "RestMix", "FocusMargin", "RestMargin", "Contribution"]].round(3).to_string())
    print(f"  -> margin {q1['focus_margin_actual']:.1%} would be {q1['focus_margin_at_rest_mix']:.1%} with the rest-of-chain mix")
    print(f"  -> the profit upside splits into a MIX effect ({q1['focus_extra_profit_same_sales']:,.0f} ILS: selling a different "
          f"category mix) and a RATE effect ({q1['focus_rate_effect']:,.0f} ILS: earning a different margin inside the same category)")
    if q1["driver_categories"]:
        print(f"  -> under-weighted categories driving the mix gap: {q1['driver_categories']}; "
              f"over-weighted: {q1['overweighted_category']}")
    else:
        print("  -> no category is under-weighted: this store's gap is NOT a mix gap, look at margins inside categories")

    print("\n" + "=" * 78, "\nQ2  MEAN VS MEDIAN RECEIPT: WHAT THE GAP SAYS AND WHAT EXPLAINS IT\n" + "=" * 78)
    print(f"  mean {q2['mean']:.2f}  median {q2['median']:.2f}  gap {q2['gap']:.2f} ({q2['gap_pct_of_median']:.0%} of median); "
          f"skew {q2['skew']:.1f}; {q2['share_below_mean']:.0%} of receipts are below the mean")
    print("  gap that remains after comparing like with like (mean - median inside comparable receipts):")
    print(q2["gap_steps"].round(3).to_string())
    print(f"  -> bulk lines alone remove {q2['gap_explained_by_bulk']:.0%}; adding basket size and the presence of "
          f"'{q2['priciest_category']}' (the priciest category) removes {q2['gap_steps'].ShareOfGapRemoved.iloc[-1]:.0%}")

    print("\n" + "=" * 78, "\nQ3  Q1 VS Q4 BY CATEGORY\n" + "=" * 78)
    print(q3["categories"][["Sales_Q1", "Sales_Q4", "SalesChangePct", "Profit_Q1", "Profit_Q4", "Margin_Q1", "Margin_Q4"]].round(3).to_string())
    print(f"  -> chain sales {q3['total_sales_change_pct']:+.1%}, profit {q3['total_profit_change_pct']:+.1%}; "
          f"biggest decline: {q3['biggest_decline']}, biggest growth: {q3['biggest_growth']}")

    subject = q4["subject"]
    print("\n" + "=" * 78, f"\nQ4  OPEN INVESTIGATION: WHAT EXPLAINS {subject.upper()}'S GAP (one hypothesis at a time)\n" + "=" * 78)
    print("  (the largest store gap in the data, chosen inside Q4 by the same rule Q1 uses - Q4 does not read Q1's output)")
    print("  1-3. share of receipts containing each category, basket size and discount rate:")
    print(q4["attach_rates"].round(3).to_string())
    print(f"  list price differs between stores for {q4['list_price_varies_across_stores_items']} items -> pricing "
          f"{'is' if q4['list_price_varies_across_stores_items'] else 'is NOT'} the reason")
    print(f"  4. {q4['customers_in_multiple_stores_share']:.0%} of identified customers shop in more than one store; "
          f"{q4['common_customers']} bought both in {subject} and elsewhere. Their basket mix:")
    print(q4["common_customer_category_mix"].round(3).to_string())
    print("     -> if the same people buy a different mix in this store, the cause is the store, not the customers")
    print("  5. is it behind in the chain's biggest trend? (share of the store's sales, by quarter)")
    for cat, table in q4["share_by_quarter_of_chain_movers"].items():
        print(f"     {cat}:\n{table.round(3).to_string()}")

    print("\n" + "=" * 78, "\nQ5  KPI BASELINES\n" + "=" * 78)
    print(r["q5"].round(3).to_string())


def report_changes(flat: dict) -> None:
    print("\n" + "=" * 78, "\nWHAT CHANGED SINCE THE LAST RUN\n" + "=" * 78)
    if not RESULTS_JSON.exists():
        print("  first run - nothing to compare with")
        return
    old = json.loads(RESULTS_JSON.read_text(encoding="utf-8"))
    diff = changes_since_last_run(flat, old)
    if not diff:
        print("  nothing: the numbers are identical to the previous run")
        return
    by_section = pd.Series([d.split(".")[0].strip() for d in diff]).value_counts()
    print(f"  {len(diff)} numbers changed ({', '.join(f'{s}: {n}' for s, n in by_section.items())}). First ones:")
    print("\n".join(diff[:15]))
    if len(diff) > 15:
        print(f"  ... and {len(diff) - 15} more (compare outputs/results.json in git)")
    print("\n  NOTE: docs/answers.md is written by hand. If these changes affect a figure or a conclusion\n"
          "  quoted there (for example the store chosen), update the text. 'pytest' lists the pinned\n"
          "  reference figures that no longer match.")


def main() -> None:
    r = compute_all()
    flat = flatten(r)
    explain(r)
    report_changes(flat)
    write_outputs(r)
    RESULTS_JSON.write_text(json.dumps(flat, indent=1, ensure_ascii=False, sort_keys=True), encoding="utf-8")


if __name__ == "__main__":
    main()
