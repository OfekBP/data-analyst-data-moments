"""Run the whole analysis and write tables + a flat results.json.

    python -m dm_analyst.run_analysis
"""
import json

import pandas as pd

from . import config
from .load import clean, load_raw
from .quality import quality_report
from .questions import (q1_stores, q2_receipt_distribution, q3_category_change,
                        q4_focus_store_drivers, q5_kpi_baselines)


def compute_all() -> dict:
    raw = load_raw()
    df = clean(raw)
    quality = quality_report(raw)
    q1, q4 = q1_stores(df), q4_focus_store_drivers(df)
    return {
        "quality": quality,
        "q1": q1,
        "q2": q2_receipt_distribution(df),
        "q3": q3_category_change(df),
        "q4": q4,
        "q5": q5_kpi_baselines(q1, q4, quality),
        "df": df,
    }


def _scalars(d: dict) -> dict:
    return {k: (float(v) if isinstance(v, float) else v) for k, v in d.items()
            if isinstance(v, (int, float, bool, str))}


def main() -> None:
    config.TABLES.mkdir(parents=True, exist_ok=True)
    r = compute_all()
    r["q1"]["stores"].round(4).to_csv(config.TABLES / "q1_stores.csv")
    r["q1"]["categories"].round(4).to_csv(config.TABLES / "q1_focus_vs_rest_categories.csv")
    r["q2"]["avg_by_lines_in_bill"].round(2).to_csv(config.TABLES / "q2_avg_receipt_by_lines.csv")
    r["q3"]["categories"].round(4).to_csv(config.TABLES / "q3_category_q1_vs_q4.csv")
    r["q3"]["dry_share_by_group"].round(4).to_csv(config.TABLES / "q3_dry_share_by_quarter.csv")
    r["q3"]["wet_share_by_group"].round(4).to_csv(config.TABLES / "q3_wet_share_by_quarter.csv")
    r["q4"]["attach_rates"].round(4).to_csv(config.TABLES / "q4_attach_rates.csv")
    r["q4"]["rest_attach_rates"].round(4).to_csv(config.TABLES / "q4_rest_of_chain_attach_rates.csv", header=["RestOfChain"])
    r["q4"]["common_customer_category_mix"].round(4).to_csv(config.TABLES / "q4_common_customers_mix.csv")
    r["q5"].round(4).to_csv(config.TABLES / "q5_kpis.csv", index=False)

    summary = {
        "quality": r["quality"],
        "q1": _scalars(r["q1"]),
        "q2": _scalars(r["q2"]),
        "q3": _scalars(r["q3"]),
        "q4": _scalars(r["q4"]),
    }
    (config.OUTPUTS / "results.json").write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")

    pd.set_option("display.width", 200)
    print("Data quality:", json.dumps(r["quality"], indent=2))
    print("\nQ1 stores:\n", r["q1"]["stores"].round(3))
    print(f"\nQ1 focus margin {r['q1']['focus_margin_actual']:.3f} -> at rest-of-chain mix "
          f"{r['q1']['focus_margin_at_rest_mix']:.3f}; extra profit {r['q1']['focus_extra_profit_same_sales']:,.0f}")
    print(f"\nQ2 mean {r['q2']['mean']:.2f} median {r['q2']['median']:.2f}; "
          f"ex-bulk {r['q2']['mean_ex_bulk']:.2f}/{r['q2']['median_ex_bulk']:.2f}")
    print("\nQ3:\n", r["q3"]["categories"][["Sales_Q1", "Sales_Q4", "SalesChangePct", "Profit_Q1", "Profit_Q4"]].round(3))
    print("\nQ4 attach:\n", r["q4"]["attach_rates"].round(3))
    print("\nQ4 common customers mix:\n", r["q4"]["common_customer_category_mix"].round(3))
    print("\nQ5:\n", r["q5"].round(3))


if __name__ == "__main__":
    main()
