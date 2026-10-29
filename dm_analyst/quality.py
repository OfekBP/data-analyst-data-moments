"""Data-quality checks. Each check is counted, not fixed silently."""
import pandas as pd

from . import config
from .load import clean


def quality_report(raw: pd.DataFrame) -> dict:
    """Return a flat dict of counts and the effect of the cleaning on the totals."""
    dup = raw.duplicated(config.LINE_KEY)
    df = clean(raw)

    gross = df["SaleQuantity"] * df["UnitPrice"]
    modal_cost = df.groupby("ItemCode")["UnitCost"].transform(lambda s: s.mode().iloc[0])
    bill_channels = df.groupby("Bill")["SalesChannel"].nunique()

    ranking_raw = raw.groupby("StoreName")["LineSales"].sum().rank(ascending=False)
    ranking_clean = df.groupby("StoreName")["LineSales"].sum().rank(ascending=False)

    return {
        "rows_raw": len(raw),
        "rows_clean": len(df),
        "duplicate_lines_removed": int(dup.sum()),
        "duplicate_sales_removed": round(float(raw.loc[dup, "LineSales"].sum()), 2),
        "sales_raw": round(float(raw["LineSales"].sum()), 2),
        "sales_clean": round(float(df["LineSales"].sum()), 2),
        "store_ranking_unchanged_by_cleaning": bool((ranking_raw == ranking_clean).all()),
        "return_lines": int((df["SaleQuantity"] < 0).sum()),
        "return_sales": round(float(df.loc[df["SaleQuantity"] < 0, "LineSales"].sum()), 2),
        "bulk_lines": int((df["SaleQuantity"] >= config.BULK_QTY).sum()),
        "bulk_sales_share": round(float(df.loc[df["SaleQuantity"] >= config.BULK_QTY, "LineSales"].sum()
                                        / df["LineSales"].sum()), 4),
        "discount_ge_gross_sale_lines": int(((df["SaleQuantity"] > 0) & (df["DiscountAmount"] > 0)
                                             & (df["DiscountAmount"] >= gross)).sum()),
        "unit_cost_above_unit_price": int((df["UnitCost"] > df["UnitPrice"]).sum()),
        "unit_cost_off_item_mode": int(((df["UnitCost"] - modal_cost).abs() > 0.01).sum()),
        "customer_code_missing": int(df["CustomerCode"].isna().sum()),
        "customer_code_unidentified_zero": int((df["CustomerCode"] == 0).sum()),
        "bills_with_both_channels": int((bill_channels > 1).sum()),
        # Arithmetic consistency of the delivered columns (expected: 0 violations).
        "line_sales_formula_violations": int(((df["SaleQuantity"] * df["UnitPrice"] - df["DiscountAmount"]
                                               - df["LineSales"]).abs() > 0.02).sum()),
        "line_cost_formula_violations": int(((df["SaleQuantity"] * df["UnitCost"] - df["LineCost"]).abs() > 0.02).sum()),
        "unit_price_varies_within_item": int((df.groupby("ItemCode")["UnitPrice"].nunique() > 1).sum()),
    }
