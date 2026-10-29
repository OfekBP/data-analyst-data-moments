"""One function per assignment question. Every function takes the cleaned frame."""
import numpy as np
import pandas as pd

from . import config

FOCUS = config.FOCUS_STORE


def _margin(profit, sales):
    return profit / sales


def bill_totals(df: pd.DataFrame) -> pd.DataFrame:
    """One row per receipt: total sales, whether it holds a bulk line, and its store."""
    g = df.groupby("Bill")
    return pd.DataFrame({
        "StoreName": g["StoreName"].first(),
        "Total": g["LineSales"].sum(),
        "HasBulkLine": g["SaleQuantity"].max() >= config.BULK_QTY,
    })


# --------------------------------------------------------------------------- Q1
def q1_stores(df: pd.DataFrame) -> dict:
    """Compare stores on size, profitability and basket, then explain the gap."""
    bills = bill_totals(df)
    g = df.groupby("StoreName")
    stores = pd.DataFrame({
        "Sales": g["LineSales"].sum(),
        "Profit": g["Profit"].sum(),
        "Bills": bills.groupby("StoreName").size(),
        "OnlineShare": df[df.SalesChannel == "Online"].groupby("StoreName")["LineSales"].sum() / g["LineSales"].sum(),
        "UnidentifiedShare": df[df.CustomerCode == 0].groupby("StoreName")["LineSales"].sum() / g["LineSales"].sum(),
    })
    stores["Margin"] = _margin(stores.Profit, stores.Sales)
    stores["AvgBill"] = stores.Sales / stores.Bills
    stores["ProfitPerBill"] = stores.Profit / stores.Bills

    # Category mix of the focus store vs. the rest of the chain, and the margin
    # the focus store would have had with the rest-of-chain mix (mix standardisation).
    is_focus = df.StoreName == FOCUS
    cat = pd.DataFrame({
        "FocusSales": df[is_focus].groupby("Category")["LineSales"].sum(),
        "FocusProfit": df[is_focus].groupby("Category")["Profit"].sum(),
        "RestSales": df[~is_focus].groupby("Category")["LineSales"].sum(),
        "RestProfit": df[~is_focus].groupby("Category")["Profit"].sum(),
    }).loc[config.CATEGORIES]
    cat["FocusMargin"] = _margin(cat.FocusProfit, cat.FocusSales)
    cat["RestMargin"] = _margin(cat.RestProfit, cat.RestSales)
    cat["FocusMix"] = cat.FocusSales / cat.FocusSales.sum()
    cat["RestMix"] = cat.RestSales / cat.RestSales.sum()

    actual = cat.FocusProfit.sum() / cat.FocusSales.sum()
    at_rest_mix = float((cat.FocusMargin * cat.RestMix).sum())
    return {
        "stores": stores,
        "categories": cat,
        "focus_margin_actual": actual,
        "focus_margin_at_rest_mix": at_rest_mix,
        "focus_extra_profit_same_sales": float(cat.FocusSales.sum() * (at_rest_mix - actual)),
    }


# --------------------------------------------------------------------------- Q2
def q2_receipt_distribution(df: pd.DataFrame) -> dict:
    """Mean vs. median receipt amount and a test of the 'bulk purchases' explanation."""
    bills = bill_totals(df)
    t = bills.Total
    ex = bills.loc[~bills.HasBulkLine, "Total"]
    lines_per_bill = df.groupby("Bill").size()
    return {
        "bills": int(len(t)),
        "mean": float(t.mean()),
        "median": float(t.median()),
        "gap": float(t.mean() - t.median()),
        "gap_pct_of_median": float(t.mean() / t.median() - 1),
        "std": float(t.std()),
        "skew": float(t.skew()),
        "share_below_mean": float((t < t.mean()).mean()),
        "p90": float(t.quantile(0.9)),
        "p99": float(t.quantile(0.99)),
        "max": float(t.max()),
        "top_decile_sales_share": float(t[t >= t.quantile(0.9)].sum() / t.sum()),
        "bills_with_bulk_line": int(bills.HasBulkLine.sum()),
        "mean_ex_bulk": float(ex.mean()),
        "median_ex_bulk": float(ex.median()),
        "gap_ex_bulk": float(ex.mean() - ex.median()),
        "gap_explained_by_bulk": float(1 - (ex.mean() - ex.median()) / (t.mean() - t.median())),
        "avg_by_lines_in_bill": t.groupby(lines_per_bill).agg(["count", "mean"]),
    }


# --------------------------------------------------------------------------- Q3
def q3_category_change(df: pd.DataFrame) -> dict:
    """Category sales and profit in Q1 vs. Q4 and the chain-level effect of the mix shift."""
    sub = df[df.Quarter.isin([1, 4])]
    sales = sub.pivot_table(index="Category", columns="Quarter", values="LineSales", aggfunc="sum").loc[config.CATEGORIES]
    profit = sub.pivot_table(index="Category", columns="Quarter", values="Profit", aggfunc="sum").loc[config.CATEGORIES]
    t = pd.DataFrame({
        "Sales_Q1": sales[1], "Sales_Q4": sales[4],
        "Profit_Q1": profit[1], "Profit_Q4": profit[4],
    })
    t["SalesChange"] = t.Sales_Q4 - t.Sales_Q1
    t["SalesChangePct"] = t.Sales_Q4 / t.Sales_Q1 - 1
    t["Mix_Q1"] = t.Sales_Q1 / t.Sales_Q1.sum()
    t["Mix_Q4"] = t.Sales_Q4 / t.Sales_Q4.sum()
    t["Margin_Q1"] = t.Profit_Q1 / t.Sales_Q1
    t["Margin_Q4"] = t.Profit_Q4 / t.Sales_Q4
    t["ProfitChange"] = t.Profit_Q4 - t.Profit_Q1

    focus = np.where(df.StoreName == FOCUS, FOCUS, "Rest of chain")
    shares = {}
    for cat in ["Dry Food", "Wet Food"]:
        s = df.assign(Group=focus, Is=(df.Category == cat) * df.LineSales)
        shares[cat] = (s.groupby(["Group", "Quarter"])["Is"].sum() / s.groupby(["Group", "Quarter"])["LineSales"].sum()).unstack()
    return {
        "categories": t,
        "total_sales_change_pct": float(t.Sales_Q4.sum() / t.Sales_Q1.sum() - 1),
        "total_profit_change_pct": float(t.Profit_Q4.sum() / t.Profit_Q1.sum() - 1),
        "margin_Q1": float(t.Profit_Q1.sum() / t.Sales_Q1.sum()),
        "margin_Q4": float(t.Profit_Q4.sum() / t.Sales_Q4.sum()),
        "dry_share_by_group": shares["Dry Food"],
        "wet_share_by_group": shares["Wet Food"],
    }


# --------------------------------------------------------------------------- Q4
def q4_focus_store_drivers(df: pd.DataFrame) -> dict:
    """Test candidate explanations for the focus store's low margin, one by one."""
    # Candidate 1: pricing/discounts. Candidate 2: basket size. Candidate 3: category attach rates.
    cats_by_bill = df.assign(
        Dry=df.Category == "Dry Food", Wet=df.Category == "Wet Food", Snacks=df.Category == "Snacks",
    ).groupby("Bill").agg(StoreName=("StoreName", "first"), Dry=("Dry", "max"), Wet=("Wet", "max"),
                          Snacks=("Snacks", "max"), Lines=("LineNum", "size"))
    attach = cats_by_bill.groupby("StoreName")[["Dry", "Wet", "Snacks"]].mean()
    attach["LinesPerBill"] = cats_by_bill.groupby("StoreName")["Lines"].mean()
    rest_attach = cats_by_bill[cats_by_bill.StoreName != FOCUS][["Dry", "Wet", "Snacks"]].mean()  # pooled, not a mean of stores

    gross = (df.SaleQuantity.clip(lower=0) * df.UnitPrice)
    attach["DiscountRate"] = df.groupby("StoreName")["DiscountAmount"].sum() / gross.groupby(df.StoreName).sum()
    # Does any item have a different list price in different stores? (expected: no)
    price_varies = int((df.groupby("ItemCode")["UnitPrice"].nunique() > 1).sum())

    # Candidate 4: a different customer base. Customers who bought both in the focus
    # store and elsewhere are the same people, so any difference is store-driven.
    ident = df[df.CustomerCode > 0].assign(IsFocus=lambda d: d.StoreName == FOCUS)
    both = ident.groupby("CustomerCode")["IsFocus"].nunique()
    both_ids = both[both == 2].index
    common = ident[ident.CustomerCode.isin(both_ids)]
    common_mix = (common.pivot_table(index="IsFocus", columns="Category", values="LineSales", aggfunc="sum")
                  .pipe(lambda x: x.div(x.sum(axis=1), axis=0)))
    multi_store = (ident.groupby("CustomerCode")["StoreCode"].nunique() > 1).mean()

    # Are wet food and snack volumes low across the board or on a few items?
    items = df[df.Category.isin(["Wet Food", "Snacks"])].pivot_table(
        index="ItemCode", columns=df.StoreName == FOCUS, values="SaleQuantity", aggfunc="sum")
    items["Expected"] = items[False] / 4  # focus store would sell a quarter of the other four stores' volume
    items["FocusVsExpected"] = items[True] / items["Expected"]

    return {
        "attach_rates": attach,
        "rest_attach_rates": rest_attach,
        "list_price_varies_across_stores_items": price_varies,
        "common_customers": int(len(both_ids)),
        "customers_in_multiple_stores_share": float(multi_store),
        "common_customer_category_mix": common_mix.rename(index={True: FOCUS, False: "Rest of chain"}).loc[:, config.CATEGORIES],
        "item_volume_vs_expected": items["FocusVsExpected"].describe(),
        "items_compared": int(len(items)),
    }


# --------------------------------------------------------------------------- Q5
def q5_kpi_baselines(q1: dict, q4: dict, quality: dict) -> pd.DataFrame:
    """Baseline values of the two recommended KPIs, for the numbers written in the answer."""
    attach, rest = q4["attach_rates"], q4["rest_attach_rates"]
    cat = q1["categories"]
    exception_lines = (quality["bulk_lines"] + quality["discount_ge_gross_sale_lines"]
                       + quality["unit_cost_above_unit_price"])
    return pd.DataFrame([
        ("1. Wet-food attach rate, focus store", "% of receipts containing Wet Food", attach.loc[FOCUS, "Wet"], "up", "monthly",
         f"rest of chain {rest.Wet:.1%}"),
        ("1. Snack attach rate, focus store", "% of receipts containing Snacks", attach.loc[FOCUS, "Snacks"], "up", "monthly",
         f"rest of chain {rest.Snacks:.1%}"),
        ("1. Gross margin, focus store", "profit / net sales", q1["focus_margin_actual"], "up", "monthly",
         f"rest of chain {cat.RestProfit.sum() / cat.RestSales.sum():.1%}"),
        ("2. Exception lines per month", "bulk lines + discount>=gross + cost>price", exception_lines / 12, "down", "monthly",
         "2025 monthly average"),
    ], columns=["KPI", "Calculation", "Baseline", "Desired direction", "Frequency", "Reference"])
