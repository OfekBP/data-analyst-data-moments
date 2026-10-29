"""One function per assignment question.

The five questions of the brief are independent, so each `qN_*` function takes ONLY the cleaned
data frame (and the data-quality report for Q5's exception KPI). No question receives another
question's result. Where two questions need the same building block (for example "which store
has the largest profit upside"), both call the same helper below on the data.

Nothing here names a store or a category: stores, categories and the drivers of a gap are derived
from the numbers, so the same code stays consistent if the figures change (the layout must not).
"""
import pandas as pd

from . import config


def _margin(profit, sales):
    return profit / sales


# ============================================================================ helpers (shared)
def categories_by_sales(df: pd.DataFrame) -> list[str]:
    return list(df.groupby("Category")["LineSales"].sum().sort_values(ascending=False).index)


def bill_totals(df: pd.DataFrame) -> pd.DataFrame:
    """One row per receipt: total sales, whether it holds a bulk line, and its store."""
    g = df.groupby("Bill")
    return pd.DataFrame({
        "StoreName": g["StoreName"].first(),
        "Total": g["LineSales"].sum(),
        "HasBulkLine": g["SaleQuantity"].max() >= config.BULK_QTY,
    })


def store_table(df: pd.DataFrame) -> pd.DataFrame:
    """Size, profitability and basket per store, plus the profit left on the table.

    ProfitUpside = Sales x (margin of the OTHER stores pooled - own margin): the extra profit the
    store would make on its current sales if it earned what the rest of the chain earns.
    """
    bills = bill_totals(df)
    g = df.groupby("StoreName")
    s = pd.DataFrame({
        "Sales": g["LineSales"].sum(),
        "Profit": g["Profit"].sum(),
        "Bills": bills.groupby("StoreName").size(),
        "OnlineShare": df[df.SalesChannel == "Online"].groupby("StoreName")["LineSales"].sum() / g["LineSales"].sum(),
        "UnidentifiedShare": df[df.CustomerCode == 0].groupby("StoreName")["LineSales"].sum() / g["LineSales"].sum(),
    })
    s["Margin"] = _margin(s.Profit, s.Sales)
    s["AvgBill"] = s.Sales / s.Bills
    s["ProfitPerBill"] = s.Profit / s.Bills
    # Trend: Q4 vs Q1 (a store can look fine on the year and still be sliding).
    q = df[df.Quarter.isin([1, 4])]
    sales_q = q.pivot_table(index="StoreName", columns="Quarter", values="LineSales", aggfunc="sum")
    profit_q = q.pivot_table(index="StoreName", columns="Quarter", values="Profit", aggfunc="sum")
    bills_q = q.pivot_table(index="StoreName", columns="Quarter", values="Bill", aggfunc="nunique")
    s["SalesChangeQ4vsQ1"] = sales_q[4] / sales_q[1] - 1
    s["ProfitChangeQ4vsQ1"] = profit_q[4] / profit_q[1] - 1
    s["BillsChangeQ4vsQ1"] = bills_q[4] / bills_q[1] - 1
    s["RestMargin"] = _margin(s.Profit.sum() - s.Profit, s.Sales.sum() - s.Sales)
    s["ProfitUpside"] = s.Sales * (s.RestMargin - s.Margin)
    return s


def pick_store(stores: pd.DataFrame) -> tuple[str, str]:
    """The store with the largest profit upside (or the configured override), and why."""
    if config.FOCUS_STORE_OVERRIDE:
        return config.FOCUS_STORE_OVERRIDE, "set manually in config.FOCUS_STORE_OVERRIDE"
    name = stores.ProfitUpside.idxmax()
    row = stores.loc[name]
    return name, (f"largest profit upside: sales {row.Sales:,.0f} at margin {row.Margin:.1%} vs {row.RestMargin:.1%} "
                  f"for the rest of the chain = {row.ProfitUpside:,.0f} ILS")


def category_gap(df: pd.DataFrame, store: str) -> dict:
    """Split a store's margin gap versus the rest of the chain into a category-mix part and a rate part.

    Contribution(category) = store sales x (rest mix - store mix) x store margin in that category,
    so the contributions add up to the profit the store would gain with the rest-of-chain mix.
    The rate effect is the remainder: earning a different margin inside the same category.
    mix effect + rate effect == the store's ProfitUpside.
    """
    is_store = df.StoreName == store
    cats = categories_by_sales(df)
    cat = pd.DataFrame({
        "FocusSales": df[is_store].groupby("Category")["LineSales"].sum(),
        "FocusProfit": df[is_store].groupby("Category")["Profit"].sum(),
        "RestSales": df[~is_store].groupby("Category")["LineSales"].sum(),
        "RestProfit": df[~is_store].groupby("Category")["Profit"].sum(),
    }).reindex(cats).fillna(0.0)
    cat["FocusMargin"] = _margin(cat.FocusProfit, cat.FocusSales)
    cat["RestMargin"] = _margin(cat.RestProfit, cat.RestSales)
    cat["FocusMix"] = cat.FocusSales / cat.FocusSales.sum()
    cat["RestMix"] = cat.RestSales / cat.RestSales.sum()
    margin = cat.FocusMargin.fillna(cat.RestMargin)
    cat["Contribution"] = cat.FocusSales.sum() * (cat.RestMix - cat.FocusMix) * margin

    actual = cat.FocusProfit.sum() / cat.FocusSales.sum()
    at_rest_mix = float((margin * cat.RestMix).sum())
    gainers = cat.Contribution[cat.Contribution > 0].sort_values(ascending=False)
    return {
        "categories": cat,
        "driver_categories": list(gainers.index[:2]),   # under-weighted in the store, biggest profit effect
        "overweighted_category": str(cat.Contribution.idxmin()),
        "margin_actual": actual,
        "margin_at_rest_mix": at_rest_mix,
        "mix_effect": float(cat.FocusSales.sum() * (at_rest_mix - actual)),
        "rate_effect": float(cat.FocusSales.sum() * (cat.RestMix * (cat.RestMargin - margin)).sum()),
    }


def attach_rates(df: pd.DataFrame, store: str) -> tuple[pd.DataFrame, pd.Series]:
    """Share of receipts containing each category: per store, and pooled over the other stores."""
    cats = categories_by_sales(df)
    lines = df.groupby("Bill").agg(StoreName=("StoreName", "first"), Lines=("LineNum", "size"))
    has = pd.crosstab(df.Bill, df.Category).reindex(columns=cats).gt(0).join(lines)
    per_store = has.groupby("StoreName")[cats].mean()
    per_store["LinesPerBill"] = has.groupby("StoreName")["Lines"].mean()
    gross = df.SaleQuantity.clip(lower=0) * df.UnitPrice
    per_store["DiscountRate"] = df.groupby("StoreName")["DiscountAmount"].sum() / gross.groupby(df.StoreName).sum()
    return per_store, has[has.StoreName != store][cats].mean()


# ============================================================================ Q1
def q1_stores(df: pd.DataFrame) -> dict:
    """Compare stores, choose the one to improve, and explain its gap by category mix."""
    stores = store_table(df)
    focus, reason = pick_store(stores)
    gap = category_gap(df, focus)
    return {
        "focus": focus,
        "focus_reason": reason,
        "stores": stores,
        "categories": gap["categories"],
        "driver_categories": gap["driver_categories"],
        "overweighted_category": gap["overweighted_category"],
        "focus_margin_actual": gap["margin_actual"],
        "focus_margin_at_rest_mix": gap["margin_at_rest_mix"],
        "focus_extra_profit_same_sales": gap["mix_effect"],
        "focus_rate_effect": gap["rate_effect"],
    }


# ============================================================================ Q2
def q2_receipt_distribution(df: pd.DataFrame) -> dict:
    """Mean vs. median receipt amount and a test of the 'bulk purchases' explanation."""
    bills = bill_totals(df)
    t = bills.Total
    ex = bills.loc[~bills.HasBulkLine, "Total"]
    lines_per_bill = df.groupby("Bill").size()
    priciest = df.groupby("Category")["LineSales"].mean().idxmax()   # category with the biggest average line
    bills["Lines"] = lines_per_bill
    bills["HasPriciest"] = df.assign(h=df.Category == priciest).groupby("Bill")["h"].max()

    def gap_within(frame, cells):
        """Receipt-weighted average of (mean - median) inside comparable groups of receipts."""
        g = frame.groupby(cells)["Total"].agg(["count", "mean", "median"])
        return float(((g["mean"] - g["median"]) * g["count"]).sum() / g["count"].sum())

    no_bulk = bills[~bills.HasBulkLine]
    overall_gap = float(t.mean() - t.median())
    steps = pd.DataFrame({
        "Gap": [overall_gap,
                float(ex.mean() - ex.median()),
                gap_within(bills, "Lines"),
                gap_within(no_bulk, ["Lines", "HasPriciest"])],
    }, index=["all receipts",
              "without bulk receipts",
              "within receipts with the same number of lines",
              f"within same lines and same '{priciest}' presence, without bulk"])
    steps["ShareOfGapRemoved"] = 1 - steps.Gap / overall_gap
    return {
        "priciest_category": str(priciest),
        "gap_steps": steps,
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


# ============================================================================ Q3
def q3_category_change(df: pd.DataFrame) -> dict:
    """Chain-wide category sales and profit in Q1 vs. Q4 and the effect of the mix shift on profit."""
    sub = df[df.Quarter.isin([1, 4])]
    cats = categories_by_sales(df)
    sales = sub.pivot_table(index="Category", columns="Quarter", values="LineSales", aggfunc="sum").reindex(cats)
    profit = sub.pivot_table(index="Category", columns="Quarter", values="Profit", aggfunc="sum").reindex(cats)
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
    return {
        "categories": t,
        "biggest_decline": str(t.SalesChange.idxmin()),
        "biggest_growth": str(t.SalesChange.idxmax()),
        "total_sales_change_pct": float(t.Sales_Q4.sum() / t.Sales_Q1.sum() - 1),
        "total_profit_change_pct": float(t.Profit_Q4.sum() / t.Profit_Q1.sum() - 1),
        "margin_Q1": float(t.Profit_Q1.sum() / t.Sales_Q1.sum()),
        "margin_Q4": float(t.Profit_Q4.sum() / t.Sales_Q4.sum()),
    }


# ============================================================================ Q4
def q4_open_investigation(df: pd.DataFrame) -> dict:
    """Open investigation: the largest store performance gap in the data, tested hypothesis by hypothesis.

    The brief lets us pick the gap; we pick the store with the largest profit upside (same helper
    as Q1, applied to the data, not Q1's result) and test: is it pricing, basket size, the
    customers, or what the store sells?
    """
    stores = store_table(df)
    subject, _ = pick_store(stores)
    gap = category_gap(df, subject)
    drivers = gap["driver_categories"]
    n_other_stores = df.StoreName.nunique() - 1

    # Hypotheses 1-3: pricing/discounts, basket size, category attach rate.
    attach, rest_attach = attach_rates(df, subject)
    price_varies = int((df.groupby("ItemCode")["UnitPrice"].nunique() > 1).sum())   # expected: none

    # Hypothesis 4: a different customer base. Customers who bought both in the subject store and
    # elsewhere are the same people, so any difference between the two baskets is store-driven.
    ident = df[df.CustomerCode > 0].assign(IsSubject=lambda d: d.StoreName == subject)
    both = ident.groupby("CustomerCode")["IsSubject"].nunique()
    both_ids = both[both == 2].index
    common = ident[ident.CustomerCode.isin(both_ids)]
    common_mix = (common.pivot_table(index="IsSubject", columns="Category", values="LineSales", aggfunc="sum")
                  .pipe(lambda x: x.div(x.sum(axis=1), axis=0)))
    multi_store = (ident.groupby("CustomerCode")["StoreCode"].nunique() > 1).mean()

    # Are the driver categories low across the board or on a few items?
    if drivers:
        items = df[df.Category.isin(drivers)].pivot_table(
            index="ItemCode", columns=df.StoreName == subject, values="SaleQuantity", aggfunc="sum")
        item_ratio = items[True] / (items[False] / n_other_stores)   # vs. an average other store
    else:  # the gap is not a category-mix gap, so there is nothing to test item by item
        item_ratio = pd.Series(dtype=float)

    # Is the subject store behind in the chain's biggest trend (the Q1-vs-Q4 movers)?
    movers = q3_category_change(df)
    group = df.StoreName.where(df.StoreName == subject, "Rest of chain")
    trend = {}
    for key in ("biggest_decline", "biggest_growth"):
        cat = movers[key]
        s = df.assign(Group=group, Is=(df.Category == cat) * df.LineSales)
        trend[cat] = (s.groupby(["Group", "Quarter"])["Is"].sum() / s.groupby(["Group", "Quarter"])["LineSales"].sum()).unstack()

    return {
        "subject": subject,
        "driver_categories": drivers,
        "mix_effect": gap["mix_effect"],
        "rate_effect": gap["rate_effect"],
        "attach_rates": attach,
        "rest_attach_rates": rest_attach,
        "list_price_varies_across_stores_items": price_varies,
        "common_customers": int(len(both_ids)),
        "customers_in_multiple_stores_share": float(multi_store),
        "common_customer_category_mix": common_mix.rename(index={True: subject, False: "Rest of chain"}).reindex(columns=categories_by_sales(df)),
        "item_volume_vs_expected": item_ratio.describe(),
        "items_compared": int(len(item_ratio)),
        "share_by_quarter_of_chain_movers": trend,
    }


# ============================================================================ Q5
def q5_kpi_baselines(df: pd.DataFrame) -> pd.DataFrame:
    """Baselines of the KPIs of the two recommended actions, computed from the data.

    Action 1 - fix the category mix of the store with the largest profit upside.
    Action 2 - stop selling below cost (price/discount/cost-master guard at the till).
    """
    stores = store_table(df)
    subject, _ = pick_store(stores)
    gap = category_gap(df, subject)
    attach, rest = attach_rates(df, subject)
    cat = gap["categories"]
    below_cost = df[(df.SaleQuantity > 0) & (df.LineSales < df.LineCost)]
    rows = [
        (f"1. {c} attach rate, {subject}", f"% of receipts containing {c}", attach.loc[subject, c], "up", "monthly",
         f"rest of chain {rest[c]:.1%}")
        for c in gap["driver_categories"]
    ]
    rows += [
        (f"1. Gross margin, {subject}", "profit / net sales", gap["margin_actual"], "up", "monthly",
         f"rest of chain {cat.RestProfit.sum() / cat.RestSales.sum():.1%}"),
        ("2. Sale lines below cost per month", "count of lines with quantity > 0 and LineSales < LineCost",
         len(below_cost) / 12, "down", "monthly", f"2025 monthly average; {below_cost.Profit.sum():,.0f} ILS lost in the year"),
    ]
    return pd.DataFrame(rows, columns=["KPI", "Calculation", "Baseline", "Desired direction", "Frequency", "Reference"])
