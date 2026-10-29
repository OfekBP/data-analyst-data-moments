"""Loading and cleaning the sales data."""
import pandas as pd

from . import config


def load_raw(path=config.RAW_XLSX) -> pd.DataFrame:
    """Read the Sales_Data sheet exactly as delivered."""
    return pd.read_excel(path, sheet_name=config.SALES_SHEET)


def clean(raw: pd.DataFrame) -> pd.DataFrame:
    """Drop duplicated receipt lines and add the derived columns used everywhere.

    Only exact duplicates of (StoreCode, BillNum, LineNum) are removed. Returns,
    bulk lines and unidentified customers are real business events and are kept.
    """
    df = raw.loc[~raw.duplicated(config.LINE_KEY)].copy()
    # LineSales is already net of the line discount and LineCost is the recorded
    # line cost, so profit is a plain difference (no second discount deduction).
    df["Profit"] = df["LineSales"] - df["LineCost"]
    df["Quarter"] = df["BillDate"].dt.quarter
    df["Bill"] = df["StoreCode"].astype(str) + "-" + df["BillNum"].astype(str)
    return df.reset_index(drop=True)


def load_clean() -> pd.DataFrame:
    return clean(load_raw())
