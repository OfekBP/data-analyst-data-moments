"""Project-wide paths and analysis parameters."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RAW_XLSX = ROOT / "data" / "raw" / "analyst_exam.xlsx"
SALES_SHEET = "Sales_Data"
OUTPUTS = ROOT / "outputs"
TABLES = OUTPUTS / "tables"
SQL_DIR = ROOT / "sql"
EXCEL_PATH = ROOT / "excel" / "DataMoments_Analyst_Calculations.xlsx"

# Primary key of a receipt line (BillNum alone is not unique across stores).
LINE_KEY = ["StoreCode", "BillNum", "LineNum"]
BILL_KEY = ["StoreCode", "BillNum"]

# A receipt line with this many units or more is treated as a bulk purchase.
# The normal quantity is 1-2 (or -1 for a return); bulk lines are 20-45 units.
BULK_QTY = 20

# Store the analysis zooms in on (answer to question 1, drill-down in question 4).
FOCUS_STORE = "Haifa"

CATEGORIES = ["Dry Food", "Wet Food", "Snacks", "Accessories", "Grooming", "Litter", "Toys"]
