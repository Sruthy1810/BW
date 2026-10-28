"""Fixed business rules (not file paths - those live in config.xlsx)."""
from decimal import Decimal

COMMISSION_RATE = Decimal("0.25")   # 25% of gross payment
DECIMALS = Decimal("0.001")         # round to 3 decimal places
FIRST_DATA_ROW = 2                  # first data row in the template

# Selects this flow's folder keys from the shared key/value config.xlsx.
PROJECT_NAME = "Firm Payments"

# Column positions (0-based) in the raw input report
COL_NAME_OR_ACCOUNT = 0
COL_DATE_OR_AMOUNT = 1
COL_DESCRIPTION = 2
COL_REF = 13

MONEY_FMT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
DATE_FMT = "mm-dd-yy"

# ---------------------------------------------------------------------------
# Output layouts. The layout is chosen automatically by matching the header row
# of the template. Each column = (field, number format).
# Fields: account, name, date, gross, net, commission, origination
# ---------------------------------------------------------------------------
LAYOUTS = {
    "Firm Payment": {
        "headers": ["Posting Date", "account number", "Customer Name", "Gross Payment",
                    "Net Payment", "Commision Amt.", "Payment Origination"],
        "columns": [("date", DATE_FMT), ("account", "General"), ("name", "General"),
                    ("gross", MONEY_FMT), ("net", MONEY_FMT), ("commission", MONEY_FMT),
                    ("origination", "General")],
    },
    "Non-Attorney Payment": {
        "headers": ["ACCOUNT NUMBER", "FULL NAME", "POSTING DATE",
                    "AMOUNT OF NON ATTORNEY PAYMENT", "COMMISION AMOUNT"],
        "columns": [("account", "General"), ("name", "General"), ("date", DATE_FMT),
                    ("gross", "General"), ("commission", "General")],
    },
}