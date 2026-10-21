import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

PAYMENT_COLUMNS = [
    "DEFENDANT'S NAME", "REF NO.", "AMOUNT PAID", "DIRECT PAYMENT",
    "CLIENT SHARE", "BW, LLP SHARE", "Notes",
]

COST_COLUMNS = ["DEFENDANT'S NAME", "REF NO.", "AMOUNT", "COST REASON", "Date", "Notes"]


def _share(amount, pct):
    """Calculate each allocated share to three decimal places."""
    value = Decimal(str(round(amount, 2))) * Decimal(str(round(pct, 6)))
    return float(value.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP))

def build_payment_sheet(data, client_pct, bw_pct):
    out = pd.DataFrame(columns=PAYMENT_COLUMNS)
    if data.empty:
        return out

    amount = (data["principal"] + data["interest"]).round(2)

    out["DEFENDANT'S NAME"] = data["name"]
    out["REF NO."] = data["ref_no"]
    out["AMOUNT PAID"] = amount
    out["DIRECT PAYMENT"] = data["direct_payment"]
    out["CLIENT SHARE"] = amount.apply(lambda a: _share(a, client_pct))   # each share rounded on its own
    out["BW, LLP SHARE"] = amount.apply(lambda a: _share(a, bw_pct))
    out["Notes"] = ""
    return out


def build_cost_sheet(data):
    out = pd.DataFrame(columns=COST_COLUMNS)
    if data.empty:
        return out
    out["DEFENDANT'S NAME"] = data["name"]
    out["REF NO."] = data["ref_no"]
    out["AMOUNT"] = data["amount"]
    out["COST REASON"] = data["reason"]
    out["Date"] = data["date"]
    out["Notes"] = data["note"]
    return out