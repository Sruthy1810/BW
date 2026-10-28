"""Turns raw report rows into payment records.

Each payment is a 2-row block:
  row 1: [customer name | posting date | 'PMT: Debtor Payment - <origination> - Release Dt ...']
  row 2: [loan no. | amount ... | 'Ref #:' | account number]
An optional 'Over Payment:' row after a block reduces that payment's gross amount.
"""
import logging
import re
from datetime import date, datetime
from decimal import Decimal


from temp3_Firm_Payments.src4.settings import (
    COL_DATE_OR_AMOUNT,
    COL_DESCRIPTION,
    COL_NAME_OR_ACCOUNT,
    COL_REF,
)

BUSINESS_WORDS = {"llc", "llp", "inc", "corp", "co", "ltd", "pc", "pllc", "lp", "company"}


def format_name(name: str) -> str:
    """'Smith, John A' -> 'John A Smith'. Names without a comma are left as they are."""
    name = re.sub(r"\s+", " ", str(name)).strip()
    if "," not in name:
        return name
    last, first = [p.strip() for p in name.split(",", 1)]
    if not last or not first:
        return name
    if first.lower().rstrip(".") in BUSINESS_WORDS:    # 'Acme, Inc.' is a business, not a person
        return name
    return f"{first} {last}"

log = logging.getLogger("remittance")


def _cell(row, idx):
    return row[idx] if idx < len(row) else None


def _text(v) -> str:
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def _decimal(v) -> Decimal:
    return Decimal(str(v)) if v not in (None, "") else Decimal("0")


def _account(v):
    s = _text(v)
    s = s[:-2] if s.endswith(".0") else s
    return int(s) if s.isdigit() else s


def _origination(desc: str) -> str:
    m = re.search(r"Debtor Payment\s*-\s*(.*?)\s*-\s*Release Dt", desc, re.I)
    return _text(m.group(1)) if m else ""


def parse_payments(rows: list[list]) -> list[dict]:
    payments, i = [], 0
    while i < len(rows):
        row = rows[i]
        desc = _text(_cell(row, COL_DESCRIPTION))

        if desc.upper().startswith("PMT:") and isinstance(_cell(row, COL_DATE_OR_AMOUNT), (datetime, date)):
            if i + 1 >= len(rows):
                log.warning("Row %d: payment header without amount row - skipped", i + 1)
                break
            amt_row = rows[i + 1]
            ref = _cell(amt_row, COL_REF)
            payment = {
                "src_row": i + 1,
                "date": _cell(row, COL_DATE_OR_AMOUNT),
                "name": format_name(_text(_cell(row, COL_NAME_OR_ACCOUNT))),
                "origination": _origination(desc),
                "account": _account(ref if ref not in (None, "") else _cell(amt_row, COL_NAME_OR_ACCOUNT)),
                "amount": _decimal(_cell(amt_row, COL_DATE_OR_AMOUNT)),
                "overpayment": Decimal("0"),
            }
            log.debug("Row %d: parsed %s", i + 1, payment)
            payments.append(payment)
            i += 2
            continue

        if _text(_cell(row, COL_NAME_OR_ACCOUNT)).lower().startswith("over payment:") and payments:
            op = _decimal(_cell(row, COL_DATE_OR_AMOUNT))
            payments[-1]["overpayment"] += op
            log.info("Row %d: over payment %s applied to %s", i + 1, op, payments[-1]["name"])
        i += 1
    return payments