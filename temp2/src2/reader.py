import re
import logging
import pandas as pd

log = logging.getLogger("remittance")
CLAIM_RE = re.compile(r"^\d{7}\.\d{3}$")


def _to_number(value):
    try:
        return float(str(value).replace("$", "").replace(",", "").strip())
    except ValueError:
        return None


def _payment_row_record(values, columns):
    if len(values) != 9:
        return None

    # A real payment row starts with a name and a plaintiff (text, not numbers)
    if _to_number(values[0]) is not None or _to_number(values[1]) is not None:
        return None

    numbers = [_to_number(value) for value in values[3:]]
    if None in numbers:
        return None

    total, principal, costs, client_share, firm_share, interest = numbers
    amounts = {
        "Payments Received": total,
        "Direct Payments": 0,
        "Principal": principal,
        "Interest": interest,
        "Fees": 0,
        "Costs": costs,
        "Other": 0,
        "Costs Expended": 0,
        "Client Proceeds": client_share,
        "Fees Due": firm_share,
    }
    return {
        "Claim": "",
        "Debtor Name": str(values[0]).strip(),
        "Ref Acct #": str(values[2]).strip(),
        **{column: amounts[column] for column in columns},
        "Plaintiff": str(values[1]).strip(),
    }


def _claim_text(value):
    if isinstance(value, float):
        return f"{value:.3f}"
    return str(value).strip()


def _pick_sheet(xls: pd.ExcelFile, wanted: str) -> str:
    names = xls.sheet_names
    log.info("Sheets found: %s", names)
    for name in names:
        if wanted.lower() in name.lower():
            return name
    log.warning("Sheet '%s' not found, scanning all sheets for debtor rows", wanted)
    best, best_count = names[0], -1
    for name in names:
        raw = xls.parse(name, header=None)
        count = sum(
            1 for v in raw.iloc[:, :3].values.flatten()
            if pd.notna(v) and CLAIM_RE.match(_claim_text(v))
        )
        if count > best_count:
            best, best_count = name, count
    return best


def _is_totals_row(vals, n):
    """Totals row: n numeric values, and the first one is not a claim number."""
    if len(vals) != n:
        return False
    if CLAIM_RE.match(_claim_text(vals[0])):
        return False
    return all(_to_number(v) is not None for v in vals)


def read_report(path: str, sheet_name: str, columns: list) -> pd.DataFrame:
    log.info("Reading file: %s", path)
    with pd.ExcelFile(path) as xls:
        sheet = _pick_sheet(xls, sheet_name)
        log.info("Using sheet: %s", sheet)
        raw = xls.parse(sheet, header=None)

    n = len(columns)
    records = []

    for idx, row in raw.iterrows():
        vals = [v for v in row.tolist() if pd.notna(v) and str(v).strip() != ""]
        if not vals:
            continue

        # Stop at the totals row: do not read it or anything after it
        if _is_totals_row(vals, n):
            log.info("Totals row found at row %d, skipping it: %s", idx + 1, vals)
            continue

        first = _claim_text(vals[0])

        if CLAIM_RE.match(first):
            if len(vals) < n + 2:
                log.warning("Skipping short row for claim %s: %s", first, vals)
                continue
            numbers = [_to_number(v) for v in vals[-n:]]
            if None in numbers:
                log.warning("Skipping row with non-numeric values, claim %s", first)
                continue
            record = {
                "Claim": first,
                "Debtor Name": str(vals[1]).strip(),
                "Ref Acct #": str(vals[2]).strip() if len(vals) > n + 2 else "",
            }
            record.update(dict(zip(columns, numbers)))
            record["Plaintiff"] = ""
            records.append(record)

        elif not first.startswith("Plaintiff:"):
            record = _payment_row_record(vals, columns)
            if record is not None:
                records.append(record)

        elif first.startswith("Plaintiff:") and records:
            records[-1]["Plaintiff"] = first.replace("Plaintiff:", "").strip()

    df = pd.DataFrame(records)
    log.info("Debtor rows read: %d", len(df))
    return df