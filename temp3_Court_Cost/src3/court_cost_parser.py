import datetime as dt
import re
import warnings
from pathlib import Path

import pandas as pd

from temp1.src1.t1_parser import cell, clean_text, to_number
from temp1.src1.t1_reader import excel_engine

START_MARKER = "cost detail report"
END_MARKER = "total costs expended"
REMITTANCE_MARKER = "remittance report"
REMITTANCE_END_MARKER = "bessine walterbach, llp"
REMITTANCE_ACCOUNT_COL = 2

DEFAULT_COLUMNS = {
    "accountnumber": 3,
    "debtorname": 2,
    "effectivedate": 0,
    "transactionamt": 5,
    "costreason": 4,
    "placementdate": 5,
    "judgmentdate": 6,
}

OUTPUT_COLUMNS = [
    "Account Number",
    "Debtor Name",
    "Effective Date",
    "Transaction Amt.",
    "Cost Reason",
    "Placement date",
    "Judgment Date"
]
JUDGMENT_HIGHLIGHT_COLUMN = "_highlight_judgment"


def _column_key(value):
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def _canonical_cost_column(value):
    key = _column_key(value)
    aliases = {
        "account": "accountnumber",
        "accountno": "accountnumber",
        "acct": "accountnumber",
        "acctno": "accountnumber",
        "acctnumber": "accountnumber",
        "amount": "transactionamt",
        "costdescription": "costreason",
        "date": "effectivedate",
    }
    return aliases.get(key, key)


def find_cost_rows(input_file, log):
    """Return rows after the Cost Detail Report marker on any worksheet."""
    sheets = pd.read_excel(
        input_file,
        sheet_name=None,
        header=None,
        dtype=object,
        engine=excel_engine(input_file),
    )
    for sheet_name, frame in sheets.items():
        for position in range(len(frame)):
            values = [str(value).lower() for value in frame.iloc[position] if pd.notna(value)]
            if any(START_MARKER in value for value in values):
                log.info(
                    "Cost Detail Report found on sheet '%s', row %d",
                    sheet_name,
                    position + 1,
                )
                return frame.iloc[position + 1:]
    raise ValueError(f"Cost Detail Report marker not found in {input_file}")


def to_date(value):
    """Convert Excel dates or common US date strings to date objects."""
    if value is None or pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value

    text = clean_text(value)
    if not text:
        return None

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=False)
    return parsed.date() if pd.notna(parsed) else None


def _account_number_text(value):
    text = clean_text(value)
    if re.fullmatch(r"\d+\.0+", text):
        text = text[:text.index(".")]
    return text


def _account_number_key(value):
    text = _account_number_text(value)
    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    return digits.zfill(10) if len(digits) < 10 else digits


def _format_cost_reason(value):
    reason = clean_text(value)
    return re.sub(
        r"\bEGRN\s*:\s*E-filed\b",
        "EGRN: E-Filed",
        reason,
        flags=re.IGNORECASE,
    )


def find_remittance_dates(input_file, log):
    """Map the Remittance Report claim/account number to its Open and Judgment dates."""
    sheets = pd.read_excel(
        input_file,
        sheet_name=None,
        header=None,
        dtype=object,
        engine=excel_engine(input_file),
    )
    for sheet_name, frame in sheets.items():
        marker_position = None
        for position in range(len(frame)):
            text = " ".join(
                str(value).lower() for value in frame.iloc[position] if pd.notna(value)
            )
            if REMITTANCE_MARKER in text:
                marker_position = position
                break

        if marker_position is None:
            continue

        dates_by_account = {}
        current_account = ""
        for position in range(marker_position + 1, len(frame)):
            row = frame.iloc[position]
            row_values = [value for value in row.tolist() if pd.notna(value)]
            row_text = " ".join(str(value) for value in row_values)
            if REMITTANCE_END_MARKER in row_text.lower():
                log.info(
                    "Remittance end marker found on sheet '%s', row %d",
                    sheet_name,
                    position + 1,
                )
                break

            account = _account_number_key(cell(row, REMITTANCE_ACCOUNT_COL))
            if account:
                current_account = account

            if not current_account:
                continue

            open_match = re.search(
                r"\bopen\s+date\s*[:\-]\s*"
                r"(\d{1,4}[/-]\d{1,2}[/-]\d{1,4})",
                row_text,
                flags=re.IGNORECASE,
            )
            judgment_match = re.search(
                r"\bjudg(?:e)?ment\s+date\s*[:\-]\s*"
                r"(\d{1,4}[/-]\d{1,2}[/-]\d{1,4})",
                row_text,
                flags=re.IGNORECASE,
            )
            if open_match or judgment_match:
                record = dates_by_account.setdefault(
                    current_account, {"open_date": None, "judgment_date": None}
                )
                if open_match:
                    record["open_date"] = to_date(open_match.group(1))
                if judgment_match:
                    record["judgment_date"] = to_date(judgment_match.group(1))

        log.info("Remittance date records found: %d", len(dates_by_account))
        return dates_by_account

    log.warning("Remittance Report marker not found; using dates from Cost Detail Report")
    return {}


def _cost_columns(rows):
    for _, row in rows.iterrows():
        normalized = {_canonical_cost_column(value): index for index, value in row.items()
                      if pd.notna(value)}
        if "accountnumber" in normalized and "debtorname" in normalized:
            return {**DEFAULT_COLUMNS, **normalized}
    return DEFAULT_COLUMNS


def parse_cost_rows(rows, log, remittance_dates=None):
    records = []
    columns = _cost_columns(rows)
    remittance_dates = remittance_dates or {}

    for index, row in rows.iterrows():
        values = [value for value in row.tolist() if pd.notna(value)]
        text = " ".join(str(value) for value in values).strip()
        if not text:
            continue

        lowered = text.lower()
        if END_MARKER in lowered:
            log.info("Cost totals row reached (row %d): stopping", index + 1)
            break

        canonical = {_canonical_cost_column(value) for value in values}
        if "accountnumber" in canonical and "debtorname" in canonical:
            continue

        account_value = cell(row, columns["accountnumber"])
        account = _account_number_text(account_value)
        lookup_account = _account_number_key(account_value)
        name = clean_text(cell(row, columns["debtorname"]))
        amount = to_number(cell(row, columns["transactionamt"]))
        if amount is None or not (account or name):
            continue

        if len(_account_number_key(account)) > 10:
            log.warning("Account number has more than 10 digits at row %d", index + 1)

        placement_date = to_date(cell(row, columns["placementdate"]))
        judgment_date = to_date(cell(row, columns["judgmentdate"]))
        highlight_judgment = False
        report_dates = remittance_dates.get(lookup_account)
        if report_dates is not None:
            has_remittance_judgment = report_dates["judgment_date"] is not None
            placement_date = (
                report_dates["judgment_date"]
                if has_remittance_judgment
                else report_dates["open_date"]
            )
            judgment_date = (
                report_dates["open_date"] if has_remittance_judgment else None
            )
            highlight_judgment = has_remittance_judgment
        else:
            log.warning(
                "No Remittance Report date record matched account %s (%s)",
                account,
                name,
            )

        records.append({
            "Account Number": account,
            "Debtor Name": name,
            "Effective Date": to_date(cell(row, columns["effectivedate"])),
            "Transaction Amt.": amount,
            "Cost Reason": _format_cost_reason(cell(row, columns["costreason"])),
            "Placement date": placement_date,
            "Judgment Date": judgment_date,
            JUDGMENT_HIGHLIGHT_COLUMN: highlight_judgment,
        })

    log.info("Court cost rows: %d", len(records))
    return pd.DataFrame(records, columns=OUTPUT_COLUMNS + [JUDGMENT_HIGHLIGHT_COLUMN])


def read_costs(input_file, log):
    path = Path(input_file)
    remittance_dates = find_remittance_dates(path, log)
    return parse_cost_rows(find_cost_rows(path, log), log, remittance_dates)
