import logging
import re

import pandas as pd

log = logging.getLogger("remittance")
CLAIM_RE = re.compile(r"^\d{7}\.\d{3}$")
PAYMENT_COLUMNS = [
    "Claim",
    "Debtor Name",
    "Ref Acct #",
    "Payments Received",
    "Direct Payments",
    "Principal",
    "Interest",
    "Fees",
    "Costs",
    "Other",
    "Costs Expended",
    "Client Proceeds",
    "Fees Due",
    "Plaintiff",
]
COST_COLUMNS = [
    "Date",
    "Claim",
    "Debtor Name",
    "Account #",
    "Cost Description",
    "Amount",
    "Notes",
]


def _claim_text(value):
    if isinstance(value, float):
        return f"{value:.3f}"
    return str(value).strip()


def _to_number(value):
    if pd.isna(value) or str(value).strip() == "":
        return 0.0
    text = str(value).strip().replace(",", "").replace("$", "")
    if text.startswith("(") and text.endswith(")"):
        text = f"-{text[1:-1]}"
    try:
        return float(text)
    except ValueError:
        return None


def _load_sheet(path: str, wanted: str) -> pd.DataFrame:
    log.info("Reading file: %s", path)
    workbook = pd.ExcelFile(path)
    log.info("Sheets found: %s", workbook.sheet_names)

    for name in workbook.sheet_names:
        if wanted.lower() in name.lower():
            log.info("Using sheet: %s", name)
            return workbook.parse(name, header=None)

    for name in workbook.sheet_names:
        raw = workbook.parse(name, header=None)
        if raw.astype(str).apply(
            lambda column: column.str.contains("Remittance Report", case=False, na=False)
        ).any().any():
            log.info("Using sheet containing the remittance report: %s", name)
            return raw

    raise ValueError(f"No sheet containing a remittance report was found in {path}")


def _find_marker(raw: pd.DataFrame, marker: str):
    for index, row in raw.iterrows():
        for value in row:
            if isinstance(value, str) and marker.lower() in value.lower():
                return index
    return None


def _read_payments(raw: pd.DataFrame, columns: list) -> pd.DataFrame:
    header_row = _find_marker(raw, "Remittance Report")
    if header_row is None:
        log.warning("Skipping payments: no Remittance Report section was found")
        return pd.DataFrame(columns=PAYMENT_COLUMNS)
    cost_header = _find_marker(raw, "Cost Detail Report")
    end_row = cost_header if cost_header is not None else len(raw)
    records = []
    last_record = None

    for _, row in raw.iloc[header_row + 1 : end_row].iterrows():
        first = row.iloc[0] if len(row) else None
        if pd.isna(first) or str(first).strip() == "":
            continue

        claim = _claim_text(first)
        if CLAIM_RE.fullmatch(claim):
            values = [_to_number(value) for value in row.iloc[3 : 3 + len(columns)]]
            if len(values) != len(columns) or any(value is None for value in values):
                log.warning("Skipping remittance row with invalid amounts for claim %s", claim)
                last_record = None
                continue

            record = {
                "Claim": claim,
                "Debtor Name": str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "",
                "Ref Acct #": str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
                **dict(zip(columns, values)),
                "Plaintiff": "",
            }
            records.append(record)
            last_record = record
        elif isinstance(first, str) and first.strip().lower().startswith("plaintiff:"):
            if last_record is not None:
                plaintiff_lines = [
                    re.sub(
                        r"^\s*Plaintiff(?:\s+Name\s+\d+)?\s*:\s*",
                        "",
                        line,
                        flags=re.IGNORECASE,
                    ).strip()
                    for line in first.splitlines()
                ]
                last_record["Plaintiff"] = " ".join(
                    line for line in plaintiff_lines if line
                )

    if not records:
        log.warning("Skipping payments: no debtor payment rows were found")
        return pd.DataFrame(columns=PAYMENT_COLUMNS)

    log.info("Payment rows read: %d", len(records))
    return pd.DataFrame(records, columns=PAYMENT_COLUMNS)


def _read_costs(raw: pd.DataFrame) -> pd.DataFrame:
    header_row = _find_marker(raw, "Cost Detail Report")
    if header_row is None:
        log.warning("Skipping costs: no Cost Detail Report section was found")
        return pd.DataFrame(columns=COST_COLUMNS)

    records = []
    last_record = None
    for _, row in raw.iloc[header_row + 1 :].iterrows():
        if len(row) < 6:
            continue
        claim = _claim_text(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
        if not CLAIM_RE.fullmatch(claim):
            if isinstance(row.iloc[0], str) and row.iloc[0].strip().lower().startswith("total costs"):
                break
            note = row.iloc[0]
            if (
                last_record is not None
                and isinstance(note, str)
                and note.strip()
                and all(pd.isna(value) for value in row.iloc[1:])
            ):
                last_record["Notes"] = note.strip()
            continue

        amount = _to_number(row.iloc[5])
        if amount is None:
            log.warning("Skipping cost row with invalid amount for claim %s", claim)
            continue
        raw_date = row.iloc[0]
        cost_date = pd.to_datetime(raw_date, errors="coerce")
        if pd.isna(cost_date):
            log.warning("Skipping cost row with invalid date for claim %s", claim)
            continue

        last_record = {
            "Date": cost_date,
            "Claim": claim,
            "Debtor Name": str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
            "Account #": str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else "",
            "Cost Description": str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else "",
            "Amount": amount,
            "Notes": "",
        }

        records.append(last_record)

    if not records:
        log.warning("Skipping costs: no cost detail rows were found")
        return pd.DataFrame(columns=COST_COLUMNS)
    log.info("Cost rows read: %d", len(records))
    return pd.DataFrame(records, columns=COST_COLUMNS)


def _read_statement_summary(raw: pd.DataFrame) -> pd.DataFrame:
    header_row = _find_marker(raw, "Remittance Report")
    if header_row is None:
        return pd.DataFrame(columns=["Section", "Metric", "Amount"])

    entries = []
    current_section = "Statement"
    for _, row in raw.iloc[:header_row].iterrows():
        if len(row) < 2 or pd.isna(row.iloc[0]):
            continue
        label = str(row.iloc[0]).strip()
        amount = _to_number(row.iloc[1])
        if label.lower().endswith("section"):
            current_section = label
        elif amount is not None and pd.notna(row.iloc[1]):
            entries.append(
                {"Section": current_section, "Metric": label, "Amount": amount}
            )
    return pd.DataFrame(entries, columns=["Section", "Metric", "Amount"])


def extract_remittance_address(path: str, sheet_name: str) -> str:
    """Return non-empty lines between Remittance Statement and Collection Receipts Section."""
    raw = _load_sheet(path, sheet_name)
    start_row = _find_marker(raw, "Remittance Statement")
    end_row = _find_marker(raw, "Collection Receipts Section")
    if start_row is None:
        raise ValueError("The source sheet does not contain a Remittance Statement section")
    if end_row is None:
        raise ValueError("The source sheet does not contain a Collection Receipts Section marker")
    if end_row <= start_row:
        raise ValueError("Collection Receipts Section must follow Remittance Statement")

    address_lines = []
    for _, row in raw.iloc[start_row + 1 : end_row].iterrows():
        values = [
            str(value).strip()
            for value in row
            if pd.notna(value) and str(value).strip()
        ]
        if values:
            address_lines.append(" ".join(values))

    if not address_lines:
        raise ValueError("No address lines were found between the statement markers")
    return "\n".join(address_lines)


def read_remittance_sections(path: str, sheet_name: str, columns: list):
    raw = _load_sheet(path, sheet_name)
    payments = _read_payments(raw, columns)
    costs = _read_costs(raw)
    statement_summary = _read_statement_summary(raw)
    return payments, costs, statement_summary


def read_report(path: str, sheet_name: str, columns: list) -> pd.DataFrame:
    """Read debtor rows from the Remittance Report section."""
    raw = _load_sheet(path, sheet_name)
    return _read_payments(raw, columns)
