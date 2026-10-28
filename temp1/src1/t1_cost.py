import datetime as dt
import warnings
import pandas as pd

from temp1.src1.t1_parser import to_number, clean_text, cell, format_name   # adjust to your layout
from temp1.src1.t1_reader import excel_engine

MARKER = "cost detail report"

# 0-based column positions (A=0, B=1, ...)
DATE_COL = 0      # A -> Date
NAME_COL = 2      # C -> defendant name
REF_COL = 3       # D -> ref no.
REASON_COL = 4    # E -> cost reason
AMOUNT_COL = 5    # F -> amount

# Where does a note line like "Court Fee" belong?
#   "previous"  -> it is the note of the data row ABOVE it
#   "following" -> it is the note of the data rows BELOW it, until the next note line
NOTE_MODE = "previous"
STOP_TEXT = "total costs expended"

STOP_TEXT = "total"          # stop at a row containing this word
SKIP_STARTS = ("page ",)     # page footers etc. are not notes


def find_cost_rows(input_file, log):
    """Find the 'Cost Detail Report' row (on any sheet) and return the rows after it."""
    sheets = pd.read_excel(input_file, sheet_name=None, header=None, dtype=object,
                           engine=excel_engine(input_file))
    for sheet_name, df in sheets.items():
        for pos in range(len(df)):
            cells = [str(v).lower() for v in df.iloc[pos].tolist() if pd.notna(v)]
            if any(MARKER in c for c in cells):
                log.info("'Cost Detail Report' found on sheet '%s', row %d", sheet_name, pos + 1)
                return df.iloc[pos + 1:]
    raise ValueError(
        f"'Cost Detail Report' marker not found in {input_file}"
    )


def to_date(value):
    """Excel date / '01/15/2024' text -> date. None if it is not a date."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str) and value.strip():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            d = pd.to_datetime(value.strip(), errors="coerce")
        if pd.notna(d):
            return d.date()
    return None


def parse_cost_rows(rows, log):
    records = []
    pending_note = ""

    for idx, row in rows.iterrows():
        text = " ".join(str(v) for v in row.tolist() if pd.notna(v)).lower()
        if not text.strip():
            continue                                     # blank row
        if STOP_TEXT in text:
            log.info("Cost totals row reached (row %d): %s - stopping", idx + 1, text[:100])
            break
        if text.startswith(SKIP_STARTS):
            continue

        a_val = cell(row, DATE_COL)
        a_text = clean_text(a_val)
        name = clean_text(cell(row, NAME_COL))
        ref = clean_text(cell(row, REF_COL))
        amount = to_number(cell(row, AMOUNT_COL))

        # ---- data row ----
        if amount is not None and (name or ref):
            date = to_date(a_val)
            note = ""
            if date is None and a_text:                  # text in column A on the same row
                note = a_text
            if not note and NOTE_MODE == "following":
                note = pending_note
            records.append({
                "name": format_name(name),
                "ref_no": ref,
                "amount": amount,
                "reason": clean_text(cell(row, REASON_COL)),
                "date": date,
                "note": note,
            })
            continue

        # ---- note line: text only in column A, everything else empty ----
        others_empty = all(pd.isna(v) for i, v in row.items() if i != DATE_COL)
        if others_empty and a_text and to_date(a_val) is None:
            if NOTE_MODE == "previous":
                if records:
                    prev = records[-1]
                    prev["note"] = f"{prev['note']}; {a_text}" if prev["note"] else a_text
            else:
                pending_note = a_text

    log.info("Cost rows: %d", len(records))
    return pd.DataFrame(records, columns=["name", "ref_no", "amount", "reason", "date", "note"])