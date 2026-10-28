import re
import logging
import pandas as pd
from temp2.src2.reader import excel_engine

log = logging.getLogger("remittance")

CLAIM_RE = re.compile(r"^\d{7}\.\d{3}$")
COST_COLUMNS = ["DEFENDANT'S NAME", "PLAINTIFF", "REF NO.", "AMOUNT",
                "COST REASON", "NOTES"]


def _claim_text(value):
    if isinstance(value, float):
        return f"{value:.3f}"
    return str(value).strip()


def _to_number(value):
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def _find_identifier(xls, identifier):
    for name in xls.sheet_names:
        raw = xls.parse(name, header=None)
        for idx in range(len(raw)):
            cells = [str(v) for v in raw.iloc[idx].tolist() if pd.notna(v)]
            if any(identifier.lower() in c.lower() for c in cells):
                log.info("Identifier '%s' found on sheet '%s', row %d",
                         identifier, name, idx + 1)
                return raw, idx
    raise ValueError(f"Identifier '{identifier}' not found in any sheet")


def read_costs(path, identifier, end_marker, plaintiff_map) -> pd.DataFrame:
    with pd.ExcelFile(path, engine=excel_engine(path)) as xls:
        raw, start = _find_identifier(xls, identifier)

    records = []
    for idx in range(start + 1, len(raw)):
        vals = [v for v in raw.iloc[idx].tolist()
                if pd.notna(v) and str(v).strip() != ""]
        if not vals:
            continue
        first = str(vals[0]).strip()

        if first.lower().startswith(end_marker.lower()):
            break

        if len(vals) >= 6 and CLAIM_RE.match(_claim_text(vals[1])):
            claim = _claim_text(vals[1])
            amount = _to_number(vals[-1])
            if amount is None:
                log.warning("Skipping cost row %d, bad amount: %s", idx + 1, vals)
                continue
            plaintiff = plaintiff_map.get(claim, "")
            if not plaintiff:
                log.warning("No plaintiff found for claim %s", claim)
            records.append({
                "DEFENDANT'S NAME": str(vals[2]).strip(),
                "PLAINTIFF": plaintiff,
                "REF NO.": str(vals[3]).strip(),
                "AMOUNT": amount,
                "COST REASON": str(vals[4]).strip(),
                "NOTES": "",
            })
        elif len(vals) == 1 and records:
            records[-1]["NOTES"] = first     # note line under a cost row

    df = pd.DataFrame(records, columns=COST_COLUMNS)
    log.info("Cost rows read: %d | total: %.2f", len(df), df["AMOUNT"].sum())
    return df