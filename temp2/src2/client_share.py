import logging
import pandas as pd
from temp2.src2.reader import excel_engine

log = logging.getLogger("remittance")


def _to_fraction(value) -> float:
    """Accepts 0.85, 85, 85.0 or '85.0%' and returns 0.85."""
    if isinstance(value, str):
        x = float(value.replace("%", "").strip()) / 100
    else:
        x = float(value)
        if x > 1:
            x = x / 100
    return round(x, 6)


def get_client_share(master_path, sheet_name, client_name, share_column) -> float:
    xls = pd.ExcelFile(master_path, engine=excel_engine(master_path))
    lookup = {s.strip().lower(): s for s in xls.sheet_names}
    real = lookup.get(sheet_name.strip().lower())
    if real is None:
        raise ValueError(f"Sheet '{sheet_name}' not found. Sheets: {xls.sheet_names}")
    df = xls.parse(real)
    df.columns = [str(c).strip() for c in df.columns]

    cols = {c.lower(): c for c in df.columns}
    key_col = cols.get("client no")
    value_col = cols.get(share_column.strip().lower())
    if key_col is None:
        raise ValueError(f"Column 'CLIENT NO' not found. Columns: {list(df.columns)}")
    if value_col is None:
        raise ValueError(f"Column '{share_column}' not found. Columns: {list(df.columns)}")

    names = df[key_col].astype(str).str.strip().str.lower()
    match = df[names == client_name.strip().lower()]

    if match.empty:
        raise ValueError(f"Client '{client_name}' not found in sheet '{sheet_name}'")
    if len(match) > 1:
        raise ValueError(f"Client '{client_name}' appears {len(match)} times in '{sheet_name}'")

    pct = _to_fraction(match.iloc[0][value_col])
    log.info("Client share for '%s' = %.2f%% (column '%s', sheet '%s')",
             client_name, pct * 100, value_col, sheet_name)
    return pct