from pathlib import Path
import pandas as pd

SHEET_NAME = "IntrestPercentage" 
NAME_COL, CLIENT_COL, BW_COL = 0, 1, 2     # A = file name, B = client %, C = BW %


def _norm(text):
    return " ".join(str(text).lower().split())


def _to_pct(value):
    """'70.00%' / 0.7 / 70 -> 0.70"""
    if isinstance(value, str):
        text = value.replace(",", "").strip()
        if text.endswith("%"):
            return float(text[:-1]) / 100
        value = text
    value = float(value)
    return value / 100 if value > 1 else value

def _find_sheet(master_file, wanted):
    """Match the sheet name ignoring case and extra spaces."""
    names = pd.ExcelFile(master_file).sheet_names
    for name in names:
        if _norm(name) == _norm(wanted):
            return name
    raise ValueError(f"Sheet '{wanted}' not found in MasterDatatable. Available sheets: {names}")


def get_shares(master_file, input_file, log, sheet_name=SHEET_NAME):
    """Find the input file name (without extension) in the given MasterDatatable sheet.
    Returns (client_pct, bw_pct) as fractions, e.g. (0.70, 0.30)."""
    key = _norm(Path(input_file).stem)               # 'cash link usa, llc'
    sheet = _find_sheet(master_file, sheet_name)
    df = pd.read_excel(master_file, sheet_name=sheet, header=None, dtype=object)

    for _, row in df.iterrows():
        if pd.isna(row[NAME_COL]):
            continue
        if _norm(row[NAME_COL]) == key:
            client, bw = _to_pct(row[CLIENT_COL]), _to_pct(row[BW_COL])
            if abs(client + bw - 1) > 0.0001:
                raise ValueError(f"Master shares for '{key}' do not add up to 100%")
            log.info("Master match '%s' (sheet '%s'): client %.2f%% | BW %.2f%%",
                     key, sheet, client * 100, bw * 100)
            return client, bw

    raise ValueError(f"'{key}' not found in sheet '{sheet}' of MasterDatatable: {master_file}")