from pathlib import Path
import pandas as pd

START_MARKER = "remittance report"


class StartMarkerNotFoundError(ValueError):
    """Raised when an input workbook does not contain the remittance report marker."""


def excel_engine(path):
    return "xlrd" if Path(path).suffix.lower() == ".xls" else None


def get_input_file(config):
    """Single input file from config key 'input_file_temp1' (used only if no folder is set)."""
    path = Path(str(config["input_file_temp1"]).strip().strip('"'))
    if not path.exists():
        raise FileNotFoundError(f"Temp1 input file not found: {path}")
    return path

def get_input_files(config):
    """All .xls/.xlsx files in input_folder_temp1 (falls back to the single input_file_temp1)."""
    folder = config.get("input_folder_temp1", "")
    if not folder:
        return [get_input_file(config)]

    folder = Path(str(folder).strip().strip('"'))
    if not folder.is_dir():
        raise FileNotFoundError(f"Temp1 input folder not found: {folder}")

    files = sorted(f for f in folder.iterdir()
                   if f.is_file() and f.suffix.lower() in (".xls", ".xlsx")
                   and not f.name.startswith("~$"))          # ~$ = temp files Excel creates
    if not files:
        raise FileNotFoundError(f"No Excel files found in: {folder}")
    return files


def read_raw(input_file):
    """Read the first sheet with no header, so every row is raw data."""
    return pd.read_excel(input_file, sheet_name=0, header=None, dtype=object,
                         engine=excel_engine(input_file))


def rows_after_marker(df):
    """Return all rows after the first row that contains 'Remittance Report'."""
    for pos in range(len(df)):
        cells = [str(v).lower() for v in df.iloc[pos].tolist() if pd.notna(v)]
        if any(START_MARKER in c for c in cells):
            return df.iloc[pos + 1:].reset_index(drop=True)
    raise StartMarkerNotFoundError("'Remittance Report' row not found in the input file")