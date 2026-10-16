import math
from pathlib import Path

import pandas as pd

MASTER_SHEET = "IntrestPercentage_Temp2"
MASTER_COLUMNS = ["CLIENT NO", "ClientShare", "BWShare", "BW Interest %"]


def get_client_percentages(master_file: str, client_name: str) -> tuple[float, float, float]:
    path = Path(master_file)
    if not path.is_file():
        raise FileNotFoundError(f"Master data workbook was not found: {path}")

    table = pd.read_excel(path, sheet_name=MASTER_SHEET)
    missing_columns = set(MASTER_COLUMNS).difference(table.columns)
    if missing_columns:
        raise ValueError(
            f"Master sheet {MASTER_SHEET!r} is missing columns: "
            f"{', '.join(sorted(missing_columns))}"
        )

    client_keys = table["CLIENT NO"].map(_normalize_client_name)
    wanted_key = _normalize_client_name(client_name)
    matches = table.loc[client_keys == wanted_key]
    if matches.empty:
        raise LookupError(
            f"Client {client_name!r} from the input filename was not found "
            f"in master sheet {MASTER_SHEET!r}"
        )
    if len(matches) != 1:
        raise ValueError(
            f"Client {client_name!r} has {len(matches)} matching rows in "
            f"master sheet {MASTER_SHEET!r}; exactly one is required"
        )

    row = matches.iloc[0]
    values = tuple(
        _read_percentage(row[column], column, client_name)
        for column in MASTER_COLUMNS[1:]
    )
    print(values)
    return values


def _normalize_client_name(value) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().casefold().split())


def _read_percentage(value, column: str, client_name: str) -> float:
    try:
        percentage = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"Invalid {column} value for client {client_name!r}: {value!r}"
        ) from exc
    if not math.isfinite(percentage) or not 0 <= percentage <= 1:
        raise ValueError(
            f"{column} for client {client_name!r} must be between 0 and 1; "
            f"got {value!r}"
        )
    return percentage
