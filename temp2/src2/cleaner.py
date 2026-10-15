import logging
import pandas as pd

log = logging.getLogger("remittance")


def keep_valid_rows(df: pd.DataFrame, col_a: int, col_b: int,
                    rule_mode: str = "BOTH", fees_col: int = 7) -> pd.DataFrame:
    a = df.iloc[:, col_a]
    b = df.iloc[:, col_b]
    mode = rule_mode.upper()

    if mode == "SUM":                 # Principal + Interest not 0, or a Fees-only payment
        fees = df.iloc[:, fees_col]
        mask = ((a + b).round(2) != 0) | ((a == 0) & (b == 0) & (fees != 0))
    elif mode == "ANY":               # either one is non-zero
        mask = (a != 0) | (b != 0)
    else:                             # BOTH: both columns non-zero
        mask = (a != 0) & (b != 0)

    valid = df[mask].reset_index(drop=True)
    log.info("Rule: %s | columns %d and %d | total: %d | valid: %d | removed: %d",
             mode, col_a, col_b, len(df), len(valid), len(df) - len(valid))
    return valid