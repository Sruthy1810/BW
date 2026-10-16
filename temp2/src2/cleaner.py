import logging
import pandas as pd

log = logging.getLogger("remittance")


def keep_valid_rows(df: pd.DataFrame, col_a: int, col_b: int,
                    rule_mode: str = "BOTH") -> pd.DataFrame:
    a_ok = df.iloc[:, col_a] != 0
    b_ok = df.iloc[:, col_b] != 0

    if rule_mode.upper() == "ANY":
        mask = a_ok | b_ok
    else:
        mask = a_ok & b_ok

    valid = df[mask].reset_index(drop=True)
    log.info("Rule: %s | columns %d and %d | total: %d | valid: %d | removed: %d",
             rule_mode.upper(), col_a, col_b,
             len(df), len(valid), len(df) - len(valid))
    return valid