from pathlib import Path
import pandas as pd

CONFIG_PATH = Path("config/config.xlsx")

DEFAULTS = [
    ("input_file", "input/Consumer Collection Management, Inc. 2.xls"),
    ("sheet_name", "Remittance Report"),
    ("output_file", "output/valid_remittance.xlsx"),
    ("log_file", "logs/app.log"),
    ("log_level", "INFO"),
    ("rule_mode", "SUM"),            # BOTH = both columns non-zero, ANY = either
    ("principal_col_index", "5"),     # 0-based column number
    ("interest_col_index", "7"),      # 0-based column number
    ("expected_principal_total", "49090.97"),
    ("expected_interest_total", "6689.97"),
        # ---- Cost sheet ----
    ("cost_sheet_name", "Cost"),
    ("cost_identifier", "Cost Detail Report"),
    ("cost_end_marker", "Total Costs Expended"),
    ("cost_title", "Consumer Collection Management, Inc. 2"),
    ("cost_thru_date", "AUTO"),       # AUTO = today's date, or type e.g. 10-13-2026
    ("expected_cost_total", "9659.09"),
        # ---- Payments sheet format ----
    ("pay_title", "CONSUMER COLLECTION MANAGEMENT, INC. 2"),
    ("pay_thru_date", "AUTO"),
    ("pay_total_col", "3"),
    ("pay_principal_col", "5"),
    ("pay_costs_col", "8"),
    ("pay_interest_col", "6"),
    ("client_share_pct", "0.85"),
]

# Order of the numeric columns on each debtor row in the Remittance Report
REPORT_COLUMNS = [
    "Payments Received", "Direct Payments", "Principal", "Interest", "Fees",
    "Costs", "Other", "Costs Expended", "Client Proceeds", "Fees Due",
]


def _create_default_config():
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(DEFAULTS, columns=["key", "value"]).to_excel(CONFIG_PATH, index=False)


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        _create_default_config()
    df = pd.read_excel(CONFIG_PATH, dtype=str)
    config = dict(zip(df["key"].str.strip(), df["value"].str.strip()))
    for key, value in DEFAULTS:          # fill anything missing
        config.setdefault(key, value)
    config["report_columns"] = REPORT_COLUMNS
    config["principal_col_index"] = int(config["principal_col_index"])
    config["interest_col_index"] = int(config["interest_col_index"])
    config["expected_principal_total"] = float(config["expected_principal_total"])
    config["expected_interest_total"] = float(config["expected_interest_total"])
    config["expected_cost_total"] = float(config["expected_cost_total"])
    for k in ("pay_total_col", "pay_principal_col", "pay_costs_col", "pay_interest_col"):
        config[k] = int(config[k])
    config["client_share_pct"] = float(config["client_share_pct"])
    return config