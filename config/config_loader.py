from pathlib import Path
import pandas as pd
import sys

if getattr(sys, "frozen", False):
    ROOT_DIR = Path(sys.executable).resolve().parent
else:
    ROOT_DIR = Path(__file__).resolve().parents[2]

CONFIG_PATH = ROOT_DIR / "config_file" / "config.xlsx"



# =====================================================================
# TEMP1  - loads ONLY the temp1 keys, ignores all of her settings
# =====================================================================
TEMP1_DEFAULTS = {
    "input_file_temp1": "",
    "input_folder_temp1": "",                # folder with the input Excel files
    "output_folder_temp1": "output",
    "output_template1": "",
    "MasterDatatable": "",
    "log_file": "logs/app.log",
    "log_level": "INFO",
}



PATH_KEYS = ["input_file_temp1", "input_folder_temp1", "output_folder_temp1",
             "output_template1", "MasterDatatable"]


def resolve_path(value) -> str:
    p = Path(str(value).strip().strip('"'))
    return str(p if p.is_absolute() else (ROOT_DIR / p).resolve())

def load_temp1_config() -> dict:
    cfg_path = CONFIG_PATH

    if not cfg_path.exists():
        raise FileNotFoundError(
            f"Config Excel file not found: {cfg_path.resolve()}"
        )

    df = pd.read_excel(cfg_path, dtype=str)

    # Normalize column headers
    df.columns = df.columns.str.strip().str.lower()
    dup = df[df["key"].str.strip().duplicated(keep=False)]
    if not dup.empty:
        print("WARNING - duplicate keys (the LAST row wins):")
        print(dup.to_string())

    if "key" not in df.columns or "value" not in df.columns:
        raise ValueError(
            f"Excel must contain 'key' and 'value' columns. "
            f"Found: {df.columns.tolist()}"
        )

    df = df.dropna(subset=["key"])

    values = dict(
        zip(
            df["key"].str.strip(),
            df["value"].fillna("").str.strip()
        )
    )

    print("Loaded config keys:", values)

    config = {}
    for key, default in TEMP1_DEFAULTS.items():
        value = values.get(key, "")
        config[key] = value if value else default

    if not config["input_folder_temp1"] and not config["input_file_temp1"]:
        raise ValueError(
            "Add input_folder_temp1 or input_file_temp1 "
            "with a non-empty value in config.xlsx"
        )

    if not config["output_template1"]:
        raise ValueError("Add output_template1 in config.xlsx")

    if not config["MasterDatatable"]:
        raise ValueError("Add MasterDatatable in config.xlsx")

    for key in PATH_KEYS:
        if config.get(key):
            config[key] = resolve_path(config[key])

    return config


# =====================================================================
# TEMP2  - loads ONLY the temp2 keys, ignores all of her settings
# =====================================================================


DEFAULTS = [
    ("input_directory", "input"),
    ("sheet_name", "Remittance Report"),
    ("output_file", "output/valid_remittance.xlsx"),
    ("template_file", "Output_Template/OutputFile_Temp2.xlsx"),
    ("master_table_file", "MasterDataTable/MasterDatatable_1.xlsx"),
    ("log_file", "logs/app.log"),
    ("log_level", "INFO"),
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

    # ---- map the shared config.xlsx keys (temp2 rows) to the names the temp2 code uses ----
    aliases = {
        "input_directory": "input_folder_temp2",
        "template_file": "output_template2",
        "master_table_file": "MasterDatatable",
    }
    for temp2_key, sheet_key in aliases.items():
        value = config.get(sheet_key)
        if isinstance(value, str) and value.strip():
            config[temp2_key] = value

    out = config.get("output_folder_temp2")
    if isinstance(out, str) and out.strip():
        p = Path(out.strip())
        # temp2 only uses the folder part of output_file
        config["output_file"] = out.strip() if p.suffix else str(p / "output.xlsx")

    for key, value in DEFAULTS:          # fill anything missing
        config.setdefault(key, value)
    config["report_columns"] = REPORT_COLUMNS
    return config