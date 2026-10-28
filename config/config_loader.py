from pathlib import Path
import pandas as pd
import sys

if getattr(sys, "frozen", False):
    ROOT_DIR = Path(sys.executable).resolve().parent
else:
    ROOT_DIR = Path(__file__).resolve().parents[2]

CONFIG_PATH = ROOT_DIR / "config_file" / "config.xlsx"

# =====================================================================
# MAIL  - settings for the trigger / exception / end mails
# =====================================================================
MAIL_DEFAULTS = {
    "mail_enabled": "false",
    "smtp_host": "smtp.gmail.com",
    "smtp_port": "587",
    "smtp_user": "",
    "mail_from": "",
    "mail_to": "",
    "mail_cc": "",
    "mail_subject_prefix": "[B&W Bot]",
    "sender_password": "",
}




def load_mail_config() -> dict:
    values = {}

    try:
        if not CONFIG_PATH.exists():
            print(f"Mail config not found: {CONFIG_PATH}")
            return {**MAIL_DEFAULTS, "mail_enabled": "false"}

        df = pd.read_excel(CONFIG_PATH, dtype=str)
        df.columns = df.columns.str.strip().str.lower()

        if "key" not in df.columns or "value" not in df.columns:
            raise ValueError("Excel must contain 'key' and 'value' columns.")

        df = df.dropna(subset=["key"])

        values = dict(
            zip(
                df["key"].str.strip(),
                df["value"].fillna("").str.strip()
            )
        )

    except Exception as exc:
        print(f"Mail config could not be read ({exc}) - mails are disabled")
        return {**MAIL_DEFAULTS, "mail_enabled": "false"}

    config = {
        key: values.get(key) or default
        for key, default in MAIL_DEFAULTS.items()
    }

    return config
# =====================================================================
# TEMP1  - loads ONLY the temp1 keys, ignores all of her settings
# =====================================================================
TEMP1_DEFAULTS = {
    "input_file_temp1": "",
    "input_folder_temp1": "",                # folder with the input Excel files
    "output_folder_temp1": "output",
    "failed_temp1": "Failed/temp_1",
    "processed_file1": "Processed/temp_1",
    "output_template1": "",
    "MasterDatatable": "",
    "log_file": "logs/app.log",
    "log_level": "INFO",
}



PATH_KEYS = ["input_file_temp1", "input_folder_temp1", "output_folder_temp1",
             "failed_temp1", "processed_file1", "output_template1",
             "MasterDatatable"]


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
    ("failed_temp2", "Failed/temp_2"),
    ("processed_file2", "Processed/temp_2"),
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


#-------------------------------------------------------------------------------------------------#
#                                TEMP3_COURT_COST                                                 #
#-------------------------------------------------------------------------------------------------#

TEMP3_COURT_COST = {
    # name used in the code        : row (key) in config.xlsx
    "input_folder":     "input_folder_temp3_Court",
    "output_folder":    "output_folder_temp3_Court",
    "output_template":  "output_template3_Court",
    "failed_folder":    "failed_temp3_Court",
    "processed_folder": "processed_file3_Court",
}

TEMP3_COURT_DEFAULTS = {
    "log_file": "logs/app.log",
    "log_level": "INFO",
}


def _read_config_values() -> dict:
    """key -> value for every row of config.xlsx (blank values become '')."""
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"Config Excel file not found: {CONFIG_PATH.resolve()}")

    df = pd.read_excel(CONFIG_PATH, dtype=str)
    df.columns = df.columns.str.strip().str.lower()
    if "key" not in df.columns or "value" not in df.columns:
        raise ValueError(f"Excel must contain 'key' and 'value' columns. Found: {df.columns.tolist()}")

    df = df.dropna(subset=["key"])
    return dict(zip(df["key"].str.strip(), df["value"].fillna("").str.strip()))


def load_temp3_Court_config() -> dict:
    values = _read_config_values()
    config = {}

    # required rows: stop with a clear message if one is missing
    missing = [row for row in TEMP3_COURT_COST.values() if not values.get(row)]
    if missing:
        raise ValueError(f"Add these rows with a value in config.xlsx: {missing}")

    for name, row in TEMP3_COURT_COST.items():
        config[name] = resolve_path(values[row])            # relative -> from the project root

    for key, default in TEMP3_COURT_DEFAULTS.items():
        config[key] = values.get(key) or default

    # folders the flow writes to: create them if they are not there yet
    for name in ("output_folder", "failed_folder", "processed_folder"):
        Path(config[name]).mkdir(parents=True, exist_ok=True)

    if not Path(config["input_folder"]).is_dir():
        raise FileNotFoundError(f"Temp3 input folder not found: {config['input_folder']}")
    if not Path(config["output_template"]).is_file():
        raise FileNotFoundError(f"Temp3 output template not found: {config['output_template']}")

    return config

#-------------------------------------------------------------------------------------------------#
#                                TEMP3_FIRM_PAYMENTS                                              #
#-------------------------------------------------------------------------------------------------#

TEMP3_FIRM_PAYMENT_CONFIG = {
    "input_folder": ("input_folder_temp3_Firm_Payments",),
    "output_folder": ("output_folder_temp3_Firm_Payments",),
    "output_template": ("output_template3_Firm_Payments",),
    "failed_folder": ("failed_temp3_Firm_Payments",),
    "processed_folder": ("processed_file3_Firm_Payments",),
}

TEMP3_NON_ATTORNEY_CONFIG = {
    "input_folder": (
        "input_folder_temp3_Non_Attorney_Payments",
        "input_folder_temp3_Non_Attorney",
    ),
    "output_folder": (
        "output_folder_temp3_Non_Attorney_Payments",
        "output_folder_temp3_Non_Attorney",
    ),
    "output_template": (
        "output_template3_Non_Attorney_Payments",
        "output_template3_Non_Attorney",
    ),
    "failed_folder": (
        "failed_temp3_Non_Attorney_Payments",
        "failed_temp3_Non_Attorney",
    ),
    "processed_folder": (
        "processed_file3_Non_Attorney_Payments",
        "processed_file3_Non_Attorney",
    ),
}

TEMP3_PAYMENT_DEFAULTS = {
    "Firm Payments": {
        "output_template": "Output/Templates/Firm Payment Template.xlsx",
        "processed_folder": "Processed/temp3_Firm_Payments",
    },
    "Non-Attorney": {
        "output_template": "Output/Templates/Non-Attorney Payments Template.xlsx",
        "processed_folder": "Processed/temp3_Non_Attorney_Payments",
    },
}


def _load_temp3_payment_config(project: str, key_map: dict) -> dict:
    values = _read_config_values()
    defaults = TEMP3_PAYMENT_DEFAULTS[project]
    resolved_values = {}
    missing = []
    for name, candidates in key_map.items():
        if isinstance(candidates, str):
            candidates = (candidates,)
        value = next((values[key] for key in candidates if values.get(key)), "")
        if not value:
            value = defaults.get(name, "")
        if value:
            resolved_values[name] = value
        else:
            missing.append(candidates[0])

    if missing:
        raise ValueError(f"Add these rows with a value in config.xlsx: {missing}")

    config = {
        name: Path(resolve_path(value))
        for name, value in resolved_values.items()
    }
    config["log_file"] = Path(
        resolve_path(values.get("log_file") or "logs/app.log")
    )
    config["log_level"] = values.get("log_level") or "INFO"

    for name in ("output_folder", "failed_folder", "processed_folder"):
        Path(config[name]).mkdir(parents=True, exist_ok=True)

    if not Path(config["input_folder"]).is_dir():
        raise FileNotFoundError(
            f"{project} input folder not found: {config['input_folder']}"
        )
    if not Path(config["output_template"]).is_file():
        raise FileNotFoundError(
            f"{project} output template not found: {config['output_template']}"
        )

    return config


def load_temp3_firm_config() -> dict:
    """Load only the Firm Payments paths from the shared config.xlsx."""
    return _load_temp3_payment_config("Firm Payments", TEMP3_FIRM_PAYMENT_CONFIG)


def load_temp3_non_attorney_config() -> dict:
    """Load only the Non-Attorney paths from the shared config.xlsx."""
    return _load_temp3_payment_config(
        "Non-Attorney",
        TEMP3_NON_ATTORNEY_CONFIG,
    )