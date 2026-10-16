from pathlib import Path
import sys

from config.config_loader import load_config, resolve_path 
from logs.logger_setup import setup_logger
from temp2.src2.master_data import get_client_percentages
from temp2.src2.reader import extract_remittance_address, read_remittance_sections
from temp2.src2.writer import write_output


def run_temp2():
    config = load_config()
    for key in ("input_directory", "output_file", "master_table_file", "template_file"):
        if config.get(key):
            config[key] = resolve_path(config[key]) 
    log = setup_logger(config["log_file"], config["log_level"])
    log.info("===== Run started =====")

    input_directory = Path(config["input_directory"])
    supported_extensions = {".xls", ".xlsx", ".xlsm"}
    input_files = sorted(
        path
        for path in input_directory.iterdir()
        if path.is_file()
        and path.suffix.lower() in supported_extensions
        and not path.name.startswith("~$")
    ) if input_directory.is_dir() else []

    if not input_files:
        message = f"No Excel workbooks found in input folder: {input_directory}"
        log.error(message)
        raise FileNotFoundError(message)

    output_directory = Path(config["output_file"]).parent
    output_directory.mkdir(parents=True, exist_ok=True)
    failed_files = []

    try:
        log.info("Found %d Excel workbook(s) in %s", len(input_files), input_directory)
        for input_file in input_files:
            log.info("Processing input workbook: %s", input_file)
            try:
                payments, costs, statement_summary = read_remittance_sections(
                    str(input_file), config["sheet_name"], config["report_columns"]
                )
                remittance_address = extract_remittance_address(
                    str(input_file), config["sheet_name"]
                )

                client_share, bw_share, bw_interest_share = get_client_percentages(
                    config["master_table_file"], input_file.stem
                )
                log.info(
                    "Master percentages for %s: client=%s, BW LLP=%s, BW interest=%s",
                    input_file.stem,
                    client_share,
                    bw_share,
                    bw_interest_share,
                )

                output_file = output_directory / f"{input_file.stem} Output File.xlsx"
                write_output(
                    output_file,
                    config["template_file"],
                    payments,
                    costs,
                    statement_summary,
                    client_share,
                    input_file,
                    bw_share,
                    bw_interest_share,
                    remittance_address,
                )
                log.info("Saved output workbook to %s", output_file)
                print(
                    f"{input_file.name}: {len(payments)} payment rows, "
                    f"{len(costs)} cost rows -> {output_file}"
                )
            except Exception:
                failed_files.append(input_file)
                log.exception("Failed to process input workbook: %s", input_file)

        if failed_files:
            log.error(
                "%d of %d input workbook(s) failed: %s",
                len(failed_files),
                len(input_files),
                ", ".join(str(path) for path in failed_files),
            )
            sys.exit(1)
        log.info("Successfully processed all %d input workbook(s)", len(input_files))
    finally:
        log.info("===== Run finished =====")


if __name__ == "__main__":
    run_temp2()