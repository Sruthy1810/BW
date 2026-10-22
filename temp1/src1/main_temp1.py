from pathlib import Path
from datetime import datetime
import pandas as pd
from notify.mail import record_files
from config.config_loader import load_temp1_config
from logs.logger_setup import setup_logger
from temp1.src1.t1_reader import (
    StartMarkerNotFoundError,
    get_input_files,
    read_raw,
    rows_after_marker,
)
from temp1.src1.t1_parser import parse_rows
from temp1.src1.t1_builder import build_payment_sheet, build_cost_sheet
from temp1.src1.t1_writer import (
    prepare_output,
    format_payment_share_precision,
    remove_empty_data_sheets,
    write_payment_sheet,
    write_cost_sheet,
    write_statement_address,
)
from temp1.src1.t1_master import get_shares
from temp1.src1.t1_cost import find_cost_rows, parse_cost_rows


def process_file(input_file, config, log):
    title = Path(input_file).stem.upper()                         
    output_name = f"{Path(input_file).stem.title()} - Output - {datetime.now():%m-%d-%Y}.xlsx"
    output_file = Path(config["output_folder_temp1"]) / output_name
    client_pct = bw_pct = None
    payment_sheet = build_payment_sheet(pd.DataFrame(), 0, 0)
    cost_sheet = build_cost_sheet(pd.DataFrame())
    section_errors = []

    try:
        raw = read_raw(input_file)
        rows = rows_after_marker(raw)
        data = parse_rows(rows, log)
        client_pct, bw_pct = get_shares(
            config["MasterDatatable"], input_file, log,
            sheet_name="IntrestPercentage")
        payment_sheet = build_payment_sheet(data, client_pct, bw_pct)
    except Exception as exc:
        section_errors.append(("payment", exc))
        log.exception("Payment section failed for %s; payment output will be blank",
                      input_file)

    try:
        cost_rows = find_cost_rows(input_file, log)
        cost_data = parse_cost_rows(cost_rows, log)
        cost_sheet = build_cost_sheet(cost_data)
    except Exception as exc:
        section_errors.append(("cost", exc))
        log.exception("Cost section failed for %s; cost output will be blank",
                      input_file)

    if len(section_errors) == 2:
        failures = ", ".join(section for section, _ in section_errors)
        raise RuntimeError(
            f"Both payment and cost sections failed: {failures}"
        ) from section_errors[0][1]

    prepare_output(config["output_template1"], output_file, log)
    write_payment_sheet(payment_sheet, output_file, log, title, client_pct, bw_pct)
    write_cost_sheet(cost_sheet, output_file, log, title)
    write_statement_address(output_file, input_file, log)
    remove_empty_data_sheets(output_file, log)
    format_payment_share_precision(output_file)

    print(f"{Path(input_file).name}: payment rows {len(payment_sheet)}, "
          f"cost rows {len(cost_sheet)} -> {output_file}")


def run_temp1():
    config = load_temp1_config()
    log = setup_logger(config["log_file"], config["log_level"])
    log.info("===== Temp1 run started =====")

    done, failed = [], []
    try:
        files = get_input_files(config)
        log.info("Input files found: %d", len(files))

        for input_file in files:
            log.info("----- Processing %s -----", Path(input_file).name)
            try:
                process_file(input_file, config, log)
                done.append(Path(input_file).name)
            except StartMarkerNotFoundError:
                log.exception("Missing remittance marker: %s", input_file)
                failed.append(Path(input_file).name)
            except Exception:
                log.exception("Failed: %s", input_file)        # log it and continue with the next file
                failed.append(Path(input_file).name)
    except Exception:
        log.exception("Temp1 run failed")
        raise
    finally:
        record_files("temp1", len(files) if "files" in locals() else 0,
                     len(done), len(failed))

        log.info("Done: %d | Failed: %d %s", len(done), len(failed), failed or "")
        log.info("===== Temp1 run finished =====")
        print(f"\nDone: {len(done)} | Failed: {len(failed)} {failed or ''}")

        if failed:
            raise RuntimeError(f"{len(failed)} file(s) failed: {failed}")

if __name__ == "__main__":
    run_temp1()