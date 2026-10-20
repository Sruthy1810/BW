from pathlib import Path
from notify.mail import record_files
from config.config_loader import load_temp1_config
from logs.logger_setup import setup_logger
from temp1.src1.t1_reader import get_input_files, read_raw, rows_after_marker
from temp1.src1.t1_parser import parse_rows
from temp1.src1.t1_builder import build_payment_sheet, build_cost_sheet
from temp1.src1.t1_writer import prepare_output, write_payment_sheet, write_cost_sheet
from temp1.src1.t1_master import get_shares
from temp1.src1.t1_cost import find_cost_rows, parse_cost_rows


def process_file(input_file, config, log):
    title = Path(input_file).stem.upper()                          # 'CASH LINK USA, LLC'
    output_file = Path(config["output_folder_temp1"]) / f"{Path(input_file).stem}.xlsx"
    client_pct, bw_pct = get_shares(config["MasterDatatable"], input_file, log, sheet_name="IntrestPercentage")

    prepare_output(config["output_template1"], output_file, log)   # copy template -> same name as input

    # ---- Payment sheet ----
    raw = read_raw(input_file)
    rows = rows_after_marker(raw)
    data = parse_rows(rows, log)
    payment_sheet = build_payment_sheet(data, client_pct, bw_pct) 
    write_payment_sheet(payment_sheet, output_file, log, title, client_pct, bw_pct)

    # ---- Cost sheet ----
    cost_rows = find_cost_rows(input_file, log)
    cost_data = parse_cost_rows(cost_rows, log)
    cost_sheet = build_cost_sheet(cost_data)
    write_cost_sheet(cost_sheet, output_file, log, title)

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