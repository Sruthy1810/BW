import shutil
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from temp3_Firm_Payments.src4.calculator import calculate
from config.config_loader import CONFIG_PATH, load_temp3_firm_config
from temp3_Firm_Payments.src4.input_reader import read_rows
from logs.logger_setup import setup_logger
from temp3_Firm_Payments.src4.output_writer import write_output
from temp3_Firm_Payments.src4.parser import parse_payments
from temp3_Firm_Payments.src4.settings import PROJECT_NAME
from notify.mail import record_files
from notify.fonts import apply_font_to_file


def _unique_output_path(output_folder: Path, input_file: Path) -> Path:
    base = f"{input_file.stem.upper()} - OUTPUT - {datetime.now():%m-%d-%Y}"
    output = output_folder / f"{base}.xlsx"
    suffix = 1
    while output.exists():
        output = output_folder / f"{base} ({suffix}).xlsx"
        suffix += 1
    return output


def _move_input(input_file: Path, folder: Path) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    destination = folder / input_file.name
    suffix = 1
    while destination.exists():
        destination = folder / f"{input_file.stem} ({suffix}){input_file.suffix}"
        suffix += 1
    shutil.move(str(input_file), str(destination))


def run_temp3_Firm() -> int:
    files = []
    done = []
    failures = []
    try:
        cfg = load_temp3_firm_config()
        log = setup_logger(str(cfg["log_file"]))
        log.info("=== Run started | project=%s | config=%s", PROJECT_NAME, CONFIG_PATH)
        input_folder = cfg["input_folder"]
        files = sorted(
            path for path in input_folder.iterdir()
            if path.is_file() and path.suffix.lower() in (".xls", ".xlsx")
            and not path.name.startswith("~$")
        )
        if not files:
            raise FileNotFoundError(f"No Excel input files found in {input_folder}")

        for input_file in files:
            try:
                log.info("Processing input: %s", input_file)
                payments = parse_payments(read_rows(input_file))
                log.info("Parsed %d payment(s)", len(payments))

                results = [calculate(payment) for payment in payments]
                for result in results:
                    log.info(
                        "%s | acct %s | %s | gross %s | commission %s | net %s | %s",
                        result["date"].strftime("%m-%d-%Y"),
                        result["account"],
                        result["name"],
                        result["gross"],
                        result["commission"],
                        result["net"],
                        result["origination"],
                    )

                output_file = _unique_output_path(cfg["output_folder"], input_file)
                write_output(cfg["output_template"], output_file, results)

                apply_font_to_file(output_file)          # <-- add here, LAST step on the output
                log.info("Arial Narrow applied to %s", output_file)
                                
                log.info(
                    "Totals: gross %s | commission %s | net %s",
                    sum(result["gross"] for result in results),
                    sum(result["commission"] for result in results),
                    sum(result["net"] for result in results),
                )
                log.info("Output written: %s", output_file)
                _move_input(input_file, cfg["processed_folder"])
                done.append(input_file.name)
            except Exception:
                log.exception("Failed to process %s", input_file)
                failures.append(input_file.name)
                _move_input(input_file, cfg["failed_folder"])

        if failures:
            log.error("Failed input files: %s", failures)
            return 1
        log.info("=== Run finished OK")
        return 0
    except Exception:
        if "log" in locals():
            log.exception("Run failed")
        return 1
    finally:
        record_files("temp3_firm_payments", len(files), len(done), len(failures))


if __name__ == "__main__":
    sys.exit(run_temp3_Firm())