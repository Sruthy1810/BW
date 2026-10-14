from pathlib import Path
from src.payment_writer import format_payment_sheet
from src.config_loader import load_config
from src.logger_setup import setup_logger
from src.reader import read_report
from src.validator import check_totals
from src.cleaner import keep_valid_rows
from src.validator import check_cost_total
from src.cost_reader import read_costs
from src.cost_writer import write_cost_sheet


def main():
    config = load_config()
    log = setup_logger(config["log_file"], config["log_level"])
    log.info("===== Run started =====")

    try:
        df = read_report(config["input_file"], config["sheet_name"],
                         config["report_columns"])

        totals_ok = check_totals(df, config["principal_col_index"],
                                 config["interest_col_index"],
                                 config["expected_principal_total"],
                                 config["expected_interest_total"])
        if not totals_ok:
            log.warning("Totals do not match the report. Some rows may have been missed.")

        valid = keep_valid_rows(df, config["principal_col_index"],
                                config["interest_col_index"], config["rule_mode"])

        Path(config["output_file"]).parent.mkdir(parents=True, exist_ok=True)
        valid.to_excel(config["output_file"], index=False)
        log.info("Saved valid data to %s", config["output_file"])
        format_payment_sheet(config["output_file"], valid, config)
                # ---- Cost sheet (added after the valid data is saved) ----
        plaintiff_map = dict(zip(df["Claim"], df["Plaintiff"]))
        cost_df = read_costs(config["input_file"], config["cost_identifier"],
                             config["cost_end_marker"], plaintiff_map)
        check_cost_total(cost_df, config["expected_cost_total"])
        write_cost_sheet(config["output_file"], config["cost_sheet_name"],
                         cost_df, config["cost_title"], config["cost_thru_date"])
        log.info("Cost sheet added to %s", config["output_file"])

        print(valid.head(20).to_string())
        print(f"\nValid rows: {len(valid)}")
    except Exception:
        log.exception("Run failed")
        raise
    finally:
        log.info("===== Run finished =====")


if __name__ == "__main__":
    main()