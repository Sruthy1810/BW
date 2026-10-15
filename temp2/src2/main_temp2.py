from pathlib import Path
from temp2.src2.payment_writer import format_payment_sheet
from config.config_loader import load_config
from temp2.src2.logger_setup import setup_logger
from temp2.src2.reader import read_report
from temp2.src2.validator import check_totals
from temp2.src2.cleaner import keep_valid_rows
from temp2.src2.validator import check_cost_total
from temp2.src2.cost_reader import read_costs
from temp2.src2.cost_writer import write_cost_sheet
from pathlib import Path
from temp2.src2.client_share import get_client_share


def run_temp2():
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
        client_name = config.get("client_name")
        if not isinstance(client_name, str) or not client_name.strip() \
                or client_name.strip().lower() == "nan":
            client_name = Path(config["input_file"]).stem
        config["client_share_pct"] = get_client_share(
            config.get("master_file", "input/MasterDatatable_1.xlsx"),
            config.get("master_sheet", "InterestPercentage_Temp2"),
            client_name,
            config.get("master_share_column", "BWShare"),
        )

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
    run_temp2()