<<<<<<< HEAD
import sys

from temp1.src1.main_temp1 import run_temp1
from main_temp2 import run_temp2

FLOWS = {
    "temp1": run_temp1,
    "temp2": run_temp2,
}
=======
from pathlib import Path
from temp2.src.payment_writer import format_payment_sheet
from temp2.src.config_loader import load_config
from temp2.src.logger_setup import setup_logger
from temp2.src.reader import read_report
from temp2.src.validator import check_totals
from temp2.src.cleaner import keep_valid_rows
from temp2.src.validator import check_cost_total
from temp2.src.cost_reader import read_costs
from temp2.src.cost_writer import write_cost_sheet
from pathlib import Path
from temp2.src.client_share import get_client_share
>>>>>>> 452d2195d3bda2f34aa98fda87bdc2da38ca013a


def main():
    # python main.py            -> runs temp1 and temp2
    # python main.py temp1      -> runs only temp1
    # python main.py temp2      -> runs only temp2
    names = [a.lower() for a in sys.argv[1:]] or list(FLOWS)

    unknown = [n for n in names if n not in FLOWS]
    if unknown:
        raise SystemExit(f"Unknown flow(s): {unknown}. Use: {list(FLOWS)}")

    failed = []
    for name in names:
        print(f"\n########## Running {name} ##########")
        try:
            FLOWS[name]()
        except Exception as exc:               # one flow failing must not stop the other
            failed.append(name)
            print(f"{name} FAILED: {exc}  (details in logs/app.log)")

<<<<<<< HEAD
    print(f"\nFinished. Succeeded: {[n for n in names if n not in failed]} | Failed: {failed}")
    if failed:
        sys.exit(1)
=======
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
>>>>>>> 452d2195d3bda2f34aa98fda87bdc2da38ca013a


if __name__ == "__main__":
    main()