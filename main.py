import logging
import os
import sys
import traceback
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
os.chdir(APP_DIR)

from config.config_loader import load_mail_config
from logs.logger_setup import start_execution_log
from notify.mail import Notifier
from temp1.src1.main_temp1 import run_temp1
from temp2.src2.main_temp2 import run_temp2
from notify.process_killer import kill_excel
FLOWS = {
    "temp1": run_temp1,
    "temp2": run_temp2,
}


def main():
    # python main.py            -> runs temp1 and temp2
    # python main.py temp1      -> runs only temp1
    # python main.py temp2      -> runs only temp2
    names = [a.lower() for a in sys.argv[1:]] or list(FLOWS)

    unknown = [n for n in names if n not in FLOWS]
    if unknown:
        raise SystemExit(f"Unknown flow(s): {unknown}. Use: {list(FLOWS)}")

    execution_log = start_execution_log()                      # this run's execution .txt
    log = logging.getLogger("remittance")
    log.info("Execution log: %s", execution_log)

    kill_excel()  

    mail_cfg = load_mail_config()
    notifier = Notifier(mail_cfg)
    notifier.trigger(names)                                    # 1) trigger mail

    results = {}
    for name in names:
        print(f"\n########## Running {name} ##########")
        log.info("########## Running %s ##########", name)
        try:
            FLOWS[name]()
            results[name] = ("Success", "")
        except Exception as exc:               # one flow failing must not stop the other
            tb = traceback.format_exc()
            traceback.print_exc()
            log.error("Flow %s FAILED: %s\n%s", name, exc, tb)  # makes sure it is in the execution file
            results[name] = ("Failed", str(exc))
            notifier.exception(name, exc, tb)                  # 2) exception mail (execution log attached)

    failed = [n for n, (status, _) in results.items() if status != "Success"]
    summary = f"Finished. Succeeded: {[n for n in names if n not in failed]} | Failed: {failed}"
    print(f"\n{summary}")
    log.info(summary)
    notifier.end(results)                                      # 3) end mail (execution log attached)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()