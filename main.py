import os
import sys
from pathlib import Path

APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
os.chdir(APP_DIR)

from temp1.src1.main_temp1 import run_temp1
from temp2.src2.main_temp2 import run_temp2

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

    failed = []
    for name in names:
        print(f"\n########## Running {name} ##########")
        try:
            FLOWS[name]()
        except Exception as exc:               # one flow failing must not stop the other
            failed.append(name)
            print(f"{name} FAILED: {exc}  (details in logs/app.log)")

    print(f"\nFinished. Succeeded: {[n for n in names if n not in failed]} | Failed: {failed}")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    finally:
        if getattr(sys, "frozen", False):          # only for the exe: stop the window closing instantly
            input("\nPress Enter to close...")