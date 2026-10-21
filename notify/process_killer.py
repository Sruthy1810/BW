import logging
import subprocess
import sys
import time

log = logging.getLogger("remittance")


def kill_excel():
    """Force-close every running Excel process (Windows). Returns True if something was killed."""
    if sys.platform != "win32":
        log.info("kill_excel skipped: not Windows")
        return False
    try:
        result = subprocess.run(
            ["taskkill", "/F", "/T", "/IM", "EXCEL.EXE"],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode == 0:
            log.warning("Excel was open - process killed before the run")
            time.sleep(2)               # give Windows a moment to release the file locks
            return True
        log.info("No open Excel process found")     # taskkill returns 128 when none is running
        return False
    except Exception as exc:
        log.warning("Could not kill Excel: %s", exc)  # never stop the bot because of this
        return False