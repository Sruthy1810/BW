import logging
from datetime import datetime
from pathlib import Path

_FMT = "%(asctime)s | %(levelname)s | %(message)s"
_EXECUTION_LOG = None            # path of this run's execution file (set by start_execution_log)


def start_execution_log(folder: str = "logs/executions") -> Path:
    """Create this run's execution file and start writing to it. Call once, at the start of main.py."""
    global _EXECUTION_LOG
    Path(folder).mkdir(parents=True, exist_ok=True)
    _EXECUTION_LOG = (Path(folder) / f"execution_{datetime.now():%Y%m%d_%H%M%S}.txt").resolve()
    _EXECUTION_LOG.touch()

    logger = logging.getLogger("remittance")
    if logger.level == logging.NOTSET:
        logger.setLevel(logging.INFO)
    _attach_execution_handler(logger)
    return _EXECUTION_LOG


def get_execution_log():
    return _EXECUTION_LOG


def _attach_execution_handler(logger: logging.Logger) -> None:
    if _EXECUTION_LOG is None:
        return
    if any(getattr(h, "_is_execution", False) for h in logger.handlers):
        return
    handler = logging.FileHandler(_EXECUTION_LOG, mode="a", encoding="utf-8")
    handler.setFormatter(logging.Formatter(_FMT))
    handler._is_execution = True
    logger.addHandler(handler)


def setup_logger(log_file: str, level: str = "INFO") -> logging.Logger:
    Path(log_file).parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("remittance")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))

    for handler in list(logger.handlers):       # close the old handlers before dropping them
        handler.close()
    logger.handlers.clear()

    fmt = logging.Formatter(_FMT)
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(fmt)
    console = logging.StreamHandler()
    console.setFormatter(fmt)

    logger.addHandler(file_handler)
    logger.addHandler(console)
    _attach_execution_handler(logger)           # keep writing to this run's execution file
    return logger