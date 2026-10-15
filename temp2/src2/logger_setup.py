import logging
from pathlib import Path


def setup_logger(log_file: str, level: str = "INFO") -> logging.Logger:
    Path(log_file).parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("remittance")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(fmt)
    console = logging.StreamHandler()
    console.setFormatter(fmt)

    logger.addHandler(file_handler)
    logger.addHandler(console)
    return logger