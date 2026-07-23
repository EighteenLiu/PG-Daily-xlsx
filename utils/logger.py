from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler

from utils.paths import app_root, ensure_app_dirs


def setup_logging() -> logging.Logger:
    ensure_app_dirs()
    logger = logging.getLogger("ledger_splitter")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    log_path = app_root() / "logs" / "app.log"
    handler = RotatingFileHandler(
        log_path,
        maxBytes=1_000_000,
        backupCount=3,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    )
    logger.addHandler(handler)
    return logger

