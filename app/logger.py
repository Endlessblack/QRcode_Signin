from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_logging(debug: bool = False) -> logging.Logger:
    log = logging.getLogger("app")
    if getattr(log, "_configured", False):  # idempotent
        return log

    level = logging.DEBUG if debug else logging.INFO
    log.setLevel(level)

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    # Console
    ch = logging.StreamHandler()
    ch.setLevel(level)
    ch.setFormatter(fmt)
    log.addHandler(ch)

    # File (rotating)
    log_dir = Path.cwd() / "logs"
    log_dir.mkdir(exist_ok=True)
    fh = RotatingFileHandler(log_dir / "app.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    fh.setLevel(level)
    fh.setFormatter(fmt)
    log.addHandler(fh)

    setattr(log, "_configured", True)
    log.debug("Logging initialized. Debug=%s", debug)
    return log

