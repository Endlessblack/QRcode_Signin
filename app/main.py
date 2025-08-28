from __future__ import annotations

import os
import sys
from pathlib import Path

from PyQt6 import QtWidgets

# Support running as a module (python -m app.main) and as a script (python app/main.py)
try:
    from .config import AppConfig  # type: ignore
    from .ui import MainWindow  # type: ignore
    from .logger import setup_logging  # type: ignore
except Exception:  # ImportError when no package context
    sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
    from app.config import AppConfig  # type: ignore
    from app.ui import MainWindow  # type: ignore
    from app.logger import setup_logging  # type: ignore


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName("QR 簽到")
    cfg = AppConfig(Path("config.json"))
    setup_logging(cfg.debug)
    win = MainWindow(cfg)
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
