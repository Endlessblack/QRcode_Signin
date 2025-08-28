from __future__ import annotations

import sys
from pathlib import Path

from PyQt6 import QtWidgets

from .config import AppConfig
from .ui import MainWindow


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName("QR 簽到")
    cfg = AppConfig(Path("config.json"))
    win = MainWindow(cfg)
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

