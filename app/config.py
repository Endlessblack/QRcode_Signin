from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict


DEFAULTS: Dict[str, Any] = {
    "google": {
        "credentials_path": "credentials.json",
        "spreadsheet_id": "",
        "worksheet_name": "Signin",
    },
    "event": {"name": "活動"},
    "camera": {"index": 0},
    "output": {"qr_folder": "output_qrcodes"},
}


class AppConfig:
    def __init__(self, path: Path | str = "config.json") -> None:
        self.path = Path(path)
        self.data: Dict[str, Any] = json.loads(json.dumps(DEFAULTS))  # deep copy
        if self.path.exists():
            self.load()

    def load(self) -> None:
        try:
            with self.path.open("r", encoding="utf-8") as f:
                incoming = json.load(f)
            self._merge(self.data, incoming)
        except Exception:
            # Keep defaults on error
            pass

    def save(self) -> None:
        try:
            with self.path.open("w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _merge(self, target: Dict[str, Any], src: Dict[str, Any]) -> None:
        for k, v in src.items():
            if isinstance(v, dict) and isinstance(target.get(k), dict):
                self._merge(target[k], v)
            else:
                target[k] = v

    # Convenience getters/setters
    @property
    def credentials_path(self) -> str:
        return str(self.data["google"]["credentials_path"])  # type: ignore

    @credentials_path.setter
    def credentials_path(self, p: str) -> None:
        self.data["google"]["credentials_path"] = p

    @property
    def spreadsheet_id(self) -> str:
        return str(self.data["google"]["spreadsheet_id"])  # type: ignore

    @spreadsheet_id.setter
    def spreadsheet_id(self, s: str) -> None:
        self.data["google"]["spreadsheet_id"] = s

    @property
    def worksheet_name(self) -> str:
        return str(self.data["google"]["worksheet_name"])  # type: ignore

    @worksheet_name.setter
    def worksheet_name(self, s: str) -> None:
        self.data["google"]["worksheet_name"] = s

    @property
    def event_name(self) -> str:
        return str(self.data["event"]["name"])  # type: ignore

    @event_name.setter
    def event_name(self, s: str) -> None:
        self.data["event"]["name"] = s

    @property
    def camera_index(self) -> int:
        try:
            return int(self.data["camera"]["index"])  # type: ignore
        except Exception:
            return 0

    @camera_index.setter
    def camera_index(self, idx: int) -> None:
        self.data["camera"]["index"] = int(idx)

    @property
    def qr_folder(self) -> str:
        return str(self.data["output"]["qr_folder"])  # type: ignore

    @qr_folder.setter
    def qr_folder(self, s: str) -> None:
        self.data["output"]["qr_folder"] = s

