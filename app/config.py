from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict
from .paths import config_file, ensure_dirs, oauth_client_file, oauth_token_file, OAUTH_CLIENT_FILENAME, OAUTH_TOKEN_FILENAME


DEFAULTS: Dict[str, Any] = {
    "google": {
        "credentials_path": "credentials.json",
        "auth_method": "service_account",  # or 'oauth'
        "oauth_client_path": "",           # fixed to ./client/client.json for OAuth
        "oauth_token_path": "token.json",  # where to store authorized user token
        "spreadsheet_id": "",
        "worksheet_name": "Signin",
    },
    "event": {"name": "活動"},
    "camera": {"index": 0},
    "output": {"qr_folder": "output_qrcodes"},
    "fields": {"extras": ["salon", "seller"]},
    "design": {
        "use_design": True,
        "width": 1080,
        "height": 1350,
        "qr_ratio": 0.7,
        "bg_color": "#FFFFFF",
        "qr_color": "#000000",
        "text_color": "#000000",
        "font_family": "",
        "font_size": 48,
        "font_weight": "regular",
        "bg_image_path": "",
        "text_margin": 40,
        "text_top_gap": 40,
        "text_bottom_margin": 40,
        "line_spacing_scale": 0.4,
        "auto_fit_text": True,
        "text_point": None,
    },
    "debug": False,
}


class AppConfig:
    def __init__(self, path: Path | str | None = None) -> None:
        # Ensure required directories exist before reading/saving
        ensure_dirs()
        self.path = Path(path) if path is not None else config_file()
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
            self.path.parent.mkdir(parents=True, exist_ok=True)
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

    @property
    def extra_fields(self) -> list[str]:
        vals = self.data.get("fields", {}).get("extras", [])
        if isinstance(vals, list):
            return [str(x) for x in vals if str(x).strip() and str(x) not in ("id", "name")]
        return []

    @extra_fields.setter
    def extra_fields(self, items: list[str]) -> None:
        cleaned = []
        for x in items:
            s = str(x).strip()
            if s and s not in ("id", "name") and s not in cleaned:
                cleaned.append(s)
        self.data.setdefault("fields", {})["extras"] = cleaned

    @property
    def debug(self) -> bool:
        return bool(self.data.get("debug", False))

    @debug.setter
    def debug(self, v: bool) -> None:
        self.data["debug"] = bool(v)

    # Design getters/setters (generic)
    def get_design(self, key: str, default: Any = None) -> Any:
        try:
            return self.data.get("design", {}).get(key, default)
        except Exception:
            return default

    def set_design(self, key: str, value: Any) -> None:
        self.data.setdefault("design", {})[key] = value

    # Google auth extras
    @property
    def auth_method(self) -> str:
        try:
            return str(self.data.get("google", {}).get("auth_method", "service_account")).lower()
        except Exception:
            return "service_account"

    @auth_method.setter
    def auth_method(self, m: str) -> None:
        self.data.setdefault("google", {})["auth_method"] = str(m)

    @property
    def oauth_client_path(self) -> str:
        # Hardcode to ./client/<OAUTH_CLIENT_FILENAME>
        return str(oauth_client_file())

    @oauth_client_path.setter
    def oauth_client_path(self, p: str) -> None:
        # Keep only the filename in config (for backward compat), but value is ignored on read
        self.data.setdefault("google", {})["oauth_client_path"] = OAUTH_CLIENT_FILENAME

    @property
    def oauth_token_path(self) -> str:
        # Hardcode to ./client/<OAUTH_TOKEN_FILENAME>
        return str(oauth_token_file())

    @oauth_token_path.setter
    def oauth_token_path(self, p: str) -> None:
        # Keep only the filename in config (for backward compat), but value is ignored on read
        self.data.setdefault("google", {})["oauth_token_path"] = OAUTH_TOKEN_FILENAME
