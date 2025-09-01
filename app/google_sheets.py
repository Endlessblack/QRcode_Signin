from __future__ import annotations

import datetime as dt
from typing import Dict, List, Any

import gspread


class GoogleSheetsClient:
    def __init__(self,
                 credentials_path: str,
                 spreadsheet_id: str,
                 worksheet_name: str,
                 auth_method: str = "service_account",
                 oauth_client_path: str | None = None,
                 oauth_token_path: str | None = None) -> None:
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.worksheet_name = worksheet_name
        self.auth_method = (auth_method or "service_account").lower()
        self.oauth_client_path = oauth_client_path or ""
        self.oauth_token_path = oauth_token_path or "token.json"
        self._client = None
        self._ws = None

    def connect(self) -> None:
        if self.auth_method == 'oauth':
            # OAuth installed-app flow; opens browser on first run to create token
            if self.oauth_client_path:
                gc = gspread.oauth(credentials_filename=self.oauth_client_path,
                                    authorized_user_filename=self.oauth_token_path)
            else:
                # Fallback to default filenames in current directory
                gc = gspread.oauth()
        else:
            gc = gspread.service_account(filename=self.credentials_path)
        sh = gc.open_by_key(self.spreadsheet_id)
        try:
            ws = sh.worksheet(self.worksheet_name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=self.worksheet_name, rows=1000, cols=26)
        self._client = gc
        self._ws = ws
        self._ensure_headers(["timestamp", "event", "id", "name", "raw"])  # basic headers

    def _ensure_headers(self, headers: List[str]) -> None:
        assert self._ws is not None
        values = self._ws.get_all_values()
        if not values:
            self._ws.append_row(headers)
        else:
            existing = values[0]
            new_headers = list(dict.fromkeys(existing + headers))
            if new_headers != existing:
                self._ws.resize(rows=max(1000, len(values)), cols=max(len(new_headers), len(existing)))
                self._ws.update('1:1', [new_headers])

    def append_signin(self, payload: Dict[str, Any]) -> None:
        if self._ws is None:
            self.connect()
        assert self._ws is not None

        # Build row based on headers
        headers = self._ws.row_values(1)
        timestamp = dt.datetime.now().astimezone().isoformat(timespec="seconds")

        # Expand payload to flat dict
        base = {
            "timestamp": timestamp,
            "event": payload.get("event", ""),
            "id": payload.get("id", ""),
            "name": payload.get("name", ""),
            "raw": payload.get("raw", ""),
        }

        extra = payload.get("extra", {})
        if isinstance(extra, dict):
            for k, v in extra.items():
                base[str(k)] = v

        # Ensure headers cover keys
        need_headers = [h for h in base.keys() if h not in headers]
        if need_headers:
            self._ensure_headers(headers + need_headers)
            headers = self._ws.row_values(1)

        row = [base.get(h, "") for h in headers]
        self._ws.append_row(row)

    # Read all records as list of dicts (header row determines keys)
    def fetch_records(self) -> List[Dict[str, Any]]:
        if self._ws is None:
            self.connect()
        assert self._ws is not None
        try:
            return self._ws.get_all_records()
        except Exception:
            # Fallback: manual conversion
            values = self._ws.get_all_values()
            if not values:
                return []
            headers = [str(h).strip() for h in (values[0] or [])]
            out: List[Dict[str, Any]] = []
            for row in values[1:]:
                d: Dict[str, Any] = {}
                for i, h in enumerate(headers):
                    d[h] = row[i] if i < len(row) else ""
                out.append(d)
            return out
