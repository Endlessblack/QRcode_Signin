from __future__ import annotations

import sys
from pathlib import Path
import shutil


def app_root() -> Path:
    """Return the base directory for reading/writing app data.

    - When frozen (PyInstaller), use the executable directory.
    - Otherwise, use current working directory so local runs behave intuitively.
    """
    try:
        if getattr(sys, "frozen", False):  # PyInstaller/py2exe
            return Path(sys.executable).resolve().parent
    except Exception:
        pass
    return Path.cwd()


def config_dir() -> Path:
    return app_root() / "setting"


def client_dir() -> Path:
    return app_root() / "client"


def config_file() -> Path:
    return config_dir() / "config.json"


# Fixed filenames for OAuth client & token
OAUTH_CLIENT_FILENAME = "client.json"
OAUTH_TOKEN_FILENAME = "token.json"
OFFLINE_QUEUE_FILENAME = "offline_queue.csv"


def oauth_client_file() -> Path:
    return client_dir() / OAUTH_CLIENT_FILENAME


def oauth_token_file() -> Path:
    return client_dir() / OAUTH_TOKEN_FILENAME


def ensure_dirs() -> None:
    # Create expected folders if missing
    try:
        config_dir().mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    try:
        client_dir().mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    # If running from a bundled EXE and the packaged client.json exists,
    # ensure it's available at ./client/client.json for the app to use.
    try:
        src = _bundled_client_json_path()
        dst = oauth_client_file()
        if src and src.exists() and not dst.exists():
            shutil.copyfile(src, dst)
    except Exception:
        pass


def _bundled_client_json_path() -> Path | None:
    """Return the path to client.json inside a PyInstaller bundle (if present).

    When bundling, include it via: --add-data "client/client.json;client"
    so it ends up at <_MEIPASS>/client/client.json during runtime.
    """
    base = Path(getattr(sys, "_MEIPASS", "")) if getattr(sys, "frozen", False) else None
    if not base:
        return None
    p = base / "client" / OAUTH_CLIENT_FILENAME
    return p if p.exists() else None


def offline_queue_file() -> Path:
    return config_dir() / OFFLINE_QUEUE_FILENAME
