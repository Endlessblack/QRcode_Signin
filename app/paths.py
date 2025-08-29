from __future__ import annotations

import sys
from pathlib import Path


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
