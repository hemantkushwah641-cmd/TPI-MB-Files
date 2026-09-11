from __future__ import annotations

import os
from pathlib import Path

APP_NAME = "TPI Measurement Book"
APP_VERSION = "2.3.1"
UPLOAD_VERSION = "1.0.0"


def get_edition() -> str:
    env = (os.environ.get("TPI_EDITION") or "").strip().lower()
    if env in ("upload", "full"):
        return env
    p = Path(__file__).resolve().parent / "edition.txt"
    try:
        if p.exists() and p.read_text(encoding="utf-8").strip().lower().startswith("upload"):
            return "upload"
    except Exception:
        pass
    return "full"


def display_name() -> str:
    return "TPI Upload" if get_edition() == "upload" else APP_NAME


def display_version() -> str:
    return UPLOAD_VERSION if get_edition() == "upload" else APP_VERSION
