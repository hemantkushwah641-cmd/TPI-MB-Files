from __future__ import annotations

import os
from pathlib import Path

APP_NAME = "TPI Measurement Book"
APP_VERSION = "2.4.2"
UPLOAD_VERSION = "2.4.2"
DOWNLOAD_VERSION = "2.4.2"


def get_edition() -> str:
    env = (os.environ.get("TPI_EDITION") or "").strip().lower()
    if env in ("upload", "full", "download"):
        return env
    p = Path(__file__).resolve().parent / "edition.txt"
    try:
        if p.exists():
            t = p.read_text(encoding="utf-8").strip().lower()
            if t.startswith("upload"):
                return "upload"
            if t.startswith("download"):
                return "download"
    except Exception:
        pass
    return "full"


def display_name() -> str:
    ed = get_edition()
    if ed == "upload":
        return "TPI Upload"
    if ed == "download":
        return "TPI Download"
    return APP_NAME


def display_version() -> str:
    ed = get_edition()
    if ed == "upload":
        return UPLOAD_VERSION
    if ed == "download":
        return DOWNLOAD_VERSION
    return APP_VERSION
