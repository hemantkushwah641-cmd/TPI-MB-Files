from __future__ import annotations

import os
from pathlib import Path

APP_NAME = "TPI Measurement Book"
APP_VERSION = "2.4.2"
UPLOAD_VERSION = "2.4.2"


def get_edition() -> str:
    return "full"


def display_name() -> str:
    return "TPI Upload" if get_edition() == "upload" else APP_NAME


def display_version() -> str:
    return UPLOAD_VERSION if get_edition() == "upload" else APP_VERSION
