from __future__ import annotations

import sys
from pathlib import Path


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def ensure_app_dirs() -> None:
    for name in ("output", "logs", "resources"):
        (app_root() / name).mkdir(parents=True, exist_ok=True)

