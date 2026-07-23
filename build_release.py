from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
VENDOR = ROOT / ".vendor"


def main() -> int:
    os.environ.setdefault("TMP", str(ROOT / "build_tmp"))
    os.environ.setdefault("TEMP", str(ROOT / "build_tmp"))
    os.environ.setdefault("PYINSTALLER_CONFIG_DIR", str(ROOT / "build_tmp" / "pyinstaller_config"))
    (ROOT / "build_tmp").mkdir(exist_ok=True)
    if VENDOR.exists():
        existing = os.environ.get("PYTHONPATH")
        os.environ["PYTHONPATH"] = (
            str(VENDOR) if not existing else str(VENDOR) + os.pathsep + existing
        )

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        "台账处理与拆分系统",
        "--paths",
        str(ROOT),
        "--add-data",
        str(ROOT / "input" / "两街一镇通告模板.txt") + os.pathsep + "input",
        "app.py",
    ]
    return subprocess.call(cmd, cwd=ROOT)


if __name__ == "__main__":
    raise SystemExit(main())
