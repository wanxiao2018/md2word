# md2word — Cross-platform build script
# Usage:  python build.py

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def main() -> int:
    if sys.platform == "win32":
        return subprocess.call([sys.executable, str(ROOT / "build_exe.py")])
    if sys.platform == "darwin":
        return subprocess.call([sys.executable, str(ROOT / "build_mac.py")])

    print("Building for Linux / other POSIX systems...")
    return subprocess.call([
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name=md2word",
        f"--paths={ROOT}",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        "--hidden-import=lxml.etree",
        str(ROOT / "main.py"),
    ])


if __name__ == "__main__":
    raise SystemExit(main())
