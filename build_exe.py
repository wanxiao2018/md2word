# md2word — PyInstaller build script (run from project root)
# Usage:  python build_exe.py

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
NAME = "md2word"


def main() -> int:
    # Delegate to the more robust runner that writes build_status.txt
    code = subprocess.call([sys.executable, str(ROOT / "run_build.py")], cwd=str(ROOT))
    exe = DIST / f"{NAME}.exe"
    if code == 0 and exe.is_file():
        shutil.copy2(exe, ROOT / f"{NAME}.exe")
        size_mb = exe.stat().st_size / (1024 * 1024)
        print(f"\nOK: {exe}")
        print(f"Also copied to: {ROOT / f'{NAME}.exe'}")
        print(f"Size: {size_mb:.1f} MB")
        return 0
    print("Build failed. See build_log.txt / build_status.txt")
    return code or 1


if __name__ == "__main__":
    raise SystemExit(main())
