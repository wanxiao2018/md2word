"""Detached PyInstaller builder. Writes build_status.txt when finished."""
from __future__ import annotations

import os
import subprocess
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
STATUS = ROOT / "build_status.txt"
LOG = ROOT / "build_log.txt"


def log(msg: str) -> None:
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def main() -> int:
    STATUS.write_text("running\n", encoding="utf-8")
    if LOG.exists():
        LOG.unlink()
    log("Build started")
    log(f"Python: {sys.executable}")
    try:
        ico = ROOT / "assets" / "app.ico"
        png = ROOT / "assets" / "app.png"
        sep = ";" if os.name == "nt" else ":"
        cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--windowed",
            "--onefile",
            "--name=md2word",
            f"--paths={ROOT}",
            "--hidden-import=win32clipboard",
            "--hidden-import=win32con",
            "--hidden-import=win32api",
            "--hidden-import=pywintypes",
            "--hidden-import=pythoncom",
            "--hidden-import=win32com",
            "--hidden-import=win32com.client",
            "--hidden-import=docx",
            "--hidden-import=lxml",
            "--hidden-import=lxml.etree",
            # Keep package lean
            "--exclude-module=pygame",
            "--exclude-module=matplotlib",
            "--exclude-module=numpy",
            "--exclude-module=torch",
            "--exclude-module=pandas",
            "--exclude-module=scipy",
            "--exclude-module=cv2",
            str(ROOT / "main.py"),
        ]
        if ico.is_file():
            cmd.insert(-1, f"--icon={ico}")
            cmd.insert(-1, f"--add-data={ico}{sep}assets")
        if png.is_file():
            cmd.insert(-1, f"--add-data={png}{sep}assets")
        header_png = ROOT / "assets" / "app-32.png"
        if header_png.is_file():
            cmd.insert(-1, f"--add-data={header_png}{sep}assets")
        log("CMD: " + " ".join(cmd))
        proc = subprocess.run(
            cmd,
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if proc.stdout:
            with LOG.open("a", encoding="utf-8") as f:
                f.write(proc.stdout)
        if proc.stderr:
            with LOG.open("a", encoding="utf-8") as f:
                f.write(proc.stderr)
        log(f"PyInstaller exit code: {proc.returncode}")
        exe = ROOT / "dist" / "md2word.exe"
        if proc.returncode == 0 and exe.is_file():
            size = exe.stat().st_size
            STATUS.write_text(f"ok\n{exe}\n{size}\n", encoding="utf-8")
            log(f"SUCCESS: {exe} ({size} bytes)")
            return 0
        STATUS.write_text(f"fail\ncode={proc.returncode}\n", encoding="utf-8")
        log("FAILED")
        return proc.returncode or 1
    except Exception:
        tb = traceback.format_exc()
        log(tb)
        STATUS.write_text("fail\nexception\n", encoding="utf-8")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
