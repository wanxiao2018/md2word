"""Microsoft Word helpers: native clipboard copy and PDF export.

Windows uses COM. macOS uses AppleScript when Word.app is installed.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import Optional


def word_installed() -> bool:
    """Does Microsoft Word appear to be installed? Does not launch Word."""
    if os.name == "nt":
        try:
            import winreg

            winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application")
            return True
        except OSError:
            return False
    if sys.platform == "darwin":
        return Path("/Applications/Microsoft Word.app").is_dir()
    return False


def _dispatch_word():
    """Return (word, owned). owned=True means we launched Word and must quit it."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    owned = False
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:  # noqa: BLE001
        word = win32com.client.DispatchEx("Word.Application")
        owned = True
        word.Visible = False
    try:
        word.DisplayAlerts = 0
    except Exception:  # noqa: BLE001
        pass
    try:
        # msoAutomationSecurityForceDisable = 3
        word.AutomationSecurity = 3
    except Exception:  # noqa: BLE001
        pass
    return word, owned


def _quit_word(word, owned: bool) -> None:
    if owned:
        try:
            word.Quit()
        except Exception:  # noqa: BLE001
            pass
    try:
        import pythoncom

        pythoncom.CoUninitialize()
    except Exception:  # noqa: BLE001
        pass


def copy_docx_to_clipboard(docx_path: Path) -> None:
    """Open a .docx in Word and copy its content to the clipboard.

    Word places native OMML equations on the clipboard, so paste into another
    document yields editable Word formulas rather than LaTeX or plain text.
    """
    docx_path = Path(docx_path).resolve()
    if not docx_path.is_file():
        raise FileNotFoundError(str(docx_path))
    if sys.platform == "darwin":
        _darwin_copy_docx(docx_path)
        return
    if os.name != "nt":
        raise RuntimeError("当前系统没有可用的 Word 自动化接口")

    word, owned = _dispatch_word()
    doc = None
    try:
        doc = word.Documents.Open(
            str(docx_path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        doc.Content.Copy()
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
        _quit_word(word, owned)


def export_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    """Export a .docx to PDF via Word (keeps OMML formulas and Chinese fonts)."""
    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    if not docx_path.is_file():
        raise FileNotFoundError(str(docx_path))
    if sys.platform == "darwin":
        _darwin_export_pdf(docx_path, pdf_path)
        return
    if os.name != "nt":
        raise RuntimeError("当前系统没有可用的 Word 自动化接口")

    word, owned = _dispatch_word()
    doc = None
    try:
        doc = word.Documents.Open(
            str(docx_path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        # wdExportFormatPDF = 17
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf_path),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1,
        )
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
        _quit_word(word, owned)


def find_soffice() -> Optional[str]:
    import shutil

    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return found
    candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
        / "LibreOffice"
        / "program"
        / "soffice.exe",
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path("/usr/local/bin/soffice"),
        Path("/opt/homebrew/bin/soffice"),
    ]
    for path in candidates:
        if path.is_file():
            return str(path)
    return None


def _posix_as(path: Path) -> str:
    return str(path.resolve()).replace("\\", "\\\\").replace('"', '\\"')


def _osascript(script: str) -> None:
    proc = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=90,
    )
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "osascript failed").strip()
        raise RuntimeError(err)


def _darwin_copy_docx(docx_path: Path) -> None:
    posix = _posix_as(docx_path)
    script = f'''
tell application "Microsoft Word"
  set theDoc to open POSIX file "{posix}"
  copy object (text object of theDoc)
  close theDoc saving no
end tell
'''
    _osascript(script)


def _darwin_export_pdf(docx_path: Path, pdf_path: Path) -> None:
    src = _posix_as(docx_path)
    dst = _posix_as(pdf_path)
    script = f'''
tell application "Microsoft Word"
  set theDoc to open POSIX file "{src}"
  save as theDoc file name "{dst}" file format format PDF
  close theDoc saving no
end tell
'''
    _osascript(script)
    if not pdf_path.is_file():
        raise RuntimeError("Word 未能导出 PDF")
