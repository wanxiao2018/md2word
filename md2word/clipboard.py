"""Cross-platform clipboard helpers for plain text and rich paste into Word."""

from __future__ import annotations

import shutil
import subprocess
import sys
from typing import Optional


def paste_shortcut() -> str:
    return "⌘V" if sys.platform == "darwin" else "Ctrl+V"


def get_text() -> str:
    if sys.platform == "win32":
        from . import clipboard_win

        return clipboard_win.get_text()
    if sys.platform == "darwin":
        return _popen_text(["pbpaste"])
    return _unix_get_text()


def set_text(text: str) -> None:
    if sys.platform == "win32":
        from . import clipboard_win

        clipboard_win.set_text(text)
        return
    if sys.platform == "darwin":
        _popen_write(["pbcopy"], text)
        return
    _unix_set_text(text)


def set_rich_for_word(
    html: Optional[str] = None,
    rtf: Optional[str] = None,
    plain: Optional[str] = None,
) -> str:
    if sys.platform == "win32":
        from . import clipboard_win

        return clipboard_win.set_rich_for_word(html=html, rtf=rtf, plain=plain)
    if sys.platform == "darwin":
        return _darwin_set_rich(html=html, plain=plain)
    return _unix_set_rich(html=html, plain=plain)


def _popen_text(cmd: list[str]) -> str:
    proc = subprocess.run(cmd, capture_output=True, timeout=10)
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or b"").decode("utf-8", errors="replace").strip()
        raise RuntimeError(err or "读取剪贴板失败")
    return proc.stdout.decode("utf-8", errors="replace")


def _popen_write(cmd: list[str], text: str) -> None:
    proc = subprocess.run(
        cmd,
        input=text.encode("utf-8"),
        capture_output=True,
        timeout=10,
    )
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or b"").decode("utf-8", errors="replace").strip()
        raise RuntimeError(err or "写入剪贴板失败")


def _applescript_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace('"', '\\"')


def _darwin_set_rich(*, html: Optional[str], plain: Optional[str]) -> str:
    formats: list[str] = []
    body = html or ""
    fallback = plain if plain is not None else ""
    if html:
        hexdata = html.encode("utf-8").hex()
        escaped = _applescript_escape(fallback)
        script = (
            f'set the clipboard to {{«class HTML»:«data HTML{hexdata}», '
            f'Unicode text:"{escaped}"}}'
        )
        proc = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            timeout=20,
        )
        if proc.returncode == 0:
            formats.append("HTML")
            if plain is not None:
                formats.append("纯文本")
            return " + ".join(formats)
        # Fall through to plain pbcopy if AppleScript rejects a large payload.
    if plain is None and not body:
        raise ValueError("没有可写入剪贴板的内容")
    _popen_write(["pbcopy"], fallback or body)
    formats.append("纯文本")
    return " + ".join(formats)


def _unix_get_text() -> str:
    if shutil.which("wl-paste"):
        return _popen_text(["wl-paste", "-n"])
    if shutil.which("xclip"):
        return _popen_text(["xclip", "-selection", "clipboard", "-o"])
    if shutil.which("xsel"):
        return _popen_text(["xsel", "--clipboard", "--output"])
    raise RuntimeError("需要 wl-clipboard、xclip 或 xsel 才能读写剪贴板")


def _unix_set_text(text: str) -> None:
    if shutil.which("wl-copy"):
        _popen_write(["wl-copy"], text)
        return
    if shutil.which("xclip"):
        _popen_write(["xclip", "-selection", "clipboard"], text)
        return
    if shutil.which("xsel"):
        _popen_write(["xsel", "--clipboard", "--input"], text)
        return
    raise RuntimeError("需要 wl-clipboard、xclip 或 xsel 才能读写剪贴板")


def _unix_set_rich(*, html: Optional[str], plain: Optional[str]) -> str:
    formats: list[str] = []
    if html and shutil.which("wl-copy"):
        _popen_write(["wl-copy", "--type", "text/html"], html)
        formats.append("HTML")
        if plain is not None:
            _popen_write(["wl-copy", "--type", "text/plain"], plain)
            formats.append("纯文本")
        return " + ".join(formats)
    if html and shutil.which("xclip"):
        _popen_write(["xclip", "-selection", "clipboard", "-t", "text/html"], html)
        formats.append("HTML")
        return " + ".join(formats)
    payload = plain if plain is not None else (html or "")
    if not payload:
        raise ValueError("没有可写入剪贴板的内容")
    _unix_set_text(payload)
    formats.append("纯文本")
    return " + ".join(formats)
