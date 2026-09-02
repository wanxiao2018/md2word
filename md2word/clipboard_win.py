"""Windows clipboard helpers for rich text paste into Word."""

from __future__ import annotations

import re
import time
from typing import Optional


def _require_win32():
    try:
        import win32clipboard  # noqa: F401
        import win32con  # noqa: F401
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("需要安装 pywin32：pip install pywin32") from exc


def _open_clipboard(retries: int = 8, delay: float = 0.05) -> None:
    """OpenClipboard can block or fail if another app holds the clipboard."""
    import win32clipboard

    last: Optional[BaseException] = None
    for _ in range(retries):
        try:
            win32clipboard.OpenClipboard()
            return
        except Exception as exc:  # noqa: BLE001
            last = exc
            time.sleep(delay)
    raise RuntimeError(f"无法打开剪贴板：{last}") from last


def get_text() -> str:
    """Read plain text from clipboard."""
    _require_win32()
    import win32clipboard
    import win32con

    _open_clipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
            data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            return data or ""
        if win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT):
            data = win32clipboard.GetClipboardData(win32con.CF_TEXT)
            if isinstance(data, bytes):
                return data.decode("utf-8", errors="replace")
            return data or ""
        return ""
    finally:
        win32clipboard.CloseClipboard()


def set_text(text: str) -> None:
    """Write plain text to clipboard."""
    _require_win32()
    import win32clipboard
    import win32con

    _open_clipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def _html_fragment(html: str) -> str:
    """Extract body content when given a full HTML document."""
    m = re.search(r"<body[^>]*>(.*)</body>", html, flags=re.I | re.S)
    if m:
        return m.group(1).strip()
    return html.strip()


def build_cf_html(html: str) -> str:
    """
    Build Windows CF_HTML clipboard payload.
    Spec: https://learn.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format
    """
    fragment = _html_fragment(html)
    # Ensure fragment is well-formed enough for Word
    if not re.search(r"<html", html, flags=re.I):
        full = (
            "<html><head><meta http-equiv=\"Content-Type\" "
            "content=\"text/html; charset=utf-8\"></head>"
            f"<body><!--StartFragment-->{fragment}<!--EndFragment--></body></html>"
        )
    else:
        # Inject fragment markers into existing document
        if "<!--StartFragment-->" not in html:
            body_m = re.search(r"(<body[^>]*>)", html, flags=re.I)
            if body_m:
                start = body_m.end()
                end_m = re.search(r"</body>", html, flags=re.I)
                end = end_m.start() if end_m else len(html)
                full = (
                    html[:start]
                    + "<!--StartFragment-->"
                    + html[start:end]
                    + "<!--EndFragment-->"
                    + html[end:]
                )
            else:
                full = (
                    "<!--StartFragment-->" + html + "<!--EndFragment-->"
                )
        else:
            full = html

    # Prefix with CF_HTML header using byte offsets (UTF-8)
    # Header placeholders first, then recompute with actual lengths.
    header_template = (
        "Version:0.9\r\n"
        "StartHTML:{start_html:010d}\r\n"
        "EndHTML:{end_html:010d}\r\n"
        "StartFragment:{start_frag:010d}\r\n"
        "EndFragment:{end_frag:010d}\r\n"
        "SourceURL:about:blank\r\n"
    )
    # Use provisional zeros to measure header size
    dummy = header_template.format(
        start_html=0, end_html=0, start_frag=0, end_frag=0
    )
    header_len = len(dummy.encode("utf-8"))

    full_bytes = full.encode("utf-8")
    # Locate fragment markers in bytes
    start_token = b"<!--StartFragment-->"
    end_token = b"<!--EndFragment-->"
    start_frag_rel = full_bytes.find(start_token)
    end_frag_rel = full_bytes.find(end_token)

    if start_frag_rel < 0 or end_frag_rel < 0:
        # No markers: whole body is fragment
        start_html = header_len
        end_html = header_len + len(full_bytes)
        start_frag = start_html
        end_frag = end_html
    else:
        start_html = header_len
        end_html = header_len + len(full_bytes)
        start_frag = header_len + start_frag_rel + len(start_token)
        end_frag = header_len + end_frag_rel

    header = header_template.format(
        start_html=start_html,
        end_html=end_html,
        start_frag=start_frag,
        end_frag=end_frag,
    )
    return header + full


def set_html(html: str, plain_fallback: Optional[str] = None) -> None:
    """Put HTML (CF_HTML) + optional plain text on the clipboard."""
    _require_win32()
    import win32clipboard
    import win32con

    cf_html = build_cf_html(html)
    # Register HTML Format
    _open_clipboard()
    try:
        win32clipboard.EmptyClipboard()
        html_format = win32clipboard.RegisterClipboardFormat("HTML Format")
        # CF_HTML is UTF-8 bytes
        win32clipboard.SetClipboardData(html_format, cf_html.encode("utf-8"))
        if plain_fallback is not None:
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain_fallback)
    finally:
        win32clipboard.CloseClipboard()


def set_rtf(rtf: str, plain_fallback: Optional[str] = None) -> None:
    """Put RTF + optional plain text on the clipboard."""
    _require_win32()
    import win32clipboard
    import win32con

    _open_clipboard()
    try:
        win32clipboard.EmptyClipboard()
        # CF_RTF is a registered format name "Rich Text Format"
        rtf_format = win32clipboard.RegisterClipboardFormat("Rich Text Format")
        # RTF is typically ASCII-compatible; encode carefully
        payload = rtf.encode("utf-8", errors="replace")
        win32clipboard.SetClipboardData(rtf_format, payload)
        if plain_fallback is not None:
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain_fallback)
    finally:
        win32clipboard.CloseClipboard()


def set_rich_for_word(
    html: Optional[str] = None,
    rtf: Optional[str] = None,
    plain: Optional[str] = None,
) -> str:
    """
    Place the richest available format(s) on the clipboard for Word paste.
    Preference: both HTML + RTF if available, else whichever exists.
    Returns a short description of what was set.
    """
    _require_win32()
    import win32clipboard
    import win32con

    if not html and not rtf and plain is None:
        raise ValueError("没有可写入剪贴板的内容")

    _open_clipboard()
    try:
        win32clipboard.EmptyClipboard()
        formats: list[str] = []

        if html:
            html_format = win32clipboard.RegisterClipboardFormat("HTML Format")
            cf_html = build_cf_html(html)
            win32clipboard.SetClipboardData(html_format, cf_html.encode("utf-8"))
            formats.append("HTML")

        if rtf:
            rtf_format = win32clipboard.RegisterClipboardFormat("Rich Text Format")
            win32clipboard.SetClipboardData(
                rtf_format, rtf.encode("utf-8", errors="replace")
            )
            formats.append("RTF")

        if plain is not None:
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain)
            formats.append("纯文本")
    finally:
        win32clipboard.CloseClipboard()

    return " + ".join(formats)
