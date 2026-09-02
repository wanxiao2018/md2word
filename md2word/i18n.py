"""UI language: English by default, Chinese when the OS locale is zh."""

from __future__ import annotations

import locale
import os
from pathlib import Path

_LANG = "en"

_CONFIG = Path.home() / ".md2word" / "lang"

_STRINGS = {
    "en": {
        "app_title": "md2word — Markdown to Word",
        "tagline": "Markdown → Word  ·  Editable equations  ·  Auto layout",
        "engines": "Engines",
        "engine_builtin": "Built-in",
        "engine_no_word": "Word not found",
        "import": "Import",
        "paste_clip": "Paste clipboard",
        "open_file": "Open file",
        "clear": "Clear",
        "watch": "Watch clipboard",
        "auto_convert": "Auto-convert while watching",
        "export": "Export",
        "convert_copy": "Convert and copy",
        "save_word": "Save Word",
        "export_pdf": "Export PDF",
        "convert_open": "Convert and open Word",
        "body_hint": "Body: 2-char indent  ·  justified  ·  1.5 line spacing  ·  editable equations  ·  no section rules",
        "md_source": "Markdown source",
        "char_count": "{n} chars",
        "ready": "Ready",
        "placeholder": (
            "Paste Markdown from an AI tool here…\n\n"
            "Supports: headings / bold / lists / code / tables / quotes / links / LaTeX\n\n"
            "Typical flow:\n"
            "1. Copy Markdown from ChatGPT / Claude / Cursor\n"
            "2. Click “Paste clipboard” or {accel}+Shift+V\n"
            "3. Click “Convert and copy”\n"
            "4. Paste in Word with {paste} (equations stay editable)\n"
        ),
        "shortcuts": (
            "Shortcuts: {accel}+Shift+V paste  ·  {accel}+Enter convert  ·  "
            "{accel}+S save Word  ·  {accel}+P export PDF  ·  Esc clear"
        ),
        "cleared": "Cleared",
        "error": "Error",
        "clip_read_fail": "Could not read the clipboard:\n{exc}",
        "clip_empty": "Clipboard is empty",
        "imported": "Pasted {n} characters from clipboard ({kind})",
        "kind_md": "Markdown",
        "kind_text": "text",
        "open_md": "Open Markdown file",
        "all_files": "All files",
        "read_fail": "Could not read the file:\n{exc}",
        "opened": "Opened: {path}",
        "hint": "Note",
        "need_content": "Paste or import Markdown first.",
        "converting": "Converting to Word equations and copying…",
        "clip_write_fail": "Could not write the clipboard:\n{err}",
        "failed": "Failed: {err}",
        "convert_fail": "Conversion failed",
        "save_docx_title": "Save as Word document",
        "word_docs": "Word documents",
        "saving_docx": "Building .docx…",
        "save_fail": "Save failed",
        "save_ok": "Saved",
        "open_in_word": "{msg}\n\nOpen in Word now?",
        "saved_no_open": "Saved, but could not open:\n{exc}",
        "export_pdf_title": "Export PDF",
        "pdf_docs": "PDF documents",
        "exporting_pdf": "Exporting PDF…",
        "export_fail": "Export failed",
        "export_ok": "Exported",
        "open_now": "{msg}\n\nOpen it now?",
        "opening": "{msg} — opening…",
        "made_no_open": "Document created:\n{path}\n\nCould not open it:\n{exc}",
        "watch_on": "Clipboard watching is on",
        "watch_off": "Clipboard watching is off",
        "watch_imported": "Watched Markdown ({n} chars), imported",
        "watch_converting": "Watch mode: converting…",
        "auto_fail": "Auto-convert failed: {msg}",
        "empty": "Nothing to convert.",
        "saved_path": "Saved: {path}",
        "no_pandoc_builtin": "Pandoc not found, using the built-in converter.",
        "saved_fallback": "Saved: {path} ({detail})",
        "convert_fail_exc": "Conversion failed: {exc}",
        "html_ok": "HTML conversion succeeded",
        "no_pandoc": "Pandoc not found",
        "html_ok_detail": "HTML conversion succeeded ({detail})",
        "docx_fail": "Could not build a Word document.",
        "copied_word": "Copied Word format (editable equations). Paste in Word with {paste}",
        "no_word": "Microsoft Word not detected",
        "copy_fail": "Copy failed: {word}; HTML fallback also failed: {html}",
        "copied_rich": "Copied rich text ({formats}). Word equation path unused ({word}). Install Word for best results.",
        "pdf_fail": "PDF export failed: Word ({word}); LibreOffice ({lo})",
        "pdf_need": "PDF export needs Microsoft Word or LibreOffice. Word: {word}",
        "saved_lo": "Saved: {path} (LibreOffice)",
        "fmt_html": "HTML",
        "fmt_rtf": "RTF",
        "fmt_plain": "plain text",
    },
    "zh": {
        "app_title": "md2word — Markdown 转 Word",
        "tagline": "Markdown → Word  ·  公式可编辑  ·  正文自动排版",
        "engines": "引擎",
        "engine_builtin": "内置转换器",
        "engine_no_word": "未检测到 Word",
        "import": "导入",
        "paste_clip": "从剪贴板导入",
        "open_file": "打开文件",
        "clear": "清空",
        "watch": "监视剪贴板",
        "auto_convert": "监视时自动转换",
        "export": "输出",
        "convert_copy": "转换并复制到剪贴板",
        "save_word": "保存 Word",
        "export_pdf": "导出 PDF",
        "convert_open": "转换并打开 Word",
        "body_hint": "正文格式：首行缩进 2 字符  ·  两端对齐  ·  1.5 倍行距  ·  公式转为 Word 可编辑公式  ·  去掉分节横线",
        "md_source": "Markdown 原文",
        "char_count": "{n} 字",
        "ready": "就绪",
        "placeholder": (
            "在这里粘贴 AI 输出的 Markdown 文本…\n\n"
            "支持：标题 / 粗体斜体 / 列表 / 代码块 / 表格 / 引用 / 链接 / LaTeX 公式\n\n"
            "推荐流程：\n"
            "1. 在 ChatGPT / Claude / Cursor 等软件中复制 Markdown\n"
            "2. 点击「从剪贴板导入」或 {accel}+Shift+V\n"
            "3. 点击「转换并复制到剪贴板」\n"
            "4. 到 Word 中 {paste} 粘贴（公式为可编辑的 Word 公式）\n"
        ),
        "shortcuts": (
            "快捷键：{accel}+Shift+V 导入  ·  {accel}+Enter 转换并复制  ·  "
            "{accel}+S 保存 Word  ·  {accel}+P 导出 PDF  ·  Esc 清空"
        ),
        "cleared": "已清空",
        "error": "错误",
        "clip_read_fail": "读取剪贴板失败：\n{exc}",
        "clip_empty": "剪贴板为空",
        "imported": "已从剪贴板导入 {n} 字符（识别为{kind}）",
        "kind_md": "Markdown",
        "kind_text": "文本",
        "open_md": "打开 Markdown 文件",
        "all_files": "所有文件",
        "read_fail": "无法读取文件：\n{exc}",
        "opened": "已打开：{path}",
        "hint": "提示",
        "need_content": "请先粘贴或导入 Markdown 内容。",
        "converting": "正在转换为 Word 公式并复制…",
        "clip_write_fail": "写入剪贴板失败：\n{err}",
        "failed": "失败：{err}",
        "convert_fail": "转换失败",
        "save_docx_title": "保存为 Word 文档",
        "word_docs": "Word 文档",
        "saving_docx": "正在生成 .docx…",
        "save_fail": "保存失败",
        "save_ok": "保存成功",
        "open_in_word": "{msg}\n\n是否立即用 Word 打开？",
        "saved_no_open": "已保存，但无法自动打开：\n{exc}",
        "export_pdf_title": "导出为 PDF",
        "pdf_docs": "PDF 文档",
        "exporting_pdf": "正在导出 PDF…",
        "export_fail": "导出失败",
        "export_ok": "导出成功",
        "open_now": "{msg}\n\n是否立即打开？",
        "opening": "{msg} — 正在打开…",
        "made_no_open": "文档已生成：\n{path}\n\n但无法自动打开：\n{exc}",
        "watch_on": "已开启剪贴板监视",
        "watch_off": "已关闭剪贴板监视",
        "watch_imported": "监视到 Markdown（{n} 字），已导入",
        "watch_converting": "监视模式：正在自动转换…",
        "auto_fail": "自动转换失败：{msg}",
        "empty": "内容为空，无法转换。",
        "saved_path": "已保存：{path}",
        "no_pandoc_builtin": "未找到 Pandoc，使用内置转换器。",
        "saved_fallback": "已保存：{path}（{detail}）",
        "convert_fail_exc": "转换失败：{exc}",
        "html_ok": "HTML 转换成功",
        "no_pandoc": "未找到 Pandoc",
        "html_ok_detail": "HTML 转换成功（{detail}）",
        "docx_fail": "生成 Word 文档失败。",
        "copied_word": "已复制 Word 格式（含可编辑公式）到剪贴板，请到 Word 中 {paste}",
        "no_word": "未检测到 Microsoft Word",
        "copy_fail": "复制失败：{word}；HTML 回退也失败：{html}",
        "copied_rich": "已复制富文本（{formats}）。未走 Word 公式通道（{word}），建议安装 Word 后重试。",
        "pdf_fail": "PDF 导出失败：Word（{word}）；LibreOffice（{lo}）",
        "pdf_need": "PDF 导出需要安装 Microsoft Word 或 LibreOffice。Word：{word}",
        "saved_lo": "已保存：{path}（LibreOffice）",
        "fmt_html": "HTML",
        "fmt_rtf": "RTF",
        "fmt_plain": "纯文本",
    },
}


def detect_lang() -> str:
    env = (os.environ.get("MD2WORD_LANG") or "").strip().lower().replace("-", "_")
    if env in {"zh", "zh_cn", "zh_tw", "cn"}:
        return "zh"
    if env in {"en", "en_us", "en_gb"}:
        return "en"
    if _CONFIG.is_file():
        saved = _CONFIG.read_text(encoding="utf-8").strip().lower()
        if saved in _STRINGS:
            return saved
    loc = ""
    try:
        loc = locale.getdefaultlocale()[0] or ""
    except Exception:  # noqa: BLE001
        loc = ""
    if not loc:
        loc = os.environ.get("LANG") or os.environ.get("LC_ALL") or ""
    if loc.lower().replace("-", "_").startswith("zh"):
        return "zh"
    return "en"


def lang() -> str:
    return _LANG


def set_lang(code: str) -> None:
    global _LANG
    _LANG = code if code in _STRINGS else "en"
    _CONFIG.parent.mkdir(parents=True, exist_ok=True)
    _CONFIG.write_text(_LANG, encoding="utf-8")


def t(key: str, **kwargs) -> str:
    table = _STRINGS.get(_LANG) or _STRINGS["en"]
    text = table.get(key) or _STRINGS["en"].get(key) or key
    if kwargs:
        return text.format(**kwargs)
    return text


_LANG = detect_lang()
