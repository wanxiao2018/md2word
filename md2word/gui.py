"""md2word desktop GUI."""

from __future__ import annotations

import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Optional

from . import clipboard, converter
from .i18n import lang as ui_lang
from .i18n import set_lang, t
from .wordcom import word_installed


APP_TITLE = "md2word"
APP_MIN_SIZE = (960, 680)
DEFAULT_EXPORT_DIR = Path.home() / "Documents" / "md2word"

HEADER_BG = "#0f2744"
HEADER_FG = "#f8fafc"
HEADER_MUTED = "#93c5fd"
PAGE_BG = "#e7eef6"
CARD_BG = "#ffffff"
ACCENT = "#1d4ed8"
ACCENT_DARK = "#1e40af"
STATUS_BG = "#dbe4f0"
BORDER = "#c9d4e3"
TEXT_MAIN = "#0f172a"
TEXT_SUB = "#334155"
TEXT_MUTED = "#64748b"
OK_FG = "#15803d"
ERR_FG = "#b91c1c"
CHIP_ON = "#86efac"
CHIP_OFF = "#94a3b8"

# Rendered as text everywhere, so the mark never depends on the ttk theme.
CHECK = "√"
IS_MAC = sys.platform == "darwin"
ACCEL = "⌘" if IS_MAC else "Ctrl"


def _ui_font() -> str:
    if IS_MAC:
        return "PingFang SC"
    if sys.platform == "win32":
        return "Microsoft YaHei UI"
    return "Noto Sans CJK SC"


def _editor_font() -> str:
    if IS_MAC:
        return "Menlo"
    if sys.platform == "win32":
        return "Consolas"
    return "DejaVu Sans Mono"


UI_FONT = _ui_font()
EDITOR_FONT = _editor_font()


def _resource_path(*parts: str) -> Path:
    """Project file path, including when frozen by PyInstaller."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        base = Path(__file__).resolve().parents[1]
    return base.joinpath(*parts)


class CheckToggle(tk.Frame):
    """Checkbox that draws a real 「√」 mark.

    ttk's themed indicator paints its own glyph (a cross under ``clam``), so the
    box and the mark are drawn here instead.
    """

    BOX = 18

    def __init__(
        self,
        parent,
        text: str,
        variable: tk.BooleanVar,
        command=None,
        bg: str = CARD_BG,
    ) -> None:
        super().__init__(parent, bg=bg, cursor="hand2")
        self._var = variable
        self._command = command
        self._bg = bg
        self._hover = False

        box = self.BOX
        self._canvas = tk.Canvas(
            self,
            width=box,
            height=box,
            bg=bg,
            bd=0,
            highlightthickness=0,
            takefocus=0,
        )
        self._canvas.pack(side=tk.LEFT)
        self._rect = self._canvas.create_rectangle(1, 1, box - 1, box - 1, width=1)
        self._mark = self._canvas.create_text(
            box / 2 + 1,
            box / 2,
            text=CHECK,
            fill="#ffffff",
            font=(UI_FONT, 11, "bold"),
        )
        self._text = tk.Label(self, text=text, bg=bg, font=(UI_FONT, 9))
        self._text.pack(side=tk.LEFT, padx=(6, 0))

        for widget in (self, self._canvas, self._text):
            widget.bind("<Button-1>", self._toggle)
            widget.bind("<Enter>", self._enter)
            widget.bind("<Leave>", self._leave)
        self._render()

    def _toggle(self, _event=None) -> None:
        self._var.set(not self._var.get())
        self._render()
        if self._command:
            self._command()

    def _enter(self, _event=None) -> None:
        self._hover = True
        self._render()

    def _leave(self, _event=None) -> None:
        self._hover = False
        self._render()

    def _render(self) -> None:
        on = self._var.get()
        if on:
            fill = ACCENT_DARK if self._hover else ACCENT
            outline = fill
        else:
            fill = self._bg
            outline = ACCENT if self._hover else "#a9b6c8"
        self._canvas.itemconfigure(self._rect, fill=fill, outline=outline)
        self._canvas.itemconfigure(self._mark, state=tk.NORMAL if on else tk.HIDDEN)
        self._text.configure(fg=TEXT_MAIN if on else TEXT_MUTED)

    def refresh(self) -> None:
        """Re-read the variable (for programmatic changes)."""
        self._render()


class Md2WordApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(t("app_title"))
        self.minsize(*APP_MIN_SIZE)
        self.geometry("1080x760")
        self.configure(bg=PAGE_BG)
        self._app_icon: Optional[tk.PhotoImage] = None
        self._header_icon: Optional[tk.PhotoImage] = None
        self._set_app_icon()

        self._watch_job: Optional[str] = None
        self._last_clip_hash: Optional[int] = None
        self._busy = False
        self._status_var = tk.StringVar(value=t("ready"))
        self._status_icon_var = tk.StringVar(value="·")
        self._count_var = tk.StringVar(value=t("char_count", n=0))
        self._watch_var = tk.BooleanVar(value=False)
        self._auto_convert_var = tk.BooleanVar(value=True)
        self._toggles: list[CheckToggle] = []
        self._action_btns: list[ttk.Button] = []
        self._pandoc_ok = converter.find_pandoc() is not None
        self._word_ok = word_installed()
        self._last_docx: Optional[Path] = None

        self._build_style()
        self._build_ui()
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.update_idletasks()
        w, h = 1080, 760
        x = (self.winfo_screenwidth() - w) // 2
        y = max(0, (self.winfo_screenheight() - h) // 3)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _set_app_icon(self) -> None:
        ico = _resource_path("assets", "app.ico")
        png = _resource_path("assets", "app.png")
        header_png = _resource_path("assets", "app-32.png")
        try:
            if ico.is_file():
                self.iconbitmap(str(ico.resolve()))
        except tk.TclError:
            pass
        try:
            if png.is_file():
                self._app_icon = tk.PhotoImage(file=str(png))
                self.iconphoto(True, self._app_icon)
            if header_png.is_file():
                self._header_icon = tk.PhotoImage(file=str(header_png))
            elif self._app_icon is not None and self._app_icon.width() > 40:
                factor = max(1, round(self._app_icon.width() / 32))
                self._header_icon = self._app_icon.subsample(factor, factor)
        except tk.TclError:
            pass

    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=PAGE_BG)
        style.configure("Card.TFrame", background=CARD_BG)
        style.configure("TLabel", background=PAGE_BG, font=(UI_FONT, 10))
        style.configure("Card.TLabel", background=CARD_BG, font=(UI_FONT, 10))
        style.configure(
            "CardTitle.TLabel",
            background=CARD_BG,
            foreground="#0f172a",
            font=(UI_FONT, 10, "bold"),
        )
        style.configure(
            "Hint.TLabel",
            background=CARD_BG,
            foreground="#64748b",
            font=(UI_FONT, 9),
        )
        style.configure("Sub.TLabel", background=PAGE_BG, foreground="#5b6470", font=(UI_FONT, 9))
        style.configure(
            "TButton",
            font=(UI_FONT, 10),
            padding=(12, 7),
            background="#f1f5fa",
            foreground=TEXT_SUB,
            bordercolor=BORDER,
            focuscolor=BORDER,
            relief="flat",
        )
        style.map(
            "TButton",
            background=[("pressed", "#dbe4f0"), ("active", "#e6edf7"), ("disabled", "#f4f6f9")],
            foreground=[("active", ACCENT), ("disabled", "#a9b6c8")],
            bordercolor=[("active", ACCENT)],
        )
        style.configure(
            "Accent.TButton",
            font=(UI_FONT, 10, "bold"),
            padding=(16, 8),
            background=ACCENT,
            foreground="#ffffff",
            bordercolor=ACCENT,
            focuscolor=ACCENT,
        )
        style.map(
            "Accent.TButton",
            background=[("!disabled", ACCENT), ("active", ACCENT_DARK), ("disabled", "#94a3b8")],
            foreground=[("!disabled", "#ffffff"), ("disabled", "#e2e8f0")],
            bordercolor=[("!disabled", ACCENT), ("active", ACCENT_DARK)],
        )
        style.configure("TLabelframe", background=CARD_BG, relief="flat")
        style.configure("TLabelframe.Label", background=CARD_BG, font=(UI_FONT, 10, "bold"))

    def _card(self, parent) -> tuple[tk.Frame, tk.Frame]:
        outer = tk.Frame(parent, bg=PAGE_BG)
        card = tk.Frame(outer, bg=CARD_BG, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill=tk.BOTH, expand=True)
        return outer, card

    def _chip(self, parent, text: str, ok: bool) -> tk.Frame:
        """A 「√ 名称」 badge for the header engine row."""
        frame = tk.Frame(parent, bg=HEADER_BG)
        tk.Label(
            frame,
            text=CHECK if ok else "·",
            bg=HEADER_BG,
            fg=CHIP_ON if ok else CHIP_OFF,
            font=(UI_FONT, 10, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            frame,
            text=text,
            bg=HEADER_BG,
            fg="#dbeafe" if ok else CHIP_OFF,
            font=(UI_FONT, 9),
        ).pack(side=tk.LEFT, padx=(4, 0))
        return frame

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=HEADER_BG)
        header.pack(fill=tk.X)
        inner = tk.Frame(header, bg=HEADER_BG)
        inner.pack(fill=tk.X, padx=22, pady=16)

        titles = tk.Frame(inner, bg=HEADER_BG)
        titles.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if self._header_icon is not None:
            tk.Label(titles, image=self._header_icon, bg=HEADER_BG, bd=0).pack(
                side=tk.LEFT, padx=(0, 12)
            )
        title_col = tk.Frame(titles, bg=HEADER_BG)
        title_col.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(
            title_col,
            text="md2word",
            bg=HEADER_BG,
            fg=HEADER_FG,
            font=(UI_FONT, 18, "bold"),
        ).pack(anchor="w")
        tk.Label(
            title_col,
            text=t("tagline"),
            bg=HEADER_BG,
            fg=HEADER_MUTED,
            font=(UI_FONT, 10),
        ).pack(anchor="w", pady=(4, 0))

        self._lang_bar(inner)
        engines = tk.Frame(inner, bg=HEADER_BG)
        engines.pack(side=tk.RIGHT, anchor="e", padx=(0, 18))
        tk.Label(
            engines,
            text=t("engines"),
            bg=HEADER_BG,
            fg="#7f9dc4",
            font=(UI_FONT, 9),
        ).pack(anchor="e")
        chips = tk.Frame(engines, bg=HEADER_BG)
        chips.pack(anchor="e", pady=(4, 0))
        self._chip(chips, "Pandoc" if self._pandoc_ok else t("engine_builtin"), self._pandoc_ok).pack(side=tk.LEFT)
        self._chip(chips, "Word" if self._word_ok else t("engine_no_word"), self._word_ok).pack(
            side=tk.LEFT, padx=(14, 0)
        )

        root = ttk.Frame(self, padding=(16, 14, 16, 10))
        root.pack(fill=tk.BOTH, expand=True)

        actions_outer, actions = self._card(root)
        actions_outer.pack(fill=tk.X, pady=(0, 12))
        pad = tk.Frame(actions, bg=CARD_BG)
        pad.pack(fill=tk.X, padx=14, pady=12)

        import_row = tk.Frame(pad, bg=CARD_BG)
        import_row.pack(fill=tk.X)
        tk.Label(import_row, text=t("import"), bg=CARD_BG, fg="#334155", font=(UI_FONT, 9, "bold")).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        self._btn(import_row, t("paste_clip"), self.import_clipboard).pack(side=tk.LEFT, padx=(0, 6))
        self._btn(import_row, t("open_file"), self.open_file).pack(side=tk.LEFT, padx=(0, 6))
        self._btn(import_row, t("clear"), self.clear_editor).pack(side=tk.LEFT, padx=(0, 16))
        watch_toggle = CheckToggle(
            import_row,
            t("watch"),
            self._watch_var,
            command=self._toggle_watch,
        )
        watch_toggle.pack(side=tk.LEFT)
        auto_toggle = CheckToggle(import_row, t("auto_convert"), self._auto_convert_var)
        auto_toggle.pack(side=tk.LEFT, padx=(16, 0))
        self._toggles = [watch_toggle, auto_toggle]

        tk.Frame(pad, bg="#e2e8f0", height=1).pack(fill=tk.X, pady=10)

        export_row = tk.Frame(pad, bg=CARD_BG)
        export_row.pack(fill=tk.X)
        tk.Label(export_row, text=t("export"), bg=CARD_BG, fg="#334155", font=(UI_FONT, 9, "bold")).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        self._btn(export_row, t("convert_copy"), self.convert_to_clipboard).pack(side=tk.LEFT, padx=(0, 6))
        self._btn(export_row, t("save_word"), self.save_docx).pack(side=tk.LEFT, padx=(0, 6))
        self._btn(export_row, t("export_pdf"), self.save_pdf).pack(side=tk.LEFT, padx=(0, 6))
        self._btn(export_row, t("convert_open"), self.convert_and_open).pack(side=tk.LEFT)

        tk.Label(
            pad,
            text=t("body_hint"),
            bg=CARD_BG,
            fg="#64748b",
            font=(UI_FONT, 9),
        ).pack(anchor="w", pady=(10, 0))

        editor_outer, editor_card = self._card(root)
        editor_outer.pack(fill=tk.BOTH, expand=True)

        editor_head = tk.Frame(editor_card, bg=CARD_BG)
        editor_head.pack(fill=tk.X, padx=14, pady=(10, 4))
        tk.Label(
            editor_head,
            text=t("md_source"),
            bg=CARD_BG,
            fg="#0f172a",
            font=(UI_FONT, 10, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            editor_head,
            textvariable=self._count_var,
            bg=CARD_BG,
            fg="#64748b",
            font=(UI_FONT, 9),
        ).pack(side=tk.RIGHT)

        editor_wrap = tk.Frame(editor_card, bg=CARD_BG)
        editor_wrap.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.editor = tk.Text(
            editor_wrap,
            wrap=tk.WORD,
            font=(EDITOR_FONT, 11),
            undo=True,
            relief=tk.FLAT,
            bg="#fbfcfe",
            fg="#1e293b",
            insertbackground="#0f172a",
            padx=12,
            pady=10,
            spacing1=2,
            spacing3=2,
            highlightthickness=1,
            highlightbackground="#d6dee9",
            highlightcolor=ACCENT,
        )
        yscroll = ttk.Scrollbar(editor_wrap, orient=tk.VERTICAL, command=self.editor.yview)
        self.editor.configure(yscrollcommand=yscroll.set)
        self.editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._placeholder = t(
            "placeholder",
            accel=ACCEL,
            paste=clipboard.paste_shortcut(),
        )
        self._placeholder_active = True
        self.editor.insert("1.0", self._placeholder)
        self.editor.configure(fg="#94a3b8")
        self.editor.bind("<FocusIn>", self._clear_placeholder)
        self.editor.bind("<<Paste>>", self._on_paste)
        self._bind_keys(self.editor, "v", self._on_paste)
        self.editor.bind("<KeyRelease>", lambda e: self._update_count())

        tips = ttk.Label(
            root,
            text=t("shortcuts", accel=ACCEL),
            style="Sub.TLabel",
        )
        tips.pack(fill=tk.X, pady=(8, 6))

        status = tk.Frame(root, bg=STATUS_BG, highlightthickness=1, highlightbackground=BORDER)
        status.pack(fill=tk.X)
        self._status_icon = tk.Label(
            status,
            textvariable=self._status_icon_var,
            bg=STATUS_BG,
            fg=TEXT_MUTED,
            font=(UI_FONT, 10, "bold"),
        )
        self._status_icon.pack(side=tk.LEFT, padx=(12, 6), pady=7)
        self._status_label = tk.Label(
            status,
            textvariable=self._status_var,
            bg=STATUS_BG,
            fg="#1e293b",
            font=(UI_FONT, 9),
            anchor="w",
            justify=tk.LEFT,
        )
        self._status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12), pady=7)

    def _lang_bar(self, parent) -> None:
        bar = tk.Frame(parent, bg=HEADER_BG)
        bar.pack(side=tk.RIGHT, anchor="e")
        current = ui_lang()
        for code, label in (("en", "EN"), ("zh", "中文")):
            active = code == current
            lbl = tk.Label(
                bar,
                text=label,
                bg=HEADER_BG,
                fg=HEADER_FG if active else HEADER_MUTED,
                font=(UI_FONT, 9, "bold" if active else "normal"),
                cursor="hand2",
            )
            lbl.pack(side=tk.LEFT, padx=(8, 0))
            lbl.bind("<Button-1>", lambda _e, c=code: self._switch_lang(c))

    def _switch_lang(self, code: str) -> None:
        if code == ui_lang() or self._busy:
            return
        md = None if self._placeholder_active else self.get_markdown()
        watch = self._watch_var.get()
        auto = self._auto_convert_var.get()
        if self._watch_job:
            try:
                self.after_cancel(self._watch_job)
            except Exception:  # noqa: BLE001
                pass
            self._watch_job = None
        set_lang(code)
        for child in self.winfo_children():
            child.destroy()
        self._action_btns.clear()
        self._toggles.clear()
        self._watch_var = tk.BooleanVar(value=watch)
        self._auto_convert_var = tk.BooleanVar(value=auto)
        self._status_var = tk.StringVar(value=t("ready"))
        self._status_icon_var = tk.StringVar(value="·")
        self._count_var = tk.StringVar(value=t("char_count", n=0))
        self.title(t("app_title"))
        self._build_style()
        self._build_ui()
        if md:
            self.set_markdown(md)
        if watch:
            self._poll_clipboard()

    def _btn(self, parent, text: str, command) -> ttk.Button:
        btn = ttk.Button(parent, text=text, command=command)
        self._action_btns.append(btn)
        return btn

    def _run_bg(self, work: Callable, done: Callable) -> None:
        def worker() -> None:
            result = None
            err = None
            try:
                result = work()
            except Exception as exc:  # noqa: BLE001
                err = exc
            self.after(0, lambda r=result, e=err: done(r, e))

        threading.Thread(target=worker, daemon=True).start()

    def _bind_keys(self, widget, key: str, handler) -> None:
        widget.bind(f"<Control-{key}>", handler)
        widget.bind(f"<Control-{key.upper()}>", handler)
        if IS_MAC:
            widget.bind(f"<Command-{key}>", handler)
            widget.bind(f"<Command-{key.upper()}>", handler)

    def _bind_shortcuts(self) -> None:
        self.bind("<Control-Return>", lambda e: self.convert_to_clipboard())
        self._bind_keys(self, "s", lambda e: self.save_docx())
        self._bind_keys(self, "p", lambda e: self.save_pdf())
        self.bind("<Control-Shift-V>", lambda e: self.import_clipboard())
        self.bind("<Control-Shift-v>", lambda e: self.import_clipboard())
        if IS_MAC:
            self.bind("<Command-Return>", lambda e: self.convert_to_clipboard())
            self.bind("<Command-Shift-V>", lambda e: self.import_clipboard())
            self.bind("<Command-Shift-v>", lambda e: self.import_clipboard())
        self.bind("<Escape>", lambda e: self.clear_editor())

    def _set_status(self, msg: str, level: str = "info") -> None:
        icon, color = {
            "ok": (CHECK, OK_FG),
            "err": ("!", ERR_FG),
            "busy": ("…", ACCENT),
        }.get(level, ("·", TEXT_MUTED))
        self._status_icon_var.set(icon)
        self._status_icon.configure(fg=color)
        self._status_label.configure(fg=ERR_FG if level == "err" else "#1e293b")
        self._status_var.set(msg)
        self.update_idletasks()

    def _set_busy(self, busy: bool, msg: Optional[str] = None) -> None:
        self._busy = busy
        self.config(cursor="watch" if busy else "")
        state = tk.DISABLED if busy else tk.NORMAL
        for btn in self._action_btns:
            btn.configure(state=state)
        if msg:
            self._set_status(msg, "busy" if busy else "info")
        self.update_idletasks()

    def _update_count(self) -> None:
        if self._placeholder_active:
            self._count_var.set(t("char_count", n=0))
            return
        text = self.editor.get("1.0", "end-1c")
        self._count_var.set(t("char_count", n=len(text.strip())))

    def _clear_placeholder(self, _event=None) -> None:
        if self._placeholder_active:
            self.editor.delete("1.0", tk.END)
            self.editor.configure(fg="#1e293b")
            self._placeholder_active = False
            self._update_count()

    def _on_paste(self, _event=None):
        self.after(10, self._clear_placeholder)
        self.after(20, self._update_count)

    def get_markdown(self) -> str:
        if self._placeholder_active:
            return ""
        return self.editor.get("1.0", "end-1c")

    def set_markdown(self, text: str) -> None:
        self._placeholder_active = False
        self.editor.configure(fg="#1e293b")
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", text)
        self.editor.edit_reset()
        self.editor.mark_set(tk.INSERT, "1.0")
        self.editor.see("1.0")
        self._update_count()

    def clear_editor(self) -> None:
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", self._placeholder)
        self.editor.configure(fg="#94a3b8")
        self._placeholder_active = True
        self._update_count()
        self._set_status(t("cleared"))

    def import_clipboard(self) -> None:
        try:
            text = clipboard.get_text()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(t("error"), t("clip_read_fail", exc=exc))
            return
        if not text.strip():
            self._set_status(t("clip_empty"), "err")
            return
        self.set_markdown(text)
        kind = t("kind_md") if converter.looks_like_markdown(text) else t("kind_text")
        self._set_status(t("imported", n=len(text), kind=kind), "ok")

    def open_file(self) -> None:
        path = filedialog.askopenfilename(
            title=t("open_md"),
            filetypes=[
                ("Markdown", "*.md *.markdown *.mdx *.txt" if not sys.platform == "win32" else "*.md;*.markdown;*.mdx;*.txt"),
                (t("all_files"), "*.*"),
            ],
        )
        if not path:
            return
        try:
            content = Path(path).read_text(encoding="utf-8")
        except UnicodeDecodeError:
            content = Path(path).read_text(encoding="gbk", errors="replace")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(t("error"), t("read_fail", exc=exc))
            return
        self.set_markdown(content)
        self._set_status(t("opened", path=path), "ok")

    def _ensure_content(self) -> Optional[str]:
        md = self.get_markdown()
        if not md.strip():
            messagebox.showinfo(t("hint"), t("need_content"))
            return None
        return md

    def convert_to_clipboard(self) -> None:
        md = self._ensure_content()
        if md is None or self._busy:
            return
        self._set_busy(True, t("converting"))
        self._run_bg(lambda: converter.copy_markdown_for_word(md), self._after_convert_clipboard)

    def _after_convert_clipboard(self, res, err) -> None:
        self._set_busy(False)
        if err is not None:
            messagebox.showerror(t("error"), t("clip_write_fail", err=err))
            self._set_status(t("failed", err=err), "err")
            return
        if not res.success:
            messagebox.showerror(t("convert_fail"), res.message)
            self._set_status(res.message, "err")
            return
        self._set_status(res.message, "ok")

    def save_docx(self) -> None:
        md = self._ensure_content()
        if md is None or self._busy:
            return
        DEFAULT_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        initial = DEFAULT_EXPORT_DIR / f"md2word_{stamp}.docx"
        path = filedialog.asksaveasfilename(
            title=t("save_docx_title"),
            defaultextension=".docx",
            initialfile=initial.name,
            initialdir=str(DEFAULT_EXPORT_DIR),
            filetypes=[(t("word_docs"), "*.docx")],
        )
        if not path:
            return
        self._set_busy(True, t("saving_docx"))
        out = Path(path)
        self._run_bg(lambda: converter.markdown_to_docx(md, out), lambda res, err: self._after_save_docx(res, err, out))

    def _after_save_docx(self, res, err, path: Path) -> None:
        self._set_busy(False)
        if err is not None:
            messagebox.showerror(t("save_fail"), str(err))
            self._set_status(str(err), "err")
            return
        if not res.success:
            messagebox.showerror(t("save_fail"), res.message)
            self._set_status(res.message, "err")
            return
        self._last_docx = path
        self._set_status(res.message, "ok")
        if messagebox.askyesno(t("save_ok"), t("open_in_word", msg=res.message)):
            try:
                converter.open_file(path)
            except Exception as exc:  # noqa: BLE001
                messagebox.showwarning(t("hint"), t("saved_no_open", exc=exc))

    def save_pdf(self) -> None:
        md = self._ensure_content()
        if md is None or self._busy:
            return
        DEFAULT_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        initial = DEFAULT_EXPORT_DIR / f"md2word_{stamp}.pdf"
        path = filedialog.asksaveasfilename(
            title=t("export_pdf_title"),
            defaultextension=".pdf",
            initialfile=initial.name,
            initialdir=str(DEFAULT_EXPORT_DIR),
            filetypes=[(t("pdf_docs"), "*.pdf")],
        )
        if not path:
            return
        self._set_busy(True, t("exporting_pdf"))
        out = Path(path)
        self._run_bg(lambda: converter.markdown_to_pdf(md, out), lambda res, err: self._after_save_pdf(res, err, out))

    def _after_save_pdf(self, res, err, path: Path) -> None:
        self._set_busy(False)
        if err is not None:
            messagebox.showerror(t("export_fail"), str(err))
            self._set_status(str(err), "err")
            return
        if not res.success:
            messagebox.showerror(t("export_fail"), res.message)
            self._set_status(res.message, "err")
            return
        self._set_status(res.message, "ok")
        if messagebox.askyesno(t("export_ok"), t("open_now", msg=res.message)):
            try:
                converter.open_file(path)
            except Exception as exc:  # noqa: BLE001
                messagebox.showwarning(t("hint"), t("saved_no_open", exc=exc))

    def convert_and_open(self) -> None:
        md = self._ensure_content()
        if md is None or self._busy:
            return
        DEFAULT_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = DEFAULT_EXPORT_DIR / f"md2word_{stamp}.docx"
        self._set_busy(True, t("saving_docx"))
        self._run_bg(lambda: converter.markdown_to_docx(md, out), lambda res, err: self._after_convert_open(res, err, out))

    def _after_convert_open(self, res, err, out: Path) -> None:
        self._set_busy(False)
        if err is not None:
            messagebox.showerror(t("convert_fail"), str(err))
            self._set_status(str(err), "err")
            return
        if not res.success or not res.output_path:
            messagebox.showerror(t("convert_fail"), res.message)
            self._set_status(res.message, "err")
            return
        self._last_docx = res.output_path
        self._set_status(t("opening", msg=res.message), "ok")
        try:
            converter.open_file(res.output_path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showwarning(t("hint"), t("made_no_open", path=res.output_path, exc=exc))

    def _toggle_watch(self) -> None:
        for toggle in self._toggles:
            toggle.refresh()
        if self._watch_var.get():
            self._last_clip_hash = None
            self._set_status(t("watch_on"), "ok")
            self._poll_clipboard()
        else:
            if self._watch_job:
                try:
                    self.after_cancel(self._watch_job)
                except Exception:  # noqa: BLE001
                    pass
                self._watch_job = None
            self._set_status(t("watch_off"))

    def _poll_clipboard(self) -> None:
        if not self._watch_var.get():
            return
        try:
            text = clipboard.get_text()
            h = hash(text) if text else None
            if text and h != self._last_clip_hash:
                if converter.looks_like_markdown(text):
                    self._last_clip_hash = h
                    self.set_markdown(text)
                    self._set_status(t("watch_imported", n=len(text)), "ok")
                    if self._auto_convert_var.get() and not self._busy:
                        self.after(50, self._auto_convert_from_watch)
                else:
                    self._last_clip_hash = h
        except Exception:
            pass
        self._watch_job = self.after(800, self._poll_clipboard)

    def _auto_convert_from_watch(self) -> None:
        md = self.get_markdown()
        if not md.strip() or self._busy:
            return
        self._set_busy(True, t("watch_converting"))
        self._run_bg(lambda: converter.copy_markdown_for_word(md), self._after_auto_convert)

    def _after_auto_convert(self, res, err) -> None:
        self._set_busy(False)
        if err is not None:
            self._set_status(t("auto_fail", msg=err), "err")
            return
        if not res.success:
            self._set_status(t("auto_fail", msg=res.message), "err")
            return
        try:
            self._last_clip_hash = hash(clipboard.get_text())
        except Exception:  # noqa: BLE001
            pass
        self._set_status(res.message, "ok")

    def _on_close(self) -> None:
        self._watch_var.set(False)
        if self._watch_job:
            try:
                self.after_cancel(self._watch_job)
            except Exception:  # noqa: BLE001
                pass
        self.destroy()


def run() -> None:
    app = Md2WordApp()
    app.mainloop()
