# Contributing

**English** | [简体中文](CONTRIBUTING.zh-CN.md)

Thanks for your interest in md2word. This is a short guide for local development.

## Setup

- Windows, macOS, or Linux
- Python 3.10 or newer
- Optional: [Pandoc](https://pandoc.org/installing.html), Microsoft Word, or LibreOffice (needed for formula paste and PDF export)

```bash
python -m pip install -r requirements-dev.txt
python main.py
```

On Windows you can also double-click `start.bat`. On macOS / Linux use `./start.sh`.

## Tests

```bash
python -m unittest tests.test_md2word -v
```

Pandoc and Word cases are skipped when those tools are not installed. Please run the suite before opening a pull request.

## Windows exe

```bat
python build_exe.py
```

The binary lands in `dist\md2word.exe`. Build logs, `build/`, `dist/`, and exe files are gitignored — do not commit them.

## Conventions

- Package code lives in `md2word/`; the entry point is `main.py`
- Sample Markdown lives in `examples/`
- Do not commit machine-specific paths, build logs, or `.docx` / `.pdf` outputs
- When you change conversion logic, add coverage in `tests/test_md2word.py`
- Windows-only COM code stays in `clipboard_win.py` and the COM branch of `wordcom.py`; other platforms use `clipboard.py` plus AppleScript / LibreOffice

## Pull requests

1. Fork and create a branch
2. Keep the diff as small as the change needs
3. Run tests
4. Open a PR that explains why, what changed, and how you verified it
