#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "[错误] 未找到 Python 3，请先安装 Python 3.10+"
  exit 1
fi

if ! python3 -c "import docx" >/dev/null 2>&1; then
  echo "[提示] 正在安装依赖..."
  python3 -m pip install -r requirements.txt
fi

exec python3 main.py
