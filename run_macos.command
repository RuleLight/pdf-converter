#!/bin/bash
set -e
cd "$(dirname "$0")"

PYTHON=""
for CANDIDATE in python3.12 python3.11 python3; do
  if command -v "$CANDIDATE" >/dev/null 2>&1 && "$CANDIDATE" -c 'import sys; raise SystemExit(sys.version_info[:2] not in [(3,11),(3,12)])'; then
    PYTHON="$CANDIDATE"
    break
  fi
done
if [ -z "$PYTHON" ]; then
  osascript -e 'display alert "PDF 转 Word" message "请先从 python.org 安装 Python 3.11 或 3.12。"'
  exit 1
fi

if [ ! -x ".venv/bin/python" ]; then
  "$PYTHON" -m venv .venv
fi

".venv/bin/python" -m pip install --disable-pip-version-check -r requirements.txt
".venv/bin/python" app.py
