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
  echo "Python 3.11 or 3.12 is required."
  exit 1
fi

if [ ! -x ".build-venv/bin/python" ]; then
  "$PYTHON" -m venv .build-venv
fi
".build-venv/bin/python" -m pip install --disable-pip-version-check -r requirements-build.txt
".build-venv/bin/pyinstaller" --noconfirm --clean --windowed --name "PDF-to-Word" \
  --collect-all pdf2docx --collect-all pytesseract app.py
echo "Build complete: dist/PDF-to-Word.app"
