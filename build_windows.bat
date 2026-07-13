@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python 3 was not found. Install Python 3.11 or 3.12 first.
  pause
  exit /b 1
)

if not exist ".build-venv\Scripts\python.exe" (
  py -3.12 -m venv .build-venv 2>nul || py -3.11 -m venv .build-venv
)
call ".build-venv\Scripts\activate.bat"
python -m pip install --disable-pip-version-check -r requirements-build.txt
if errorlevel 1 exit /b 1

pyinstaller --noconfirm --clean --windowed --name "PDF-to-Word" --collect-all pdf2docx --collect-all pytesseract app.py
echo.
echo Build complete: dist\PDF-to-Word\PDF-to-Word.exe
pause

