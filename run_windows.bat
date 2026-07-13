@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python 3 was not found. Install Python 3.11 or 3.12 from python.org first.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating the local environment...
  py -3.12 -m venv .venv 2>nul || py -3.11 -m venv .venv
)

call ".venv\Scripts\activate.bat"
python -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
  echo Installation failed. Check your internet connection and try again.
  pause
  exit /b 1
)
start "" ".venv\Scripts\pythonw.exe" app.py

