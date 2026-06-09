@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv" (
  py -3 -m venv .venv
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m playwright install chromium

if not exist "config.json" copy "config.example.json" "config.json"

echo.
echo Setup complete.
echo Put Gmail OAuth desktop credentials at: data\gmail_credentials.json
pause
