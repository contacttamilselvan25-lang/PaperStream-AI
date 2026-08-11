@echo off
cd /d "%~dp0"
echo Installing required parts (only takes time the first run)...
pip install -r requirements.txt
echo.
echo Starting Paper Extractor...
python app.py
if errorlevel 1 (
  echo.
  echo Something went wrong. Please send a screenshot of this window to get help.
  pause
)