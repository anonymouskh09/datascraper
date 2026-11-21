@echo off
cd /d %~dp0
if not exist .venv (
  python -m venv .venv
  call .venv\Scripts\activate
  pip install -r requirements.txt
  python -m playwright install chromium
) else (
  call .venv\Scripts\activate
)
python scraper.py
python cleaner.py
echo.
echo Done. Files are in the "output" folder.
pause
