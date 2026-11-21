#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv
  . .venv/bin/activate
  pip install -r requirements.txt
  python -m playwright install chromium
else
  . .venv/bin/activate
fi
python scraper.py
python cleaner.py
echo "Done. Files are in the output folder."
