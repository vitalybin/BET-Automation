#!/usr/bin/env bash
set -euo pipefail

if [ ! -f "puralox/app.py" ] || [ ! -f "puralox/config.py" ] || [ ! -f "Dockerfile" ]; then
  echo "ERROR: Bitte dieses Script im Root-Verzeichnis des BET-Automation-Repositories ausführen."
  exit 1
fi

STAMP="$(date +%Y%m%d-%H%M%S)"
cp puralox/app.py "puralox/app.py.bak-${STAMP}"
cp puralox/config.py "puralox/config.py.bak-${STAMP}"

python3 - <<'PY'
from pathlib import Path
import re

app_path = Path("puralox/app.py")
text = app_path.read_text(encoding="utf-8")
original = text

text = re.sub(
    r'ELABFTW_URL\s*=\s*os\.getenv\(\s*["\']ELABFTW_URL["\']\s*,\s*["\']https://localhost/api/v2["\']\s*\)',
    'ELABFTW_URL = os.getenv("ELABFTW_URL", "https://dtpa-akg.de/api/v2")',
    text,
)

text = re.sub(
    r'ELABFTW_TOKEN\s*=\s*os\.getenv\(\s*["\']ELABFTW_TOKEN["\']\s*,\s*["\'][^"\']+["\']\s*\)',
    'ELABFTW_TOKEN = os.getenv("ELABFTW_TOKEN", "")',
    text,
)

text = re.sub(
    r'disable_ssl\s*=\s*os\.getenv\(\s*["\']ELABFTW_DISABLE_SSL["\']\s*,\s*["\']true["\']\s*\)\.lower\(\)\s*==\s*["\']true["\']',
    'disable_ssl = os.getenv("ELABFTW_DISABLE_SSL", "false").lower() == "true"',
    text,
)

if text == original:
    print("WARNUNG: puralox/app.py wurde nicht geändert. Prüfe manuell, ob die Werte bereits angepasst sind.")
else:
    app_path.write_text(text, encoding="utf-8")
    print("OK: puralox/app.py angepasst.")

config_path = Path("puralox/config.py")
config_path.write_text('''import os
from dotenv import load_dotenv

load_dotenv()

# Puralox paths
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "/usr/src/app/uploads")
DB_NAME = os.getenv("DB_NAME", "/usr/src/app/data/Purlox.db")

# External eLabFTW API configuration
ELABFTW_URL = os.getenv("ELABFTW_URL", "https://dtpa-akg.de/api/v2")
ELABFTW_TOKEN = os.getenv("ELABFTW_TOKEN", "")

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Ensure DB directory exists if DB_NAME points to a nested path
_db_dir = os.path.dirname(DB_NAME)
if _db_dir:
    os.makedirs(_db_dir, exist_ok=True)
''', encoding="utf-8")
print("OK: puralox/config.py ersetzt.")
PY

mkdir -p data uploads metadata

if [ ! -f ".env" ]; then
  cp .env.example .env
  echo "OK: .env aus .env.example erzeugt. Bitte ELABFTW_TOKEN eintragen."
else
  echo "INFO: .env existiert bereits und wurde nicht überschrieben."
fi

echo "Fertig. Start danach mit: docker compose -f compose.puralox-xdi.yml up -d --build"
