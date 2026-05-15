#!/usr/bin/env bash
set -euo pipefail

# Startskript für Puralox-XDI standalone
# Nutzung:
#   bash start_puralox_xdi.sh
# Optional:
#   NO_CACHE=1 bash start_puralox_xdi.sh

cd "$(dirname "$0")"

COMPOSE_FILE="compose.puralox-xdi.yml"
ENV_EXAMPLE=".env.example"
ENV_FILE=".env"

if [[ ! -f "$COMPOSE_FILE" ]]; then
  echo "FEHLER: $COMPOSE_FILE nicht gefunden. Bitte Skript im Repo-Root ausführen."
  exit 1
fi

if ! command -v docker >/dev/null 2>&1; then
  echo "FEHLER: docker wurde nicht gefunden."
  exit 1
fi

if ! docker info >/dev/null 2>&1; then
  echo "FEHLER: Docker läuft nicht oder ist nicht erreichbar."
  echo "Starte Docker Desktop bzw. den Docker-Daemon und versuche es erneut."
  exit 1
fi

if [[ ! -f "$ENV_FILE" ]]; then
  if [[ -f "$ENV_EXAMPLE" ]]; then
    cp "$ENV_EXAMPLE" "$ENV_FILE"
    echo ".env wurde aus .env.example erstellt."
  else
    cat > "$ENV_FILE" <<'EOF'
ELABFTW_URL=https://dtpa-akg.de/api/v2
ELABFTW_TOKEN=
ELABFTW_DISABLE_SSL=false

UPLOAD_FOLDER=/usr/src/app/uploads
DB_NAME=/usr/src/app/data/Purlox.db
EOF
    echo ".env wurde neu erstellt."
  fi
else
  echo ".env existiert bereits und wird aktualisiert."
fi

set_env_value() {
  local key="$1"
  local value="$2"
  local tmp="${ENV_FILE}.tmp"

  if grep -qE "^[[:space:]]*${key}[[:space:]]*=" "$ENV_FILE"; then
    awk -v key="$key" -v value="$value" '
      BEGIN { done=0 }
      $0 ~ "^[[:space:]]*" key "[[:space:]]*=" {
        print key "=" value
        done=1
        next
      }
      { print }
      END {
        if (done == 0) print key "=" value
      }
    ' "$ENV_FILE" > "$tmp"
    mv "$tmp" "$ENV_FILE"
  else
    printf "\n%s=%s\n" "$key" "$value" >> "$ENV_FILE"
  fi
}

echo
read -r -s -p "ELABFTW_TOKEN / XdiExtractor-Key eingeben: " ELABFTW_TOKEN
echo

if [[ -z "${ELABFTW_TOKEN// }" ]]; then
  echo "FEHLER: Token darf nicht leer sein."
  exit 1
fi

set_env_value "ELABFTW_URL" "https://dtpa-akg.de/api/v2"
set_env_value "ELABFTW_TOKEN" "$ELABFTW_TOKEN"
set_env_value "ELABFTW_DISABLE_SSL" "false"
set_env_value "UPLOAD_FOLDER" "/usr/src/app/uploads"
set_env_value "DB_NAME" "/usr/src/app/data/Purlox.db"

mkdir -p data uploads metadata

if [[ ! -f "data/Purlox.db" && -f "Purlox.db" ]]; then
  cp "Purlox.db" "data/Purlox.db"
  echo "data/Purlox.db wurde aus Purlox.db erstellt."
fi

echo
echo "Docker Compose Konfiguration wird geprüft..."
docker compose -f "$COMPOSE_FILE" config >/dev/null

echo
if [[ "${NO_CACHE:-0}" == "1" ]]; then
  echo "Baue Image ohne Cache..."
  docker compose -f "$COMPOSE_FILE" build --no-cache
  echo "Starte Container..."
  docker compose -f "$COMPOSE_FILE" up -d
else
  echo "Baue und starte Container..."
  docker compose -f "$COMPOSE_FILE" up -d --build
fi

echo
docker ps --filter "name=puralox-xdi"

echo
echo "Fertig."
echo "Puralox-XDI: http://localhost:5000"
echo "Logs: docker logs -f puralox-xdi"
