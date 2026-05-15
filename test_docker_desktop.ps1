param(
    [switch]$Start
)

$ErrorActionPreference = "Stop"

function Write-Ok($msg) { Write-Host "OK: $msg" -ForegroundColor Green }
function Write-Info($msg) { Write-Host "INFO: $msg" -ForegroundColor Cyan }
function Write-Fail($msg) { Write-Host "ERROR: $msg" -ForegroundColor Red }

if (!(Test-Path "compose.puralox-xdi.yml")) {
    Write-Fail "compose.puralox-xdi.yml nicht gefunden. Bitte im Root-Verzeichnis des BET-Automation-Repositories ausführen."
    exit 1
}

if (!(Get-Command docker -ErrorAction SilentlyContinue)) {
    Write-Fail "Docker wurde nicht gefunden. Bitte Docker Desktop starten/installieren und PowerShell neu öffnen."
    exit 1
}

Write-Info "Prüfe Docker Desktop..."
docker info | Out-Null
Write-Ok "Docker ist erreichbar."

if (!(Test-Path ".env") -and (Test-Path ".env.example")) {
    Copy-Item ".env.example" ".env" -Force
    Write-Ok ".env aus .env.example erzeugt. Für ELN-Test bitte ELABFTW_TOKEN eintragen."
}

Write-Info "Prüfe Compose-Konfiguration..."
docker compose -f compose.puralox-xdi.yml config | Out-Null
Write-Ok "Compose-Konfiguration ist gültig."

Write-Info "Baue Image..."
docker compose -f compose.puralox-xdi.yml build
Write-Ok "Build abgeschlossen."

if ($Start) {
    Write-Info "Starte Container..."
    docker compose -f compose.puralox-xdi.yml up -d
    Write-Ok "Container gestartet. Öffne http://localhost:5000"
    Write-Host "Logs: docker logs -f puralox-xdi"
    Write-Host "Stop: docker compose -f compose.puralox-xdi.yml down"
} else {
    Write-Host ""
    Write-Host "Zum Starten ausführen:"
    Write-Host "  docker compose -f compose.puralox-xdi.yml up -d"
    Write-Host "Oder dieses Skript mit -Start ausführen:"
    Write-Host "  powershell -ExecutionPolicy Bypass -File .\test_docker_desktop.ps1 -Start"
}
