param(
    [string]$RepoPath = (Get-Location).Path
)

$ErrorActionPreference = "Stop"

function Write-Ok($msg) { Write-Host "OK: $msg" -ForegroundColor Green }
function Write-Info($msg) { Write-Host "INFO: $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "WARNUNG: $msg" -ForegroundColor Yellow }
function Write-Fail($msg) { Write-Host "ERROR: $msg" -ForegroundColor Red }

$RepoPath = (Resolve-Path $RepoPath).Path
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Set-Location $RepoPath

if (!(Test-Path "puralox\app.py") -or !(Test-Path "puralox\config.py") -or !(Test-Path "Dockerfile")) {
    Write-Fail "Bitte dieses Skript im Root-Verzeichnis des BET-Automation-Repositories ausführen oder -RepoPath angeben."
    Write-Host "Beispiel: powershell -ExecutionPolicy Bypass -File .\apply_puralox_xdi_patch.ps1 -RepoPath C:\Pfad\zu\BET-Automation"
    exit 1
}

# Copy deployment files from patch folder to repository root if available.
$filesToCopy = @(
    @{ Source = "compose.puralox-xdi.yml"; Destination = "compose.puralox-xdi.yml" },
    @{ Source = ".env.example"; Destination = ".env.example" }
)

foreach ($file in $filesToCopy) {
    $src = Join-Path $ScriptDir $file.Source
    $dst = Join-Path $RepoPath $file.Destination
    if (Test-Path $src) {
        $srcResolved = (Resolve-Path $src).Path
        $dstFull = [System.IO.Path]::GetFullPath($dst)
        if ($srcResolved -ne $dstFull) {
            Copy-Item $srcResolved $dstFull -Force
            Write-Ok "$($file.Destination) kopiert."
        } else {
            Write-Info "$($file.Destination) ist bereits im Repository-Root vorhanden."
        }
    } else {
        Write-Warn "$($file.Source) wurde neben dem Skript nicht gefunden. Datei wird nicht kopiert."
    }
}

# Backup existing files.
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
Copy-Item "puralox\app.py" "puralox\app.py.bak-$stamp" -Force
Copy-Item "puralox\config.py" "puralox\config.py.bak-$stamp" -Force
Write-Ok "Backups erzeugt: puralox\app.py.bak-$stamp und puralox\config.py.bak-$stamp"

# Patch puralox/app.py without touching XDI extraction or ELN transfer logic.
$appPath = "puralox\app.py"
$app = Get-Content $appPath -Raw
$original = $app

$app = [regex]::Replace(
    $app,
    'ELABFTW_URL\s*=\s*os\.getenv\(\s*["'']ELABFTW_URL["'']\s*,\s*["'']https://localhost/api/v2["'']\s*\)',
    'ELABFTW_URL = os.getenv("ELABFTW_URL", "https://dtpa-akg.de/api/v2")'
)

$app = [regex]::Replace(
    $app,
    'ELABFTW_TOKEN\s*=\s*os\.getenv\(\s*["'']ELABFTW_TOKEN["'']\s*,\s*["''][^"'']+["'']\s*\)',
    'ELABFTW_TOKEN = os.getenv("ELABFTW_TOKEN", "")'
)

$app = [regex]::Replace(
    $app,
    'disable_ssl\s*=\s*os\.getenv\(\s*["'']ELABFTW_DISABLE_SSL["'']\s*,\s*["'']true["'']\s*\)\.lower\(\)\s*==\s*["'']true["'']',
    'disable_ssl = os.getenv("ELABFTW_DISABLE_SSL", "false").lower() == "true"'
)

if ($app -eq $original) {
    Write-Warn "puralox\app.py wurde nicht geändert. Prüfe manuell, ob die Werte bereits angepasst sind."
} else {
    Set-Content -Path $appPath -Value $app -Encoding UTF8 -NoNewline
    Write-Ok "puralox\app.py angepasst."
}

# Replace config.py with standalone-safe config.
$configPath = "puralox\config.py"
$config = @'
import os
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
'@
Set-Content -Path $configPath -Value $config -Encoding UTF8 -NoNewline
Write-Ok "puralox\config.py ersetzt."

# Create local runtime directories.
New-Item -ItemType Directory -Force -Path "data", "uploads", "metadata" | Out-Null
Write-Ok "Runtime-Ordner data, uploads, metadata vorhanden."

# Create .env if missing.
if (!(Test-Path ".env")) {
    if (Test-Path ".env.example") {
        Copy-Item ".env.example" ".env" -Force
        Write-Ok ".env aus .env.example erzeugt. Bitte ELABFTW_TOKEN eintragen."
    } else {
        Write-Warn ".env.example fehlt; .env wurde nicht erzeugt."
    }
} else {
    Write-Info ".env existiert bereits und wurde nicht überschrieben."
}

Write-Host ""
Write-Ok "Patch fertig."
Write-Host "Nächste Schritte:"
Write-Host "  notepad .env"
Write-Host "  docker compose -f compose.puralox-xdi.yml config"
Write-Host "  docker compose -f compose.puralox-xdi.yml up -d --build"
