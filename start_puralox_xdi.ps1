param(
    [switch]$NoCache,
    [switch]$DownFirst
)

$ErrorActionPreference = "Stop"

# Startskript für Puralox-XDI standalone unter PowerShell / Docker Desktop
# Nutzung:
#   powershell -ExecutionPolicy Bypass -File .\start_puralox_xdi.ps1
# Optional:
#   powershell -ExecutionPolicy Bypass -File .\start_puralox_xdi.ps1 -NoCache
#   powershell -ExecutionPolicy Bypass -File .\start_puralox_xdi.ps1 -DownFirst

Set-Location $PSScriptRoot

$ComposeFile = "compose.puralox-xdi.yml"
$EnvExample = ".env.example"
$EnvFile = ".env"

if (-not (Test-Path $ComposeFile)) {
    throw "$ComposeFile nicht gefunden. Bitte Skript im Repo-Root ausführen."
}

if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    throw "docker wurde nicht gefunden. Bitte Docker Desktop installieren/starten."
}

try {
    docker info | Out-Null
} catch {
    throw "Docker läuft nicht oder ist nicht erreichbar. Starte Docker Desktop und versuche es erneut."
}

if (-not (Test-Path $EnvFile)) {
    if (Test-Path $EnvExample) {
        Copy-Item $EnvExample $EnvFile -Force
        Write-Host ".env wurde aus .env.example erstellt."
    } else {
        @"
ELABFTW_URL=https://dtpa-akg.de/api/v2
ELABFTW_TOKEN=
ELABFTW_DISABLE_SSL=false

UPLOAD_FOLDER=/usr/src/app/uploads
DB_NAME=/usr/src/app/data/Purlox.db
"@ | Set-Content $EnvFile -Encoding UTF8
        Write-Host ".env wurde neu erstellt."
    }
} else {
    Write-Host ".env existiert bereits und wird aktualisiert."
}

function Set-EnvValue {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Key,
        [Parameter(Mandatory=$true)][string]$Value
    )

    $lines = @()
    if (Test-Path $Path) {
        $lines = Get-Content $Path
    }

    $pattern = "^\s*$([regex]::Escape($Key))\s*="
    $found = $false
    $newLines = foreach ($line in $lines) {
        if ($line -match $pattern) {
            "$Key=$Value"
            $found = $true
        } else {
            $line
        }
    }

    if (-not $found) {
        $newLines += "$Key=$Value"
    }

    Set-Content -Path $Path -Value $newLines -Encoding UTF8
}

Write-Host ""
$secureToken = Read-Host "ELABFTW_TOKEN / XdiExtractor-Key eingeben" -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken)
try {
    $token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
}

if ([string]::IsNullOrWhiteSpace($token)) {
    throw "Token darf nicht leer sein."
}

Set-EnvValue -Path $EnvFile -Key "ELABFTW_URL" -Value "https://dtpa-akg.de/api/v2"
Set-EnvValue -Path $EnvFile -Key "ELABFTW_TOKEN" -Value $token
Set-EnvValue -Path $EnvFile -Key "ELABFTW_DISABLE_SSL" -Value "false"
Set-EnvValue -Path $EnvFile -Key "UPLOAD_FOLDER" -Value "/usr/src/app/uploads"
Set-EnvValue -Path $EnvFile -Key "DB_NAME" -Value "/usr/src/app/data/Purlox.db"

New-Item -ItemType Directory -Force ".\data" | Out-Null
New-Item -ItemType Directory -Force ".\uploads" | Out-Null
New-Item -ItemType Directory -Force ".\metadata" | Out-Null

if ((-not (Test-Path ".\data\Purlox.db")) -and (Test-Path ".\Purlox.db")) {
    Copy-Item ".\Purlox.db" ".\data\Purlox.db" -Force
    Write-Host "data\Purlox.db wurde aus Purlox.db erstellt."
}

Write-Host ""
Write-Host "Docker Compose Konfiguration wird geprüft..."
docker compose -f $ComposeFile config | Out-Null

if ($DownFirst) {
    Write-Host ""
    Write-Host "Stoppe vorhandene Container..."
    docker compose -f $ComposeFile down
}

Write-Host ""
if ($NoCache) {
    Write-Host "Baue Image ohne Cache..."
    docker compose -f $ComposeFile build --no-cache
    Write-Host "Starte Container..."
    docker compose -f $ComposeFile up -d
} else {
    Write-Host "Baue und starte Container..."
    docker compose -f $ComposeFile up -d --build
}

Write-Host ""
docker ps --filter "name=puralox-xdi"

Write-Host ""
Write-Host "Fertig."
Write-Host "Puralox-XDI: http://localhost:5000"
Write-Host "Logs: docker logs -f puralox-xdi"
