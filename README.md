# Puralox-XDI Standalone Patch für Ubuntu + externe eLabFTW-API

Ziel: Nur `puralox-xdi` wird auf der Ubuntu-Maschine per Docker gestartet. eLabFTW bleibt extern unter:

```text
https://dtpa-akg.de/
```

Die bestehende XDI-Extraktion und bestehende Datenübermittlung zu eLabFTW werden nicht verändert.

## Enthaltene Dateien

```text
compose.puralox-xdi.yml       # neues Docker-Compose-File nur für Puralox-XDI
.env.example                  # Env-Vorlage für externe eLabFTW-API
apply_puralox_xdi_patch.sh    # passt app.py/config.py minimal an
puralox/config.py             # neue config.py, falls du manuell kopieren willst
```

## Anwendung

1. Repository vorbereiten:

```bash
git clone https://github.com/vitalybin/BET-Automation.git
cd BET-Automation
git checkout elabftw-xdi
```

2. ZIP in das Repository-Root entpacken oder die Dateien aus dem ZIP dorthin kopieren.

Danach sollte im Repository-Root unter anderem vorhanden sein:

```text
compose.puralox-xdi.yml
.env.example
apply_puralox_xdi_patch.sh
puralox/config.py
```

3. Patch anwenden:

```bash
bash apply_puralox_xdi_patch.sh
```

Das Script erzeugt Backups:

```text
puralox/app.py.bak-YYYYMMDD-HHMMSS
puralox/config.py.bak-YYYYMMDD-HHMMSS
```

4. `.env` bearbeiten:

```bash
nano .env
```

Eintragen:

```env
ELABFTW_URL=https://dtpa-akg.de/api/v2
ELABFTW_TOKEN=<hier-den-XdiExtractor-Token-einfügen>
ELABFTW_DISABLE_SSL=false
```

Wichtig: `XdiExtractor` ist vermutlich der Name des API-Keys in eLabFTW. In `ELABFTW_TOKEN` muss der tatsächliche generierte Token-Wert stehen.

5. Container bauen und starten:

```bash
docker compose -f compose.puralox-xdi.yml up -d --build
```

6. Logs prüfen:

```bash
docker logs -f puralox-xdi
```

7. Puralox-XDI öffnen:

```text
http://<ubuntu-server-ip>:5000
```



## Anwendung unter Windows PowerShell / Docker Desktop

Bash-Skripte laufen in PowerShell nicht nativ. Für Windows ist deshalb zusätzlich dieses Skript enthalten:

```text
apply_puralox_xdi_patch.ps1
```

Empfohlener Ablauf in PowerShell:

```powershell
git clone https://github.com/vitalybin/BET-Automation.git
cd BET-Automation
git checkout elabftw-xdi
git checkout -b elabftw-xdi-puralox-standalone
```

Patch-ZIP in einen temporären Ordner entpacken, z. B.:

```powershell
$zip = "$env:USERPROFILE\Downloads\puralox-xdi-standalone-patch.zip"
$target = "$env:TEMP\puralox-xdi-standalone-patch"
Remove-Item $target -Recurse -Force -ErrorAction SilentlyContinue
Expand-Archive -Path $zip -DestinationPath $target -Force
$patchRoot = "$target\puralox-xdi-standalone-patch"
```

Dann das PowerShell-Patch-Skript ausführen und das Repository als Ziel übergeben:

```powershell
powershell -ExecutionPolicy Bypass -File "$patchRoot\apply_puralox_xdi_patch.ps1" -RepoPath (Get-Location).Path
```

Danach `.env` bearbeiten:

```powershell
notepad .env
```

Docker Desktop lokal testen:

```powershell
docker compose -f compose.puralox-xdi.yml config
docker compose -f compose.puralox-xdi.yml build
docker compose -f compose.puralox-xdi.yml up -d
```

Oder mit dem enthaltenen Testskript:

```powershell
powershell -ExecutionPolicy Bypass -File "$patchRoot\test_docker_desktop.ps1" -Start
```

Wenn `test_docker_desktop.ps1` aus dem entpackten Patch-Ordner gestartet wird, vorher ins Repository wechseln:

```powershell
cd C:\Pfad\zu\BET-Automation
powershell -ExecutionPolicy Bypass -File "$patchRoot\test_docker_desktop.ps1" -Start
```

Öffnen:

```text
http://localhost:5000
```

Stoppen:

```powershell
docker compose -f compose.puralox-xdi.yml down
```


## Was geändert wird

In `puralox/app.py`:

```python
ELABFTW_URL = os.getenv("ELABFTW_URL", "https://dtpa-akg.de/api/v2")
ELABFTW_TOKEN = os.getenv("ELABFTW_TOKEN", "")
disable_ssl = os.getenv("ELABFTW_DISABLE_SSL", "false").lower() == "true"
```

In `puralox/config.py`:

```python
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "/usr/src/app/uploads")
DB_NAME = os.getenv("DB_NAME", "/usr/src/app/data/Purlox.db")
ELABFTW_URL = os.getenv("ELABFTW_URL", "https://dtpa-akg.de/api/v2")
ELABFTW_TOKEN = os.getenv("ELABFTW_TOKEN", "")
```

## Was nicht geändert wird

- keine Änderung an `xdi_processor.py`
- keine Änderung an der XDI-Dateiextraktion
- keine Änderung an der bestehenden ELN-Datenübermittlung
- kein lokales eLabFTW
- kein MySQL-Container
- keine Raspberry-Pi-spezifische Konfiguration
