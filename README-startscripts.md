# Puralox-XDI Startskripte

Diese Skripte gehören in den Root-Ordner des Projekts `BET-Automation`.

## PowerShell / Docker Desktop

```powershell
cd C:\Users\Vitaly\Documents\BET-Automation
powershell -ExecutionPolicy Bypass -File .\start_puralox_xdi.ps1
```

Optional ohne Docker-Cache:

```powershell
powershell -ExecutionPolicy Bypass -File .\start_puralox_xdi.ps1 -NoCache
```

Optional vorher Container stoppen:

```powershell
powershell -ExecutionPolicy Bypass -File .\start_puralox_xdi.ps1 -DownFirst
```

## Bash / Ubuntu / WSL / Git Bash

```bash
cd /pfad/zu/BET-Automation
bash start_puralox_xdi.sh
```

Optional ohne Docker-Cache:

```bash
NO_CACHE=1 bash start_puralox_xdi.sh
```

## Was die Skripte machen

1. Prüfen, ob Docker erreichbar ist.
2. `.env` aus `.env.example` erzeugen, falls `.env` noch fehlt.
3. `ELABFTW_TOKEN` im Terminal abfragen.
4. `.env` aktualisieren:
   - `ELABFTW_URL=https://dtpa-akg.de/api/v2`
   - `ELABFTW_TOKEN=<Eingabe>`
   - `ELABFTW_DISABLE_SSL=false`
   - `UPLOAD_FOLDER=/usr/src/app/uploads`
   - `DB_NAME=/usr/src/app/data/Purlox.db`
5. Lokale Ordner `data`, `uploads`, `metadata` anlegen.
6. Falls nötig `Purlox.db` nach `data/Purlox.db` kopieren.
7. `docker compose -f compose.puralox-xdi.yml up -d --build` ausführen.

Danach öffnen:

```text
http://localhost:5000
```

Logs:

```powershell
docker logs -f puralox-xdi
```
