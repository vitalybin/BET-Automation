# Puralox — BET Data Processor & eLabFTW Integrator

Puralox is a laboratory automation platform built with Flask. Specialized in BET analysis processing, it automates the extraction, storage, visualization, and synchronization of data across two primary instrument outputs: Excel reports and BET PDF reports.

---

## 📋 Table of Contents

1. [Features](#🚀-core-features)
2. [Architecture & Design](#🛠️-architecture--design)
3. [Prerequisites](#🔧-prerequisites)
4. [eLabFTW Installation & Docker-Compose](#🖥️-elabftw-installation--docker-compose)
5. [Installation & Setup](#⚙️-installation--setup)
6. [Project Structure](#📂-project-structure)
7. [Excel Input Format](#📊-excel-input-format)
8. [Running Locally](#▶️-how-to-run)
9. [Web UI Walkthrough](#🌐-web-ui-walkthrough)
10. [API Endpoints](#🔌-api-endpoints)
11. [Configuration Reference](#🔧-configuration-reference)
12. [Database Schema](#🗄️-database-schema)
13. [Customization & Extension](#🔨-customization--extension)
14. [Troubleshooting & FAQs](#❓-troubleshooting--faqs)
15. [Contributing](#🤝-contributing)
16. [License](#📜-license)

---

## 🚀 Core Features

- **Automated BET Data Extraction**
  - **Excel Parsing:** Extracts metadata, BET parameters, and isotherm data points from `.xlsx` files.
  - **PDF Parsing (New):** Extracts comprehensive data from instrument PDF reports using `PyMuPDF`.
- **Scientific Visualization**
  - Generates isotherm, t-plot, and BJH plots using `matplotlib`.
  - Detailed web UI with `DataTables` integration.
- **eLabFTW Synchronization**
  - Full v2 API integration using `ElnClient`.
  - Rich-text summaries, PDF attachments, and auto-tagging.

---

## 🛠️ Architecture & Design

Puralox follows a decoupled, service-oriented architecture with a 1:1 mapping between classes and modules for maximum maintainability.

### Class Diagram (Mermaid)

```mermaid
classDiagram
    direction TB

    class PuraloxApp {
        +Flask app
        +DatabaseManager db
        +ExcelProcessor processor
        +PdfProcessor pdf_processor
        +TemplateProcessor template_processor
        +MetadataBuilder metadata_builder
        +ElnClient eln_client
        +run()
    }

    class DatabaseManager {
        +str db_path
        +fetchall_dict(sql, params) list
        +fetchone_dict(sql, params) dict
        +execute(sql, params) int
        +executemany(sql, seq) int
        +table_exists(name) bool
    }

    class BaseImporter {
        <<abstract>>
        +import_file(path, name)* int
    }

    class ExcelProcessor {
        +DatabaseManager db
        +import_file(path, name) int
    }

    class PdfProcessor {
        +DatabaseManager db
        +import_file(path, name) int
    }

    class BetPdfParser {
        +parse(pdf_source)$ dict
        +parse_general(text)$ dict
    }

    class ElnClient {
        +str base_url
        +create_experiment(title, body) str
        +upload_plot(exp_id, file_id, ...)
    }

    class MeasurementIdBuilder {
        +build(file_id, name, ...)$ str
    }

    %% Inheritance
    ExcelProcessor --|> BaseImporter : extends
    PdfProcessor   --|> BaseImporter : extends

    %% Composition
    PuraloxApp *-- DatabaseManager
    PuraloxApp *-- ExcelProcessor
    PuraloxApp *-- PdfProcessor
    PuraloxApp *-- TemplateProcessor
    PuraloxApp *-- MetadataBuilder
    PuraloxApp *-- ElnClient

    %% Usage
    ExcelProcessor --> DatabaseManager
    PdfProcessor   --> DatabaseManager
    PdfProcessor   ..> BetPdfParser : uses
    ExcelProcessor ..> MeasurementIdBuilder : calls
    PdfProcessor   ..> MeasurementIdBuilder : calls
```

### Component Breakdown

- **`PuraloxApp`**: Main Flask orchestrator and router.
- **`DatabaseManager`**: Centralized SQLite interaction layer.
- **`ElnClient`**: Dedicated wrapper for the eLabFTW API.
- **`BetPdfParser`**: Specialized scientific PDF extraction logic.
- **`ExcelProcessor` & `PdfProcessor`**: Importers inheriting from **`BaseImporter`**.
- **`MeasurementIdBuilder`**: Logic for consistent measurement IDs.
- **`TemplateProcessor`**: Generates rich ELN HTML summaries.
- **`MetadataBuilder`**: Generates metadata Excel sheets.

---

## 🔧 Prerequisites

- **Python 3.10+**
- **PyMuPDF (`fitz`)**, **Pandas**, **Matplotlib**
- **SQLite3**
- **eLabFTW** server (v2 API) & personal access token

---

## 🖥️ eLabFTW Installation & Docker-Compose

Official `docker-compose.yml` snippet for eLabFTW:

```yaml
services:
  web:
    image: elabftw/elabimg:stable
    container_name: elabftw
    restart: always
    environment:
      - SITE_URL=https://localhost
      - DISABLE_HTTPS=false
    ports:
      - '443:443'
    networks:
      - elabftw-net

  mysql:
    image: mysql:8.0
    container_name: mysql
    restart: always
    environment:
      - MYSQL_DATABASE=elabftw
      - MYSQL_USER=elabftw
      - MYSQL_PASSWORD=your_password
    networks:
      - elabftw-net

networks:
  elabftw-net:
```

---

## ⚙️ Installation & Setup

```bash
git clone https://github.com/vitalybin/BET-Automation.git
cd BET-Automation
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

---

## 📂 Project Structure

```text
puralox/
├── app.py                  # PuraloxApp Class
├── base_importer.py        # BaseImporter Class
├── database_manager.py     # DatabaseManager Class
├── excel_processor.py      # ExcelProcessor Class
├── pdf_processor.py        # PdfProcessor Class
├── bet_pdf_parser.py       # BetPdfParser Class
├── eln_client.py           # ElnClient Class
├── measurement_id_builder.py# MeasurementIdBuilder Class
├── metadata_builder.py     # MetadataBuilder Class
├── template_processor.py   # TemplateProcessor Class
├── config.py               # Config & Env
├── static/                 # Assets
└── templates/              # HTML Views
```

---

## 📊 Excel Input Format

Your Excel must have a sheet named **BET** with:

| Field | Cell |
|-------|------|
| file_name | C2 |
| date_of_measurement | C3 |
| time_of_measurement | C4 |
| comment4 (equipment) | C8 |
| serial_number | C9 |

---

## ▶️ How to Run

### Command Line
```bash
python run.py
```
Default: `http://localhost:2200`

---

## 🔌 API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/` | Upload form |
| GET | `/files` | List processed files |
| POST | `/push/<id>` | Push to eLabFTW |
| GET | `/api/data/<id>`| Raw JSON data |

---

## 🔧 Configuration Reference

| Variable | Purpose |
|----------|---------|
| `ELABFTW_URL` | eLabFTW API base URL |
| `ELABFTW_TOKEN` | Personal access token |
| `DB_NAME` | SQLite DB filename |

---

## 🗄️ Database Schema

- **`file_info`**: Core metadata.
- **`bet_parameters`**: Scientific parameters.
- **`technical_info`**: Instrument technical data.
- **`bet_data_points`**: Raw isotherm data points.

---

## ❓ Troubleshooting & FAQs

- **SSL errors**: Set `ELABFTW_DISABLE_SSL=true` in `.env`.
- **Permission denied**: Check `uploads/` folder permissions.

---

## 📜 License

All rights of Code and Development belong to KIT (Karlsruhe Institute of Technology).

---

*Development and logic by Rachit Jain (`rachitjain24`)*
