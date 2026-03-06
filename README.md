# Puralox — BET Data Processor & eLabFTW Integrator

A Flask application to:

1. Parse BET experiment data from Excel (`.xlsx`) files  
2. **(New)** Parse comprehensive BET data from instrument PDF reports using PyMuPDF  
3. Store metadata, BET parameters, technical info and isotherm data points in SQLite  
4. Visualize and inspect data in a web UI  
5. Generate PDF/DOCX reports containing BET plots (Isotherm, t-Plot, BJH)  
6. Push experiments—with rich-text summary, PDF attachment and tag—into your eLabFTW instance via its v2 API  

---

## 📋 Table of Contents

1. [Features](#features)  
2. [Architecture & Design](#architecture--design)  
3. [Prerequisites](#prerequisites)  
4. [eLabFTW Installation & Docker-Compose](#elabftw-installation--docker-compose)  
5. [Project Structure](#project-structure)  
6. [Installation & Setup](#installation--setup)  
7. [Excel Input Format](#excel-input-format)  
8. [Running Locally](#running-locally)  
9. [Web UI Walkthrough](#web-ui-walkthrough)  
10. [API Endpoints](#api-endpoints)  
11. [PDF Reporting & Plotting](#pdf-reporting--plotting)  
12. [Docker & Docker-Compose (Puralox)](#docker--docker-compose-puralox)  
13. [Configuration Reference](#configuration-reference)  
14. [Database Schema](#database-schema)  
15. [Customization & Extension](#customization--extension)  
16. [Troubleshooting & FAQs](#troubleshooting--faqs)  
17. [Contributing](#contributing)  
18. [License](#license)  

---

## 🚀 Features

- **Excel & PDF → SQLite**  
  - Parses `.xlsx` and scientific PDF reports into structured database tables.  
- **Rich Web UI**  
  - Upload files, list processed results, and inspect tables with DataTables.  
- **Scientific Plotting**  
  - Generates Isotherm, t-Plot, and BJH graphs using Matplotlib.  
- **eLabFTW Integration**  
  - Creates & patches experiments via API, attaches reports, and applies auto-tags.  
- **Clean Architecture**  
  - 1:1 class-to-module mapping for scale and maintainability.  

---

## 🛠️ Architecture & Design

Puralox follows a decoupled, service-oriented architecture. Data extraction is handled by dedicated parsers, while business logic and database persistence are managed by processors.

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

---

## 🔧 Prerequisites

- **Python 3.10+**  
- **SQLite** (bundled)  
- **eLabFTW** server (v2 API) & access token  
- **PyMuPDF (`fitz`)** (Required for PDF parsing)  
- **Matplotlib/Pandas/NumPy**

---

## 🖥️ eLabFTW Installation & Docker-Compose

Follow [https://doc.elabftw.net/](https://doc.elabftw.net/) for full installation. Sample `docker-compose.yml` for eLabFTW:

```yaml
networks:
  elabftw-net:

services:
  web:
    image: elabftw/elabimg:stable
    container_name: elabftw
    restart: always
    environment:
      - DB_HOST=mysql
      - DB_NAME=elabftw
      - SITE_URL=https://localhost
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
    networks:
      - elabftw-net
```

---

## 📂 Project Structure

Each core logic component is isolated in its own module:

```text
puralox-app/
├── puralox/
│   ├── app.py                  # PuraloxApp (Routes & Logic)
│   ├── base_importer.py        # Abstract Base Class
│   ├── database_manager.py     # DatabaseManager (SQLite helper)
│   ├── excel_processor.py      # ExcelProcessor (Excel → DB)
│   ├── pdf_processor.py        # PdfProcessor (PDF → DB)
│   ├── bet_pdf_parser.py       # BetPdfParser (PDF Extraction)
│   ├── eln_client.py           # ElnClient (eLabFTW API)
│   ├── measurement_id_builder.py# MeasurementIdBuilder (IDs)
│   ├── metadata_builder.py     # MetadataBuilder (Excel gen)
│   ├── template_processor.py   # TemplateProcessor (HTML summary)
│   ├── static/                 # Assets
│   └── templates/              # HTML views
├── run.py                      # Entry point
└── README.md                   # ← You are here
```

---

## ⚙️ Installation & Setup

```bash
git clone https://github.com/vitalybin/BET-Automation.git
cd BET-Automation
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---

## 📊 Excel Input Format

Target sheet: **BET**
- **C2-C10**: Metadata (Filename, Date, Operator)
- **Rows 12-28**: BET parameters
- **Rows 32+**: Plot data points

---

## ▶️ Running Locally

```bash
python run.py
```
Default URL: `http://localhost:2200`

---

## 🌐 Web UI Walkthrough

1. **Home**: Upload `.xlsx` or `.pdf`.
2. **Files**: List all processed experiments.
3. **Detail**: View scientific data, and click "Push to eLabFTW".

---

## 🔌 API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/` | Upload form |
| POST | `/push/<id>` | Push to ELN |
| GET | `/api/data/<id>`| Raw data (JSON) |

---

## ❓ Troubleshooting & FAQs

- **SSL issues**: Set `ELABFTW_DISABLE_SSL=true`.
- **ModuleNotFound (fitz)**: Install `pymupdf`.

---

## 📜 License

All rights belong to KIT (Karlsruhe Institute of Technology).

---

*Development and logic by Rachit Jain (`rachitjain24`)*
