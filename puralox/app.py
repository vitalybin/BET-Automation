#!/usr/bin/env python3
# puralox/app.py
# Standalone Puralox-XDI app for Ubuntu/Docker + external eLabFTW.
# Changes vs. elabftw-xdi: no hardcoded token, external URL default, SSL default on,
# and ELN experiment update uses requests.patch() instead of ExperimentsApi.patch_experiment().

import json
import logging
import os
import time
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from dotenv import load_dotenv
from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    send_from_directory,
    url_for,
)

from .config import DB_NAME, UPLOAD_FOLDER
from .db_manager import DatabaseManager
from .metadata_builder import MetadataBuilder
from .nomenclature import build_measurement_id
from .xdi_processor import XdiProcessor

# Optional imports kept for compatibility with older templates/routes.
try:
    from .excel_processor import ExcelProcessor
except Exception:  # pragma: no cover
    ExcelProcessor = None

try:
    from .bet_integration import extract_all_with_prints
except Exception:  # pragma: no cover
    extract_all_with_prints = None


load_dotenv()

ELABFTW_URL = os.getenv("ELABFTW_URL", "https://dtpa-akg.de/api/v2").rstrip("/")
ELABFTW_TOKEN = os.getenv("ELABFTW_TOKEN")
ELABFTW_DISABLE_SSL = os.getenv("ELABFTW_DISABLE_SSL", "false").lower() == "true"

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logging.getLogger("urllib3").setLevel(logging.WARNING)


class PuraloxApp:
    def __init__(self):
        base = os.path.abspath(os.path.dirname(__file__))
        self.base = base
        self.app = Flask(__name__, template_folder=os.path.join(base, "..", "templates"))
        self.app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

        self.metadata_dir = os.path.join(base, "..", "metadata")
        os.makedirs(self.metadata_dir, exist_ok=True)
        os.makedirs(self.app.config["UPLOAD_FOLDER"], exist_ok=True)

        self.db = DatabaseManager(DB_NAME)
        self.xdi_processor = XdiProcessor(self.db)
        self.metadata_builder = MetadataBuilder(self.db, self.metadata_dir)
        self.processor = ExcelProcessor(self.db) if ExcelProcessor else None
        self.verify_ssl = not ELABFTW_DISABLE_SSL

        self._ensure_optional_tables()
        self._test_elabftw_connection_nonfatal()
        self._register_routes()

    # ─── DB helpers ────────────────────────────────────────────────────
    def _ensure_optional_tables(self) -> None:
        try:
            self.xdi_processor.ensure_tables()
        except Exception:
            logging.exception("Failed to ensure XDI tables")

        # Existing DB has file_info. During very first startup it can be missing;
        # do not crash here, upload processing will create/fill it through processors.
        try:
            cols = self.db.fetchall_dict("PRAGMA table_info(file_info)")
            if cols:
                colnames = {c["name"] for c in cols}
                if "comment5" not in colnames:
                    logging.info("Adding comment5 column to file_info")
                    self.db.execute("ALTER TABLE file_info ADD COLUMN comment5 TEXT")
        except Exception:
            logging.exception("Failed to ensure comment5 column on file_info")

    def _table_exists(self, name: str) -> bool:
        rows = self.db.fetchall_dict(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (name,),
        )
        return bool(rows)

    def _has_rows_for_file(self, table: str, file_id: int) -> bool:
        try:
            if not self._table_exists(table):
                return False
            rows = self.db.fetchall_dict(
                f"SELECT 1 FROM {table} WHERE file_info_id=? LIMIT 1",
                (file_id,),
            )
            return bool(rows)
        except Exception:
            logging.exception("Failed checking rows in %s for file_id=%s", table, file_id)
            return False

    def _set_measurement_id_for_file(self, file_id: int) -> None:
        try:
            rows = self.db.fetchall_dict("SELECT * FROM file_info WHERE id=?", (file_id,))
            if not rows:
                return
            fi = rows[0]
            measurement_id = build_measurement_id(
                file_id=file_id,
                file_name=fi.get("file_name", ""),
                date_of_measurement=fi.get("date_of_measurement", ""),
                time_of_measurement=fi.get("time_of_measurement", ""),
                operator=fi.get("comment2", ""),
                instrument=fi.get("comment4", ""),
                serial_number=fi.get("serial_number", ""),
                comment1=fi.get("comment1", ""),
                comment3=fi.get("comment3", ""),
            )
            self.db.execute("UPDATE file_info SET comment5=? WHERE id=?", (measurement_id, file_id))
        except Exception as e:
            logging.warning("Failed to set measurement_id for file_id=%s: %s", file_id, e)

    # ─── eLabFTW helpers ────────────────────────────────────────────────
    def _auth_headers(self, content_type: str | None = None) -> dict[str, str]:
        headers = {"Authorization": ELABFTW_TOKEN or ""}
        if content_type:
            headers["Content-Type"] = content_type
        return headers

    def _test_elabftw_connection_nonfatal(self) -> None:
        if not ELABFTW_TOKEN:
            logging.warning("ELABFTW_TOKEN is not set; ELN push will fail until token is configured.")
            return

        url = f"{ELABFTW_URL}/experiments?limit=1"
        for attempt in range(1, 4):
            try:
                logging.debug("Testing eLabFTW connection (attempt %d): %s", attempt, url)
                resp = requests.get(url, headers=self._auth_headers(), verify=self.verify_ssl, timeout=15)
                if resp.ok:
                    logging.info("✅ eLabFTW API OK.")
                    return
                logging.warning("eLabFTW test returned [%s]: %s", resp.status_code, resp.text[:500])
            except Exception as e:
                logging.warning("Connection attempt %d failed: %s", attempt, e)
                time.sleep(2 ** attempt)

    def _create_experiment(self) -> int:
        resp = requests.post(
            f"{ELABFTW_URL}/experiments",
            headers=self._auth_headers("application/json"),
            json={},
            verify=self.verify_ssl,
            timeout=30,
        )
        if resp.status_code not in (200, 201):
            raise RuntimeError(f"Create experiment failed [{resp.status_code}]: {resp.text}")

        location = resp.headers.get("Location") or resp.headers.get("location")
        if location:
            return int(location.rstrip("/").split("/")[-1])

        # Fallback for installations returning JSON.
        try:
            data = resp.json()
            exp_id = data.get("id") or data.get("entity", {}).get("id")
            if exp_id:
                return int(exp_id)
        except Exception:
            pass

        raise RuntimeError(f"Could not determine created experiment id. Headers={dict(resp.headers)} Body={resp.text[:500]}")

    def _patch_experiment(self, exp_id: int, title: str, body: str) -> None:
        # This is the important fix: do not call ExperimentsApi.patch_experiment().
        # Newer elabapi-python versions can raise:
        # TypeError: got multiple values for argument 'body'
        resp = requests.patch(
            f"{ELABFTW_URL}/experiments/{exp_id}",
            headers=self._auth_headers("application/json"),
            json={"title": title, "body": body},
            verify=self.verify_ssl,
            timeout=30,
        )
        if not resp.ok:
            raise RuntimeError(f"Patch experiment failed [{resp.status_code}]: {resp.text}")

    def _upload_file_to_experiment(self, exp_id: int, path: str, upload_name: str | None = None) -> None:
        if not path or not os.path.exists(path):
            return
        upload_name = upload_name or os.path.basename(path)
        with open(path, "rb") as fh:
            resp = requests.post(
                f"{ELABFTW_URL}/experiments/{exp_id}/uploads",
                headers=self._auth_headers(),
                files={"file": (upload_name, fh)},
                verify=self.verify_ssl,
                timeout=60,
            )
        if not resp.ok:
            raise RuntimeError(f"Upload failed [{resp.status_code}]: {resp.text}")

    def _tag_experiment(self, exp_id: int, tag: str) -> None:
        resp = requests.post(
            f"{ELABFTW_URL}/experiments/{exp_id}/tags",
            headers=self._auth_headers("application/json"),
            json={"tag": tag},
            verify=self.verify_ssl,
            timeout=30,
        )
        if not resp.ok:
            logging.warning("Tagging failed [%s]: %s", resp.status_code, resp.text)

    # ─── Routes ────────────────────────────────────────────────────────
    def _register_routes(self) -> None:
        @self.app.route("/", methods=["GET", "POST"])
        def upload():
            if request.method == "POST":
                f = request.files.get("file")
                if not f:
                    return "No file provided.", 400

                dst = os.path.join(self.app.config["UPLOAD_FOLDER"], f.filename)
                logging.debug("Saving uploaded file to %s", dst)
                f.save(dst)

                ext = os.path.splitext(f.filename)[1].lower()
                try:
                    if ext in (".xdi", ".txt"):
                        new_file_id = self.xdi_processor.process_file(dst, original_filename=f.filename)
                        self._set_measurement_id_for_file(new_file_id)
                    elif ext in (".xlsx", ".xls") and self.processor:
                        new_file_id = self.processor.process_file(dst)
                        self._set_measurement_id_for_file(new_file_id)
                    else:
                        return "Unsupported file type. Upload .xdi or .txt for Puralox-XDI.", 400

                    try:
                        self.metadata_builder.generate_metadata(new_file_id)
                    except Exception:
                        logging.exception("metadata generation failed")

                    return redirect(url_for("list_files"))
                except Exception:
                    logging.exception("Upload processing failed")
                    return jsonify({"error": "Failed to process file"}), 500

            return render_template("upload.html")

        @self.app.route("/files")
        def list_files():
            try:
                rows = self.db.fetchall_dict(
                    "SELECT id, file_name, date_of_measurement, time_of_measurement, comment5 "
                    "FROM file_info ORDER BY id DESC"
                )
            except Exception:
                logging.exception("Could not load file list")
                rows = []

            items: list[dict[str, Any]] = []
            for r in rows:
                fname = r.get("file_name") or ""
                if self._has_rows_for_file("xdi_points", r["id"]):
                    vurl = url_for("view_xdi", file_id=r["id"])
                    ftype = "XDI"
                    region = "KIT Campus Nord"
                else:
                    vurl = url_for("view_xdi", file_id=r["id"])
                    ftype = "File"
                    region = "KIT Campus Nord"

                items.append({
                    "id": r["id"],
                    "file_name": fname,
                    "measurement_id": r.get("comment5") or "",
                    "date": r.get("date_of_measurement", ""),
                    "time": r.get("time_of_measurement", ""),
                    "type": ftype,
                    "region": region,
                    "view_url": vurl,
                    "metadata_url": url_for("download_metadata_for_file", file_id=r["id"]),
                })
            return render_template("list_files.html", items=items)

        @self.app.route("/uploads")
        def uploads():
            return redirect(url_for("list_files"))

        @self.app.route("/view/xdi/<int:file_id>")
        def view_xdi(file_id: int):
            fi_rows = self.db.fetchall_dict("SELECT * FROM file_info WHERE id=?", (file_id,))
            if not fi_rows:
                return "Not found", 404
            fi = fi_rows[0]

            scan_rows = self.db.fetchall_dict(
                "SELECT * FROM xdi_scans WHERE file_info_id=? ORDER BY id DESC LIMIT 1",
                (file_id,),
            )
            scan = scan_rows[0] if scan_rows else {}
            cols = self.db.fetchall_dict(
                "SELECT col_index, col_name FROM xdi_columns WHERE file_info_id=? ORDER BY col_index ASC",
                (file_id,),
            )
            pts = self.db.fetchall_dict(
                "SELECT energy, mu, bkg FROM xdi_points WHERE file_info_id=? ORDER BY id ASC",
                (file_id,),
            )
            df_json_rows = self.db.fetchall_dict(
                "SELECT df_json FROM xdi_dataframe WHERE file_info_id=? LIMIT 1",
                (file_id,),
            )
            df_json = df_json_rows[0]["df_json"] if df_json_rows else None

            return render_template(
                "view_xdi.html",
                file_id=file_id,
                info=fi,
                scan=scan,
                cols=cols,
                pts=pts[:500],
                pts_data=pts,
                df_json=df_json,
                metadata_url=url_for("download_metadata_for_file", file_id=file_id),
                sample_region="KIT Campus Nord",
            )

        # Compatibility fallbacks.
        @self.app.route("/view/excel/<int:file_id>")
        def view_excel(file_id: int):
            return redirect(url_for("view_xdi", file_id=file_id))

        @self.app.route("/view/pdf/<int:file_id>")
        def view_pdf(file_id: int):
            return redirect(url_for("view_xdi", file_id=file_id))

        @self.app.route("/eln/<int:file_id>", methods=["POST"])
        def eln_create(file_id: int):
            return self._eln_create_local_json(file_id, template_id=None)

        @self.app.route("/eln/<int:file_id>/<int:template_id>", methods=["POST"])
        def eln_create_legacy(file_id: int, template_id: int):
            return self._eln_create_local_json(file_id, template_id=template_id)

        @self.app.route("/metadata/<path:filename>")
        def download_metadata(filename):
            return send_from_directory(self.metadata_dir, filename, as_attachment=True)

        @self.app.route("/metadata/file/<int:file_id>")
        def download_metadata_for_file(file_id: int):
            md_path = self.metadata_builder.generate_metadata(file_id)
            return send_from_directory(self.metadata_dir, os.path.basename(md_path), as_attachment=True)

        @self.app.route("/api")
        def api_info():
            return render_template("api.html")

        @self.app.route("/api/elab/experiments")
        def api_elab_experiments():
            try:
                resp = requests.get(
                    f"{ELABFTW_URL}/experiments?limit=10",
                    headers=self._auth_headers(),
                    verify=self.verify_ssl,
                    timeout=30,
                )
                return jsonify(resp.json() if resp.text else []), resp.status_code
            except Exception as e:
                logging.exception("api_elab_experiments failed")
                return jsonify({"error": str(e)}), 500

    # ─── ELN push core ────────────────────────────────────────────────
    def _eln_create_local_json(self, file_id: int, template_id: int | None = None):
        try:
            if not ELABFTW_TOKEN:
                return jsonify({"ok": False, "stage": "config", "error": "ELABFTW_TOKEN is missing"}), 500

            fi_rows = self.db.fetchall_dict("SELECT * FROM file_info WHERE id=?", (file_id,))
            if not fi_rows:
                return jsonify({"ok": False, "stage": "load", "error": "file_info not found"}), 404
            fi = fi_rows[0]

            scan_rows = self.db.fetchall_dict(
                "SELECT * FROM xdi_scans WHERE file_info_id=? ORDER BY id DESC LIMIT 1",
                (file_id,),
            )
            scan = scan_rows[0] if scan_rows else {}
            cols = self.db.fetchall_dict(
                "SELECT col_index, col_name FROM xdi_columns WHERE file_info_id=? ORDER BY col_index ASC",
                (file_id,),
            )
            pts = self.db.fetchall_dict(
                "SELECT energy, mu, bkg FROM xdi_points WHERE file_info_id=? ORDER BY id ASC",
                (file_id,),
            )

            title = request.form.get("title") or f"XDI Measurement: {fi.get('file_name') or file_id}"
            measurement_id = fi.get("comment5") or fi.get("file_name") or f"file_{file_id}"

            energy_values = [p.get("energy") for p in pts if p.get("energy") is not None]
            mu_values = [p.get("mu") for p in pts if p.get("mu") is not None]
            energy_min = min(energy_values) if energy_values else "—"
            energy_max = max(energy_values) if energy_values else "—"

            body = f"""
<h1>XDI Measurement Report</h1>

<h2>Metadata</h2>
<ul>
  <li><strong>Measurement ID:</strong> {measurement_id}</li>
  <li><strong>File:</strong> {fi.get('file_name', '')}</li>
  <li><strong>Date:</strong> {fi.get('date_of_measurement', '')} {fi.get('time_of_measurement', '')}</li>
  <li><strong>Operator:</strong> {fi.get('comment2', '')}</li>
  <li><strong>Instrument:</strong> {fi.get('comment4', '')}</li>
  <li><strong>Serial #:</strong> {fi.get('serial_number', '')}</li>
  <li><strong>Region:</strong> KIT Campus Nord</li>
</ul>

<h2>XDI Scan</h2>
<ul>
  <li><strong>Scan title:</strong> {scan.get('title', '—')}</li>
  <li><strong>Scan date:</strong> {scan.get('scan_date', '—')}</li>
  <li><strong>Number of points:</strong> {len(pts)}</li>
  <li><strong>Energy range:</strong> {energy_min} – {energy_max}</li>
  <li><strong>Columns:</strong> {', '.join(c.get('col_name', '') for c in cols) or '—'}</li>
</ul>

<h2>Preview</h2>
<table border="1" cellspacing="0" cellpadding="4">
<tr><th>#</th><th>Energy</th><th>Mu</th><th>Bkg</th></tr>
{''.join(f"<tr><td>{i+1}</td><td>{p.get('energy','')}</td><td>{p.get('mu','')}</td><td>{p.get('bkg','')}</td></tr>" for i, p in enumerate(pts[:20]))}
</table>
"""

            logging.info("ELN push for XDI file_id=%s", file_id)
            exp_id = self._create_experiment()
            self._patch_experiment(exp_id, title, body)

            # Upload original XDI file if present in upload folder.
            file_name = fi.get("file_name") or ""
            original_path = os.path.join(self.app.config["UPLOAD_FOLDER"], file_name)
            try:
                self._upload_file_to_experiment(exp_id, original_path, file_name)
            except Exception:
                logging.exception("Original XDI upload failed (non-fatal)")

            # Upload generated metadata Excel if present.
            try:
                md_path = self.metadata_builder.generate_metadata(file_id)
                self._upload_file_to_experiment(exp_id, md_path, os.path.basename(md_path))
            except Exception:
                logging.exception("Metadata upload failed (non-fatal)")

            try:
                self._tag_experiment(exp_id, "XDI_result")
            except Exception:
                logging.exception("Tagging failed (non-fatal)")

            return jsonify({
                "ok": True,
                "exp_id": exp_id,
                "experiment_url": f"{ELABFTW_URL}/experiments/{exp_id}",
                "template_id": template_id,
            }), 201

        except Exception as e:
            logging.exception("eln_create_local_json failure")
            return jsonify({"ok": False, "stage": "unknown", "error": str(e)}), 500

    def run(self):
        self.app.run(host="0.0.0.0", port=5000, debug=True)


puralox_app = PuraloxApp()
app = puralox_app.app

if __name__ == "__main__":
    puralox_app.run()
