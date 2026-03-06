#!/usr/bin/env python3
# puralox/app.py

import os
import re
import time
import logging
import warnings
from io import BytesIO

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from flask import (
    Flask, request, redirect, url_for,
    render_template, jsonify, send_from_directory
)
from dotenv import load_dotenv

import elabapi_python
from elabapi_python import ExperimentsApi
from elabapi_python.rest import ApiException

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Image as RLImage

from .config import UPLOAD_FOLDER, DB_NAME
from .db_manager import DatabaseManager
from .excel_processor import ExcelProcessor
from .pdf_processor import PdfProcessor
from .template_processor import TemplateProcessor
from .nomenclature import build_measurement_id
from .metadata_builder import MetadataBuilder   # <— separate module for metadata Excel
import requests

# ─── CONFIG ────────────────────────────────────────────────────────────
load_dotenv()
ELABFTW_URL   = os.getenv("ELABFTW_URL", "https://localhost/api/v2")
ELABFTW_TOKEN = os.getenv(
    "ELABFTW_TOKEN",
    "15-c80599971b8e2592a5fadaa45f143f8201828540bbb3b0cf3731316c65c885c50e30bb08ff94069b006115"
)
# ──────────────────────────────────────────────────────────────────────

# ─── LOGGING ───────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s"
)
logging.getLogger("urllib3").setLevel(logging.DEBUG)
warnings.filterwarnings("ignore", category=Warning, module="urllib3")
# ──────────────────────────────────────────────────────────────────────


class PuraloxApp:

    def __init__(self):
        base = os.path.abspath(os.path.dirname(__file__))
        self.app = Flask(
            __name__,
            template_folder=os.path.join(base, "..", "templates")
        )
        self.app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

        self.metadata_dir = os.path.join(base, "..", "metadata")
        os.makedirs(self.metadata_dir, exist_ok=True)

        self.db        = DatabaseManager(DB_NAME)
        self.processor = ExcelProcessor(self.db)

        # Importers (inheritance: BaseImporter -> ExcelProcessor, PdfProcessor)
        self.pdf_processor = PdfProcessor(self.db, self.app.config["UPLOAD_FOLDER"])

        # Template builder used by ELN push (so UML matches runtime)
        self.template_processor = TemplateProcessor(self.db)

        # NEW: dedicated builder for metadata Excel
        self.metadata_builder = MetadataBuilder(self.db, self.metadata_dir)

        self._ensure_optional_tables()
        self._configure_elabftw()
        self._register_routes()

    # ─── DB helpers ────────────────────────────────────────────────────
    def _ensure_optional_tables(self):
        # Extra tables used by PDF → DB and metadata
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS bet_summaries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER NOT NULL,
                key TEXT NOT NULL,
                value TEXT
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS isotherm_points (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER NOT NULL,
                p_over_p0 REAL,
                vol_cc_g_stp REAL
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS tplot_points (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER NOT NULL,
                thickness_nm REAL,
                volume_cc_g_stp REAL
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS bjh_desorption (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER NOT NULL,
                diameter_nm REAL,
                porevol_ccg REAL,
                porearea_m2g REAL,
                dv_d REAL,
                ds_d REAL,
                dv_logd REAL,
                ds_logd REAL
            )
        """)

        # --- ensure comment5 exists on file_info (for Measurement ID) ---
        try:
            cols = self.db.fetchall_dict("PRAGMA table_info(file_info)")
            colnames = {c["name"] for c in cols}
            if "comment5" not in colnames:
                logging.info("Adding comment5 column to file_info")
                self.db.execute("ALTER TABLE file_info ADD COLUMN comment5 TEXT")
        except Exception:
            logging.exception("Failed to ensure comment5 column on file_info")

    def _table_exists(self, name: str) -> bool:
        rows = self.db.fetchall_dict(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (name,)
        )
        return bool(rows)

    # helper to set Measurement ID for Excel uploads
    def _set_measurement_id_for_file(self, file_id: int) -> None:
        try:
            rows = self.db.fetchall_dict(
                "SELECT * FROM file_info WHERE id=?",
                (file_id,),
            )
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
            self.db.execute(
                "UPDATE file_info SET comment5=? WHERE id=?",
                (measurement_id, file_id)
            )
        except Exception as e:
            logging.warning("Failed to set measurement_id for file_id=%s: %s", file_id, e)

    # ─── ELN template helper ───────────────────────────────────────────
    def _get_templates(self):
        """
        Fetch templates from eLabFTW. If the token or URL is wrong,
        we just return [] instead of crashing.
        """
        url = f"{ELABFTW_URL}/experiments_templates?limit=100"
        logging.debug("GET %s", url)
        try:
            resp = requests.get(
                url,
                headers={"Authorization": ELABFTW_TOKEN},
                verify=self.verify_ssl
            )
        except Exception as e:
            logging.error("Template HTTP error: %s", e)
            return []

        logging.debug("→ template status=%s", resp.status_code)
        if not resp.ok:
            logging.error("Template fetch failed [%s]: %s",
                          resp.status_code, resp.text)
            return []

        try:
            return resp.json()
        except Exception:
            logging.exception("Failed to parse templates JSON")
            return []

    # ─── ELN configuration ─────────────────────────────────────────────
    def _configure_elabftw(self):
        cfg = elabapi_python.Configuration()
        cfg.host       = ELABFTW_URL
        disable_ssl = os.getenv("ELABFTW_DISABLE_SSL", "true").lower() == "true"
        cfg.verify_ssl = not disable_ssl
        client = elabapi_python.ApiClient(cfg)
        client.set_default_header("Authorization", ELABFTW_TOKEN)

        self.exp_api    = ExperimentsApi(client)
        self.verify_ssl = cfg.verify_ssl

        for attempt in range(1, 6):
            try:
                logging.debug("Testing eLabFTW connection (attempt %d)", attempt)
                self.exp_api.read_experiments(limit=1)
                logging.info("✅ eLabFTW API OK.")
                return
            except Exception as e:
                logging.warning("Connection attempt %d failed: %s", attempt, e)
                time.sleep(2 ** attempt)

    # ─── ROUTES ────────────────────────────────────────────────────────
    def _register_routes(self):

        # ── Upload: Excel or PDF ──────────────────────────────────────
        @self.app.route("/", methods=["GET", "POST"])
        def upload():
            if request.method == "POST":
                f = request.files.get("file")
                if not f:
                    return "No file provided.", 400

                os.makedirs(self.app.config["UPLOAD_FOLDER"], exist_ok=True)
                dst = os.path.join(self.app.config["UPLOAD_FOLDER"], f.filename)
                logging.debug("Saving uploaded file to %s", dst)
                f.save(dst)

                ext = os.path.splitext(f.filename)[1].lower()
                try:
                    if ext in (".xlsx", ".xls"):
                        # Excel import (via common BaseImporter interface)
                        new_file_id = self.processor.import_file(dst, original_filename=f.filename)

                    elif ext == ".pdf":
                        # PDF import (via PdfProcessor)
                        new_file_id = self.pdf_processor.import_file(dst, original_filename=f.filename)

                    else:
                        return "Unsupported file type. Upload .xlsx or .pdf", 400

                    # Create unified metadata for both Excel + PDF (via separate module)
                    try:
                        self.metadata_builder.generate_metadata(new_file_id)
                    except Exception:
                        logging.exception("metadata generation failed")

                    # Go to unified list
                    return redirect(url_for("list_files"))

                except Exception:
                    logging.exception("Upload processing failed")
                    return jsonify({"error": "Failed to process file"}), 500

            return render_template("upload.html")

        # ── Unified list page (list_files.html) ───────────────────────
        @self.app.route("/files")
        def list_files():
            rows = self.db.fetchall_dict(
                "SELECT id, file_name, date_of_measurement, time_of_measurement, comment5 "
                "FROM file_info ORDER BY id DESC"
            )
            items = []
            for r in rows:
                fname = r.get("file_name") or ""
                ext = os.path.splitext(fname)[1].lower()
                if ext == ".pdf":
                    vurl = url_for("view_pdf", file_id=r["id"])
                    ftype = "PDF"
                else:
                    vurl = url_for("view_excel", file_id=r["id"])
                    ftype = "Excel"
                items.append({
                    "id": r["id"],
                    "file_name": fname,
                    "measurement_id": r.get("comment5") or "",
                    "date": r.get("date_of_measurement", ""),
                    "time": r.get("time_of_measurement", ""),
                    "type": ftype,
                    "view_url": vurl,
                    "metadata_url": url_for("download_metadata_for_file", file_id=r["id"]),
                })
            return render_template("list_files.html", items=items)

        # Alias: /uploads -> /files
        @self.app.route("/uploads")
        def uploads():
            return redirect(url_for("list_files"))

        # ── Excel detail view ─────────────────────────────────────────
        @self.app.route("/view/excel/<int:file_id>")
        def view_excel(file_id: int):
            # Load DB rows into DataFrames
            to_df = lambda q: pd.DataFrame(self.db.fetchall_dict(q))

            info_df = to_df(f"SELECT * FROM file_info WHERE id={file_id}")
            bet_df  = to_df(f"SELECT * FROM bet_parameters WHERE file_info_id={file_id}")
            tech_df = to_df(f"SELECT * FROM technical_info WHERE file_info_id={file_id}")
            cols_df = to_df(f"SELECT * FROM bet_plot_columns WHERE file_info_id={file_id}")
            pts_df  = to_df(f"SELECT * FROM bet_data_points WHERE file_info_id={file_id}")

            # Convert to plain dicts / list-of-dicts for Jinja
            info = info_df.iloc[0].to_dict() if not info_df.empty else {}
            bet  = bet_df.iloc[0].to_dict() if not bet_df.empty else {}
            tech = tech_df.iloc[0].to_dict() if not tech_df.empty else {}
            cols = cols_df.to_dict(orient="records") if not cols_df.empty else []
            pts_rows = pts_df.to_dict(orient="records") if not pts_df.empty else []

            # For interactive plot in template (clean, sorted)
            pts_data = []
            if pts_rows:
                pts_sorted = sorted(
                    pts_rows,
                    key=lambda r: r.get("no") if r.get("no") is not None else 0
                )
                pts_data = [
                    {
                        "no": r.get("no"),
                        "p_p0": r.get("p_p0"),
                        "p_va_p0_p": r.get("p_va_p0_p"),
                    }
                    for r in pts_sorted
                ]

            md_url = url_for("download_metadata_for_file", file_id=file_id)
            sample_region = "KIT Campus South"  # Excel = South

            return render_template(
                "view_excel.html",
                file_id=file_id,
                info=info,          # dict
                bet=bet,            # dict
                tech=tech,          # dict
                cols=cols,          # list[dict]
                pts=pts_rows,       # list[dict] for table
                pts_data=pts_data,  # list[dict] for chart
                metadata_url=md_url,
                sample_region=sample_region,
            )

        # ── PDF detail view (summary + plot + core/non-core metadata) ─
        @self.app.route("/view/pdf/<int:file_id>")
        def view_pdf(file_id: int):
            # Full file_info so template can show Measurement ID etc.
            fi_rows = self.db.fetchall_dict(
                "SELECT * FROM file_info WHERE id=?",
                (file_id,),
            )
            fi = fi_rows[0] if fi_rows else {}

            pdf_filename = fi.get("file_name") or f"file_{file_id}.pdf"

            # Count points
            cnt_rows = self.db.fetchall_dict(
                "SELECT COUNT(*) AS c FROM bet_data_points WHERE file_info_id=?",
                (file_id,),
            )
            points_count = cnt_rows[0]["c"] if cnt_rows else 0

            # All summaries from bet_summaries (for metadata tables)
            kv_rows = []
            if self._table_exists("bet_summaries"):
                kv_rows = self.db.fetchall_dict(
                    "SELECT key, value FROM bet_summaries "
                    "WHERE file_info_id=? ORDER BY key",
                    (file_id,),
                )

            # ------ classify into core vs extra (non-core) ------
            # Use the same "core" concept as the metadata Excel
            core_keys = {
                "general:Sample weight",
                "general:Analysis gas",
                "general:Bath Temp",
                "general:OutgasTemp",
                "multipoint_bet_summary:Surface Area",
                "isotherm_summary:Surface Area",
                "isotherm_summary: Surface Area",
                "tplot_summary:Pore Volume",
                "tplot_summary: Pore Volume",
                "tplot_summary:Pore Diameter Dv(d)",
                "tplot_summary: Pore Diameter Dv(d)",
                "general:OperatorPrimary",
                "general:Operators",
                "general:Instrument",
            }

            core_fields = []
            extra_fields = []

            for r in kv_rows:
                key = r["key"]
                val = r["value"]
                row = {"key": key, "value": val}
                if key in core_keys:
                    core_fields.append(row)
                else:
                    extra_fields.append(row)

            # For debugging
            logging.debug(
                "PDF view for file_id=%s: kv_rows=%s, core_fields=%s, extra_fields=%s, bet_present=%s",
                file_id,
                len(kv_rows),
                len(core_fields),
                len(extra_fields),
                bool(points_count),
            )

            # Points for plot & BET table
            pts_rows = self.db.fetchall_dict(
                "SELECT no, p_p0, p_va_p0_p FROM bet_data_points "
                "WHERE file_info_id=? ORDER BY no",
                (file_id,),
            )
            pts_data = pts_rows  # list[dict]

            # Default plot ranges (for UI sliders / inputs)
            x_vals = [r["p_p0"] for r in pts_rows if r["p_p0"] is not None]
            y_vals = [r["p_va_p0_p"] for r in pts_rows if r["p_va_p0_p"] is not None]

            default_x_min = min(x_vals) if x_vals else None
            default_x_max = max(x_vals) if x_vals else None
            default_y_min = min(y_vals) if y_vals else None
            default_y_max = max(y_vals) if y_vals else None

            # Header object for template (used as header.*)
            measurement_id = fi.get("comment5") or fi.get("file_name") or f"File {file_id}"
            header = {
                "measurement_id": measurement_id,
                "date": fi.get("date_of_measurement", ""),
                "time": fi.get("time_of_measurement", ""),
                "operator": fi.get("comment2", ""),
                "instrument": fi.get("comment4", ""),
                "serial_number": fi.get("serial_number", ""),
                "version": fi.get("version", ""),
            }

            # bundle: still available if you use it elsewhere
            bundle = {
                "file_id": file_id,
                "points_count": points_count,
                "summaries": [{"key": r["key"], "value": r["value"]} for r in kv_rows],
            }

            md_url = url_for("download_metadata_for_file", file_id=file_id)
            sample_region = "KIT Campus Nord"  # PDF = North

            return render_template(
                "view_pdf_extract.html",
                pdf_filename=pdf_filename,
                header=header,
                bundle=bundle,
                file_info=fi,
                pts_data=pts_data,
                metadata_url=md_url,
                sample_region=sample_region,
                default_x_min=default_x_min,
                default_x_max=default_x_max,
                default_y_min=default_y_min,
                default_y_max=default_y_max,
                core_fields=core_fields,
                extra_fields=extra_fields,
            )

        # ── ELN push (template_id optional) ───────────────────────────
        @self.app.route("/eln/<int:file_id>", methods=["POST"])
        def eln_create(file_id: int):
            return self._eln_create_local_json(file_id, template_id=None)

        # Legacy URL /eln/<file_id>/<template_id>
        @self.app.route("/eln/<int:file_id>/<int:template_id>", methods=["POST"])
        def eln_create_legacy(file_id: int, template_id: int):
            logging.info("ELN push (legacy URL) with template_id=%s", template_id)
            return self._eln_create_local_json(file_id, template_id=template_id)

        # ── Metadata download ─────────────────────────────────────────
        @self.app.route("/metadata/<filename>")
        def download_metadata(filename):
            return send_from_directory(self.metadata_dir, filename, as_attachment=True)

        @self.app.route("/metadata/file/<int:file_id>")
        def download_metadata_for_file(file_id: int):
            # Use new module to (re)generate before download
            md_path = self.metadata_builder.generate_metadata(file_id)
            fname = os.path.basename(md_path)
            return send_from_directory(self.metadata_dir, fname, as_attachment=True)

        # ── API info page ─────────────────────────────────────────────
        @self.app.route("/api")
        def api_info():
            return render_template("api.html")

        # ELN: list templates for frontend dropdown
        @self.app.route("/api/elab/templates")
        def api_elab_templates():
            try:
                tpls = self._get_templates()
                out = [{"id": t["id"], "title": t.get("title", f"Template {t['id']}")} for t in tpls]
                return jsonify(out), 200
            except Exception as e:
                logging.exception("api_elab_templates failed")
                return jsonify({"error": str(e)}), 500

        # Optional quick ELN experiment list API
        @self.app.route("/api/elab/experiments")
        def api_elab_experiments():
            try:
                exps = self.exp_api.read_experiments(limit=10)
                return jsonify([e.to_dict() for e in exps]), 200
            except ApiException as e:
                return jsonify({"error": "API Exception", "details": str(e)}), 500

    # ─── ELN push core (build HTML directly from DB, JSON only) ──────
    def _eln_create_local_json(self, file_id: int, template_id: int | None = None):
        try:
            # ---- template_id (optional, informational) ----
            form_tid = request.form.get("template_id")
            if form_tid and str(form_tid).strip():
                try:
                    template_id = int(form_tid)
                except ValueError:
                    pass

            title = request.form.get("title") or f"File {file_id}"

            # Optional plot range coming from Excel/PDF view
            plot_xmin = request.form.get("plot_xmin")
            plot_xmax = request.form.get("plot_xmax")
            try:
                plot_xmin = float(plot_xmin) if plot_xmin not in (None, "", "null") else None
            except ValueError:
                plot_xmin = None
            try:
                plot_xmax = float(plot_xmax) if plot_xmax not in (None, "", "null") else None
            except ValueError:
                plot_xmax = None

            logging.info(
                "ELN push for file_id=%s with template_id=%s (range: %s – %s)",
                file_id, template_id, plot_xmin, plot_xmax
            )

            # ---- Load data from DB ----
            fi_rows = self.db.fetchall_dict("SELECT * FROM file_info WHERE id=?", (file_id,))
            if not fi_rows:
                return jsonify({"ok": False, "stage": "load", "error": "file_info not found"}), 404
            fi = fi_rows[0]

            bet_rows = self.db.fetchall_dict(
                "SELECT * FROM bet_parameters WHERE file_info_id=?", (file_id,)
            )
            bet = bet_rows[0] if bet_rows else {}

            tech_rows = self.db.fetchall_dict(
                "SELECT * FROM technical_info WHERE file_info_id=?", (file_id,)
            )
            tech = tech_rows[0] if tech_rows else {}

            pts = self.db.fetchall_dict(
                "SELECT * FROM bet_data_points WHERE file_info_id=? ORDER BY no",
                (file_id,)
            )

            # ---- Build ELN HTML (moved into TemplateProcessor) ----
            html_body = self.template_processor.build_eln_html(file_id)

            # Print some debug on console for you
            measurement_id = (fi.get("comment5") or fi.get("file_name") or f"file_{file_id}")
            print("\n[ELN] Creating experiment for file_id:", file_id)
            print("[ELN] Measurement ID:", measurement_id)
            print("[ELN] HTML body (first 500 chars):")
            print(html_body[:500])
            print("------- END OF PREVIEW -------\n")


            # Print some debug on console for you
            print("\n[ELN] Creating experiment for file_id:", file_id)
            print("[ELN] Measurement ID:", measurement_id)
            print("[ELN] HTML body (first 500 chars):")
            print(html_body[:500])
            print("------- END OF PREVIEW -------\n")

            # ---- Create experiment in ELN ----
            try:
                _, status, headers = self.exp_api.post_experiment_with_http_info(body={})
            except Exception as e:
                logging.exception("post_experiment failed")
                return jsonify({"ok": False, "stage": "create", "error": str(e)}), 500

            if status != 201:
                return jsonify({"ok": False, "stage": "create", "error": f"Create failed ({status})"}), 500

            exp_id = headers["Location"].rstrip("/").split("/")[-1]

            # ---- Patch with title + body ----
            try:
                self.exp_api.patch_experiment(exp_id, body={"title": title, "body": html_body})
            except Exception as e:
                logging.exception("patch_experiment failed")
                return jsonify({"ok": False, "stage": "patch", "error": str(e), "exp_id": exp_id}), 500

            # ---- Attach BET plot (best-effort, honoring optional range) ----
            pts_for_plot = self.db.fetchall_dict(
                "SELECT p_p0, p_va_p0_p FROM bet_data_points WHERE file_info_id=?",
                (file_id,)
            )
            if pts_for_plot:
                try:
                    x_all = np.array([r["p_p0"] for r in pts_for_plot if r["p_p0"] is not None])
                    y_all = np.array([r["p_va_p0_p"] for r in pts_for_plot if r["p_va_p0_p"] is not None])

                    if len(x_all) and len(y_all):
                        # Apply optional range from UI
                        mask = np.ones_like(x_all, dtype=bool)
                        if plot_xmin is not None:
                            mask &= x_all >= plot_xmin
                        if plot_xmax is not None:
                            mask &= x_all <= plot_xmax

                        x = x_all[mask]
                        y = y_all[mask]

                        # Fallback: if filter removes everything, use full range
                        if not len(x) or not len(y):
                            x, y = x_all, y_all

                        fig, ax = plt.subplots()
                        ax.scatter(x, y, s=20)
                        ax.set_title("BET Plot")
                        ax.set_xlabel("p/p0")
                        ax.set_ylabel("p/va_p0_p")
                        img_buf = BytesIO()
                        fig.savefig(img_buf, format="PNG", bbox_inches="tight")
                        plt.close(fig)
                        img_buf.seek(0)

                        pdf_buf = BytesIO()
                        doc = SimpleDocTemplate(pdf_buf, pagesize=letter)
                        doc.build([RLImage(img_buf, width=400, height=300)])
                        pdf_buf.seek(0)

                        files = {
                            "file": (f"BET_Plot_{file_id}.pdf", pdf_buf.read(), "application/pdf")
                        }
                        up_url = f"{ELABFTW_URL}/experiments/{exp_id}/uploads"
                        resp = requests.post(
                            up_url,
                            headers={"Authorization": ELABFTW_TOKEN},
                            files=files,
                            verify=self.verify_ssl
                        )
                        logging.debug(
                            "Upload plot resp: %s %s (range: %s – %s)",
                            resp.status_code,
                            resp.text,
                            plot_xmin,
                            plot_xmax,
                        )
                except Exception:
                    logging.exception("Plot upload failed (non-fatal)")

            # ---- Add BET_result tag (best-effort, template-aware) ----
            try:
                tag_name = "BET_result"
                if template_id is not None:
                    tag_name = f"BET_result_template_{template_id}"
                tag_resp = requests.post(
                    f"{ELABFTW_URL}/experiments/{exp_id}/tags",
                    headers={
                        "Authorization": ELABFTW_TOKEN,
                        "Content-Type": "application/json"
                    },
                    json={"tag": tag_name},
                    verify=self.verify_ssl
                )
                logging.debug("Tag resp: %s %s", tag_resp.status_code, tag_resp.text)
            except Exception:
                logging.exception("Tagging failed (non-fatal)")

            return jsonify({
                "ok": True,
                "exp_id": exp_id,
                "experiment_url": f"{ELABFTW_URL}/experiments/{exp_id}",
                "template_id": template_id
            }), 201

        except Exception as e:
            logging.exception("eln_create_local_json outer failure")
            return jsonify({"ok": False, "stage": "unknown", "error": str(e)}), 500

    # ─── PDF bundle → DB ──────────────────────────────────────────────
    def _insert_pdf_bundle_into_db(self, bundle: dict, original_filename: str | None = None) -> int:
        """LEGACY wrapper: PDF insertion is now handled by PdfProcessor."""
        return self.pdf_processor._insert_pdf_bundle_into_db(bundle, original_filename=original_filename)



    # ─── Run ───────────────────────────────────────────────────────────
    def run(self):
        self.app.run(host="0.0.0.0", port=5000, debug=True)


if __name__ == "__main__":
    PuraloxApp().run()
