#!/usr/bin/env python3
# puralox/app.py

import os
import logging
import warnings

import numpy as np
import pandas as pd

from flask import (
    Flask, request, redirect, url_for,
    render_template, jsonify, send_from_directory
)
from dotenv import load_dotenv

from elabapi_python.rest import ApiException

from .config import UPLOAD_FOLDER, DB_NAME, ELABFTW_URL, ELABFTW_TOKEN
from .database_manager import DatabaseManager
from .excel_processor import ExcelProcessor
from .pdf_processor import PdfProcessor
from .template_processor import TemplateProcessor
from .metadata_builder import MetadataBuilder
from .eln_client import ElnClient

# ─── LOGGING ───────────────────────────────────────────────────────────
load_dotenv()
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

        # ── Core services (dependency injection) ───────────────────────
        self.db               = DatabaseManager(DB_NAME)
        self.processor        = ExcelProcessor(self.db)
        self.pdf_processor    = PdfProcessor(self.db, self.app.config["UPLOAD_FOLDER"])
        self.template_processor = TemplateProcessor(self.db)
        self.metadata_builder = MetadataBuilder(self.db, self.metadata_dir)

        # ── ELN integration client ──────────────────────────────────────
        disable_ssl = os.getenv("ELABFTW_DISABLE_SSL", "true").lower() == "true"
        self.eln_client = ElnClient(ELABFTW_URL, ELABFTW_TOKEN, disable_ssl=disable_ssl)

        self._ensure_optional_tables()
        self._register_routes()

    # ─── DB schema helpers ─────────────────────────────────────────────
    def _ensure_optional_tables(self):
        """Create tables that are not part of the base schema."""
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

        # Ensure comment5 column exists on file_info (for Measurement ID)
        try:
            cols = self.db.fetchall_dict("PRAGMA table_info(file_info)")
            colnames = {c["name"] for c in cols}
            if "comment5" not in colnames:
                logging.info("Adding comment5 column to file_info")
                self.db.execute("ALTER TABLE file_info ADD COLUMN comment5 TEXT")
        except Exception:
            logging.exception("Failed to ensure comment5 column on file_info")

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
                        new_file_id = self.processor.import_file(dst, original_filename=f.filename)
                    elif ext == ".pdf":
                        new_file_id = self.pdf_processor.import_file(dst, original_filename=f.filename)
                    else:
                        return "Unsupported file type. Upload .xlsx or .pdf", 400

                    try:
                        self.metadata_builder.generate_metadata(new_file_id)
                    except Exception:
                        logging.exception("metadata generation failed")

                    return redirect(url_for("list_files"))

                except Exception:
                    logging.exception("Upload processing failed")
                    return jsonify({"error": "Failed to process file"}), 500

            return render_template("upload.html")

        # ── Unified list page ─────────────────────────────────────────
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
            to_df = lambda q: pd.DataFrame(self.db.fetchall_dict(q))

            info_df = to_df(f"SELECT * FROM file_info WHERE id={file_id}")
            bet_df  = to_df(f"SELECT * FROM bet_parameters WHERE file_info_id={file_id}")
            tech_df = to_df(f"SELECT * FROM technical_info WHERE file_info_id={file_id}")
            cols_df = to_df(f"SELECT * FROM bet_plot_columns WHERE file_info_id={file_id}")
            pts_df  = to_df(f"SELECT * FROM bet_data_points WHERE file_info_id={file_id}")

            info = info_df.iloc[0].to_dict() if not info_df.empty else {}
            bet  = bet_df.iloc[0].to_dict() if not bet_df.empty else {}
            tech = tech_df.iloc[0].to_dict() if not tech_df.empty else {}
            cols = cols_df.to_dict(orient="records") if not cols_df.empty else []
            pts_rows = pts_df.to_dict(orient="records") if not pts_df.empty else []

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
            sample_region = "KIT Campus South"

            return render_template(
                "view_excel.html",
                file_id=file_id,
                info=info,
                bet=bet,
                tech=tech,
                cols=cols,
                pts=pts_rows,
                pts_data=pts_data,
                metadata_url=md_url,
                sample_region=sample_region,
            )

        # ── PDF detail view ───────────────────────────────────────────
        @self.app.route("/view/pdf/<int:file_id>")
        def view_pdf(file_id: int):
            fi_rows = self.db.fetchall_dict(
                "SELECT * FROM file_info WHERE id=?", (file_id,)
            )
            fi = fi_rows[0] if fi_rows else {}

            pdf_filename = fi.get("file_name") or f"file_{file_id}.pdf"

            cnt_rows = self.db.fetchall_dict(
                "SELECT COUNT(*) AS c FROM bet_data_points WHERE file_info_id=?",
                (file_id,),
            )
            points_count = cnt_rows[0]["c"] if cnt_rows else 0

            kv_rows = []
            if self.db.table_exists("bet_summaries"):
                kv_rows = self.db.fetchall_dict(
                    "SELECT key, value FROM bet_summaries "
                    "WHERE file_info_id=? ORDER BY key",
                    (file_id,),
                )

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

            core_fields  = [{"key": r["key"], "value": r["value"]} for r in kv_rows if r["key"] in core_keys]
            extra_fields = [{"key": r["key"], "value": r["value"]} for r in kv_rows if r["key"] not in core_keys]

            logging.debug(
                "PDF view for file_id=%s: kv_rows=%s, core_fields=%s, extra_fields=%s, bet_present=%s",
                file_id, len(kv_rows), len(core_fields), len(extra_fields), bool(points_count),
            )

            pts_rows = self.db.fetchall_dict(
                "SELECT no, p_p0, p_va_p0_p FROM bet_data_points "
                "WHERE file_info_id=? ORDER BY no",
                (file_id,),
            )
            pts_data = pts_rows

            x_vals = [r["p_p0"] for r in pts_rows if r["p_p0"] is not None]
            y_vals = [r["p_va_p0_p"] for r in pts_rows if r["p_va_p0_p"] is not None]

            measurement_id = fi.get("comment5") or fi.get("file_name") or f"File {file_id}"
            header = {
                "measurement_id": measurement_id,
                "date":          fi.get("date_of_measurement", ""),
                "time":          fi.get("time_of_measurement", ""),
                "operator":      fi.get("comment2", ""),
                "instrument":    fi.get("comment4", ""),
                "serial_number": fi.get("serial_number", ""),
                "version":       fi.get("version", ""),
            }
            bundle = {
                "file_id":      file_id,
                "points_count": points_count,
                "summaries":    [{"key": r["key"], "value": r["value"]} for r in kv_rows],
            }

            md_url = url_for("download_metadata_for_file", file_id=file_id)
            sample_region = "KIT Campus Nord"

            return render_template(
                "view_pdf_extract.html",
                pdf_filename=pdf_filename,
                header=header,
                bundle=bundle,
                file_info=fi,
                pts_data=pts_data,
                metadata_url=md_url,
                sample_region=sample_region,
                default_x_min=min(x_vals) if x_vals else None,
                default_x_max=max(x_vals) if x_vals else None,
                default_y_min=min(y_vals) if y_vals else None,
                default_y_max=max(y_vals) if y_vals else None,
                core_fields=core_fields,
                extra_fields=extra_fields,
            )

        # ── ELN push ──────────────────────────────────────────────────
        @self.app.route("/eln/<int:file_id>", methods=["POST"])
        def eln_create(file_id: int):
            return self._eln_push(file_id, template_id=None)

        @self.app.route("/eln/<int:file_id>/<int:template_id>", methods=["POST"])
        def eln_create_legacy(file_id: int, template_id: int):
            logging.info("ELN push (legacy URL) with template_id=%s", template_id)
            return self._eln_push(file_id, template_id=template_id)

        # ── Metadata download ─────────────────────────────────────────
        @self.app.route("/metadata/<filename>")
        def download_metadata(filename):
            return send_from_directory(self.metadata_dir, filename, as_attachment=True)

        @self.app.route("/metadata/file/<int:file_id>")
        def download_metadata_for_file(file_id: int):
            md_path = self.metadata_builder.generate_metadata(file_id)
            fname = os.path.basename(md_path)
            return send_from_directory(self.metadata_dir, fname, as_attachment=True)

        # ── API info ──────────────────────────────────────────────────
        @self.app.route("/api")
        def api_info():
            return render_template("api.html")

        @self.app.route("/api/elab/templates")
        def api_elab_templates():
            try:
                tpls = self.eln_client.fetch_templates()
                out = [{"id": t["id"], "title": t.get("title", f"Template {t['id']}")} for t in tpls]
                return jsonify(out), 200
            except Exception as e:
                logging.exception("api_elab_templates failed")
                return jsonify({"error": str(e)}), 500

        @self.app.route("/api/elab/experiments")
        def api_elab_experiments():
            try:
                return jsonify(self.eln_client.fetch_experiments(limit=10)), 200
            except ApiException as e:
                return jsonify({"error": "API Exception", "details": str(e)}), 500

    # ─── ELN push core ─────────────────────────────────────────────────
    def _eln_push(self, file_id: int, template_id: int | None = None):
        """Orchestrate an ELN push: build HTML, create experiment, upload plot, tag."""
        try:
            # ---- Resolve template_id from form (optional) ----
            form_tid = request.form.get("template_id")
            if form_tid and str(form_tid).strip():
                try:
                    template_id = int(form_tid)
                except ValueError:
                    pass

            title = request.form.get("title") or f"File {file_id}"

            # ---- Optional plot range from UI ----
            def _parse_float(key):
                val = request.form.get(key)
                if val in (None, "", "null"):
                    return None
                try:
                    return float(val)
                except ValueError:
                    return None

            plot_xmin = _parse_float("plot_xmin")
            plot_xmax = _parse_float("plot_xmax")

            logging.info(
                "ELN push for file_id=%s with template_id=%s (range: %s – %s)",
                file_id, template_id, plot_xmin, plot_xmax,
            )

            # ---- Verify file exists ----
            fi_rows = self.db.fetchall_dict("SELECT * FROM file_info WHERE id=?", (file_id,))
            if not fi_rows:
                return jsonify({"ok": False, "stage": "load", "error": "file_info not found"}), 404
            fi = fi_rows[0]

            # ---- Build ELN HTML via TemplateProcessor ----
            html_body = self.template_processor.build_eln_html(file_id)

            measurement_id = fi.get("comment5") or fi.get("file_name") or f"file_{file_id}"
            logging.info("[ELN] Measurement ID: %s | HTML preview: %.200s", measurement_id, html_body)

            # ---- Create experiment via ElnClient ----
            try:
                exp_id = self.eln_client.create_experiment(title, html_body)
            except Exception as e:
                logging.exception("eln_client.create_experiment failed")
                return jsonify({"ok": False, "stage": "create", "error": str(e)}), 500

            # ---- Upload BET plot ----
            pts_for_plot = self.db.fetchall_dict(
                "SELECT p_p0, p_va_p0_p FROM bet_data_points WHERE file_info_id=?",
                (file_id,),
            )
            self.eln_client.upload_plot(exp_id, file_id, pts_for_plot, x_min=plot_xmin, x_max=plot_xmax)

            # ---- Add tag ----
            tag_name = f"BET_result_template_{template_id}" if template_id is not None else "BET_result"
            self.eln_client.add_tag(exp_id, tag_name)

            return jsonify({
                "ok": True,
                "exp_id": exp_id,
                "experiment_url": f"{ELABFTW_URL}/experiments/{exp_id}",
                "template_id": template_id,
            }), 201

        except Exception as e:
            logging.exception("_eln_push outer failure")
            return jsonify({"ok": False, "stage": "unknown", "error": str(e)}), 500

    # ─── Run ───────────────────────────────────────────────────────────
    def run(self):
        self.app.run(host="0.0.0.0", port=5000, debug=True)


if __name__ == "__main__":
    PuraloxApp().run()
