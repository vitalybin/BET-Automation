#!/usr/bin/env python3
# puralox/app.py

import os
import re
import time
import logging
import warnings

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from io import BytesIO
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
from .bet_integration import extract_all_with_prints
import requests

# ─── CONFIG ────────────────────────────────────────────────────────────
load_dotenv()
ELABFTW_URL   = "https://localhost/api/v2"
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

    def _table_exists(self, name: str) -> bool:
        rows = self.db.fetchall_dict(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (name,)
        )
        return bool(rows)

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
        cfg.verify_ssl = False
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
                        # Existing Excel import flow
                        new_file_id = self.processor.process_file(dst)

                    elif ext == ".pdf":
                        # PDF → parse → DB
                        out_dir = os.path.join(self.app.config["UPLOAD_FOLDER"], "pdf_out")
                        os.makedirs(out_dir, exist_ok=True)
                        bundle = extract_all_with_prints(dst, out_dir)
                        new_file_id = self._insert_pdf_bundle_into_db(bundle)

                    else:
                        return "Unsupported file type. Upload .xlsx or .pdf", 400

                    # Create unified metadata for both Excel + PDF
                    try:
                        self._generate_metadata(new_file_id)
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
                "SELECT id, file_name, date_of_measurement, time_of_measurement "
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
            info = to_df(f"SELECT * FROM file_info WHERE id={file_id}")
            bet  = to_df(f"SELECT * FROM bet_parameters WHERE file_info_id={file_id}")
            tech = to_df(f"SELECT * FROM technical_info WHERE file_info_id={file_id}")
            cols = to_df(f"SELECT * FROM bet_plot_columns WHERE file_info_id={file_id}")
            pts  = to_df(f"SELECT * FROM bet_data_points WHERE file_info_id={file_id}")
            md_url = url_for("download_metadata_for_file", file_id=file_id)
            return render_template(
                "view_excel.html",
                file_id=file_id,
                info=info,
                bet=bet,
                tech=tech,
                cols=cols,
                pts=pts,
                metadata_url=md_url
            )

        # ── PDF detail view (summary) ─────────────────────────────────
        @self.app.route("/view/pdf/<int:file_id>")
        def view_pdf(file_id: int):
            rows = self.db.fetchall_dict(
                "SELECT file_name FROM file_info WHERE id=?",
                (file_id,)
            )
            pdf_filename = rows[0]["file_name"] if rows else f"file_{file_id}.pdf"

            cnt_rows = self.db.fetchall_dict(
                "SELECT COUNT(*) AS c FROM bet_data_points WHERE file_info_id=?",
                (file_id,)
            )
            points_count = cnt_rows[0]["c"] if cnt_rows else 0

            summaries = []
            if self._table_exists("bet_summaries"):
                kv_rows = self.db.fetchall_dict(
                    "SELECT key, value FROM bet_summaries WHERE file_info_id=? ORDER BY key",
                    (file_id,)
                )
                summaries = [{"key": r["key"], "value": r["value"]} for r in kv_rows]

            bundle = {
                "file_id": file_id,
                "points_count": points_count,
                "summaries": summaries
            }
            md_url = url_for("download_metadata_for_file", file_id=file_id)
            return render_template(
                "view_pdf_extract.html",
                pdf_filename=pdf_filename,
                bundle=bundle,
                metadata_url=md_url
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
            md_path = self._generate_metadata(file_id)
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
        """
        Creates an ELN experiment for given file_id.

        Accepts an optional template_id (from URL or form), but currently
        does NOT change the generated content. It's just logged so we can
        use it later.

        Always returns JSON:
          { "ok": True, "exp_id": "...", "experiment_url": "..." }
          or
          { "ok": False, "error": "...", "stage": "..." }
        """
        try:
            # ---- template_id (optional, informational) ----
            form_tid = request.form.get("template_id")
            if form_tid and str(form_tid).strip():
                try:
                    template_id = int(form_tid)
                except ValueError:
                    # ignore bad input, keep existing template_id
                    pass

            logging.info(
                "ELN push for file_id=%s with template_id=%s (currently ignored for content)",
                file_id, template_id
            )

            title = request.form.get("title") or f"File {file_id}"

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

            # ---- Derive some fields ----
            measurement_id = fi.get("comment5") or fi.get("file_name") or f"file_{file_id}"
            scientist = fi.get("comment2") or "—"

            m_sample = re.search(r"(\d{4}-\d{4})", measurement_id)
            sample_id = m_sample.group(1) if m_sample else "—"

            comment3 = fi.get("comment3") or ""
            parts = [p.strip() for p in comment3.split(",")] if comment3 else []
            temp = parts[0] if len(parts) > 0 else "—"
            duration = parts[1] if len(parts) > 1 else "—"
            env = parts[2] if len(parts) > 2 else "—"

            try:
                pvals = [r["p_p0"] for r in pts if r["p_p0"] is not None]
                pmin = min(pvals) if pvals else "—"
                pmax = max(pvals) if pvals else "—"
            except Exception:
                pmin = pmax = "—"

            mass = tech.get("mass", "—")
            internal_device_id = tech.get("internal_device_id", "—")

            specific_surf_area = bet.get("as_bet", "—") or bet.get("Specific surface area", "—")
            pore_volume = bet.get("total_pore_volume", "—") or bet.get("Total pore volume", "—")

            # ---- Small HTML tables ----
            def make_table(data_rows):
                if not data_rows:
                    return "<p>No data</p>"
                headers = list(data_rows[0].keys())
                html_rows = [
                    "<tr>" + "".join(f"<th>{h}</th>" for h in headers) + "</tr>"
                ]
                for row in data_rows:
                    html_rows.append(
                        "<tr>" + "".join(f"<td>{row.get(h,'')}</td>" for h in headers) + "</tr>"
                    )
                return (
                    "<table border='1' cellspacing='0' cellpadding='4'>"
                    + "".join(html_rows)
                    + "</table>"
                )

            bet_table_html = make_table(bet_rows[:1]) if bet_rows else "<p>No BET parameters</p>"
            pts_table_html = make_table(pts[:15]) if pts else "<p>No BET data points</p>"

            # ---- Final HTML body for ELN ----
            html_body = f"""
            <h1>Experiment: {measurement_id}</h1>

            <h2>🧪 Experiment Meta Data</h2>
            <ul>
              <li><strong>Measurement ID:</strong> {measurement_id}</li>
              <li><strong>Date:</strong> {fi.get('date_of_measurement','')} {fi.get('time_of_measurement','')}</li>
              <li><strong>Operator:</strong> {fi.get('comment2','')}</li>
              <li><strong>Instrument:</strong> {fi.get('comment4','')}</li>
              <li><strong>Internal Device ID:</strong> {internal_device_id}</li>
              <li><strong>Serial #:</strong> {fi.get('serial_number','')}</li>
              <li><strong>Version:</strong> {fi.get('version','')}</li>
              <li><strong>Scientist (Sample Preparation):</strong> {scientist}</li>
              <li><strong>Sample ID:</strong> {sample_id}</li>
              <li><strong>Pretreatment conditions:</strong> {comment3}</li>
            </ul>

            <h2>🧬 Experimental Procedure</h2>
            <p>
              {mass} g of the sample <strong>{scientist}_{sample_id}</strong>
              were pretreated at <strong>{temp}</strong> for <strong>{duration}</strong>
              in <strong>{env}</strong>.<br>
              For the evaluation of the BET isotherm, <strong>{len(pts)}</strong> points
              in a pressure range of <strong>{pmin}</strong> to <strong>{pmax}</strong> were considered.
            </p>

            <h2>📊 Results</h2>
            <p>
              The sample exhibited a specific surface area of
              <strong>{specific_surf_area}</strong> and a pore volume of
              <strong>{pore_volume}</strong>.
            </p>

            <h3>📌 BET Parameters</h3>
            {bet_table_html}

            <h3>📈 BET Data Points (First 15)</h3>
            {pts_table_html}
            """

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

            # ---- Attach BET plot (best-effort) ----
            pts_for_plot = self.db.fetchall_dict(
                "SELECT p_p0, p_va_p0_p FROM bet_data_points WHERE file_info_id=?",
                (file_id,)
            )
            if pts_for_plot:
                try:
                    x = np.array([r["p_p0"] for r in pts_for_plot if r["p_p0"] is not None])
                    y = np.array([r["p_va_p0_p"] for r in pts_for_plot if r["p_va_p0_p"] is not None])
                    if len(x) and len(y):
                        fig, ax = plt.subplots()
                        ax.scatter(x, y, s=20)
                        if len(x) >= 2:
                            fit = np.polyfit(x, y, 1)
                            ax.plot(x, np.poly1d(fit)(x), linestyle="--")
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
                        logging.debug("Upload plot resp: %s %s", resp.status_code, resp.text)
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

    # ─── Metadata Excel (unified Excel + PDF) ────────────────────────
    def _summary_map(self, file_id: int) -> dict:
        if not self._table_exists("bet_summaries"):
            return {}
        rows = self.db.fetchall_dict(
            "SELECT key, value FROM bet_summaries WHERE file_info_id=?",
            (file_id,)
        )
        return {r["key"]: r["value"] for r in rows}

    def _generate_metadata(self, file_id: int):
        fi_rows = self.db.fetchall_dict("SELECT * FROM file_info WHERE id=?", (file_id,))
        if not fi_rows:
            raise RuntimeError(f"file_info row not found for id={file_id}")
        fi = fi_rows[0]
        bet_rows = self.db.fetchall_dict(
            "SELECT * FROM bet_parameters WHERE file_info_id=?", (file_id,)
        )
        bet = bet_rows[0] if bet_rows else {}
        pts = self.db.fetchall_dict(
            "SELECT p_p0 FROM bet_data_points WHERE file_info_id=?",
            (file_id,)
        )
        kv  = self._summary_map(file_id)

        # P/Po stats
        try:
            pvals = [r["p_p0"] for r in pts if r["p_p0"] is not None]
            pmin = min(pvals) if pvals else ""
            pmax = max(pvals) if pvals else ""
            pcount = len(pvals)
        except Exception:
            pmin = pmax = ""
            pcount = 0

        def pick(*candidates):
            for c in candidates:
                if c is None:
                    continue
                s = str(c).strip()
                if s and s.lower() != "none":
                    return s
            return ""

        sample_weight   = pick(bet.get("sample_weight"), kv.get("general:Sample weight"))
        adsorptive      = pick(bet.get("adsorptive"), kv.get("general:Analysis gas"))
        apparatus_temp  = pick(bet.get("apparatus_temperature"), kv.get("general:Bath Temp"))
        adsorption_temp = pick(bet.get("adsorption_temperature"), kv.get("general:OutgasTemp"))

        specific_surface_area = pick(
            bet.get("as_bet"),
            kv.get("multipoint_bet_summary: Surface Area"),
            kv.get("isotherm_summary: Surface Area"),
            kv.get("isotherm_summary:Surface Area"),
        )
        total_pore_volume = pick(
            bet.get("total_pore_volume"),
            kv.get("tplot_summary: Pore Volume"),
            kv.get("tplot_summary:Pore Volume"),
        )
        avg_pore_diam = pick(
            bet.get("average_pore_diameter"),
            kv.get("tplot_summary: Pore Diameter Dv(d)"),
            kv.get("tplot_summary:Pore Diameter Dv(d)"),
        )

        md = {
            "File Name":                      fi.get("file_name", ""),
            "timestamp":                      f"{fi.get('date_of_measurement','')}T{fi.get('time_of_measurement','')}Z",
            "sample":                         fi.get("file_name", ""),
            "operator of experiment":         pick(fi.get("comment2"), kv.get("general:OperatorPrimary"), kv.get("general:Operators")),
            "pretreatment conditions":        fi.get("comment3", ""),
            "manufacturer":                   "BEL, BEL",
            "measurement device":             pick(fi.get("comment4"), kv.get("general:Instrument"), "BELprep II, BELsorp II mini instrument"),
            "serial number":                  fi.get("serial_number", ""),
            "Version":                        fi.get("version", ""),

            "Sample weight [g]":              sample_weight,
            "Standard volume [cm3]":          pick(bet.get("standard_volume")),
            "Dead volume [cm3]":              pick(bet.get("dead_volume")),
            "Equilibrium time [sec]":         pick(bet.get("equilibrium_time")),
            "Adsorptive":                     adsorptive,
            "Apparatus temperature [C]":      apparatus_temp,
            "Adsorption temperature [K]":     adsorption_temp,

            "BET points (count)":             pcount,
            "BET P/Po min":                   pmin,
            "BET P/Po max":                   pmax,

            "Specific surface area":          specific_surface_area,
            "Total pore volume":              total_pore_volume,
            "Average pore diameter":          avg_pore_diam,
        }

        df = pd.DataFrame(md.items(), columns=["Field", "Value"])
        base = os.path.splitext(fi.get("file_name", "file"))[0] or f"file_{file_id}"
        md_file = os.path.join(self.metadata_dir, f"{base}_metadata.xlsx")
        df.to_excel(md_file, index=False)
        logging.info("Unified Metadata Excel written to %s", md_file)
        return md_file

    # ─── PDF bundle → DB ──────────────────────────────────────────────
    def _insert_pdf_bundle_into_db(self, bundle: dict) -> int:
        gen = (bundle.get("general") or {})
        iso = (bundle.get("isotherm_summary") or {})
        mp  = (bundle.get("multipoint_bet_summary") or {})
        tp  = (bundle.get("tplot_summary") or {})

        file_name = gen.get("Filename") or gen.get("Sample ID") or "BET_PDF_Report.pdf"

        date_str, time_str = "", ""
        if isinstance(gen.get("Dates"), list) and gen["Dates"]:
            date_full = str(gen["Dates"][0])
            parts = date_full.split()
            if len(parts) >= 1:
                date_str = parts[0]
            if len(parts) >= 2:
                time_str = parts[1]

        # operator: pick one name only
        operator_primary = gen.get("OperatorPrimary") or \
            (gen.get("Operators")[0] if isinstance(gen.get("Operators"), list) and gen["Operators"] else "")

        file_info = {
            "file_name": file_name,
            "date_of_measurement": date_str,
            "time_of_measurement": time_str,
            "comment1": gen.get("Comment", ""),
            "comment2": operator_primary,
            "comment3": "",
            "comment4": gen.get("Instrument", ""),
            "serial_number": "",
            "version": ""
        }
        fid = self.db.execute(
            """
            INSERT INTO file_info
               (file_name, date_of_measurement, time_of_measurement,
                comment1, comment2, comment3, comment4, serial_number, version)
            VALUES (:file_name, :date_of_measurement, :time_of_measurement,
                    :comment1, :comment2, :comment3, :comment4, :serial_number, :version)
            """,
            file_info
        )

        def _num(s):
            if not s:
                return None
            m = re.search(r"[-+]?\d+(\.\d+)?", str(s))
            return float(m.group(0)) if m else None

        bet_params = {
            "file_info_id": fid,
            "sample_weight": _num(gen.get("Sample weight")),
            "standard_volume": None,
            "dead_volume": None,
            "equilibrium_time": None,
            "adsorptive": gen.get("Analysis gas", ""),
            "apparatus_temperature": None,
            "adsorption_temperature": None,
            "starting_point": None,
            "end_point": None,
            "slore": _num(iso.get("Isotherm Slope")),
            "intercept": _num(iso.get("Isotherm Intercept")),
            "correlation_coefficient": _num(iso.get("Isotherm r")),
            "vm": None,
            "as_bet": _num(iso.get("Surface Area") or mp.get("Surface Area")),
            "c_value": _num(iso.get("C constant")),
            "total_pore_volume": _num(tp.get("Pore Volume")),
            "average_pore_diameter": _num(tp.get("Pore Diameter Dv(d)"))
        }
        cols = ", ".join(bet_params.keys())
        ph   = ", ".join(":" + k for k in bet_params)
        self.db.execute(f"INSERT INTO bet_parameters ({cols}) VALUES ({ph})", bet_params)

        # Save summaries (for PDF detail view + metadata)
        def _insert_summary(prefix: str, dct: dict):
            if not dct:
                return
            for k, v in dct.items():
                self.db.execute(
                    "INSERT INTO bet_summaries (file_info_id, key, value) VALUES (?, ?, ?)",
                    (fid, f"{prefix}:{k}", str(v) if v is not None else None)
                )
        _insert_summary("general", gen)
        _insert_summary("isotherm_summary", iso)
        _insert_summary("multipoint_bet_summary", mp)
        _insert_summary("tplot_summary", tp)

        # Insert isotherm points from CSV (if available)
        csvs = (bundle.get("csvs") or {})
        iso_csv = csvs.get("isotherm")
        if iso_csv and os.path.exists(iso_csv):
            try:
                df_pts = pd.read_csv(iso_csv)
                for idx, r in df_pts.iterrows():
                    ppo = float(r.get("P_over_P0"))
                    vol = float(r.get("Vol_cc_g_STP"))
                    self.db.execute(
                        "INSERT INTO bet_data_points (file_info_id, no, p_p0, p_va_p0_p) "
                        "VALUES (?, ?, ?, ?)",
                        (fid, idx + 1, ppo, vol)
                    )
                # Plot column names
                self.db.execute(
                    "INSERT INTO bet_plot_columns (file_info_id, col_index, col_name) VALUES (?, ?, ?)",
                    (fid, 1, "P_over_P0")
                )
                self.db.execute(
                    "INSERT INTO bet_plot_columns (file_info_id, col_index, col_name) VALUES (?, ?, ?)",
                    (fid, 2, "Vol_cc_g_STP")
                )
            except Exception as e:
                logging.exception("Failed to import isotherm CSV into DB: %s", e)

        # minimal technical_info row
        self.db.execute(
            """
            INSERT INTO technical_info
               (file_info_id, saturated_vapor_pressure, adsorption_cross_section,
                wall_adsorption_correction1, wall_adsorption_correction2,
                num_adsorption_points, num_desorption_points)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (fid, None, None, "", "", None, None)
        )
        return fid

    # ─── Run ───────────────────────────────────────────────────────────
    def run(self):
        self.app.run(host="0.0.0.0", port=5000, debug=True)


if __name__ == "__main__":
    PuraloxApp().run()
