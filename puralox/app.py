#!/usr/bin/env python3
# puralox/app.py

import os
import time
import logging
import warnings

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from io import BytesIO
from datetime import datetime
from flask import Flask, request, redirect, url_for, render_template, jsonify
from dotenv import load_dotenv

import elabapi_python
from elabapi_python import ExperimentsApi
from elabapi_python.rest import ApiException

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

from .config import UPLOAD_FOLDER, DB_NAME
from .db_manager import DatabaseManager
from .excel_processor import ExcelProcessor
import requests

# ─── CONFIG ────────────────────────────────────────────────────────────
load_dotenv()
ELABFTW_URL      = "https://localhost/api/v2"
ELABFTW_TOKEN    = "6-58fefeb5b740b8334164e94dbac6faf1f52f07f10f911f43b2fa2e5f376b1e06b38be33679fbeb3d0cd66"
SCIENTIST_NAME   = "Rachit Jain"
EQUIPMENT_NAME   = "Agilent Devices"
MEASUREMENT_DESC = "BET Parameters and Data Points (Reference Below)"
# ──────────────────────────────────────────────────────────────────────

class PuraloxApp:

    def __init__(self):
        base = os.path.abspath(os.path.dirname(__file__))
        self.app = Flask(__name__, template_folder=os.path.join(base, "..", "templates"))
        self.app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

        self.db        = DatabaseManager(DB_NAME)
        self.processor = ExcelProcessor(self.db)
        self._configure_elabftw()
        self._register_routes()

    def _configure_elabftw(self):
        cfg = elabapi_python.Configuration()
        cfg.host       = ELABFTW_URL
        cfg.verify_ssl = False
        client = elabapi_python.ApiClient(cfg)
        client.set_default_header("Authorization", ELABFTW_TOKEN)

        self.exp_api    = ExperimentsApi(client)
        self.verify_ssl = cfg.verify_ssl
        warnings.filterwarnings("ignore", category=Warning, module="urllib3")

        for attempt in range(1, 6):
            try:
                self.exp_api.read_experiments(limit=1)
                print("✅ eLabFTW API OK.")
                return
            except Exception as e:
                if attempt == 5:
                    logging.error(f"Cannot connect to eLabFTW: {e}")
                    raise
                time.sleep(2 ** attempt)

    def _register_routes(self):
        @self.app.route("/", methods=["GET", "POST"])
        def upload():
            if request.method == "POST":
                f = request.files.get("file")
                if f and f.filename.endswith(".xlsx"):
                    dst = os.path.join(self.app.config["UPLOAD_FOLDER"], f.filename)
                    f.save(dst)
                    self.processor.process_file(dst)
                    return redirect(url_for("list_files"))
            return render_template("upload.html")

        @self.app.route("/files")
        def list_files():
            df = pd.DataFrame(self.db.fetchall_dict("SELECT * FROM file_info"))
            return render_template("list_files.html", df=df)

        @self.app.route("/file/<int:file_id>")
        def view_file(file_id):
            to_df = lambda q: pd.DataFrame(self.db.fetchall_dict(q))
            return render_template("view_file.html",
                file_id  = file_id,
                info     = to_df(f"SELECT * FROM file_info WHERE id={file_id}"),
                bet      = to_df(f"SELECT * FROM bet_parameters WHERE file_info_id={file_id}"),
                tech     = to_df(f"SELECT * FROM technical_info WHERE file_info_id={file_id}"),
                columns  = to_df(f"SELECT * FROM bet_plot_columns WHERE file_info_id={file_id}"),
                points   = to_df(f"SELECT * FROM bet_data_points WHERE file_info_id={file_id}")
            )

        @self.app.route("/push/<int:file_id>", methods=["POST"])
        def push_to_elab(file_id):
            # 1) fetch data
            fi   = self.db.fetchall_dict(f"SELECT * FROM file_info WHERE id={file_id}")[0]
            bet  = self.db.fetchall_dict(f"SELECT * FROM bet_parameters WHERE file_info_id={file_id}")
            tech = self.db.fetchall_dict(f"SELECT * FROM technical_info WHERE file_info_id={file_id}")
            pts  = self.db.fetchall_dict(f"SELECT * FROM bet_data_points WHERE file_info_id={file_id}")
            equipment = fi.get("comment4") or EQUIPMENT_NAME

            # 2) build HTML body
            html  = f"<h1>Experiment: {fi['file_name']}</h1>"
            html += f"<p><strong>ID:</strong> {file_id}</p>"
            html += f"<p><strong>Date:</strong> {fi['date_of_measurement']} {fi['time_of_measurement']}</p>"
            html += f"<p><strong>Serial #:</strong> {fi['serial_number']}</p>"
            html += f"<p><strong>Version:</strong> {fi.get('version','—')}</p>"
            html += f"<p><strong>Scientist:</strong> {SCIENTIST_NAME}</p>"
            html += "<h2>Goal</h2>"
            html += f"<p>The goal of the experiment <b>{fi['file_name']}</b> referred to the project BET Data Analysis. " \
                    f"It was done by <b>{SCIENTIST_NAME}</b> on <b>{fi['date_of_measurement']} {fi['time_of_measurement']}</b>.</p>"
            html += "<h2>Work Flow</h2>"
            html += f"<p>Executed on <b>{EQUIPMENT_NAME}</b> to measure <b>{MEASUREMENT_DESC}</b>.</p>"
            html += "<h2>Results</h2>"
            html += "<p>Shows BET parameters, technical info, and sample data points below.</p>"

            html += "<h3>BET Parameters</h3><ul>"
            for row in bet:
                for k,v in row.items():
                    html += f"<li>{k}: {v}</li>"
            html += "</ul>"

            html += "<h3>Technical Info</h3><ul>"
            for row in tech:
                for k,v in row.items():
                    html += f"<li>{k}: {v}</li>"
            html += "</ul>"

            html += "<h3>Data Points (first 10)</h3><table border='1'><tr><th>No</th><th>p/p0</th><th>p/va_p0_p</th></tr>"
            for r in pts[:10]:
                html += f"<tr><td>{r['no']}</td><td>{r['p_p0']}</td><td>{r['p_va_p0_p']}</td></tr>"
            html += "</table>"

            # 3) create empty experiment
            _, status, headers = self.exp_api.post_experiment_with_http_info(body={})
            if status != 201:
                return jsonify(error=f"Create failed ({status})"), 500
            exp_id = headers["Location"].rstrip("/").split("/")[-1]

            # 4) patch title and body
            html = html.replace(f"<strong>ID:</strong> {file_id}", f"<strong>ID:</strong> {exp_id}")
            self.exp_api.patch_experiment(exp_id, body={"title": fi["file_name"], "body": html})

            # 5) generate plot
            x = np.array([r['p_p0']       for r in pts])
            y = np.array([r['p_va_p0_p']   for r in pts])
            coeffs = np.polyfit(x, y, 1)
            fit_fn = np.poly1d(coeffs)
            fig, ax = plt.subplots()
            ax.scatter(x, y, s=20)
            ax.plot(x, fit_fn(x), linestyle='--')
            ax.set_title("BET Plot")
            ax.set_xlabel("p/p0")
            ax.set_ylabel("p/va_p0_p")
            img_buf = BytesIO()
            fig.savefig(img_buf, format="PNG", bbox_inches="tight")
            plt.close(fig)
            img_buf.seek(0)

            # 6) build PDF with only image
            pdf_buf = BytesIO()
            doc = SimpleDocTemplate(pdf_buf, pagesize=letter)
            doc.build([ RLImage(img_buf, width=400, height=300) ])
            pdf_buf.seek(0)

            # 7) upload PDF
            files = {'file': (f"BET_Plot_{fi['file_name']}.pdf", pdf_buf.read(), 'application/pdf')}
            attach_url = f"{ELABFTW_URL}/experiments/{exp_id}/uploads"
            resp = requests.post(attach_url, headers={"Authorization": ELABFTW_TOKEN}, files=files, verify=self.verify_ssl)
            if not resp.ok:
                return jsonify(error="PDF upload failed", status=resp.status_code, resp=resp.text), 500

            # 8) tag
            tag_url = f"{ELABFTW_URL}/experiments/{exp_id}/tags"
            requests.post(tag_url, headers={"Authorization": ELABFTW_TOKEN, "Content-Type":"application/json"},
                          json={"tag":"BET_result"}, verify=self.verify_ssl)

            # 9) return JSON
            new = self.exp_api.get_experiment(exp_id)
            return jsonify(new.to_dict()), 201

        @self.app.route("/api/elab/experiments", methods=["GET"])
        def fetch_elab_experiments():
            try:
                exps = self.exp_api.read_experiments(limit=10)
                return jsonify([e.to_dict() for e in exps]), 200
            except ApiException as e:
                return jsonify(error=str(e)), 500

        @self.app.route("/api")
        def api_info():
            return render_template("api.html")

    def run(self):
        self.app.run(host="0.0.0.0", port=5000, debug=True)


if __name__ == "__main__":
    PuraloxApp().run()
