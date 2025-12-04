# puralox/pdf_processor.py
import os
from datetime import datetime
import pandas as pd

from .db_manager import DatabaseManager
from .config import DB_NAME
from .bet_integration import extract_all_with_prints

class PdfProcessor:
    """
    PDF → DB importer for BET reports.
    Writes only to the database (no document generation).
    """

    def __init__(self):
        self.db = DatabaseManager(DB_NAME)
        self._ensure_tables()

    def _ensure_tables(self):
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS file_info (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT, date_of_measurement TEXT, time_of_measurement TEXT,
                comment1 TEXT, comment2 TEXT, comment3 TEXT, comment4 TEXT,
                serial_number TEXT, version TEXT, comment5 TEXT
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS bet_plot_columns (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER, col_index INTEGER, col_name TEXT
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS bet_data_points (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER, no INTEGER,
                p_p0 REAL, p_va_p0_p REAL, va_cc_g_stp REAL
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS bet_summaries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER, key TEXT, value TEXT
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS isotherm_points (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER, p_p0 REAL, vol_cc_g_stp REAL
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS tplot_points (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER, thickness_nm REAL, volume_cc_g_stP REAL
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS bjh_desorption (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER,
                diameter_nm REAL, porevol_ccg REAL, porearea_m2g REAL,
                dv_d REAL, ds_d REAL, dv_logd REAL, ds_logd REAL
            )
        """)

    @staticmethod
    def _dt_from_bundle(general: dict):
        dt = ""
        if isinstance(general.get("Dates"), list) and general["Dates"]:
            dt = general["Dates"][0]
        elif isinstance(general.get("Date"), str):
            dt = general["Date"]
        date_str, time_str = "", ""
        if dt:
            try:
                if " " in dt:
                    d, t = dt.split(" ", 1)
                    date_str, time_str = d.strip(), t.strip()
                else:
                    date_str = dt.strip()
            except Exception:
                pass
        return date_str, time_str

    def process_pdf(self, pdf_path: str) -> int:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = os.path.join(os.path.dirname(pdf_path), f"_pdf_parse_{stamp}")
        bundle = extract_all_with_prints(pdf_path, out_dir=out_dir, skip_docx=True)

        general = bundle.get("general", {}) or {}
        file_name = os.path.basename(pdf_path)
        date_str, time_str = self._dt_from_bundle(general)

        file_info_id = self.db.execute_returning_id(
            "INSERT INTO file_info (file_name, date_of_measurement, time_of_measurement) VALUES (?, ?, ?)",
            (file_name, date_str, time_str)
        )

        # headers for plotting
        self.db.executemany(
            "INSERT INTO bet_plot_columns (file_info_id, col_index, col_name) VALUES (?, ?, ?)",
            [
                (file_info_id, 1, "P/Po"),
                (file_info_id, 2, "Vol_cc_g_STP"),
                (file_info_id, 3, "BET_transform"),
            ]
        )

        # multipoint_bet → bet_data_points
        try:
            path = (bundle.get("csvs") or {}).get("multipoint_bet")
            if path and os.path.exists(path):
                df = pd.read_csv(path)
                rows = []
                for i, r in df.iterrows():
                    ppo = r.get("P_over_P0")
                    vol = r.get("Vol_cc_g_STP")
                    bet = r.get("BET_transform")
                    if pd.notna(ppo) and pd.notna(bet):
                        rows.append((file_info_id, i+1, float(ppo), float(bet), None if pd.isna(vol) else float(vol)))
                if rows:
                    self.db.executemany(
                        "INSERT INTO bet_data_points (file_info_id, no, p_p0, p_va_p0_p, va_cc_g_stp) VALUES (?, ?, ?, ?, ?)",
                        rows
                    )
        except Exception:
            pass

        # isotherm points
        try:
            path = (bundle.get("csvs") or {}).get("isotherm")
            if path and os.path.exists(path):
                df = pd.read_csv(path)
                rows = [(file_info_id, float(r["P_over_P0"]), float(r["Vol_cc_g_STP"])) for _, r in df.iterrows()]
                if rows:
                    self.db.executemany(
                        "INSERT INTO isotherm_points (file_info_id, p_p0, vol_cc_g_stp) VALUES (?, ?, ?)", rows)
        except Exception:
            pass

        # t-plot
        try:
            path = (bundle.get("csvs") or {}).get("tplot")
            if path and os.path.exists(path):
                df = pd.read_csv(path)
                vcol = "Volume_cc_g_STP" if "Volume_cc_g_STP" in df.columns else "Volume_cc_G_STP"
                rows = [(file_info_id, float(r["Thickness_nm"]), float(r[vcol])) for _, r in df.iterrows()]
                if rows:
                    self.db.executemany(
                        "INSERT INTO tplot_points (file_info_id, thickness_nm, volume_cc_g_stP) VALUES (?, ?, ?)", rows)
        except Exception:
            pass

        # BJH desorption
        try:
            path = (bundle.get("csvs") or {}).get("bjh")
            if path and os.path.exists(path):
                df = pd.read_csv(path)
                rows = []
                for _, r in df.iterrows():
                    rows.append((
                        file_info_id,
                        float(r["Diameter_nm"]),
                        float(r["PoreVol_ccg"]),
                        float(r["PoreArea_m2g"]),
                        float(r["dV_d"]),
                        float(r["dS_d"]),
                        float(r["dV_logd"]),
                        float(r["dS_logd"]),
                    ))
                if rows:
                    self.db.executemany(
                        """INSERT INTO bjh_desorption
                           (file_info_id, diameter_nm, porevol_ccg, porearea_m2g, dv_d, ds_d, dv_logd, ds_logd)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        rows
                    )
        except Exception:
            pass

        # summaries for ELN body
        kv = []
        for sec_name in ("isotherm_summary", "multipoint_bet_summary", "tplot_summary"):
            sec = bundle.get(sec_name) or {}
            for k, v in sec.items():
                kv.append((file_info_id, f"{sec_name}:{k}", str(v)))
        for gk in ["Operators","Dates","Sample ID","Filename","Sample Description","Comment",
                   "Sample weight","Sample Volume","Outgas Time","OutgasTemp","Analysis gas","Bath Temp"]:
            if gk in general and general[gk] not in (None, "", []):
                kv.append((file_info_id, f"general:{gk}", str(general[gk])))
        if kv:
            self.db.executemany(
                "INSERT INTO bet_summaries (file_info_id, key, value) VALUES (?, ?, ?)", kv)

        return file_info_id
