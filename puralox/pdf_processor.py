# puralox/pdf_processor.py
import os
import re
import logging
import pandas as pd

from .db_manager import DatabaseManager
from .bet_integration import extract_all_with_prints
from .nomenclature import build_measurement_id
from .excel_processor import BaseImporter


class PdfProcessor(BaseImporter):
    """
    PDF → DB importer for BET reports.

    Refactor goal:
      - Make PDF import a real, reusable class used by PuraloxApp (UML matches runtime).
      - Keep the existing PDF functionality the same as the previous app.py implementation.
      - Support dependency injection (share the same DatabaseManager instance).
    """

    def __init__(self, db: DatabaseManager, upload_folder: str):
        self.db = db
        self.upload_folder = upload_folder

    def import_file(self, pdf_path: str, original_filename: str | None = None) -> int:
        out_dir = os.path.join(self.upload_folder, "pdf_out")
        os.makedirs(out_dir, exist_ok=True)

        bundle = extract_all_with_prints(pdf_path, out_dir=out_dir)

        return self._insert_pdf_bundle_into_db(
            bundle=bundle,
            original_filename=original_filename or os.path.basename(pdf_path),
        )

    # Backward compatible alias (if anything calls PdfProcessor.process_pdf)
    def process_pdf(self, pdf_path: str) -> int:
        return self.import_file(pdf_path, original_filename=os.path.basename(pdf_path))

    # ---- Core insertion logic (ported from the old PuraloxApp._insert_pdf_bundle_into_db) ----
    def _insert_pdf_bundle_into_db(self, bundle: dict, original_filename: str | None = None) -> int:
        gen = (bundle.get("general") or {})
        iso = (bundle.get("isotherm_summary") or {})
        mp = (bundle.get("multipoint_bet_summary") or {})
        tp = (bundle.get("tplot_summary") or {})

        # Instrument's own measurement file (often .qps)
        measurement_filename = gen.get("Filename") or gen.get("Sample ID") or ""

        # For type detection we want the uploaded PDF name here
        file_name = original_filename or measurement_filename or "BET_PDF_Report.pdf"

        date_str, time_str = "", ""
        if isinstance(gen.get("Dates"), list) and gen["Dates"]:
            date_full = str(gen["Dates"][0])
            parts = date_full.split()
            if len(parts) >= 1:
                date_str = parts[0]
            if len(parts) >= 2:
                time_str = parts[1]

        # operator: pick one name only
        operator_primary = gen.get("OperatorPrimary") or (
            gen.get("Operators")[0] if isinstance(gen.get("Operators"), list) and gen["Operators"] else ""
        )

        # pretreatment / measurement conditions from PDF fields
        parts = []
        if gen.get("OutgasTemp"):
            parts.append(str(gen["OutgasTemp"]))
        if gen.get("Outgas Time"):
            parts.append(str(gen["Outgas Time"]))
        if gen.get("Analysis gas"):
            parts.append(str(gen["Analysis gas"]))
        pretreat_str = ", ".join(parts)

        file_info = {
            "file_name": file_name,
            "date_of_measurement": date_str,
            "time_of_measurement": time_str,
            "comment1": measurement_filename or gen.get("Comment", ""),
            "comment2": operator_primary,
            "comment3": pretreat_str,
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

        # Measurement ID for PDFs too
        measurement_id = build_measurement_id(
            file_id=fid,
            file_name=file_info["file_name"],
            date_of_measurement=file_info["date_of_measurement"],
            time_of_measurement=file_info["time_of_measurement"],
            operator=file_info["comment2"],
            instrument=file_info["comment4"],
            serial_number=file_info["serial_number"],
            comment1=file_info["comment1"],
            comment3=file_info["comment3"],
        )
        self.db.execute(
            "UPDATE file_info SET comment5=? WHERE id=?",
            (measurement_id, fid)
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
        ph = ", ".join(":" + k for k in bet_params)
        self.db.execute(f"INSERT INTO bet_parameters ({cols}) VALUES ({ph})", bet_params)

        # Save summaries (for PDF detail view + metadata)
        def _insert_summary(prefix: str, dct: dict):
            if not dct:
                return
            for k, v in dct.items():
                # Special handling for general:Dates to avoid huge lists
                if prefix == "general" and k == "Dates":
                    if isinstance(v, (list, tuple)):
                        uniq = []
                        for d in v:
                            if d not in uniq:
                                uniq.append(d)
                        if len(uniq) == 0:
                            v_str = ""
                        elif len(uniq) == 1:
                            v_str = uniq[0]
                        else:
                            v_str = f"{uniq[0]} – {uniq[-1]}"
                    else:
                        v_str = str(v) if v is not None else None
                else:
                    v_str = str(v) if v is not None else None

                self.db.execute(
                    "INSERT INTO bet_summaries (file_info_id, key, value) VALUES (?, ?, ?)",
                    (fid, f"{prefix}:{k}", v_str)
                )

        _insert_summary("general", gen)
        _insert_summary("isotherm_summary", iso)
        _insert_summary("multipoint_bet_summary", mp)
        _insert_summary("tplot_summary", tp)

        # Insert isotherm points from CSV (if available) into bet_data_points for your existing plot/view logic
        csvs = (bundle.get("csvs") or {})
        iso_csv = csvs.get("isotherm")
        if iso_csv and os.path.exists(iso_csv):
            try:
                df_pts = pd.read_csv(iso_csv)
                for idx, r in df_pts.iterrows():
                    ppo = float(r.get("P_over_P0"))
                    vol = float(r.get("Vol_cc_g_STP"))
                    self.db.execute(
                        "INSERT INTO bet_data_points (file_info_id, no, p_p0, p_va_p0_p) VALUES (?, ?, ?, ?)",
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

        # Minimal technical_info row (same as previous behavior)
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
