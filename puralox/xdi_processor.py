# puralox/xdi_processor.py

import os
import re
import collections
from typing import Dict, Any, List, Tuple

import numpy as np
import pandas as pd


class XdiProcessor:
    """
    XDI support for .xdi and .txt only.

    Stores:
      - file_info (existing)
      - xdi_scans (new)
      - xdi_columns (new)
      - xdi_points (new) : energy, mu, bkg
      - xdi_dataframe (new): full dataframe as JSON (orient='split')
    """

    XDI_EXTS = {".xdi", ".txt"}

    def __init__(self, db_manager):
        self.db = db_manager

    def ensure_tables(self) -> None:
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS xdi_scans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER NOT NULL,
                scan_no INTEGER,
                scan_title TEXT,
                scan_date TEXT,
                energy_col TEXT,
                mu_col TEXT
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS xdi_columns (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER NOT NULL,
                col_index INTEGER NOT NULL,
                col_name TEXT NOT NULL
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS xdi_points (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER NOT NULL,
                energy REAL,
                mu REAL,
                bkg REAL
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS xdi_dataframe (
                file_info_id INTEGER PRIMARY KEY,
                df_json TEXT NOT NULL
            )
        """)

    def process_file(self, filepath: str, original_filename: str) -> int:
        ext = os.path.splitext(original_filename)[1].lower()
        if ext not in self.XDI_EXTS:
            raise ValueError(f"Unsupported XDI extension: {ext}")

        meta, df, cols = self._parse_xdi(filepath)

        energy_col = self._pick_energy_col(df.columns)
        mu_col = self._pick_mu_col(df.columns)

        # bkg computed from mu column (numeric coercion)
        mu_numeric = pd.to_numeric(df[mu_col], errors="coerce")
        bkg = self._compute_bkg(mu_numeric.to_numpy())

        date_str, time_str = self._split_date_time(meta.get("date", ""))
        scan_title = meta.get("scan_title", "") or "XDI"

        file_info = {
            "file_name": original_filename,
            "date_of_measurement": date_str,
            "time_of_measurement": time_str,
            "comment1": scan_title,
            "comment2": "",
            "comment3": "",
            "comment4": "XDI",
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

        self.db.execute(
            """
            INSERT INTO xdi_scans
                (file_info_id, scan_no, scan_title, scan_date, energy_col, mu_col)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (fid, meta.get("scan_no"), scan_title, meta.get("date", ""), energy_col, mu_col)
        )

        for i, c in enumerate(cols):
            self.db.execute(
                "INSERT INTO xdi_columns (file_info_id, col_index, col_name) VALUES (?, ?, ?)",
                (fid, i, str(c))
            )

        energies = pd.to_numeric(df[energy_col], errors="coerce").to_numpy()
        mus = pd.to_numeric(df[mu_col], errors="coerce").to_numpy()

        for e, m, bb in zip(energies, mus, bkg):
            if np.isnan(e) or np.isnan(m):
                continue
            self.db.execute(
                "INSERT INTO xdi_points (file_info_id, energy, mu, bkg) VALUES (?, ?, ?, ?)",
                (fid, float(e), float(m), float(bb))
            )

        # store full dataframe as JSON
        df_json = df.to_json(orient="split")
        self.db.execute(
            "INSERT OR REPLACE INTO xdi_dataframe (file_info_id, df_json) VALUES (?, ?)",
            (fid, df_json)
        )

        return fid

    def _parse_xdi(self, filepath: str) -> Tuple[Dict[str, Any], pd.DataFrame, List[str]]:
        raw = open(filepath, "rb").read().decode("utf-8", errors="replace")
        lines = raw.splitlines()

        meta: Dict[str, Any] = {"scan_no": None, "scan_title": "", "date": ""}
        cols: List[str] | None = None
        rows: List[List[str]] = []
        in_data = False

        for line in lines:
            line = line.rstrip("\n")

            if line.startswith("#S "):
                if meta["scan_no"] is not None:
                    break
                m = re.match(r"#S\s+(\d+)\s*(.*)", line)
                if m:
                    meta["scan_no"] = int(m.group(1))
                    meta["scan_title"] = (m.group(2) or "").strip()

            elif line.startswith("#D "):
                meta["date"] = line[3:].strip()

            elif line.startswith("#L "):
                cols = re.split(r"\s+", line[3:].strip())
                in_data = True

            elif in_data:
                if not line or line.startswith("#"):
                    continue
                parts = re.split(r"\s+", line.strip())
                rows.append(parts)

        if not cols or not rows:
            raise ValueError("XDI parse failed: missing #L header or data rows.")

        # Ensure unique column names if duplicates exist
        counter = collections.Counter()
        unique_cols: List[str] = []
        for c in cols:
            counter[c] += 1
            unique_cols.append(c if counter[c] == 1 else f"{c}_{counter[c]}")

        df = pd.DataFrame(rows, columns=unique_cols)

        # ✅ SAFE numeric conversion:
        # convert to numeric using errors='coerce', but keep original strings if nothing numeric parsed
        for c in df.columns:
            s_num = pd.to_numeric(df[c], errors="coerce")
            if s_num.notna().sum() > 0:
                df[c] = s_num
            else:
                df[c] = df[c].astype(str)

        return meta, df, unique_cols

    def _pick_energy_col(self, columns) -> str:
        if "Energy" in columns:
            return "Energy"
        for c in columns:
            if str(c).lower().startswith("energy"):
                return str(c)
        return str(list(columns)[0])

    def _pick_mu_col(self, columns) -> str:
        if "Absorption" in columns:
            return "Absorption"
        for c in columns:
            cl = str(c).lower()
            if "absorp" in cl or cl == "mu" or "μ" in str(c):
                return str(c)
        return str(list(columns)[-1])

    def _compute_bkg(self, y: np.ndarray) -> np.ndarray:
        s = pd.Series(y)
        n = len(s)
        w = max(11, min(101, (n // 50) * 2 + 1))
        bkg = s.rolling(window=w, center=True, min_periods=1).median()
        bkg = bkg.rolling(window=w, center=True, min_periods=1).mean()
        return bkg.to_numpy()

    def _split_date_time(self, date_str: str) -> Tuple[str, str]:
        if not date_str:
            return "", ""
        parts = str(date_str).split()
        if len(parts) >= 2:
            return parts[0], " ".join(parts[1:])
        return str(date_str), ""
