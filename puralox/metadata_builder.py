#!/usr/bin/env python3
# puralox/metadata_builder.py
"""
Separate module responsible for building the unified metadata Excel
(core + non-core fields) for a given file_id.
"""

import os
import logging
import pandas as pd

logger = logging.getLogger(__name__)


class MetadataBuilder:
    def __init__(self, db_manager, metadata_dir: str):
        """
        db_manager: instance of DatabaseManager (or compatible)
        metadata_dir: directory where metadata Excel files are written
        """
        self.db = db_manager
        self.metadata_dir = metadata_dir
        os.makedirs(self.metadata_dir, exist_ok=True)

    # local helper; independent of PuraloxApp._table_exists
    def _table_exists(self, name: str) -> bool:
        rows = self.db.fetchall_dict(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (name,)
        )
        return bool(rows)

    def _summary_map(self, file_id: int) -> dict:
        """
        Return a dict { key -> value } of all entries from bet_summaries
        for the given file_id.
        """
        if not self._table_exists("bet_summaries"):
            return {}
        rows = self.db.fetchall_dict(
            "SELECT key, value FROM bet_summaries WHERE file_info_id=?",
            (file_id,)
        )
        return {r["key"]: r["value"] for r in rows}

    def generate_metadata(self, file_id: int) -> str:
        """
        Build the unified metadata Excel for the given file_id.

        Returns:
            Full filesystem path to the written Excel file.
        """
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

        kv = self._summary_map(file_id)

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
            kv.get("multipoint_bet_summary:Surface Area"),
            kv.get("isotherm_summary:Surface Area"),
            kv.get("isotherm_summary: Surface Area"),
        )
        total_pore_volume = pick(
            bet.get("total_pore_volume"),
            kv.get("tplot_summary:Pore Volume"),
            kv.get("tplot_summary: Pore Volume"),
        )
        avg_pore_diam = pick(
            bet.get("average_pore_diameter"),
            kv.get("tplot_summary:Pore Diameter Dv(d)"),
            kv.get("tplot_summary: Pore Diameter Dv(d)"),
        )

        # Core metadata items (same as before)
        md_items: list[tuple[str, str]] = [
            ("File Name",                      fi.get("file_name", "")),
            ("timestamp",                      f"{fi.get('date_of_measurement','')}T{fi.get('time_of_measurement','')}Z"),
            ("sample",                         fi.get("file_name", "")),
            ("operator of experiment",         pick(fi.get("comment2"), kv.get("general:OperatorPrimary"), kv.get("general:Operators"))),
            ("pretreatment conditions",        fi.get("comment3", "")),
            ("manufacturer",                   "BEL, BEL"),
            ("measurement device",             pick(fi.get("comment4"), kv.get("general:Instrument"), "BELprep II, BELsorp II mini instrument")),
            ("serial number",                  fi.get("serial_number", "")),
            ("Version",                        fi.get("version", "")),

            ("Sample weight [g]",              sample_weight),
            ("Standard volume [cm3]",          pick(bet.get("standard_volume"))),
            ("Dead volume [cm3]",              pick(bet.get("dead_volume"))),
            ("Equilibrium time [sec]",         pick(bet.get("equilibrium_time"))),
            ("Adsorptive",                     adsorptive),
            ("Apparatus temperature [C]",      apparatus_temp),
            ("Adsorption temperature [K]",     adsorption_temp),

            ("BET points (count)",             pcount),
            ("BET P/Po min",                   pmin),
            ("BET P/Po max",                   pmax),

            ("Specific surface area",          specific_surface_area),
            ("Total pore volume",              total_pore_volume),
            ("Average pore diameter",          avg_pore_diam),
        ]

        # Mark non-core summary fields too (unchanged logic)
        used_kv_keys = {
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

        for k, v in kv.items():
            if k in used_kv_keys:
                continue
            label = f"(non-core) {k}"
            md_items.append((label, v))

        df = pd.DataFrame(md_items, columns=["Field", "Value"])
        base = os.path.splitext(fi.get("file_name", "file"))[0] or f"file_{file_id}"
        md_file = os.path.join(self.metadata_dir, f"{base}_metadata.xlsx")
        df.to_excel(md_file, index=False)
        logger.info("Unified Metadata Excel written to %s", md_file)
        return md_file
