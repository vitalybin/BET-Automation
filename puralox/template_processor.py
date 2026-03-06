# puralox/template_processor.py
import re
from jinja2 import Template

from .db_manager import DatabaseManager
from .config import DB_NAME


# This template matches the HTML body you previously built inside PuraloxApp._eln_create_local_json
_ELN_BODY_TEMPLATE = """
<h1>BET Measurement Report</h1>

<h2>Meta Data</h2>
<ul>
  <li><strong>Measurement ID:</strong> {{ measurement_id }}</li>
  <li><strong>Date:</strong> {{ date }} {{ time }}</li>
  <li><strong>Operator:</strong> {{ operator }}</li>
  <li><strong>Instrument:</strong> {{ instrument }}</li>
  <li><strong>Internal Device ID:</strong> {{ internal_device_id }}</li>
  <li><strong>Serial #:</strong> {{ serial_number }}</li>
  <li><strong>Version:</strong> {{ version }}</li>
  <li><strong>Scientist (Sample Preparation):</strong> {{ scientist }}</li>
  <li><strong>Sample ID:</strong> {{ sample_id }}</li>
  <li><strong>Measurement Conditions:</strong> {{ comment3 }}</li>
</ul>

<h2>Experimental Procedure</h2>
<p>
  {{ mass }} g of the sample <strong>{{ scientist }}_{{ sample_id }}</strong>
  were pretreated under the following conditions:
  <strong>{{ comment3 }}</strong>.<br>
  For the evaluation of the BET isotherm, <strong>{{ points_count }}</strong> points
  in a pressure range of <strong>{{ pmin }}</strong> to <strong>{{ pmax }}</strong> were considered.
</p>

<h2>Results</h2>
<p>
  The sample exhibited a specific surface area of
  <strong>{{ specific_surf_area }}</strong> and a pore volume of
  <strong>{{ pore_volume }}</strong>.
</p>

<h3>BET Parameters</h3>
{{ bet_table_html | safe }}

<h3>BET Data Points (First 15)</h3>
{{ pts_table_html | safe }}
"""


class TemplateProcessor:
    """
    Builds ELN HTML bodies from the database.

    Refactor goal:
      - Keep the ELN HTML output the same as before.
      - Make TemplateProcessor actually used by PuraloxApp (so UML matches runtime).
      - Support dependency injection (share the same DatabaseManager instance).
    """

    def __init__(self, db: DatabaseManager | None = None, verify_ssl: bool = False):
        self.db = db if db is not None else DatabaseManager(DB_NAME)
        self.verify_ssl = verify_ssl

    # public API used by PuraloxApp
    def build_eln_html(self, file_id: int) -> str:
        ctx = self._build_eln_context(file_id)
        return Template(_ELN_BODY_TEMPLATE).render(**ctx)

    # -------------------- Context builders --------------------
    def _build_eln_context(self, file_id: int) -> dict:
        fi_rows = self.db.fetchall_dict("SELECT * FROM file_info WHERE id=?", (file_id,))
        if not fi_rows:
            raise RuntimeError(f"file_info not found for id={file_id}")
        fi = fi_rows[0]

        bet_rows = self.db.fetchall_dict("SELECT * FROM bet_parameters WHERE file_info_id=?", (file_id,))
        bet = bet_rows[0] if bet_rows else {}

        tech_rows = self.db.fetchall_dict("SELECT * FROM technical_info WHERE file_info_id=?", (file_id,))
        tech = tech_rows[0] if tech_rows else {}

        pts = self.db.fetchall_dict("SELECT * FROM bet_data_points WHERE file_info_id=? ORDER BY no", (file_id,))

        measurement_id = fi.get("comment5") or fi.get("file_name") or f"file_{file_id}"
        scientist = fi.get("comment2") or "—"

        m_sample = re.search(r"(\d{4}-\d{4})", measurement_id)
        sample_id = m_sample.group(1) if m_sample else "—"

        comment3 = fi.get("comment3") or ""
        try:
            pvals = [r.get("p_p0") for r in pts if r.get("p_p0") is not None]
            pmin = min(pvals) if pvals else "—"
            pmax = max(pvals) if pvals else "—"
        except Exception:
            pmin = pmax = "—"

        mass = tech.get("mass", "—")
        internal_device_id = tech.get("internal_device_id", "—")

        specific_surf_area = bet.get("as_bet", "—") or bet.get("Specific surface area", "—")
        pore_volume = bet.get("total_pore_volume", "—") or bet.get("Total pore volume", "—")

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

        return {
            "measurement_id": measurement_id,
            "date": fi.get("date_of_measurement", ""),
            "time": fi.get("time_of_measurement", ""),
            "operator": fi.get("comment2", ""),
            "instrument": fi.get("comment4", ""),
            "internal_device_id": internal_device_id,
            "serial_number": fi.get("serial_number", ""),
            "version": fi.get("version", ""),
            "scientist": scientist,
            "sample_id": sample_id,
            "comment3": comment3,
            "mass": mass,
            "points_count": len(pts),
            "pmin": pmin,
            "pmax": pmax,
            "specific_surf_area": specific_surf_area,
            "pore_volume": pore_volume,
            "bet_table_html": bet_table_html,
            "pts_table_html": pts_table_html,
        }
