# puralox/template_processor.py
import re
from jinja2 import Template
from .db_manager import DatabaseManager
from .config import DB_NAME

_DEFAULT_TEMPLATE = """
<h1>Experiment: {{ measurement_id }}</h1>

<h2>🧪 Experiment Meta Data</h2>
<ul>
  <li><strong>Date:</strong> {{ date_of_measurement }} {{ time_of_measurement }}</li>
  <li><strong>Operator:</strong> {{ operator }}</li>
  <li><strong>Instrument:</strong> {{ instrument }}</li>
  <li><strong>Serial #:</strong> {{ serial_number }}</li>
  <li><strong>Version:</strong> {{ version }}</li>
  <li><strong>Scientist:</strong> {{ scientist }}</li>
  <li><strong>Sample ID:</strong> {{ sample_id }}</li>
  <li><strong>Conditions:</strong> {{ temp }}, {{ duration }}, {{ env }}</li>
</ul>

<h2>📊 Results</h2>
<p>The sample exhibited a specific surface area of <strong>{{ specificSurfArea }}</strong>
and a pore volume of <strong>{{ poreVolume }}</strong>.</p>
<p>BET points considered: <strong>{{ points }}</strong> (P/Po {{ Pmin }}–{{ Pmax }}). The BET plot is attached as a separate figure: <strong>{{ figure }}</strong>.</p>

<h3>📌 BET Parameters</h3>
{{ bet_table | safe }}

<h3>📈 BET Data Points (First 15)</h3>
{{ points_table | safe }}
"""

class TemplateProcessor:
    """
    Local-only renderer. No remote template fetch. Always builds HTML from DB using
    the default inline template above.
    """
    def __init__(self, verify_ssl=False):
        self.db = DatabaseManager(DB_NAME)
        self.verify_ssl = verify_ssl

    def render_default(self, file_id: int) -> str:
        ctx = self._build_context(int(file_id))
        return Template(_DEFAULT_TEMPLATE).render(**ctx)

    # -------------------- Context --------------------
    def _build_context(self, file_id: int) -> dict:
        fi   = self.db.fetchall_dict(f"SELECT * FROM file_info WHERE id={int(file_id)}")
        if not fi:
            raise RuntimeError(f"file_info row not found for id={file_id}")
        fi = fi[0]
        bet  = self.db.fetchall_dict(f"SELECT * FROM bet_parameters WHERE file_info_id={int(file_id)}")
        tech = self.db.fetchall_dict(f"SELECT * FROM technical_info WHERE file_info_id={int(file_id)}")
        pts  = self.db.fetchall_dict(f"SELECT * FROM bet_data_points WHERE file_info_id={int(file_id)}")

        # Single operator: Excel path uses comment2; for PDF we stored OperatorPrimary into comment2
        operator = fi.get("comment2", "") or "—"

        # Measurement ID (fallback)
        measurement_id = fi.get("comment5", "11TDR_0021-0008_CBE01_15Pot_0001_20210612-142626.dat")

        scientist_match = re.match(r"([A-Za-z0-9]+)_", measurement_id)
        serial_match = re.match(r"[A-Za-z0-9]+_[^_]+_([A-Za-z0-9]+)", measurement_id)
        scientist = scientist_match.group(1) if scientist_match else "—"
        serial_number = serial_match.group(1) if serial_match else "—"

        sample_id_match = re.search(r"(\d{4}-\d{4})", measurement_id)
        sample_id = sample_id_match.group(1) if sample_id_match else "—"

        comment3 = fi.get("comment3", "") or ""
        parts = [p.strip() for p in comment3.split(",")]
        temp = parts[0] if len(parts) > 0 else "—"
        duration = parts[1] if len(parts) > 1 else "—"
        env = parts[2] if len(parts) > 2 else "—"

        # P/Po range
        try:
            pvals = [r["p_p0"] for r in pts if r["p_p0"] is not None]
            pmin = min(pvals) if pvals else "—"
            pmax = max(pvals) if pvals else "—"
        except Exception:
            pmin = pmax = "—"

        def make_table(data):
            if not data:
                return "<p>No data</p>"
            headers = list(data[0].keys())
            rows = ["<tr>" + "".join(f"<th>{h}</th>" for h in headers) + "</tr>"]
            for row in data:
                rows.append("<tr>" + "".join(f"<td>{row.get(h,'')}</td>" for h in headers) + "</tr>")
            return "<table border='1' cellspacing='0' cellpadding='4'>" + "".join(rows) + "</table>"

        def pick(*items):
            for it in items:
                if it is None:
                    continue
                s = str(it).strip()
                if s and s.lower() != "none":
                    return s
            return "—"

        ctx = {
            "measurement_id": measurement_id,
            "date_of_measurement": fi.get("date_of_measurement", ""),
            "time_of_measurement": fi.get("time_of_measurement", ""),
            "operator": operator,
            "instrument": fi.get("comment4", ""),
            "serial_number": serial_number,
            "version": fi.get("version", "—"),
            "scientist": scientist,
            "sample_id": sample_id,
            "comment3": comment3,
            "internal_device_id": (tech[0].get("internal_device_id") if tech else "—"),
            "mass": (tech[0].get("mass") if tech else "—"),
            "temp": temp,
            "duration": duration,
            "env": env,
            "points": len(pts),
            "Pmin": pmin,
            "Pmax": pmax,
            "specificSurfArea": pick((bet[0].get("Specific surface area") if bet else None),
                                     (bet[0].get("as_bet") if bet else None)),
            "poreVolume": pick((bet[0].get("Total pore volume") if bet else None),
                               (bet[0].get("total_pore_volume") if bet else None)),
            "figure": "the attached BET plot",
            "bet_table": make_table(bet),
            "points_table": make_table(pts[:15]),
        }
        return ctx
