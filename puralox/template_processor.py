import re
import requests
from jinja2 import Template
from .db_manager import DatabaseManager
from .config import DB_NAME, ELABFTW_URL, ELABFTW_TOKEN

class TemplateProcessor:
    def __init__(self, verify_ssl=False):
        self.db = DatabaseManager(DB_NAME)
        self.verify_ssl = verify_ssl

    def fetch_template_body(self, template_id):
        url = f"{ELABFTW_URL}/experiments_templates/{template_id}"
        headers = {"Authorization": ELABFTW_TOKEN}
        response = requests.get(url, headers=headers, verify=self.verify_ssl)
        if not response.ok:
            raise Exception(f"Failed to fetch template: {response.text}")
        return response.json().get("body", "")

    def render_template_with_data(self, template_str, file_id):
        fi   = self.db.fetchall_dict(f"SELECT * FROM file_info WHERE id={file_id}")[0]
        bet  = self.db.fetchall_dict(f"SELECT * FROM bet_parameters WHERE file_info_id={file_id}")
        tech = self.db.fetchall_dict(f"SELECT * FROM technical_info WHERE file_info_id={file_id}")
        pts  = self.db.fetchall_dict(f"SELECT * FROM bet_data_points WHERE file_info_id={file_id}")

        # Measurement ID manually from comment5 (or hardcoded fallback)
        measurement_id = fi.get("comment5", "11TDR_0021-0008_CBE01_15Pot_0001_20210612-142626.dat")

        # Parse scientist and serial number from measurement ID
        scientist_match = re.match(r"([A-Za-z0-9]+)_", measurement_id)
        serial_match = re.match(r"[A-Za-z0-9]+_[^_]+_([A-Za-z0-9]+)", measurement_id)

        scientist = scientist_match.group(1) if scientist_match else "—"
        serial_number = serial_match.group(1) if serial_match else "—"

        # Parse comment3
        comment3 = fi.get("comment3", "")
        parts = [p.strip() for p in comment3.split(",")]
        temp = parts[0] if len(parts) > 0 else "—"
        duration = parts[1] if len(parts) > 1 else "—"
        env = parts[2] if len(parts) > 2 else "—"

        # Sample ID from measurement ID
        sample_id_match = re.search(r"(\d{4}-\d{4})", measurement_id)
        sample_id = sample_id_match.group(1) if sample_id_match else "—"

        # Pmin and Pmax
        try:
            pmin = min([r["p_p0"] for r in pts])
            pmax = max([r["p_p0"] for r in pts])
        except Exception:
            pmin = pmax = "—"

        # Table formatter
        def make_table(data):
            if not data:
                return "<p>No data</p>"
            headers = data[0].keys()
            rows = ["<tr>" + "".join(f"<th>{h}</th>" for h in headers) + "</tr>"]
            for row in data:
                rows.append("<tr>" + "".join(f"<td>{row[h]}</td>" for h in headers) + "</tr>")
            return "<table border='1' cellspacing='0' cellpadding='4'>" + "".join(rows) + "</table>"

        # Context
        context = {
            "measurement_id": measurement_id,
            "date_of_measurement": fi.get("date_of_measurement", ""),
            "time_of_measurement": fi.get("time_of_measurement", ""),
            "operator": fi.get("comment2", ""),
            "instrument": fi.get("comment4", ""),
            "serial_number": serial_number,
            "version": fi.get("version", "—"),
            "scientist": scientist,
            "sample_id": sample_id,
            "comment3": comment3,
            "internal_device_id": tech[0].get("internal_device_id") if tech else "—",
            "mass": tech[0].get("mass") if tech else "—",
            "temp": temp,
            "duration": duration,
            "env": env,
            "points": len(pts),
            "Pmin": pmin,
            "Pmax": pmax,
            "specificSurfArea": bet[0].get("Specific surface area") if bet else "—",
            "poreVolume": bet[0].get("Total pore volume") if bet else "—",
            "figure": "the attached BET plot",
            "bet_table": make_table(bet),
            "points_table": make_table(pts[:15])
        }

        # Build HTML
        html = f"""
        <h1>Experiment: {context['measurement_id']}</h1>

        <h2>🧪 Experiment Meta Data</h2>
        <ul>
            <li><strong>Measurement ID:</strong> {context['measurement_id']}</li>
            <li><strong>Date:</strong> {context['date_of_measurement']} {context['time_of_measurement']}</li>
            <li><strong>Operator:</strong> {context['operator']}</li>
            <li><strong>Instrument:</strong> {context['instrument']}</li>
            <li><strong>Internal Device ID:</strong> {context['internal_device_id']}</li>
            <li><strong>Serial #:</strong> {context['serial_number']} (Demo Number)</li>
            <li><strong>Version:</strong> {context['version']}</li>
            <li><strong>Scientist (Sample Preparation):</strong> {context['scientist']}</li>
            <li><strong>Sample ID:</strong> {context['sample_id']}</li>
            <li><strong>Measurement Conditions:</strong> {context['temp']}, {context['duration']}, {context['env']}</li>
        </ul>

        <h2>🧬 Experimental Procedure</h2>
        <p>
            {context['mass']} g of the sample <strong>{context['scientist']}_{context['sample_id']}</strong>
            were pretreated at <strong>{context['temp']}</strong> for <strong>{context['duration']}</strong>
            in <strong>{context['env']}</strong> using a BELprep II (BEL) instrument.<br>
            The N₂ physisorption itself was performed using a BELsorp II mini instrument (BEL).<br>
            For the evaluation of the BET isotherm, <strong>{context['points']}</strong> points
            in a pressure range of <strong>{context['Pmin']}</strong> to <strong>{context['Pmax']}</strong> were considered.
        </p>

        <h2>📊 Results</h2>
        <p>The sample exhibited a specific surface area of <strong>{context['specificSurfArea']}</strong>
           and a pore volume of <strong>{context['poreVolume']}</strong>.</p>
        <p>The BET plot is attached as a separate figure: <strong>{context['figure']}</strong>.</p>

        <h3>📌 BET Parameters</h3>
        {context['bet_table']}

        <h3>📈 BET Data Points (First 15)</h3>
        {context['points_table']}
        """

        return html
