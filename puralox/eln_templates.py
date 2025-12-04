# puralox/eln_templates.py

SUMMARY = """
<h1>Experiment: {{ measurement_id }}</h1>
<h2>Meta</h2>
<ul>
  <li><b>Date:</b> {{ date_of_measurement }} {{ time_of_measurement }}</li>
  <li><b>Operator:</b> {{ operator }}</li>
  <li><b>Instrument:</b> {{ instrument }}</li>
  <li><b>Serial #:</b> {{ serial_number }}</li>
  <li><b>Scientist:</b> {{ scientist }}</li>
  <li><b>Sample ID:</b> {{ sample_id }}</li>
</ul>
<h2>Results</h2>
<ul>
  <li><b>Specific surface area:</b> {{ specificSurfArea }}</li>
  <li><b>Total pore volume:</b> {{ poreVolume }}</li>
  <li><b>BET points:</b> {{ points }} (P/Po {{ Pmin }}–{{ Pmax }})</li>
</ul>
<p>BET plot: {{ figure }}</p>
"""

DETAILED = """
<h1>Experiment: {{ measurement_id }}</h1>
<h2>Meta</h2>
<ul>
  <li><b>Date:</b> {{ date_of_measurement }} {{ time_of_measurement }}</li>
  <li><b>Operator:</b> {{ operator }}</li>
  <li><b>Instrument:</b> {{ instrument }}</li>
  <li><b>Serial #:</b> {{ serial_number }}</li>
  <li><b>Scientist:</b> {{ scientist }}</li>
  <li><b>Sample ID:</b> {{ sample_id }}</li>
  <li><b>Conditions:</b> {{ temp }}, {{ duration }}, {{ env }}</li>
</ul>
<h2>Parameters</h2>
{{ bet_table | safe }}
<h2>First 15 points</h2>
{{ points_table | safe }}
"""
