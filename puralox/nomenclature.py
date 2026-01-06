# puralox/nomenclature.py
import re
from datetime import datetime


def _clean(s):
    if s is None:
        return ""
    return str(s).strip()


def _find_sample_id(*texts):
    """
    Try to find a sample ID of the form 0000-0000 in any of the given strings.
    """
    for t in texts:
        if not t:
            continue
        m = re.search(r"(\d{4}-\d{4})", str(t))
        if m:
            return m.group(1)
    return None


def _short_code_from_instrument(instr: str | None) -> str:
    instr = _clean(instr)
    if not instr:
        return "DEV01"
    token = re.split(r"\s+", instr)[0]
    token = re.sub(r"[^A-Za-z0-9]", "", token)
    return token or "DEV01"


def _short_code_from_operator(op: str | None) -> str:
    op = _clean(op)
    if not op:
        return "SCIENT"
    m = re.search(r"([A-Za-z0-9]{3,})", op)
    if m:
        return m.group(1)
    token = re.split(r"\s+", op)[0]
    return token or "SCIENT"


def _parse_timestamp(date_str: str | None, time_str: str | None) -> str:
    """
    Build YYYYMMDD-HHMMSS from date_of_measurement + time_of_measurement.
    Fallback to current UTC if parsing fails.
    """
    date_str = _clean(date_str)
    time_str = _clean(time_str)
    combo = (date_str + " " + time_str).strip()

    dt = None
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y",
    ):
        if not combo:
            continue
        try:
            dt = datetime.strptime(combo, fmt)
            break
        except Exception:
            continue

    if dt is None:
        dt = datetime.utcnow()
    return dt.strftime("%Y%m%d-%H%M%S")


def build_measurement_id(
    file_id: int,
    file_name: str | None = None,
    date_of_measurement: str | None = None,
    time_of_measurement: str | None = None,
    operator: str | None = None,
    instrument: str | None = None,
    serial_number: str | None = None,
    comment1: str | None = None,
    comment3: str | None = None,
) -> str:
    """
    Build a Measurement ID like:
      11TDR_0021-0008_CBE01_BET01_0001_20210612-142626.dat
    """
    op_code = _short_code_from_operator(operator)
    sample_id = _find_sample_id(file_name, comment1, comment3) or "0000-0000"

    serial_number = _clean(serial_number)
    if serial_number:
        dev_code = re.split(r"\s+", serial_number)[0]
        dev_code = re.sub(r"[^A-Za-z0-9]", "", dev_code) or "DEV01"
    else:
        dev_code = _short_code_from_instrument(instrument)

    run_code = "BET01"
    index_code = f"{int(file_id):04d}"
    ts = _parse_timestamp(date_of_measurement, time_of_measurement)

    measurement_id = f"{op_code}_{sample_id}_{dev_code}_{run_code}_{index_code}_{ts}.dat"
    return measurement_id
