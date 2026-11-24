import json
import math
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests
import xlsxwriter


EC_TZ = ZoneInfo("America/Guayaquil")
FIELDS = "timestamp,vehicle.mileage"
INPUT_FILE = Path("MANTENIMIENTO.xlsx")
OUTPUT_FILE = Path("MANTENIMIENTO_KM_ACTUAL.xlsx")
# El kilometraje base corresponde al registro de recepción del vehículo.
# No se usa KM_MANTENIMIENTO para evitar mezclar valores de la orden.
KM_BASE_COLUMNS = ("KM_RECEPCION",)
MAX_WORKERS = 18
CURRENT_CHUNK_SIZE = 80
BASE_LOOKBACK_DAYS = 14
BASE_LOOKAHEAD_DAYS = 14
DEFAULT_CALC_ID = 1931826
CALC_COUNTER = "distancia_en_movimiento"
DEFAULT_REASONABLE_DIFFERENCE_KM = 500


HTTP = requests.Session()
HTTP.trust_env = False


def load_config():
    with open("config.json", "r", encoding="utf-8") as f:
        return json.load(f)


def colnum(cell_ref):
    letters = "".join(re.findall(r"[A-Z]+", cell_ref or ""))
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def read_xlsx_rows(path):
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(ns + "si"):
                shared.append("".join(t.text or "" for t in si.iter(ns + "t")))

        headers = None
        rows = []
        for _event, elem in ET.iterparse(z.open("xl/worksheets/sheet1.xml"), events=("end",)):
            if elem.tag != ns + "row":
                continue
            rid = int(elem.attrib.get("r", "0"))
            values = {}
            for c in elem.findall(ns + "c"):
                idx = colnum(c.attrib.get("r", ""))
                v = c.find(ns + "v")
                value = None if v is None else v.text
                if c.attrib.get("t") == "s" and value is not None:
                    value = shared[int(value)]
                elif c.attrib.get("t") == "inlineStr":
                    value = "".join(t.text or "" for t in c.iter(ns + "t"))
                values[idx] = value
            max_col = max(values.keys() or [0])
            row = [values.get(i) for i in range(1, max_col + 1)]
            if rid == 1:
                headers = [str(x or "").strip() for x in row]
            elif headers and any(x not in (None, "") for x in row):
                item = {headers[i]: row[i] if i < len(row) else None for i in range(len(headers))}
                rows.append(item)
            elem.clear()
    return headers, rows


def parse_float(value):
    if value in (None, ""):
        return None
    try:
        return float(str(value).replace(",", ".").strip())
    except ValueError:
        return None


def parse_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)) or str(value).strip().replace(".", "", 1).isdigit():
        # Excel serial date, compatible with Windows 1900 date system.
        serial = float(value)
        base = datetime(1899, 12, 30, tzinfo=EC_TZ)
        return base + timedelta(days=serial)
    text = str(value).strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.replace(tzinfo=EC_TZ)
        except ValueError:
            pass
    try:
        dt = datetime.fromisoformat(text)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=EC_TZ)
        return dt.astimezone(EC_TZ)
    except ValueError:
        return None


def number_or_none(value):
    try:
        n = float(value)
        return n if math.isfinite(n) else None
    except (TypeError, ValueError):
        return None


def classify_difference(value, limit):
    if value is None:
        return "SIN DATOS"
    return "RAZONABLE" if abs(value) <= limit else "REVISAR"


def km_flespi_en_fecha_base(km_actual, km_recorrido):
    """Estima el odómetro Flespi en la fecha base restando el recorrido posterior."""
    if km_actual is None or km_recorrido is None:
        return None
    return km_actual - max(0, km_recorrido)


def first_km_base(row, headers):
    for name in KM_BASE_COLUMNS:
        if name in headers:
            value = parse_float(row.get(name))
            if value is not None:
                return value
    return None


def fetch_current_batch(ids, token):
    """Metodo de respaldo: telemetria actual (vehicle.mileage) directo del dispositivo."""
    result = {}
    headers = {"Authorization": f"FlespiToken {token}"}
    ids = [str(x) for x in ids if x]
    for i in range(0, len(ids), CURRENT_CHUNK_SIZE):
        chunk = ids[i:i + CURRENT_CHUNK_SIZE]
        url = f"https://flespi.io/gw/devices/{','.join(chunk)}/telemetry/{FIELDS}"
        try:
            response = HTTP.get(url, headers=headers, timeout=25)
            if not response.ok:
                for device_id in chunk:
                    result[device_id] = {"error": f"HTTP {response.status_code}"}
                continue
            for item in response.json().get("result", []):
                device_id = str(item.get("id"))
                telemetry = item.get("telemetry") or {}
                mileage = telemetry.get("vehicle.mileage") or {}
                timestamp = telemetry.get("timestamp") or {}
                result[device_id] = {
                    "km": number_or_none(mileage.get("value")),
                    "timestamp": timestamp.get("value") or mileage.get("ts"),
                    "error": None,
                }
        except requests.RequestException as error:
            for device_id in chunk:
                result[device_id] = {"error": str(error)}
        print(f"Actuales Flespi: {min(i + CURRENT_CHUNK_SIZE, len(ids))}/{len(ids)}")
    return result


def fetch_one_message(device_id, start_utc, end_utc, reverse, token):
    query = {
        "from": int(start_utc.timestamp()),
        "to": int(end_utc.timestamp()),
        "count": 1,
        "reverse": reverse,
        "fields": FIELDS,
    }
    url = f"https://flespi.io/gw/devices/{device_id}/messages?data={quote(json.dumps(query, separators=(',', ':')))}"
    response = HTTP.get(url, headers={"Authorization": f"FlespiToken {token}"}, timeout=20)
    if not response.ok:
        return None, f"HTTP {response.status_code}"
    messages = response.json().get("result", [])
    if not messages:
        return None, None
    row = messages[0]
    return {
        "km": number_or_none(row.get("vehicle.mileage")),
        "timestamp": row.get("timestamp"),
    }, None


def fetch_base_mileage(key, token):
    """Metodo de respaldo: busca el mensaje historico mas cercano a la fecha de recepcion."""
    device_id, fecha = key
    if not device_id or not fecha:
        return {"km": None, "timestamp": None, "error": "SIN ID_FLESPI O FECHA"}
    cutoff = fecha.astimezone(timezone.utc)
    try:
        before, error = fetch_one_message(
            device_id,
            cutoff - timedelta(days=BASE_LOOKBACK_DAYS),
            cutoff,
            True,
            token,
        )
        if before and before.get("km") is not None:
            return {**before, "error": None, "lado": "ANTES"}
        after, error2 = fetch_one_message(
            device_id,
            cutoff,
            cutoff + timedelta(days=BASE_LOOKAHEAD_DAYS),
            False,
            token,
        )
        if after and after.get("km") is not None:
            return {**after, "error": None, "lado": "DESPUES"}
        return {"km": None, "timestamp": None, "error": error or error2 or "SIN MENSAJES CERCA DE FECHA"}
    except requests.RequestException as error:
        return {"km": None, "timestamp": None, "error": str(error)}


def fetch_intervals(device_id, calc_id, token):
    """Trae todos los intervalos precalculados del calculador Flespi para un dispositivo."""
    url = f"https://flespi.io/gw/calcs/{calc_id}/devices/{device_id}/intervals/all"
    try:
        response = HTTP.get(url, headers={"Authorization": f"FlespiToken {token}"}, timeout=25)
        if not response.ok:
            return None, f"HTTP {response.status_code}"
        return response.json().get("result", []), None
    except requests.RequestException as error:
        return None, str(error)


def km_recorrido_from_intervals(intervals, fecha_base, ahora, counter=CALC_COUNTER):
    """Suma el contador del calculador entre fecha_base y ahora.

    Los intervalos del calc estan partidos por mes calendario, no coinciden
    con la fecha exacta de recepcion del taller. El tramo que solo se
    superpone parcialmente con [fecha_base, ahora] se interpola de forma
    proporcional al tiempo, para no contar dias fuera del rango.
    """
    base_ts = fecha_base.timestamp()
    now_ts = ahora.timestamp()
    total = 0.0
    found = False
    for interval in intervals or []:
        begin = interval.get("begin")
        end = interval.get("end")
        value = interval.get(counter)
        if begin is None or end is None or value is None:
            continue
        duration = end - begin
        if duration <= 0:
            continue
        overlap_start = max(begin, base_ts)
        overlap_end = min(end, now_ts)
        if overlap_end <= overlap_start:
            continue
        fraction = (overlap_end - overlap_start) / duration
        total += value * fraction
        found = True
    return total if found else None


def fetch_daily_mileage(key, ahora, token):
    """Suma mileage() en intervalos diarios, usando la configuracion del reporte Flespi."""
    device_id, fecha = key
    if not device_id or not fecha:
        return {"km": None, "days": 0, "error": "SIN ID_FLESPI O FECHA"}
    query = {
        "from": int(fecha.timestamp()),
        "to": int(ahora.timestamp()),
        "timezone": "America/Guayaquil",
        "selectors": [{"type": "datetime", "split": "day"}],
        "counters": [
            {
                "type": "expression",
                "name": "distance_km",
                "expression": "mileage()",
                "method": "summary",
            },
        ],
    }
    url = f"https://flespi.io/gw/devices/{device_id}/messages/calculate"
    try:
        response = HTTP.get(
            url,
            params={"data": json.dumps(query, separators=(",", ":"))},
            headers={"Authorization": f"FlespiToken {token}"},
            timeout=60,
        )
        if not response.ok:
            return {"km": None, "days": 0, "error": f"HTTP {response.status_code}"}
        intervals = response.json().get("result", [])
        distances = [number_or_none(item.get("distance_km")) for item in intervals]
        distances = [value for value in distances if value is not None]
        if not distances:
            return {"km": None, "days": 0, "error": "SIN INTERVALOS DIARIOS"}
        return {"km": sum(distances), "days": len(distances), "error": None}
    except (requests.RequestException, ValueError) as error:
        return {"km": None, "days": 0, "error": str(error)}


def write_output(headers, rows, output_file=OUTPUT_FILE):
    out_headers = headers + [
        "FECHA_FLESPI_BASE",
        "KM_FLESPI_FECHA_RECEPCION",
        "FECHA_FLESPI_ACTUAL",
        "KM_RECORRIDO_FLESPI",
        "KM_DEBERIA_ESTAR_HOY",
        "KM_ACTUAL_ESTIMADO",
        "KM_FLESPI_ACTUAL",
        "DIFERENCIA_KM",
        "ESTADO_DIFERENCIA",
        "ODOMETRO_ACTUALIZADO",
        "KM_SOLICITADO_FLESPI",
        "ESTADO_ACTUALIZACION_FLESPI",
        "METODO_CALCULO",
        "ORIGEN_KM",
        "ESTADO_CALCULO",
        "ACCION_RECOMENDADA",
        "OBSERVACION",
    ]
    workbook = xlsxwriter.Workbook(str(output_file))
    ws = workbook.add_worksheet("KM_ACTUAL")
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    num_fmt = workbook.add_format({"num_format": "#,##0.00"})
    update_fmt = workbook.add_format({"bg_color": "#D1FADF", "font_color": "#05603A", "bold": True})
    review_fmt = workbook.add_format({"bg_color": "#FEE4E2", "font_color": "#B42318", "bold": True})
    for col, name in enumerate(out_headers):
        ws.write(0, col, name, header_fmt)
    for r, row in enumerate(rows, start=1):
        for c, name in enumerate(out_headers):
            value = row.get(name)
            if isinstance(value, (int, float)) and name.startswith(("KM_",)):
                ws.write_number(r, c, value, num_fmt)
            else:
                ws.write(r, c, value)
    ws.autofilter(0, 0, len(rows), len(out_headers) - 1)
    ws.freeze_panes(1, 0)
    ws.set_column(0, len(out_headers) - 1, 18)
    action_col = out_headers.index("ACCION_RECOMENDADA")
    ws.set_column(action_col, action_col, 24)
    ws.conditional_format(1, action_col, len(rows), action_col, {
        "type": "text", "criteria": "containing", "value": "ACTUALIZAR", "format": update_fmt,
    })
    ws.conditional_format(1, action_col, len(rows), action_col, {
        "type": "text", "criteria": "containing", "value": "REVISAR", "format": review_fmt,
    })
    workbook.close()


def finalize_output(headers, rows, results, ahora, output_file, method_name, difference_limit):
    enriched = []
    for index, row in enumerate(rows):
        out = {key: row.get(key) for key in headers}
        fecha = row["_fecha"]
        km_taller = row["_km_base_taller"]
        entry = results[index]
        status = entry["status"]
        obs = entry["obs"]
        km_recorrido = entry["km_recorrido"]
        origen = entry["origen"]

        km_estimado = None
        if km_recorrido is not None and km_taller is not None:
            if km_recorrido < -1:
                status = "KM_REVISAR"
                obs.append(f"Distancia negativa segun {origen}")
            km_estimado = km_taller + max(0, km_recorrido)
        elif status == "OK":
            status = "SIN_DATOS_CALC"

        km_flespi_actual = entry.get("km_flespi_actual")
        km_flespi_recepcion = km_flespi_en_fecha_base(km_flespi_actual, km_recorrido)
        diferencia = km_estimado - km_flespi_actual if km_estimado is not None and km_flespi_actual is not None else None
        difference_status = classify_difference(diferencia, difference_limit)
        if status == "OK" and difference_status == "REVISAR":
            obs.append(f"Diferencia alta: {abs(diferencia):.2f} km; limite razonable: {difference_limit:.2f} km")
        elif status == "OK" and difference_status == "SIN DATOS":
            obs.append("No existe KM_FLESPI_ACTUAL para comparar")
        recommended_action = "ACTUALIZAR" if status == "OK" and difference_status == "RAZONABLE" else "REVISAR"

        out.update({
            "FECHA_FLESPI_BASE": fecha.strftime("%Y-%m-%d %H:%M:%S") if fecha else None,
            "KM_FLESPI_FECHA_RECEPCION": round(km_flespi_recepcion, 2) if km_flespi_recepcion is not None else None,
            "FECHA_FLESPI_ACTUAL": ahora.astimezone(EC_TZ).strftime("%Y-%m-%d %H:%M:%S"),
            "KM_RECORRIDO_FLESPI": round(km_recorrido, 2) if km_recorrido is not None else None,
            "KM_DEBERIA_ESTAR_HOY": round(km_estimado, 2) if km_estimado is not None else None,
            "KM_ACTUAL_ESTIMADO": round(km_estimado, 2) if km_estimado is not None else None,
            "KM_FLESPI_ACTUAL": round(km_flespi_actual, 2) if km_flespi_actual is not None else None,
            "DIFERENCIA_KM": round(diferencia, 2) if diferencia is not None else None,
            "ESTADO_DIFERENCIA": difference_status,
            "ODOMETRO_ACTUALIZADO": "NO",
            "KM_SOLICITADO_FLESPI": None,
            "ESTADO_ACTUALIZACION_FLESPI": None,
            "METODO_CALCULO": method_name,
            "ORIGEN_KM": origen,
            "ESTADO_CALCULO": status,
            "ACCION_RECOMENDADA": recommended_action,
            "OBSERVACION": " | ".join(obs),
        })
        enriched.append(out)

    write_output(headers, enriched, output_file)
    summary = {}
    for row in enriched:
        summary[row["ESTADO_CALCULO"]] = summary.get(row["ESTADO_CALCULO"], 0) + 1
    print("Resumen:", summary)
    print("Archivo generado:", output_file.resolve())


def main():
    input_file = Path(sys.argv[1]) if len(sys.argv) > 1 else INPUT_FILE
    output_file = Path(sys.argv[2]) if len(sys.argv) > 2 else OUTPUT_FILE
    calculation_mode = sys.argv[3] if len(sys.argv) > 3 else "calculadora"
    if calculation_mode not in {"calculadora", "reporte_diario"}:
        raise ValueError("Metodo de calculo no valido")
    cfg = load_config()
    token = cfg["flespi"]["token"]
    calc_id = cfg.get("flespi", {}).get("calc_id", DEFAULT_CALC_ID)
    difference_limit = float(cfg.get("diferencia_razonable_km", DEFAULT_REASONABLE_DIFFERENCE_KM))
    headers, rows = read_xlsx_rows(input_file)
    print(f"Filas leidas: {len(rows)}")
    km_base_columns = [name for name in KM_BASE_COLUMNS if name in headers]
    if not km_base_columns:
        raise ValueError(f"No se encontro columna base de kilometraje: {', '.join(KM_BASE_COLUMNS)}")
    print(f"Columnas KM base disponibles: {', '.join(km_base_columns)}")
    print(f"Calculador Flespi: {calc_id} (contador: {CALC_COUNTER})")

    ahora = datetime.now(timezone.utc)
    for row in rows:
        row["_id_flespi"] = str(row.get("ID_FLESPI") or "").strip()
        row["_fecha"] = parse_date(row.get("FECHA_RECEPCION"))
        row["_km_base_taller"] = first_km_base(row, headers)

    ids = sorted({row["_id_flespi"] for row in rows if row["_id_flespi"]})
    print(f"IDs Flespi unicos: {len(ids)}")
    print(f"Metodo de calculo: {calculation_mode}")
    current_by_id = fetch_current_batch(ids, token)

    if calculation_mode == "reporte_diario":
        keys = sorted(
            {(row["_id_flespi"], row["_fecha"]) for row in rows if row["_id_flespi"] and row["_fecha"]},
            key=lambda item: (item[0], item[1].isoformat()),
        )
        reports_by_key = {}
        started = time.time()
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(fetch_daily_mileage, key, ahora, token): key for key in keys}
            for index, future in enumerate(as_completed(futures), start=1):
                key = futures[future]
                reports_by_key[key] = future.result()
                if index % 25 == 0 or index == len(keys):
                    print(f"Reportes diarios Flespi: {index}/{len(keys)} | {round(time.time() - started, 1)}s")

        results = {}
        for index, row in enumerate(rows):
            device_id = row["_id_flespi"]
            fecha = row["_fecha"]
            km_taller = row["_km_base_taller"]
            status = "OK"
            obs = []
            if not device_id:
                status = "SIN ID_FLESPI"
            if not fecha:
                status = "SIN FECHA"
            if km_taller is None:
                status = "SIN KM_RECEPCION"
            report = reports_by_key.get((device_id, fecha)) or {}
            if report.get("error"):
                obs.append(report["error"])
            if report.get("days"):
                obs.append(f"{report['days']} dias con distancia")
            km_recorrido = report.get("km") if status == "OK" else None
            if km_recorrido is None and status == "OK":
                status = "SIN_DATOS_CALC"
            results[index] = {
                "km_recorrido": km_recorrido,
                "km_flespi_actual": current_by_id.get(device_id, {}).get("km"),
                "origen": "REPORTE_DIARIO_MILEAGE" if km_recorrido is not None else None,
                "status": status,
                "obs": obs,
            }

        finalize_output(headers, rows, results, ahora, output_file, "REPORTE_DIARIO_MILEAGE", difference_limit)
        return

    print(f"Bases historicas unicas: {len(ids)}")
    intervals_by_id = {}
    started = time.time()
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(fetch_intervals, device_id, calc_id, token): device_id for device_id in ids}
        for i, future in enumerate(as_completed(futures), start=1):
            device_id = futures[future]
            result, error = future.result()
            intervals_by_id[device_id] = {"intervals": result, "error": error}
            if i % 25 == 0 or i == len(ids):
                print(f"Historicos Flespi: {i}/{len(ids)} | {round(time.time() - started, 1)}s")

    # Primera pasada: intenta resolver todo con el calculador Flespi (GPS).
    pending = []
    results = {}
    for index, row in enumerate(rows):
        device_id = row["_id_flespi"]
        fecha = row["_fecha"]
        km_taller = row["_km_base_taller"]
        entry = intervals_by_id.get(device_id) or {}
        obs = []
        status = "OK"
        if not device_id:
            status = "SIN ID_FLESPI"
        if not fecha:
            status = "SIN FECHA"
        if km_taller is None:
            status = "SIN KM_RECEPCION"
        if entry.get("error"):
            obs.append(entry["error"])

        km_recorrido = None
        if status == "OK" and fecha is not None:
            km_recorrido = km_recorrido_from_intervals(entry.get("intervals"), fecha, ahora)

        if km_recorrido is not None:
            results[index] = {
                "km_recorrido": km_recorrido,
                "origen": "CALC_GPS",
                "status": status,
                "obs": list(obs),
            }
        else:
            if status == "OK":
                obs.append("Sin cobertura del calculador Flespi para este dispositivo/fecha")
            results[index] = {"km_recorrido": None, "origen": None, "status": status, "obs": obs}
            if status == "OK" and device_id and fecha:
                pending.append(index)

    # Segunda pasada (respaldo): telemetria/mensajes crudos solo para lo que el calculador no cubrio.
    if pending:
        print(f"Filas sin cobertura del calculador, usando respaldo OBD: {len(pending)}")
        fallback_ids = sorted({rows[i]["_id_flespi"] for i in pending})
        fallback_keys = sorted(
            {(rows[i]["_id_flespi"], rows[i]["_fecha"]) for i in pending},
            key=lambda item: (item[0], item[1].isoformat()),
        )
        base_by_key = {}
        started = time.time()
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(fetch_base_mileage, key, token): key for key in fallback_keys}
            for i, future in enumerate(as_completed(futures), start=1):
                key = futures[future]
                base_by_key[key] = future.result()
                if i % 25 == 0 or i == len(fallback_keys):
                    print(f"Historicos Flespi: {i}/{len(fallback_keys)} | {round(time.time() - started, 1)}s")

        for index in pending:
            row = rows[index]
            device_id = row["_id_flespi"]
            fecha = row["_fecha"]
            current = current_by_id.get(device_id) or {}
            base = base_by_key.get((device_id, fecha)) or {}
            km_current = current.get("km")
            km_base_flespi = base.get("km")
            entry = results[index]
            if current.get("error"):
                entry["obs"].append(current["error"])
            if base.get("error"):
                entry["obs"].append(base["error"])
            if km_current is not None and km_base_flespi is not None:
                entry["km_recorrido"] = km_current - km_base_flespi
                entry["origen"] = "TELEMETRIA_OBD"
            else:
                entry["status"] = "SIN_DATOS_CALC"

    enriched = []
    for index, row in enumerate(rows):
        out = {k: row.get(k) for k in headers}
        fecha = row["_fecha"]
        km_taller = row["_km_base_taller"]
        entry = results[index]
        status = entry["status"]
        obs = entry["obs"]
        km_recorrido = entry["km_recorrido"]
        origen = entry["origen"]

        km_estimado = None
        if km_recorrido is not None and km_taller is not None:
            if km_recorrido < -1:
                status = "KM_REVISAR"
                obs.append(f"Distancia negativa segun {origen}")
            km_estimado = km_taller + max(0, km_recorrido)
        elif status == "OK":
            status = "SIN_DATOS_CALC"

        km_flespi_actual = current_by_id.get(row["_id_flespi"], {}).get("km")
        km_flespi_recepcion = km_flespi_en_fecha_base(km_flespi_actual, km_recorrido)
        diferencia = km_estimado - km_flespi_actual if km_estimado is not None and km_flespi_actual is not None else None
        difference_status = classify_difference(diferencia, difference_limit)
        if status == "OK" and difference_status == "REVISAR":
            obs.append(f"Diferencia alta: {abs(diferencia):.2f} km; limite razonable: {difference_limit:.2f} km")
        elif status == "OK" and difference_status == "SIN DATOS":
            obs.append("No existe KM_FLESPI_ACTUAL para comparar")
        recommended_action = "ACTUALIZAR" if status == "OK" and difference_status == "RAZONABLE" else "REVISAR"

        out.update({
            "FECHA_FLESPI_BASE": fecha.strftime("%Y-%m-%d %H:%M:%S") if fecha else None,
            "KM_FLESPI_FECHA_RECEPCION": round(km_flespi_recepcion, 2) if km_flespi_recepcion is not None else None,
            "FECHA_FLESPI_ACTUAL": ahora.astimezone(EC_TZ).strftime("%Y-%m-%d %H:%M:%S"),
            "KM_RECORRIDO_FLESPI": round(km_recorrido, 2) if km_recorrido is not None else None,
            "KM_DEBERIA_ESTAR_HOY": round(km_estimado, 2) if km_estimado is not None else None,
            "KM_ACTUAL_ESTIMADO": round(km_estimado, 2) if km_estimado is not None else None,
            "KM_FLESPI_ACTUAL": round(km_flespi_actual, 2) if km_flespi_actual is not None else None,
            "DIFERENCIA_KM": round(diferencia, 2) if diferencia is not None else None,
            "ESTADO_DIFERENCIA": difference_status,
            "ODOMETRO_ACTUALIZADO": "NO",
            "KM_SOLICITADO_FLESPI": None,
            "ESTADO_ACTUALIZACION_FLESPI": None,
            "METODO_CALCULO": "CALCULADORA_FLESPI",
            "ORIGEN_KM": origen,
            "ESTADO_CALCULO": status,
            "ACCION_RECOMENDADA": recommended_action,
            "OBSERVACION": " | ".join(obs),
        })
        enriched.append(out)

    write_output(headers, enriched, output_file)
    summary = {}
    for row in enriched:
        summary[row["ESTADO_CALCULO"]] = summary.get(row["ESTADO_CALCULO"], 0) + 1
    print("Resumen:", summary)
    print("Archivo generado:", output_file.resolve())


if __name__ == "__main__":
    main()

