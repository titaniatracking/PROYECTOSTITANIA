import json
import math
import sqlite3
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
import xlsxwriter

from calcular_km_actual_mantenimiento import parse_date, parse_float, read_xlsx_rows


APP_DIR = Path(__file__).resolve().parent
EC_TZ = ZoneInfo("America/Guayaquil")
MAX_WORKERS = 12
MAX_DAILY_KM = 800.0
CACHE_HOURS = 8
CACHE_PATH = APP_DIR / "odometro_real_cache.sqlite3"
HTTP = requests.Session()
HTTP.trust_env = False


def load_config():
    return json.loads((APP_DIR / "config.json").read_text(encoding="utf-8"))


def number(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def effective_start(row):
    dates = [
        parse_date(row.get("FECHA_CREACION_VEHICULO")),
        parse_date(row.get("FECHA_INSTALACION_DISPOSITIVO")),
        parse_date(row.get("FECHA_INICIO_DISPOSITIVO")),
        parse_date(row.get("FECHA_CREACION_ASIGNACION")),
    ]
    dates = [value for value in dates if value is not None]
    return max(dates) if dates else None


def cache_connection():
    connection = sqlite3.connect(CACHE_PATH, timeout=30)
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS odometro_cache (
          cache_key TEXT PRIMARY KEY,
          calculated_at TEXT NOT NULL,
          payload TEXT NOT NULL
        )
        """
    )
    return connection


def read_cache(cache_key):
    try:
        with cache_connection() as connection:
            row = connection.execute(
                "SELECT calculated_at, payload FROM odometro_cache WHERE cache_key = ?", (cache_key,)
            ).fetchone()
        if not row:
            return None
        calculated_at = datetime.fromisoformat(row[0])
        if datetime.now(timezone.utc) - calculated_at > timedelta(hours=CACHE_HOURS):
            return None
        result = json.loads(row[1])
        result["cache"] = True
        return result
    except (sqlite3.Error, ValueError, json.JSONDecodeError):
        return None


def write_cache(cache_key, payload):
    try:
        clean = {**payload, "cache": False}
        with cache_connection() as connection:
            connection.execute(
                "INSERT OR REPLACE INTO odometro_cache(cache_key, calculated_at, payload) VALUES(?,?,?)",
                (cache_key, datetime.now(timezone.utc).isoformat(), json.dumps(clean, ensure_ascii=False)),
            )
    except sqlite3.Error:
        pass


def fetch_current_batch(device_ids, token):
    result = {}
    ids = [str(value) for value in device_ids if value]
    headers = {"Authorization": f"FlespiToken {token}"}
    for index in range(0, len(ids), 80):
        chunk = ids[index:index + 80]
        url = f"https://flespi.io/gw/devices/{','.join(chunk)}/telemetry/timestamp,vehicle.mileage"
        try:
            response = HTTP.get(url, headers=headers, timeout=30)
            if not response.ok:
                for device_id in chunk:
                    result[device_id] = {"km": None, "error": f"HTTP {response.status_code}"}
                continue
            for item in response.json().get("result", []):
                telemetry = item.get("telemetry") or {}
                mileage = telemetry.get("vehicle.mileage") or {}
                timestamp = telemetry.get("timestamp") or {}
                result[str(item.get("id"))] = {
                    "km": number(mileage.get("value")),
                    "timestamp": timestamp.get("value") or mileage.get("ts"),
                    "error": None,
                }
        except requests.RequestException as error:
            for device_id in chunk:
                result[device_id] = {"km": None, "error": str(error)}
    return result


def fetch_settings_batch(device_ids, token):
    result = {}
    ids = [str(value) for value in device_ids if value]
    headers = {"Authorization": f"FlespiToken {token}"}
    for index in range(0, len(ids), 80):
        chunk = ids[index:index + 80]
        url = f"https://flespi.io/gw/devices/{','.join(chunk)}/settings/odometer_fmb"
        try:
            response = HTTP.get(url, headers=headers, timeout=30)
            if not response.ok:
                continue
            for item in response.json().get("result", []):
                current = item.get("current") or {}
                pending = item.get("pending")
                pending_properties = (pending or {}).get("properties") if isinstance(pending, dict) else None
                result[str(item.get("device_id"))] = {
                    "value": number(current.get("value")),
                    "updated": number(item.get("updated")),
                    "pending": pending_properties or pending,
                    "status": item.get("status"),
                }
        except requests.RequestException:
            continue
    return result


def calculate_device(device_id, start_date, now, token):
    if not device_id or not start_date:
        return {"error": "SIN ID FLESPI O FECHA EFECTIVA"}
    cache_key = f"{device_id}|{start_date.isoformat()}|{now.astimezone(EC_TZ).date().isoformat()}"
    cached = read_cache(cache_key)
    if cached:
        return cached

    query = {
        "from": int(start_date.timestamp()),
        "to": int(now.timestamp()),
        "timezone": "America/Guayaquil",
        "selectors": [{"type": "datetime", "split": "day"}],
        "counters": [
            {"type": "expression", "name": "odo_first", "expression": "vehicle.mileage", "method": "first"},
            {"type": "expression", "name": "odo_last", "expression": "vehicle.mileage", "method": "last"},
            {"type": "expression", "name": "gps_km", "expression": "mileage()", "method": "summary"},
        ],
    }
    url = f"https://flespi.io/gw/devices/{device_id}/messages/calculate"
    response = None
    for attempt in range(3):
        try:
            response = HTTP.get(
                url,
                params={"data": json.dumps(query, separators=(",", ":"))},
                headers={"Authorization": f"FlespiToken {token}"},
                timeout=90,
            )
            if response.status_code != 429:
                break
            time.sleep(2 + attempt * 2)
        except requests.RequestException as error:
            if attempt == 2:
                return {"error": str(error)}
            time.sleep(1 + attempt)
    if response is None or not response.ok:
        code = response.status_code if response is not None else "SIN RESPUESTA"
        return {"error": f"HTTP {code}"}

    intervals = response.json().get("result", [])
    total = 0.0
    days_with_data = 0
    matched_days = 0
    gps_jump_days = 0
    odometer_jump_days = 0
    resets = 0
    used_gps_days = 0
    first_odometer = None
    last_odometer = None

    for item in intervals:
        odo_first = number(item.get("odo_first"))
        odo_last = number(item.get("odo_last"))
        gps_km = number(item.get("gps_km"))
        if odo_first is None and odo_last is None and gps_km is None:
            continue
        days_with_data += 1
        if first_odometer is None and odo_first is not None:
            first_odometer = odo_first
        if odo_last is not None:
            last_odometer = odo_last

        odo_delta = None if odo_first is None or odo_last is None else odo_last - odo_first
        odo_valid = odo_delta is not None and 0 <= odo_delta <= MAX_DAILY_KM
        gps_valid = gps_km is not None and 0 <= gps_km <= MAX_DAILY_KM

        if odo_delta is not None and odo_delta < -1:
            resets += 1
        elif odo_delta is not None and odo_delta > MAX_DAILY_KM:
            odometer_jump_days += 1
        if gps_km is not None and gps_km > MAX_DAILY_KM:
            gps_jump_days += 1

        if odo_valid:
            total += odo_delta
            if gps_valid and odo_delta > 5 and 0.65 <= gps_km / odo_delta <= 1.35:
                matched_days += 1
        elif gps_valid:
            total += gps_km
            used_gps_days += 1

    if not days_with_data:
        return {"error": "SIN HISTORIAL DESDE LA FECHA EFECTIVA"}

    confidence_ratio = matched_days / max(1, days_with_data)
    if days_with_data >= 30 and confidence_ratio >= 0.35 and resets == 0:
        confidence = "ALTA"
    elif days_with_data >= 7 and confidence_ratio >= 0.15:
        confidence = "MEDIA"
    else:
        confidence = "BAJA"

    result = {
        "error": None,
        "distance": total,
        "first_odometer": first_odometer,
        "last_odometer": last_odometer,
        "days": days_with_data,
        "matched_days": matched_days,
        "gps_jump_days": gps_jump_days,
        "odometer_jump_days": odometer_jump_days,
        "resets": resets,
        "used_gps_days": used_gps_days,
        "confidence": confidence,
        "cache": False,
    }
    write_cache(cache_key, result)
    return result


def write_result(headers, rows, output_path):
    result_fields = [
        "FECHA_EFECTIVA_CALCULO", "KM_INICIAL_USADO", "KM_RECORRIDO_DEPURADO",
        "KM_ODOMETRO_APROXIMADO", "KM_ODOMETRO_ACTUAL", "DIFERENCIA_APROX_VS_ACTUAL",
        "METODO_APROXIMACION",
        "COMANDO_ODOMETRO", "FECHA_ULTIMO_COMANDO", "VALOR_ULTIMO_COMANDO",
        "ESTADO_CONFIGURACION_FLESPI", "VALOR_CONFIGURADO_FLESPI", "FECHA_CONFIGURACION_FLESPI",
        "DIAS_ANALIZADOS", "DIAS_COINCIDENCIA_GPS", "DIAS_SALTO_GPS",
        "DIAS_SALTO_ODOMETRO", "REINICIOS_DETECTADOS", "DIAS_RESPALDO_GPS",
        "CONFIANZA", "ESTADO", "OBSERVACION",
    ]
    output_headers = [*headers, *[field for field in result_fields if field not in headers]]
    workbook = xlsxwriter.Workbook(str(output_path))
    summary = workbook.add_worksheet("Resumen")
    detail = workbook.add_worksheet("Odometros")
    title = workbook.add_format({"bold": True, "font_size": 18, "font_color": "#17324D"})
    header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#0F766E", "border": 1, "text_wrap": True})
    number_fmt = workbook.add_format({"num_format": "#,##0.00"})
    good = workbook.add_format({"bg_color": "#DCFCE7", "font_color": "#166534"})
    review = workbook.add_format({"bg_color": "#FEF3C7", "font_color": "#92400E"})
    bad = workbook.add_format({"bg_color": "#FEE2E2", "font_color": "#991B1B"})

    counts = {}
    for row in rows:
        counts[row.get("ESTADO") or "SIN ESTADO"] = counts.get(row.get("ESTADO") or "SIN ESTADO", 0) + 1
    summary.write("A1", "Resumen de odómetros aproximados", title)
    summary.write("A3", "Indicador", header)
    summary.write("B3", "Cantidad", header)
    metrics = [("Total procesados", len(rows)), *sorted(counts.items())]
    for index, (label, value) in enumerate(metrics, start=4):
        summary.write(index - 1, 0, label)
        summary.write_number(index - 1, 1, value)
    summary.set_column("A:A", 36)
    summary.set_column("B:B", 14)

    for column, name in enumerate(output_headers):
        detail.write(0, column, name, header)
    for output_row, row in enumerate(rows, start=1):
        for column, name in enumerate(output_headers):
            value = row.get(name)
            if isinstance(value, (int, float)) and (name.startswith("KM_") or name.startswith("DIFERENCIA")):
                detail.write_number(output_row, column, value, number_fmt)
            else:
                detail.write(output_row, column, value)
    detail.freeze_panes(1, 0)
    detail.autofilter(0, 0, len(rows), len(output_headers) - 1)
    detail.set_column(0, len(output_headers) - 1, 19)
    for name, width in {"CHASIS": 22, "MODELO": 30, "OBSERVACION": 70, "VALIDACION_ANIO_VEHICULO": 30}.items():
        if name in output_headers:
            index = output_headers.index(name)
            detail.set_column(index, index, width)
    if "ESTADO" in output_headers and rows:
        index = output_headers.index("ESTADO")
        detail.conditional_format(1, index, len(rows), index, {"type": "text", "criteria": "containing", "value": "CONFIABLE", "format": good})
        detail.conditional_format(1, index, len(rows), index, {"type": "text", "criteria": "containing", "value": "REVISAR", "format": review})
        detail.conditional_format(1, index, len(rows), index, {"type": "text", "criteria": "containing", "value": "SIN DATOS", "format": bad})
    workbook.close()


def main():
    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    config = load_config()
    token = config["flespi"]["token"]
    headers, rows = read_xlsx_rows(input_path)
    now = datetime.now(timezone.utc)
    for row in rows:
        row["_device_id"] = str(row.get("ID_FLESPI") or "").strip()
        row["_start"] = effective_start(row)
    ids = sorted({row["_device_id"] for row in rows if row["_device_id"]})
    current = fetch_current_batch(ids, token)
    settings = fetch_settings_batch(ids, token)
    keys = sorted({(row["_device_id"], row["_start"]) for row in rows if row["_device_id"] and row["_start"]}, key=lambda value: (value[0], value[1].isoformat()))
    reports = {}
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(calculate_device, device_id, start, now, token): (device_id, start) for device_id, start in keys}
        for index, future in enumerate(as_completed(futures), start=1):
            reports[futures[future]] = future.result()
            if index % 10 == 0 or index == len(keys):
                print(f"Procesados: {index}/{len(keys)}", flush=True)

    output_rows = []
    for row in rows:
        output = {key: row.get(key) for key in headers}
        device_id = row["_device_id"]
        start = row["_start"]
        report = reports.get((device_id, start), {})
        current_km = (current.get(device_id) or {}).get("km")
        command_executed = str(row.get("COMANDO_ODOMETRO_EJECUTADO") or "").strip() in {"1", "SI", "True", "true"}
        setting = settings.get(device_id) or {}
        setting_value = setting.get("value")
        setting_updated = setting.get("updated")
        setting_after_start = bool(setting_updated and start and setting_updated >= start.timestamp())
        setting_pending = bool(setting.get("pending"))
        flespi_confirmed = setting_value is not None and setting_after_start and not setting_pending
        initial_input = parse_float(row.get("KM_INICIAL"))
        detected_base = report.get("first_odometer")
        initial_km = initial_input if initial_input is not None else max(0, detected_base or 0)
        distance = report.get("distance")
        reconstructed = initial_km + distance if distance is not None else None
        if (command_executed or flespi_confirmed) and current_km is not None:
            approximate = current_km
            method = "ODOMETRO CONFIGURADO POR COMANDO"
        else:
            approximate = reconstructed
            method = "RECORRIDO DEPURADO DESDE FECHA EFECTIVA"

        notes = []
        if report.get("error"):
            notes.append(report["error"])
        if report.get("gps_jump_days"):
            notes.append(f"{report['gps_jump_days']} dias con saltos GPS excluidos")
        if report.get("odometer_jump_days"):
            notes.append(f"{report['odometer_jump_days']} dias con saltos de odometro excluidos")
        if report.get("resets"):
            notes.append(f"{report['resets']} reinicios detectados")
        if report.get("cache"):
            notes.append("Resultado reutilizado desde cache local")
        if initial_input is None:
            notes.append("Sin KM_INICIAL; se usa la primera lectura historica disponible")

        if report.get("error") or approximate is None:
            status = "SIN DATOS"
        elif command_executed or flespi_confirmed:
            status = "COMANDO CONFIRMADO"
        elif report.get("confidence") == "ALTA":
            status = "APROXIMADO CONFIABLE"
        else:
            status = "REVISAR APROXIMACION"

        if command_executed:
            command_status = "EJECUTADO SEGUN VAPOR"
        elif flespi_confirmed:
            command_status = "CONFIRMADO EN FLESPI"
        elif setting_pending:
            command_status = "PENDIENTE EN FLESPI"
        else:
            command_status = "SIN EVIDENCIA"
        if setting_pending:
            flespi_status = "PENDIENTE"
        elif flespi_confirmed:
            flespi_status = "CONFIRMADO"
        elif setting_value is not None:
            flespi_status = "CONFIGURACION ANTERIOR A ASIGNACION"
        else:
            flespi_status = "SIN DATOS"
        setting_date = (
            datetime.fromtimestamp(setting_updated, timezone.utc).astimezone(EC_TZ).strftime("%Y-%m-%d %H:%M:%S")
            if setting_updated else None
        )
        difference = approximate - current_km if approximate is not None and current_km is not None else None
        output.update({
            "FECHA_EFECTIVA_CALCULO": start.astimezone(EC_TZ).strftime("%Y-%m-%d %H:%M:%S") if start else None,
            "KM_INICIAL_USADO": round(initial_km, 2) if initial_km is not None else None,
            "KM_RECORRIDO_DEPURADO": round(distance, 2) if distance is not None else None,
            "KM_ODOMETRO_APROXIMADO": round(approximate, 2) if approximate is not None else None,
            "KM_ODOMETRO_ACTUAL": round(current_km, 2) if current_km is not None else None,
            "DIFERENCIA_APROX_VS_ACTUAL": round(difference, 2) if difference is not None else None,
            "METODO_APROXIMACION": method,
            "COMANDO_ODOMETRO": command_status,
            "FECHA_ULTIMO_COMANDO": row.get("FECHA_ULTIMO_COMANDO"),
            "VALOR_ULTIMO_COMANDO": row.get("VALOR_ULTIMO_COMANDO"),
            "ESTADO_CONFIGURACION_FLESPI": flespi_status,
            "VALOR_CONFIGURADO_FLESPI": round(setting_value, 2) if setting_value is not None else None,
            "FECHA_CONFIGURACION_FLESPI": setting_date,
            "DIAS_ANALIZADOS": report.get("days"),
            "DIAS_COINCIDENCIA_GPS": report.get("matched_days"),
            "DIAS_SALTO_GPS": report.get("gps_jump_days"),
            "DIAS_SALTO_ODOMETRO": report.get("odometer_jump_days"),
            "REINICIOS_DETECTADOS": report.get("resets"),
            "DIAS_RESPALDO_GPS": report.get("used_gps_days"),
            "CONFIANZA": report.get("confidence"),
            "ESTADO": status,
            "OBSERVACION": " | ".join(notes),
        })
        output_rows.append(output)
    write_result(headers, output_rows, output_path)
    print(f"Archivo generado: {output_path}", flush=True)


if __name__ == "__main__":
    main()
