import json
import os
import re
import subprocess
import sys
import uuid
from pathlib import Path

import pymysql
import xlsxwriter
from flask import Flask, jsonify, render_template_string, request, send_file
from werkzeug.utils import secure_filename

from calcular_km_actual_mantenimiento import read_xlsx_rows


APP_DIR = Path(__file__).resolve().parent
JOBS_DIR = APP_DIR / "web_jobs_odometro_real"
CALC_SCRIPT = APP_DIR / "calcular_odometro_real.py"
OUTPUT_NAME = "ODOMETROS_REALES_APROXIMADOS.xlsx"
TEMPLATE_NAME = "PLANTILLA_ODOMETRO_REAL.xlsx"
MAX_ROWS = 5000
app = Flask(__name__)
jobs = {}


def normalize_chassis(value):
    return re.sub(r"\s+", "", str(value or "").strip()).upper()


def db_connection():
    config = json.loads((APP_DIR / "config.json").read_text(encoding="utf-8"))
    mysql = config["mysql"]
    return pymysql.connect(
        host=mysql["host"], user=mysql["user"], password=mysql["password"],
        port=int(mysql.get("port", 3306)), database=config.get("db_origen", "vapor"),
        charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor,
        connect_timeout=20, read_timeout=90,
    )


def resolve_chassis(chassis_values):
    chassis = sorted({normalize_chassis(value) for value in chassis_values if normalize_chassis(value)})
    if not chassis:
        return {}
    placeholders = ",".join(["%s"] * len(chassis))
    sql = f"""
      SELECT * FROM (
        SELECT
          UPPER(TRIM(v.chasis)) AS chasis_key,
          v.plate AS placa,
          v.model AS modelo,
          v.year AS anio_vehiculo,
          v.created_at AS fecha_creacion_vehiculo,
          d.id AS device_uuid,
          d.imei,
          d.flespi_device_id,
          d.created_at AS fecha_creacion_dispositivo,
          vd.created_at AS fecha_creacion_asignacion,
          vd.installation_date AS fecha_instalacion_dispositivo,
          vd.start_date AS fecha_inicio_dispositivo,
          vd.vehicle_device_status,
          ROW_NUMBER() OVER (
            PARTITION BY UPPER(TRIM(v.chasis))
            ORDER BY (vd.deleted_at IS NULL) DESC,
              COALESCE(vd.installation_date, vd.start_date, vd.created_at) DESC,
              vd.updated_at DESC
          ) AS rn
        FROM vehicles v
        INNER JOIN vehicle_devices vd ON vd.vehicle_id = v.id
        INNER JOIN devices d ON d.id = vd.device_id AND d.deleted_at IS NULL
        WHERE UPPER(TRIM(v.chasis)) IN ({placeholders})
      ) source
      WHERE rn = 1
    """
    with db_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(sql, chassis)
            devices = {normalize_chassis(row["chasis_key"]): row for row in cursor.fetchall()}
            uuids = [row["device_uuid"] for row in devices.values() if row.get("device_uuid")]
            if uuids:
                command_placeholders = ",".join(["%s"] * len(uuids))
                cursor.execute(
                    f"""
                    SELECT * FROM (
                      SELECT device_id, command_value, command_executed, created_at,
                        ROW_NUMBER() OVER (PARTITION BY device_id ORDER BY created_at DESC) AS rn
                      FROM command_logs
                      WHERE device_id IN ({command_placeholders})
                        AND command_type LIKE 'Actualizar Od%%metro'
                        AND deleted_at IS NULL
                    ) commands
                    WHERE rn = 1
                    """,
                    uuids,
                )
                commands = {row["device_id"]: row for row in cursor.fetchall()}
                for device in devices.values():
                    command = commands.get(device.get("device_uuid")) or {}
                    device["comando_odometro_ejecutado"] = command.get("command_executed")
                    device["fecha_ultimo_comando"] = command.get("created_at")
                    device["valor_ultimo_comando"] = command.get("command_value")
            return devices


def write_workbook(headers, rows, output_path):
    workbook = xlsxwriter.Workbook(str(output_path))
    worksheet = workbook.add_worksheet("Entrada")
    header = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    for column, name in enumerate(headers):
        worksheet.write(0, column, name, header)
    for output_row, row in enumerate(rows, start=1):
        for column, name in enumerate(headers):
            value = row.get(name)
            if hasattr(value, "isoformat"):
                value = value.isoformat(sep=" ")
            worksheet.write(output_row, column, value)
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(rows), len(headers) - 1)
    worksheet.set_column(0, len(headers) - 1, 22)
    workbook.close()


def enrich_input(input_path, output_path):
    headers, rows = read_xlsx_rows(input_path)
    headers = headers or []
    if "CHASIS" not in headers:
        raise ValueError("El archivo debe incluir la columna CHASIS")
    if not rows:
        raise ValueError("El archivo no contiene chasis")
    if len(rows) > MAX_ROWS:
        raise ValueError(f"El archivo supera el limite de {MAX_ROWS} filas")
    fields = [
        "PLACA", "MODELO", "ANIO_VEHICULO", "IMEI", "ID_FLESPI",
        "FECHA_CREACION_VEHICULO", "FECHA_CREACION_DISPOSITIVO",
        "FECHA_CREACION_ASIGNACION", "FECHA_INSTALACION_DISPOSITIVO",
        "FECHA_INICIO_DISPOSITIVO", "ESTADO_ASIGNACION",
        "COMANDO_ODOMETRO_EJECUTADO", "FECHA_ULTIMO_COMANDO", "VALOR_ULTIMO_COMANDO",
        "ESTADO_BUSQUEDA_CHASIS",
    ]
    output_headers = [*headers, *[field for field in fields if field not in headers]]
    devices = resolve_chassis([row.get("CHASIS") for row in rows])
    resolved = 0
    for row in rows:
        device = devices.get(normalize_chassis(row.get("CHASIS")))
        if not device:
            row["ESTADO_BUSQUEDA_CHASIS"] = "SIN DISPOSITIVO"
            continue
        resolved += 1
        mappings = {
            "PLACA": "placa", "MODELO": "modelo", "ANIO_VEHICULO": "anio_vehiculo",
            "IMEI": "imei", "ID_FLESPI": "flespi_device_id",
            "FECHA_CREACION_VEHICULO": "fecha_creacion_vehiculo",
            "FECHA_CREACION_DISPOSITIVO": "fecha_creacion_dispositivo",
            "FECHA_CREACION_ASIGNACION": "fecha_creacion_asignacion",
            "FECHA_INSTALACION_DISPOSITIVO": "fecha_instalacion_dispositivo",
            "FECHA_INICIO_DISPOSITIVO": "fecha_inicio_dispositivo",
            "ESTADO_ASIGNACION": "vehicle_device_status",
            "COMANDO_ODOMETRO_EJECUTADO": "comando_odometro_ejecutado",
            "FECHA_ULTIMO_COMANDO": "fecha_ultimo_comando",
            "VALOR_ULTIMO_COMANDO": "valor_ultimo_comando",
        }
        for target, source in mappings.items():
            row[target] = row.get(target) or device.get(source)
        row["ESTADO_BUSQUEDA_CHASIS"] = "ENCONTRADO"
    write_workbook(output_headers, rows, output_path)
    return {"rows": len(rows), "resolved": resolved, "unresolved": len(rows) - resolved}


def ensure_template():
    path = APP_DIR / TEMPLATE_NAME
    if path.exists():
        return path
    write_workbook(["CHASIS", "KM_INICIAL"], [{"CHASIS": "EJEMPLO_CHASIS", "KM_INICIAL": None}], path)
    return path


def parse_progress(log_text):
    matches = re.findall(r"Procesados:\s*(\d+)/(\d+)", log_text)
    if not matches:
        return 0, 0, 0
    done, total = map(int, matches[-1])
    return done, total, int(done * 100 / max(1, total))


PAGE = r"""
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Odómetro real aproximado</title>
  <style>
    :root{--ink:#102a43;--muted:#62748a;--blue:#155eef;--teal:#0f8b7f;--line:#d7e1ea;--bg:#f3f7fa;--red:#b42318}
    *{box-sizing:border-box} body{margin:0;background:var(--bg);font-family:Arial,sans-serif;color:var(--ink)}
    .sidebar{position:fixed;inset:0 auto 0 0;width:248px;padding:24px 16px;background:#0f2742;color:#fff;display:flex;flex-direction:column;z-index:10}
    .brand{padding:4px 10px 22px;border-bottom:1px solid #ffffff24}.brand b{font-size:18px}.brand small{display:block;margin-top:8px;color:#a9bdd1}
    .navLabel{margin:22px 10px 9px;color:#91a8bd;font-size:12px;font-weight:700;text-transform:uppercase}.sideNav{display:grid;gap:7px}
    .sideNav a{display:flex;align-items:center;gap:11px;padding:12px 13px;border-radius:10px;color:#d9e8f5;text-decoration:none;font-weight:700;font-size:14px}
    .sideNav a:hover{background:#ffffff12}.sideNav a.active{color:#fff;background:linear-gradient(135deg,#155eef,#087f5b)}
    .navIcon{width:28px;height:28px;display:grid;place-items:center;border-radius:8px;background:#ffffff16;font-size:12px}.foot{margin-top:auto;color:#a9bdd1;font-size:12px}
    main{margin-left:248px;padding:28px;max-width:1500px}.top{display:flex;justify-content:space-between;align-items:center;margin-bottom:18px}h1{margin:0;font-size:27px}.status{color:var(--muted)}
    .hero{background:linear-gradient(120deg,#102a43,#155eef 55%,#0f8b7f);color:white;padding:25px;border-radius:16px;box-shadow:0 12px 30px #155eef25}.hero h2{margin:0 0 8px}.hero p{margin:0;color:#e5efff;max-width:950px;line-height:1.5}
    .steps{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:18px}.step{background:#ffffff17;border:1px solid #ffffff30;padding:12px;border-radius:10px;font-size:13px}.step b{display:block;margin-bottom:5px}
    .panel{background:#fff;border:1px solid var(--line);border-radius:12px;padding:20px;margin-top:16px}form{display:grid;grid-template-columns:1fr auto auto;gap:12px;align-items:center}
    input[type=file]{padding:15px;border:1px dashed #aebdcc;border-radius:9px;width:100%;background:#fbfdff}button,.download{border:0;border-radius:8px;padding:13px 17px;background:var(--blue);color:#fff;font-weight:700;text-decoration:none;cursor:pointer}.download{background:var(--teal)}button:disabled{opacity:.55}
    .muted{color:var(--muted);font-size:14px;line-height:1.5}.progress{height:18px;border-radius:999px;background:#e6edf3;overflow:hidden;margin:14px 0}.bar{height:100%;width:0;background:linear-gradient(90deg,#155eef,#14b8a6);transition:.2s}
    .metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.metric{border:1px solid var(--line);padding:13px;border-radius:10px;background:#fbfdff}.metric b{display:block;font-size:20px}.metric span{font-size:12px;color:var(--muted)}pre{background:#0b1728;color:#cfe1f3;padding:16px;border-radius:10px;min-height:150px;white-space:pre-wrap;max-height:330px;overflow:auto}.error{color:var(--red);font-weight:700}
    @media(max-width:900px){.sidebar{position:static;width:auto}.foot,.navLabel{display:none}.sideNav{grid-template-columns:repeat(4,1fr)}main{margin:0;padding:18px}.steps,.metrics{grid-template-columns:repeat(2,1fr)}form{grid-template-columns:1fr}}
  </style>
</head>
<body>
<aside class="sidebar"><div class="brand"><b>KM · Odómetros</b><small>TITANIA · CENTRO OPERATIVO</small></div><div class="navLabel">Módulos</div><nav class="sideNav">
  <a href="http://127.0.0.1:8062"><span class="navIcon">IN</span>Inicio</a>
  <a href="http://127.0.0.1:8060"><span class="navIcon">RC</span>Recepción</a>
  <a href="http://127.0.0.1:8061"><span class="navIcon">DT</span>Distancia</a>
  <a class="active" href="http://127.0.0.1:8063"><span class="navIcon">OR</span>Odómetro real</a>
</nav><div class="foot">Análisis masivo sin modificar dispositivos</div></aside>
<main><div class="top"><h1>Odómetro real aproximado</h1><div class="status" id="status">Sin proceso activo</div></div>
<section class="hero"><h2>Análisis masivo de hasta 5.000 chasis</h2><p>Reconstruye el recorrido desde la creación/asignación, elimina saltos GPS y de odómetro, revisa si existe un comando de actualización y compara el aproximado contra el valor actual.</p>
<div class="steps"><div class="step"><b>1. Carga</b>CHASIS y KM_INICIAL opcional</div><div class="step"><b>2. Vinculación</b>Vehículo, dispositivo y fechas</div><div class="step"><b>3. Depuración</b>Saltos, reinicios y respaldo GPS</div><div class="step"><b>4. Resultado</b>Actual vs. aproximado y confianza</div></div></section>
<section class="panel"><form id="form"><input id="file" type="file" accept=".xlsx" required><button type="submit" id="start">Analizar archivo</button><a class="download" href="/api/template">Descargar plantilla</a></form><p class="muted">No envía comandos ni actualiza Flespi. El Excel incluye evidencia del último comando de odómetro registrado en Vapor.</p></section>
<section class="panel"><div id="message">Esperando archivo…</div><div class="progress"><div class="bar" id="bar"></div></div><div class="metrics"><div class="metric"><b id="percent">0%</b><span>Avance</span></div><div class="metric"><b id="done">0</b><span>Procesados</span></div><div class="metric"><b id="total">0</b><span>Total</span></div><div class="metric"><b id="phase">-</b><span>Estado</span></div></div><div id="download" style="margin-top:18px"></div></section>
<section class="panel"><pre id="log">Sin actividad.</pre></section></main>
<script>
const form=document.getElementById('form'), file=document.getElementById('file'), start=document.getElementById('start'); let timer=null;
form.addEventListener('submit',async e=>{e.preventDefault();if(!file.files.length)return;start.disabled=true;document.getElementById('download').innerHTML='';const data=new FormData();data.append('file',file.files[0]);
 const response=await fetch('/api/jobs',{method:'POST',body:data});const body=await response.json();if(!response.ok){start.disabled=false;document.getElementById('message').innerHTML=`<span class="error">${body.error||'No se pudo iniciar'}</span>`;return}document.getElementById('status').textContent=`Proceso ${body.job_id}`;poll(body.job_id);timer=setInterval(()=>poll(body.job_id),1500)});
async function poll(id){const response=await fetch(`/api/jobs/${id}`),body=await response.json();document.getElementById('bar').style.width=`${body.percent||0}%`;document.getElementById('percent').textContent=`${body.percent||0}%`;document.getElementById('done').textContent=body.processed||0;document.getElementById('total').textContent=body.total||0;document.getElementById('phase').textContent=body.done?'Terminado':'Procesando';document.getElementById('message').textContent=body.message||'';document.getElementById('log').textContent=body.log||'Procesando…';if(body.done){clearInterval(timer);start.disabled=false;if(body.success)document.getElementById('download').innerHTML=`<a class="download" href="/api/jobs/${id}/download">Descargar Excel analizado</a>`}}
</script></body></html>
"""


@app.after_request
def no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response


@app.get("/")
def index():
    return render_template_string(PAGE)


@app.get("/api/template")
def template():
    return send_file(ensure_template(), as_attachment=True, download_name=TEMPLATE_NAME)


@app.post("/api/jobs")
def create_job():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename or Path(uploaded.filename).suffix.lower() != ".xlsx":
        return jsonify({"error": "Selecciona un archivo .xlsx"}), 400
    job_id = uuid.uuid4().hex[:10]
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    raw_path = job_dir / (secure_filename(uploaded.filename) or "CHASIS.xlsx")
    input_path = job_dir / "ENTRADA_ENRIQUECIDA.xlsx"
    output_path = job_dir / OUTPUT_NAME
    log_path = job_dir / "progreso.log"
    error_path = job_dir / "errores.log"
    uploaded.save(raw_path)
    try:
        resolution = enrich_input(raw_path, input_path)
    except Exception as error:
        return jsonify({"error": str(error)}), 400
    log_file = log_path.open("w", encoding="utf-8", buffering=1)
    error_file = error_path.open("w", encoding="utf-8", buffering=1)
    process = subprocess.Popen(
        [sys.executable, "-u", "-B", str(CALC_SCRIPT), str(input_path), str(output_path)],
        cwd=APP_DIR, stdout=log_file, stderr=error_file,
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
    )
    jobs[job_id] = {"process": process, "output": output_path, "log": log_path, "error": error_path, "log_file": log_file, "error_file": error_file, "resolution": resolution}
    return jsonify({"job_id": job_id, **resolution})


@app.get("/api/jobs/<job_id>")
def job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Proceso no encontrado"}), 404
    code = job["process"].poll()
    done = code is not None
    if done and not job.get("closed"):
        job["log_file"].close(); job["error_file"].close(); job["closed"] = True
    log_text = job["log"].read_text(encoding="utf-8", errors="replace") if job["log"].exists() else ""
    error_text = job["error"].read_text(encoding="utf-8", errors="replace") if job["error"].exists() else ""
    processed, total, percent = parse_progress(log_text)
    if done:
        percent = 100 if code == 0 else percent
    resolution = job["resolution"]
    message = (
        f"Procesando {processed}/{total or resolution['resolved']} dispositivos"
        if not done else
        (f"Terminado: {resolution['resolved']} encontrados y {resolution['unresolved']} sin dispositivo" if code == 0 else "El proceso terminó con error")
    )
    combined = log_text + (("\nERRORES:\n" + error_text) if error_text.strip() else "")
    return jsonify({"done": done, "success": code == 0 and job["output"].exists(), "percent": percent, "processed": processed, "total": total or resolution["resolved"], "message": message, "log": combined[-12000:]})


@app.get("/api/jobs/<job_id>/download")
def download(job_id):
    job = jobs.get(job_id)
    if not job or not job["output"].exists():
        return jsonify({"error": "Archivo no disponible"}), 404
    return send_file(job["output"], as_attachment=True, download_name=OUTPUT_NAME)


if __name__ == "__main__":
    JOBS_DIR.mkdir(parents=True, exist_ok=True)
    ensure_template()
    app.run(host="127.0.0.1", port=8063, debug=False)
