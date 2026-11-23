import json
import os
import re
import subprocess
import sys
import threading
import uuid
from collections import Counter
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import requests
import pymysql
from flask import Flask, jsonify, render_template_string, request, send_file
from werkzeug.utils import secure_filename
import xlsxwriter

from calcular_km_actual_mantenimiento import KM_BASE_COLUMNS, first_km_base, parse_date, parse_float, read_xlsx_rows


APP_DIR = Path(__file__).resolve().parent
JOBS_DIR = APP_DIR / "web_jobs_mantenimiento"
CALC_SCRIPT = APP_DIR / "calcular_km_actual_mantenimiento.py"
OUTPUT_NAME = "MANTENIMIENTO_KM_ACTUAL.xlsx"
TEMPLATE_NAME = "PLANTILLA_MANTENIMIENTO_ODOMETROS.xlsx"
ALLOWED_EXTENSIONS = {".xlsx"}

app = Flask(__name__)
jobs = {}
update_jobs = {}
update_lock = threading.Lock()


def normalize_chassis(value):
    """Normaliza el chasis para buscarlo exactamente en Vapor."""
    return re.sub(r"\s+", "", str(value or "").strip()).upper()


def resolve_devices_by_chassis(chassis_values):
    """Obtiene el equipo activo (o la última asignación) por chasis desde Vapor."""
    chassis = sorted({normalize_chassis(value) for value in chassis_values if normalize_chassis(value)})
    if not chassis:
        return {}

    config = json.loads((APP_DIR / "config.json").read_text(encoding="utf-8"))
    mysql = config["mysql"]
    placeholders = ",".join(["%s"] * len(chassis))
    sql = f"""
      SELECT * FROM (
        SELECT UPPER(TRIM(v.chasis)) chasis_key, v.plate AS placa, v.model AS modelo, d.imei, d.flespi_device_id,
          ROW_NUMBER() OVER (
            PARTITION BY UPPER(TRIM(v.chasis))
            ORDER BY (vd.deleted_at IS NULL) DESC,
              COALESCE(vd.installation_date, vd.created_at) DESC, vd.updated_at DESC
          ) AS rn
        FROM vehicles v
        INNER JOIN vehicle_devices vd ON vd.vehicle_id = v.id
        INNER JOIN devices d ON d.id = vd.device_id AND d.deleted_at IS NULL
        WHERE UPPER(TRIM(v.chasis)) IN ({placeholders})
      ) asignaciones
      WHERE rn = 1
    """
    connection = pymysql.connect(
        host=mysql["host"], user=mysql["user"], password=mysql["password"],
        port=int(mysql.get("port", 3306)), database=config.get("db_origen", "vapor"),
        charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor, connect_timeout=15,
    )
    try:
        with connection.cursor() as cursor:
            cursor.execute(sql, chassis)
            return {normalize_chassis(row["chasis_key"]): row for row in cursor.fetchall()}
    finally:
        connection.close()


def write_workbook(headers, rows, output_path, sheet_name="MANTENIMIENTO"):
    workbook = xlsxwriter.Workbook(str(output_path))
    worksheet = workbook.add_worksheet(sheet_name)
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    for column, name in enumerate(headers):
        worksheet.write(0, column, name, header_format)
    for output_row, row in enumerate(rows, start=1):
        for column, name in enumerate(headers):
            worksheet.write(output_row, column, row.get(name))
    worksheet.autofilter(0, 0, len(rows), len(headers) - 1)
    worksheet.freeze_panes(1, 0)
    worksheet.set_column(0, len(headers) - 1, 18)
    workbook.close()


def enrich_workbook_by_chassis(input_path, output_path):
    """Completa ID_FLESPI, IMEI y datos del vehículo si el archivo trae CHASIS."""
    headers, rows = read_xlsx_rows(input_path)
    headers = headers or []
    if "CHASIS" not in headers:
        return input_path, {"resolved": 0, "unresolved": 0, "message": "El archivo no tiene columna CHASIS."}

    device_fields = ["ID_FLESPI", "IMEI", "PLACA", "MODELO", "ESTADO_BUSQUEDA_CHASIS"]
    enriched_headers = [*headers, *[field for field in device_fields if field not in headers]]
    devices = resolve_devices_by_chassis([row.get("CHASIS") for row in rows])
    resolved = 0
    unresolved = 0
    for row in rows:
        device = devices.get(normalize_chassis(row.get("CHASIS")))
        if device:
            resolved += 1
            row["ID_FLESPI"] = row.get("ID_FLESPI") or device.get("flespi_device_id")
            row["IMEI"] = row.get("IMEI") or device.get("imei")
            row["PLACA"] = row.get("PLACA") or device.get("placa")
            row["MODELO"] = row.get("MODELO") or device.get("modelo")
            row["ESTADO_BUSQUEDA_CHASIS"] = "ENCONTRADO"
        else:
            unresolved += 1
            row["ESTADO_BUSQUEDA_CHASIS"] = "SIN DISPOSITIVO EN VAPOR"
    write_workbook(enriched_headers, rows, output_path)
    return output_path, {"resolved": resolved, "unresolved": unresolved, "message": f"{resolved} chasis vinculados con dispositivo; {unresolved} sin coincidencia."}


@app.after_request
def disable_browser_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


PAGE = """
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Odometros mantenimiento</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f5f7fb;
      --panel: #ffffff;
      --ink: #172033;
      --muted: #647089;
      --line: #d9dfeb;
      --blue: #155eef;
      --green: #168a4a;
      --red: #b42318;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Arial, sans-serif;
      background: var(--bg);
      color: var(--ink);
    }
    .sidebar {
      position: fixed;
      inset: 0 auto 0 0;
      width: 248px;
      padding: 24px 16px;
      background: #0f2742;
      color: white;
      display: flex;
      flex-direction: column;
      z-index: 10;
      box-shadow: 10px 0 30px rgba(15,39,66,.12);
    }
    .brand { padding: 4px 10px 22px; border-bottom: 1px solid rgba(255,255,255,.14); }
    .brandMark { display:flex; align-items:center; gap:10px; font-weight:800; font-size:18px; }
    .brandIcon { width:38px; height:38px; display:grid; place-items:center; border-radius:11px; background:linear-gradient(135deg,#1c7ed6,#20c997); }
    .brand small { display:block; margin-top:7px; color:#a9bdd1; font-size:12px; letter-spacing:.8px; }
    .navLabel { margin:22px 10px 9px; color:#91a8bd; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:1px; }
    .sideNav { display:grid; gap:7px; }
    .sideNav a { display:flex; align-items:center; gap:11px; padding:12px 13px; border-radius:10px; color:#d9e8f5; text-decoration:none; font-size:14px; font-weight:700; border:1px solid transparent; }
    .sideNav a:hover { background:rgba(255,255,255,.08); color:white; }
    .sideNav a.active { color:white; background:linear-gradient(135deg,#155eef,#087f5b); border-color:rgba(255,255,255,.2); box-shadow:0 8px 18px rgba(0,0,0,.18); }
    .navIcon { width:28px; height:28px; display:grid; place-items:center; border-radius:8px; background:rgba(255,255,255,.1); font-size:13px; }
    .sidebarFoot { margin-top:auto; padding:14px 10px 2px; border-top:1px solid rgba(255,255,255,.14); color:#a9bdd1; font-size:12px; line-height:1.5; }
    .wrap {
      width: auto;
      max-width: 1440px;
      margin: 28px 28px 28px 276px;
    }
    header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 18px;
    }
    h1 { font-size: 24px; margin: 0; letter-spacing: 0; }
    .status { color: var(--muted); font-size: 14px; }
    .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
      margin-bottom: 16px;
    }
    form {
      display: grid;
      grid-template-columns: minmax(260px, 1fr) minmax(240px, auto) auto auto;
      gap: 12px;
      align-items: center;
    }
    input[type=file] {
      width: 100%;
      border: 1px dashed #aeb8cb;
      border-radius: 8px;
      padding: 14px;
      background: #fbfcff;
    }
    select {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      background: white;
      min-height: 42px;
    }
    button, .download {
      border: 0;
      border-radius: 7px;
      padding: 12px 16px;
      background: var(--blue);
      color: white;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
    }
    button:disabled { opacity: .55; cursor: wait; }
    .download { background: var(--green); }
    .progressbar {
      width: 100%;
      height: 18px;
      border-radius: 999px;
      overflow: hidden;
      background: #e8edf6;
      border: 1px solid var(--line);
      margin: 12px 0;
    }
    .progressbar div {
      width: 0%;
      height: 100%;
      background: linear-gradient(90deg, #155eef, #168a4a);
      transition: width .25s ease;
    }
    .grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 10px;
      margin-top: 12px;
    }
    .metric {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      background: #fbfcff;
    }
    .metric b { display: block; font-size: 18px; margin-bottom: 3px; }
    .metric span { color: var(--muted); font-size: 12px; }
    .validation {
      display: none;
      margin-top: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fbfcff;
      padding: 14px;
    }
    .validation.show { display: block; }
    .validation.good { border-color: #add9c1; background: #f0fff6; }
    .validation.bad { border-color: #f4b4ad; background: #fff6f5; }
    .validation h3 { margin: 0 0 10px; font-size: 16px; }
    .validation ul { margin: 8px 0 0; padding-left: 20px; }
    .validation li { margin: 4px 0; }
    .validationGrid {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 8px;
      margin-top: 10px;
    }
    .pill {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px;
      background: white;
    }
    .pill b { display: block; font-size: 17px; }
    .pill span { color: var(--muted); font-size: 12px; }
    .selection { display: none; }
    .selection.show { display: block; }
    .selectionTools {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
      margin: 12px 0;
    }
    .selectionTools button { padding: 8px 12px; min-height: 36px; }
    .selectionTools .secondary { background: #475467; }
    .selectionTools input { flex: 1; min-width: 220px; padding: 9px; border: 1px solid var(--line); border-radius: 7px; }
    .tableWrap { max-height: 430px; overflow: auto; border: 1px solid var(--line); border-radius: 8px; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 9px; border-bottom: 1px solid var(--line); text-align: left; white-space: nowrap; }
    th { position: sticky; top: 0; z-index: 1; background: #eef3fb; }
    tr.review { background: #fff6f5; }
    .badge { display: inline-block; border-radius: 999px; padding: 4px 8px; font-size: 11px; font-weight: 700; }
    .badge.update { color: #05603a; background: #d1fadf; }
    .badge.review { color: #b42318; background: #fee4e2; }
    pre {
      min-height: 280px;
      max-height: 460px;
      overflow: auto;
      white-space: pre-wrap;
      margin: 0;
      padding: 14px;
      border-radius: 8px;
      background: #101828;
      color: #e8edf6;
      font-size: 13px;
      line-height: 1.45;
    }
    #updateProgress { display: none; margin-top: 16px; }
    #updateProgress.show { display: block; }
    #updateLog { min-height: 150px; max-height: 300px; margin-top: 10px; }
    .error { color: var(--red); font-weight: 700; }
    @media (max-width: 760px) {
      form { grid-template-columns: 1fr; }
      .grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      header { align-items: flex-start; flex-direction: column; }
    }
    @media (max-width: 900px) {
      .sidebar { position:static; width:auto; min-height:0; padding:14px; }
      .brand { padding:0 4px 12px; }
      .brand small, .navLabel, .sidebarFoot { display:none; }
      .sideNav { grid-template-columns:repeat(3,minmax(0,1fr)); }
      .sideNav a { justify-content:center; }
      .wrap { width:min(1120px,calc(100% - 28px)); margin:18px auto; }
    }
  </style>
</head>
<body>
  <aside class="sidebar">
    <div class="brand">
      <div class="brandMark"><span class="brandIcon">KM</span><span>Odómetros</span></div>
      <small>TITANIA · CENTRO OPERATIVO</small>
    </div>
    <div class="navLabel">Módulos</div>
    <nav class="sideNav" aria-label="Módulos de odómetros">
      <a href="http://127.0.0.1:8062"><span class="navIcon">⌂</span><span>Inicio</span></a>
      <a class="active" href="http://127.0.0.1:8060"><span class="navIcon">RC</span><span>Recepción</span></a>
      <a href="http://127.0.0.1:8061"><span class="navIcon">DT</span><span>Distancia</span></a>
    </nav>
    <div class="sidebarFoot">Módulo activo<br><b style="color:white">Mantenimiento por recepción</b></div>
  </aside>
  <main class="wrap">
    <header>
      <h1>Odometros mantenimiento</h1>
      <div class="status" id="jobLabel">Sin proceso activo</div>
    </header>

    <section class="panel">
      <form id="uploadForm">
        <input id="fileInput" name="file" type="file" accept=".xlsx" required>
        <select id="calculationMode" aria-label="Metodo de calculo">
          <option value="calculadora">Calculadora Flespi</option>
          <option value="reporte_diario">Configuracion: distancia diaria mileage()</option>
        </select>
        <button id="validateButton" type="button">Validar Excel</button>
        <button id="startButton" type="submit" disabled>Analizar y generar Excel</button>
      </form>
      <div id="validationBox" class="validation"></div>
      <p>
        <a class="download" href="/api/template/download">Descargar plantilla XLSX ejemplo</a>
      </p>
      <p class="muted">Si el archivo trae <b>CHASIS</b>, el sistema busca automáticamente IMEI e ID Flespi en Vapor. Para calcular el odómetro debe incluir <b>FECHA_RECEPCION</b> y <b>KM_RECEPCION</b>.</p>
    </section>

    <section class="panel selection" id="selectionPanel">
      <h2 style="margin-top:0;font-size:18px">Paso 1: selecciona los odómetros para analizar</h2>
      <p class="muted">Esta acción solamente calcula y genera el Excel. No realiza cambios en Flespi.</p>
      <div class="selectionTools">
        <button type="button" id="selectAllButton">Marcar para analizar</button>
        <button type="button" class="secondary" id="clearSelectionButton">Quitar selección</button>
        <input id="rowSearch" type="search" placeholder="Buscar placa, chasis, modelo o ID Flespi">
        <b id="selectionCount">0 seleccionados</b>
        <button type="button" id="analyzeSelectionButton" class="download" disabled>Analizar selección y descargar Excel</button>
      </div>
      <div class="tableWrap">
        <table>
          <thead><tr><th></th><th>Accion</th><th>Placa</th><th>Chasis</th><th>Modelo</th><th>Fecha</th><th>KM base</th><th>ID Flespi</th><th>Mensaje</th></tr></thead>
          <tbody id="selectionRows"></tbody>
        </table>
      </div>
    </section>

    <section class="panel">
      <div id="state">Esperando archivo...</div>
      <div class="progressbar"><div id="bar"></div></div>
      <div class="grid">
        <div class="metric"><b id="phase">-</b><span>Fase</span></div>
        <div class="metric"><b id="current">0/0</b><span>Actuales Flespi</span></div>
        <div class="metric"><b id="historical">0/0</b><span id="methodProgressLabel">Historicos Flespi</span></div>
        <div class="metric"><b id="percent">0%</b><span>Avance estimado</span></div>
      </div>
      <p id="downloadBox"></p>
    </section>

    <section class="panel selection" id="resultsPanel">
      <h2 style="margin-top:0;font-size:18px">Paso 2 (opcional): actualizar Flespi</h2>
      <p>El Excel calculado ya está disponible para descargar. Si además deseas cambiar Flespi, selecciona los registros aprobados y pulsa el botón de actualización. <b>RAZONABLE:</b> diferencia absoluta hasta 500 km; <b>REVISAR:</b> diferencia mayor.</p>
      <div class="selectionTools">
        <button type="button" id="selectResultsButton">Marcar ACTUALIZAR</button>
        <button type="button" class="secondary" id="clearResultsButton">Desmarcar todos</button>
        <select id="differenceFilter" aria-label="Filtrar por estado de diferencia">
          <option value="TODOS">Todas las diferencias</option>
          <option value="RAZONABLE">RAZONABLE</option>
          <option value="REVISAR">REVISAR</option>
          <option value="REVISAR_POSITIVA">REVISAR - diferencia positiva</option>
          <option value="REVISAR_NEGATIVA">REVISAR - diferencia negativa</option>
          <option value="SIN DATOS">SIN DATOS</option>
        </select>
        <b id="resultSelectionCount">0 seleccionados</b>
        <button type="button" id="applyUpdatesButton" style="margin-left:auto;background:#168a4a">Actualizar odometros seleccionados</button>
      </div>
      <div class="tableWrap">
        <table>
          <thead><tr><th></th><th>Accion</th><th>Placa</th><th>Chasis</th><th>KM base</th><th>KM calculado</th><th>KM_FLESPI_ACTUAL</th><th>Diferencia</th><th>Estado diferencia</th><th>Estado calculo</th><th>Mensaje</th></tr></thead>
          <tbody id="resultRows"></tbody>
        </table>
      </div>
      <div id="updateProgress">
        <p><b id="updateState">Preparando actualizacion...</b> — <span id="updatePercent">0%</span></p>
        <div class="progressbar"><div id="updateBar"></div></div>
        <div class="grid">
          <div class="metric"><b id="updateDone">0/0</b><span>Procesados</span></div>
          <div class="metric"><b id="updateSuccess">0</b><span>Actualizados</span></div>
          <div class="metric"><b id="updatePending">0</b><span>Pendientes</span></div>
          <div class="metric"><b id="updateFailed">0</b><span>Errores</span></div>
        </div>
        <pre id="updateLog"></pre>
      </div>
    </section>

    <section class="panel">
      <pre id="log"></pre>
    </section>
  </main>

  <script>
    const form = document.getElementById("uploadForm");
    const fileInput = document.getElementById("fileInput");
    const validateButton = document.getElementById("validateButton");
    const calculationMode = document.getElementById("calculationMode");
    const button = document.getElementById("startButton");
    const validationBox = document.getElementById("validationBox");
    const state = document.getElementById("state");
    const bar = document.getElementById("bar");
    const logBox = document.getElementById("log");
    const jobLabel = document.getElementById("jobLabel");
    const phase = document.getElementById("phase");
    const current = document.getElementById("current");
    const historical = document.getElementById("historical");
    const methodProgressLabel = document.getElementById("methodProgressLabel");
    const percent = document.getElementById("percent");
    const downloadBox = document.getElementById("downloadBox");
    const selectionPanel = document.getElementById("selectionPanel");
    const selectionRows = document.getElementById("selectionRows");
    const selectionCount = document.getElementById("selectionCount");
    const analyzeSelectionButton = document.getElementById("analyzeSelectionButton");
    const selectAllButton = document.getElementById("selectAllButton");
    const clearSelectionButton = document.getElementById("clearSelectionButton");
    const rowSearch = document.getElementById("rowSearch");
    const resultsPanel = document.getElementById("resultsPanel");
    const resultRows = document.getElementById("resultRows");
    const resultSelectionCount = document.getElementById("resultSelectionCount");
    const selectResultsButton = document.getElementById("selectResultsButton");
    const clearResultsButton = document.getElementById("clearResultsButton");
    const applyUpdatesButton = document.getElementById("applyUpdatesButton");
    const differenceFilter = document.getElementById("differenceFilter");
    const updateProgress = document.getElementById("updateProgress");
    const updateState = document.getElementById("updateState");
    const updatePercent = document.getElementById("updatePercent");
    const updateBar = document.getElementById("updateBar");
    const updateDone = document.getElementById("updateDone");
    const updateSuccess = document.getElementById("updateSuccess");
    const updatePending = document.getElementById("updatePending");
    const updateFailed = document.getElementById("updateFailed");
    const updateLog = document.getElementById("updateLog");
    let timer = null;
    let activeJob = null;
    let validatedFileName = "";
    let rowOptions = [];
    let selectedRows = new Set();
    let resultOptions = [];
    let selectedResults = new Set();
    let updateTimer = null;

    fileInput.addEventListener("change", () => {
      validatedFileName = "";
      button.disabled = true;
      validationBox.className = "validation";
      validationBox.innerHTML = "";
      selectionPanel.classList.remove("show");
      rowOptions = [];
      selectedRows = new Set();
      resultsPanel.classList.remove("show");
      resultOptions = [];
      selectedResults = new Set();
      updateProgress.classList.remove("show");
      if (updateTimer) clearInterval(updateTimer);
      updateTimer = null;
      state.textContent = "Archivo seleccionado. Valida antes de procesar.";
    });

    validateButton.addEventListener("click", validateFile);
    calculationMode.addEventListener("change", () => {
      methodProgressLabel.textContent = calculationMode.value === "reporte_diario" ? "Reportes diarios Flespi" : "Historicos Flespi";
    });

    async function validateFile() {
      if (!fileInput.files.length) return;
      validateButton.disabled = true;
      button.disabled = true;
      validationBox.className = "validation show";
      validationBox.innerHTML = "<h3>Validando archivo...</h3>";
      const data = new FormData();
      data.append("file", fileInput.files[0]);
      const response = await fetch("/api/validate", { method: "POST", body: data });
      const body = await response.json();
      validateButton.disabled = false;
      renderValidation(body);
      if (response.ok && body.can_process) {
        validatedFileName = fileInput.files[0].name;
        renderSelection(body.row_options || []);
        state.textContent = "Archivo validado. Selecciona que odometros deseas actualizar.";
      } else {
        validatedFileName = "";
        button.disabled = true;
        state.textContent = "Corrige el archivo antes de procesar.";
      }
    }

    function renderValidation(body) {
      const valid = body.can_process === true;
      const errors = body.errors || [];
      const warnings = body.warnings || [];
      const chassisResolution = body.chassis_resolution || null;
      validationBox.className = `validation show ${valid ? "good" : "bad"}`;
      validationBox.innerHTML = `
        <h3>${valid ? "Control terminado" : "Control con errores"}</h3>
        <div class="validationGrid">
          <div class="pill"><b>${body.rows || 0}</b><span>Filas</span></div>
          <div class="pill"><b>${body.unique_ids || 0}</b><span>IDs Flespi unicos</span></div>
          <div class="pill"><b>${body.valid_rows || 0}</b><span>Filas validas</span></div>
          <div class="pill"><b>${errors.length}</b><span>Errores</span></div>
          <div class="pill"><b>${warnings.length}</b><span>Alertas</span></div>
        </div>
        ${chassisResolution ? `<p><b>Consulta por chasis:</b> ${escapeHtml(chassisResolution.message || "")}</p>` : ""}
        ${errors.length ? `<b class="error">Errores:</b><ul>${errors.slice(0, 25).map(item => `<li>${escapeHtml(item)}</li>`).join("")}</ul>` : ""}
        ${warnings.length ? `<b>Alertas:</b><ul>${warnings.slice(0, 25).map(item => `<li>${escapeHtml(item)}</li>`).join("")}</ul>` : ""}
      `;
    }

    function renderSelection(options) {
      rowOptions = options;
      selectedRows = new Set(options.filter(item => item.can_update).map(item => item.index));
      selectionPanel.classList.add("show");
      rowSearch.value = "";
      renderSelectionRows();
      updateSelectionCount();
    }

    function renderSelectionRows() {
      const query = rowSearch.value.trim().toLowerCase();
      const visible = rowOptions.filter(item => !query || [item.plate, item.chassis, item.model, item.device_id].join(" ").toLowerCase().includes(query));
      selectionRows.innerHTML = visible.map(item => `
        <tr class="${item.can_update ? "" : "review"}">
          <td><input type="checkbox" data-index="${item.index}" ${selectedRows.has(item.index) ? "checked" : ""} ${item.can_update ? "" : "disabled"}></td>
          <td><span class="badge ${item.action === "ACTUALIZAR" ? "update" : "review"}">${item.action}</span></td>
          <td>${escapeHtml(item.plate)}</td><td>${escapeHtml(item.chassis)}</td><td>${escapeHtml(item.model)}</td>
          <td>${escapeHtml(item.date)}</td><td>${escapeHtml(item.km_base)}</td><td>${escapeHtml(item.device_id)}</td>
          <td>${escapeHtml(item.message)}</td>
        </tr>`).join("");
      selectionRows.querySelectorAll("input[type=checkbox]").forEach(box => box.addEventListener("change", () => {
        const index = Number(box.dataset.index);
        if (box.checked) selectedRows.add(index); else selectedRows.delete(index);
        updateSelectionCount();
      }));
    }

    function updateSelectionCount() {
      selectionCount.textContent = `${selectedRows.size} seleccionados`;
      button.disabled = selectedRows.size === 0;
      analyzeSelectionButton.disabled = selectedRows.size === 0;
      button.textContent = selectedRows.size ? `Analizar ${selectedRows.size} seleccionados` : "Analizar y generar Excel";
      analyzeSelectionButton.textContent = selectedRows.size ? `Analizar ${selectedRows.size} y descargar Excel` : "Analizar selección y descargar Excel";
    }

    selectAllButton.addEventListener("click", () => {
      selectedRows = new Set(rowOptions.filter(item => item.can_update).map(item => item.index));
      renderSelectionRows(); updateSelectionCount();
    });
    clearSelectionButton.addEventListener("click", () => {
      selectedRows.clear(); renderSelectionRows(); updateSelectionCount();
    });
    rowSearch.addEventListener("input", renderSelectionRows);
    analyzeSelectionButton.addEventListener("click", () => button.click());

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!fileInput.files.length) return;
      if (validatedFileName !== fileInput.files[0].name) {
        state.innerHTML = `<span class="error">Primero valida el Excel seleccionado.</span>`;
        return;
      }
      button.disabled = true;
      downloadBox.innerHTML = "";
      logBox.textContent = "";
      state.textContent = "Subiendo archivo...";
      const data = new FormData();
      data.append("file", fileInput.files[0]);
      data.append("selected_rows", JSON.stringify(Array.from(selectedRows).sort((a, b) => a - b)));
      data.append("calculation_mode", calculationMode.value);
      const response = await fetch("/api/jobs", { method: "POST", body: data });
      const body = await response.json();
      if (!response.ok) {
        button.disabled = false;
        state.innerHTML = `<span class="error">${body.error || "No se pudo iniciar"}</span>`;
        return;
      }
      activeJob = body.job_id;
      jobLabel.textContent = `Proceso ${activeJob}`;
      poll();
      timer = setInterval(poll, 1500);
    });

    async function poll() {
      if (!activeJob) return;
      const response = await fetch(`/api/jobs/${activeJob}`);
      const body = await response.json();
      if (!response.ok) {
        state.innerHTML = `<span class="error">${body.error || "No se pudo consultar"}</span>`;
        clearInterval(timer);
        button.disabled = false;
        return;
      }
      state.textContent = body.message;
      logBox.textContent = body.log || "";
      logBox.scrollTop = logBox.scrollHeight;
      phase.textContent = body.phase;
      current.textContent = `${body.current_done}/${body.current_total}`;
      historical.textContent = `${body.historical_done}/${body.historical_total}`;
      percent.textContent = `${body.percent}%`;
      bar.style.width = `${body.percent}%`;
      if (body.done) {
        clearInterval(timer);
        button.disabled = false;
        if (body.success) {
          downloadBox.innerHTML = `<a class="download" href="/api/jobs/${activeJob}/download">Descargar Excel con cálculo</a><span class="muted"> El archivo fue analizado; Flespi aún no ha sido actualizado.</span>`;
          await loadResults();
        }
      }
    }

    async function loadResults() {
      const response = await fetch(`/api/jobs/${activeJob}/results`);
      const body = await response.json();
      if (!response.ok) {
        state.innerHTML = `<span class="error">${body.error || "No se pudieron cargar los resultados"}</span>`;
        return;
      }
      resultOptions = body.rows || [];
      selectedResults = new Set(resultOptions.filter(item => item.action === "ACTUALIZAR" && item.can_update && !item.applied).map(item => item.index));
      resultsPanel.classList.add("show");
      differenceFilter.value = "TODOS";
      renderResultRows();
      updateResultSelectionCount();
    }

    function renderResultRows() {
      const filterValue = differenceFilter.value;
      const visibleResults = resultOptions.filter(item => {
        if (filterValue === "TODOS") return true;
        if (filterValue === "REVISAR_POSITIVA") return item.difference_status === "REVISAR" && Number(item.km_difference) > 0;
        if (filterValue === "REVISAR_NEGATIVA") return item.difference_status === "REVISAR" && Number(item.km_difference) < 0;
        return item.difference_status === filterValue;
      });
      resultRows.innerHTML = visibleResults.map(item => `
        <tr class="${item.action === "REVISAR" ? "review" : ""}">
          <td><input type="checkbox" data-result-index="${item.index}" ${selectedResults.has(item.index) ? "checked" : ""} ${item.can_update ? "" : "disabled"}></td>
          <td><span class="badge ${item.can_update ? "update" : "review"}">${item.action}</span></td>
          <td>${escapeHtml(item.plate)}</td><td>${escapeHtml(item.chassis)}</td>
          <td>${escapeHtml(item.km_base)}</td><td>${escapeHtml(item.km_estimated)}</td><td>${escapeHtml(item.km_applied)}</td><td>${escapeHtml(item.km_difference)}</td>
          <td><span class="badge ${item.difference_status === "RAZONABLE" ? "update" : "review"}">${escapeHtml(item.difference_status)}</span></td>
          <td>${escapeHtml(item.status)}</td><td>${escapeHtml(item.message)}</td>
        </tr>`).join("");
      resultRows.querySelectorAll("input[type=checkbox]").forEach(box => box.addEventListener("change", () => {
        const index = Number(box.dataset.resultIndex);
        if (box.checked) selectedResults.add(index); else selectedResults.delete(index);
        updateResultSelectionCount();
      }));
    }

    function updateResultSelectionCount() {
      resultSelectionCount.textContent = `${selectedResults.size} seleccionados`;
      applyUpdatesButton.disabled = selectedResults.size === 0;
    }

    selectResultsButton.addEventListener("click", () => {
      selectedResults = new Set(resultOptions.filter(item => item.action === "ACTUALIZAR" && item.can_update).map(item => item.index));
      renderResultRows(); updateResultSelectionCount();
    });
    clearResultsButton.addEventListener("click", () => {
      selectedResults.clear(); renderResultRows(); updateResultSelectionCount();
    });
    differenceFilter.addEventListener("change", renderResultRows);
    applyUpdatesButton.addEventListener("click", async () => {
      if (!activeJob || !selectedResults.size) return;
      if (!window.confirm(`Se enviaran ${selectedResults.size} actualizaciones reales a Flespi. Los kilometrajes se redondearan a enteros. ¿Continuar?`)) return;
      applyUpdatesButton.disabled = true;
      updateProgress.classList.add("show");
      updateLog.textContent = "Preparando envio a Flespi...";
      updateBar.style.width = "0%";
      state.textContent = "Enviando actualizaciones a Flespi...";
      const response = await fetch(`/api/jobs/${activeJob}/apply`, {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({rows: Array.from(selectedResults).sort((a, b) => a - b)}),
      });
      const body = await response.json();
      if (!response.ok) {
        state.innerHTML = `<span class="error">${body.error || "No se pudo iniciar la actualizacion en Flespi"}</span>`;
        applyUpdatesButton.disabled = false;
        return;
      }
      await pollUpdate();
      updateTimer = setInterval(pollUpdate, 1000);
    });

    async function pollUpdate() {
      const response = await fetch(`/api/jobs/${activeJob}/apply-status`);
      const body = await response.json();
      if (!response.ok) {
        clearInterval(updateTimer);
        updateTimer = null;
        state.innerHTML = `<span class="error">${body.error || "No se pudo consultar la actualizacion"}</span>`;
        applyUpdatesButton.disabled = false;
        return;
      }
      updateState.textContent = body.message;
      updatePercent.textContent = `${body.percent}%`;
      updateBar.style.width = `${body.percent}%`;
      updateDone.textContent = `${body.done_count}/${body.total}`;
      updateSuccess.textContent = body.success;
      updatePending.textContent = body.pending;
      updateFailed.textContent = body.failed;
      updateLog.textContent = (body.log || []).join("\\n");
      updateLog.scrollTop = updateLog.scrollHeight;
      if (body.done) {
        clearInterval(updateTimer);
        updateTimer = null;
        state.textContent = body.message;
        applyUpdatesButton.disabled = false;
        downloadBox.innerHTML = `<a class="download" href="/api/jobs/${activeJob}/download">Descargar Excel actualizado</a>`;
        await loadResults();
        updateProgress.classList.add("show");
      }
    }

    function escapeHtml(value) {
      return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }
  </script>
</body>
</html>
"""


def parse_progress(log_text, returncode):
    current_matches = re.findall(r"Actuales Flespi:\s*(\d+)/(\d+)", log_text)
    historical_total_match = re.search(r"Bases historicas unicas:\s*(\d+)", log_text)
    historical_matches = re.findall(r"Historicos Flespi:\s*(\d+)/(\d+)", log_text)
    daily_matches = re.findall(r"Reportes diarios Flespi:\s*(\d+)/(\d+)", log_text)

    current_done = current_total = historical_done = historical_total = 0
    if current_matches:
        current_done, current_total = map(int, current_matches[-1])
    if historical_total_match:
        historical_total = int(historical_total_match.group(1))
    if historical_matches:
        historical_done, historical_total = map(int, historical_matches[-1])
    if daily_matches:
        historical_done, historical_total = map(int, daily_matches[-1])

    if daily_matches and returncode is None:
        phase = "Reporte diario"
    elif returncode is None:
        phase = "Actuales" if not historical_total_match else "Historicos"
    elif returncode == 0:
        phase = "Terminado"
    else:
        phase = "Error"

    total_units = (current_total or 0) + (historical_total or 0)
    done_units = current_done + historical_done
    percent = int((done_units / total_units) * 100) if total_units else 0
    if returncode == 0:
        percent = 100

    return {
        "phase": phase,
        "current_done": current_done,
        "current_total": current_total,
        "historical_done": historical_done,
        "historical_total": historical_total,
        "percent": max(0, min(100, percent)),
    }


def read_text(path):
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8", errors="replace")


def write_selected_workbook(headers, rows, selected_indices, output_path):
    workbook = xlsxwriter.Workbook(str(output_path))
    worksheet = workbook.add_worksheet("MANTENIMIENTO")
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    for column, name in enumerate(headers):
        worksheet.write(0, column, name, header_format)
    for output_row, source_index in enumerate(selected_indices, start=1):
        source = rows[source_index]
        for column, name in enumerate(headers):
            worksheet.write(output_row, column, source.get(name))
    worksheet.autofilter(0, 0, len(selected_indices), len(headers) - 1)
    worksheet.freeze_panes(1, 0)
    worksheet.set_column(0, len(headers) - 1, 18)
    workbook.close()


def result_action_counts(path):
    if not path.exists():
        return {"actualizar": 0, "revisar": 0}
    _headers, rows = read_xlsx_rows(path)
    return {
        "actualizar": sum(row.get("ACCION_RECOMENDADA") == "ACTUALIZAR" for row in rows),
        "revisar": sum(row.get("ACCION_RECOMENDADA") == "REVISAR" for row in rows),
    }


def write_result_workbook(headers, rows, output_path):
    if "KM_FLESPI_ACTUAL" not in headers:
        headers = [*headers, "KM_FLESPI_ACTUAL"]
    workbook = xlsxwriter.Workbook(str(output_path))
    worksheet = workbook.add_worksheet("KM_ACTUAL")
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    number_format = workbook.add_format({"num_format": "#,##0.00"})
    update_format = workbook.add_format({"bg_color": "#D1FADF", "font_color": "#05603A", "bold": True})
    review_format = workbook.add_format({"bg_color": "#FEE4E2", "font_color": "#B42318", "bold": True})
    applied_format = workbook.add_format({"bg_color": "#D9EAF7", "font_color": "#155EEF", "bold": True, "num_format": "#,##0.00"})
    for column, name in enumerate(headers):
        worksheet.write(0, column, name, header_format)
    for output_row, row in enumerate(rows, start=1):
        for column, name in enumerate(headers):
            value = row.get(name)
            numeric_value = parse_float(value) if name.startswith("KM_") else None
            if numeric_value is not None:
                worksheet.write_number(output_row, column, numeric_value, applied_format if name == "KM_FLESPI_ACTUAL" else number_format)
            else:
                worksheet.write(output_row, column, value)
    worksheet.autofilter(0, 0, len(rows), len(headers) - 1)
    worksheet.freeze_panes(1, 0)
    worksheet.set_column(0, len(headers) - 1, 18)
    if "ACCION_RECOMENDADA" in headers:
        action_column = headers.index("ACCION_RECOMENDADA")
        worksheet.set_column(action_column, action_column, 24)
        worksheet.conditional_format(1, action_column, len(rows), action_column, {
            "type": "text", "criteria": "containing", "value": "ACTUALIZAR", "format": update_format,
        })
        worksheet.conditional_format(1, action_column, len(rows), action_column, {
            "type": "text", "criteria": "containing", "value": "REVISAR", "format": review_format,
        })
    worksheet.set_column(headers.index("KM_FLESPI_ACTUAL"), headers.index("KM_FLESPI_ACTUAL"), 22)
    workbook.close()


def append_update_log(job_id, message):
    with update_lock:
        update_jobs[job_id]["log"].append(message)


def run_flespi_updates(job_id, selected_indices):
    state = update_jobs[job_id]
    job = jobs[job_id]
    session = requests.Session()
    session.trust_env = False
    try:
        config = json.loads((APP_DIR / "config.json").read_text(encoding="utf-8"))
        token = config["flespi"]["token"]
        headers, rows = read_xlsx_rows(job["output"])
        auth_headers = {"Authorization": f"FlespiToken {token}", "Content-Type": "application/json"}
        append_update_log(job_id, f"Iniciando {len(selected_indices)} actualizaciones. Todos los valores se enviaran como enteros.")

        for position, index in enumerate(selected_indices, start=1):
            row = rows[index]
            plate = str(row.get("PLACA") or "SIN PLACA")
            device_id = str(row.get("ID_FLESPI") or "").strip()
            estimated = parse_float(row.get("KM_ACTUAL_ESTIMADO"))
            target_km = int(Decimal(str(estimated)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
            row["KM_SOLICITADO_FLESPI"] = target_km
            append_update_log(
                job_id,
                f"[{position}/{len(selected_indices)}] {plate} | ID {device_id} | {estimated:.2f} -> {target_km} km",
            )

            update_status = "ERROR"
            try:
                url = f"https://flespi.io/gw/devices/{device_id}/settings/odometer_fmb"
                payload = {
                    "address": "connection",
                    "properties": {"source": 0, "value": target_km},
                }
                response = session.put(url, headers=auth_headers, json=payload, timeout=30)
                if not response.ok:
                    detail = response.text.replace("\n", " ")[:300]
                    raise RuntimeError(f"HTTP {response.status_code}: {detail}")

                verify = session.get(url, headers={"Authorization": f"FlespiToken {token}"}, timeout=20)
                setting = (verify.json().get("result") or [{}])[0] if verify.ok else {}
                current_value = parse_float((setting.get("current") or {}).get("value"))
                pending_data = setting.get("pending") or {}
                pending_properties = pending_data.get("properties") or pending_data
                pending_value = parse_float(pending_properties.get("value"))
                if current_value is not None and int(round(current_value)) == target_km:
                    update_status = "ACTUALIZADO"
                    row["KM_FLESPI_ACTUAL"] = target_km
                    row["DIFERENCIA_KM"] = round(estimated - target_km, 2)
                    row["ESTADO_DIFERENCIA"] = "RAZONABLE"
                    row["ODOMETRO_ACTUALIZADO"] = "SI"
                    with update_lock:
                        state["success"] += 1
                    append_update_log(job_id, f"    OK confirmado por el dispositivo: {target_km} km")
                else:
                    update_status = "PENDIENTE"
                    row["ODOMETRO_ACTUALIZADO"] = "PENDIENTE"
                    with update_lock:
                        state["pending"] += 1
                    pending_text = int(round(pending_value)) if pending_value is not None else target_km
                    append_update_log(job_id, f"    PENDIENTE en Flespi: {pending_text} km; espera conexion/confirmacion del dispositivo")
            except Exception as error:
                row["ODOMETRO_ACTUALIZADO"] = "ERROR"
                with update_lock:
                    state["failed"] += 1
                append_update_log(job_id, f"    ERROR: {error}")

            row["ESTADO_ACTUALIZACION_FLESPI"] = update_status
            with update_lock:
                state["done_count"] = position
                state["percent"] = int(position * 100 / len(selected_indices))
                state["message"] = f"Actualizando Flespi: {position}/{len(selected_indices)}"

        write_result_workbook(headers, rows, job["output"])
        with update_lock:
            state["done"] = True
            state["percent"] = 100
            state["message"] = (
                f"Actualizacion terminada: {state['success']} confirmados, "
                f"{state['pending']} pendientes y {state['failed']} errores."
            )
        append_update_log(job_id, state["message"])
    except Exception as error:
        append_update_log(job_id, f"ERROR GENERAL: {error}")
        with update_lock:
            state["done"] = True
            state["failed"] += max(0, state["total"] - state["done_count"])
            state["message"] = f"La actualizacion se detuvo: {error}"


def validate_workbook(path):
    errors = []
    fatal_errors = []
    warnings = []
    try:
        headers, rows = read_xlsx_rows(path)
    except Exception as exc:
        return {
            "valid": False,
            "rows": 0,
            "valid_rows": 0,
            "unique_ids": 0,
            "can_process": False,
            "row_options": [],
            "errors": [f"No se pudo leer el XLSX: {exc}"],
            "warnings": [],
        }

    headers = headers or []
    rows = rows or []
    required_columns = ["ID_FLESPI", "FECHA_RECEPCION"]
    for column in required_columns:
        if column not in headers:
            fatal_errors.append(f"Falta columna obligatoria: {column}")

    km_base_columns = [name for name in KM_BASE_COLUMNS if name in headers]
    if not km_base_columns:
        fatal_errors.append("Falta la columna obligatoria de kilometraje: KM_RECEPCION")

    if not rows:
        fatal_errors.append("El archivo no contiene filas de datos")

    ids = []
    key_counter = Counter()
    valid_rows = 0
    row_options = []
    today = datetime.now().astimezone().date()

    for index, row in enumerate(rows, start=2):
        row_errors = []
        device_id = str(row.get("ID_FLESPI") or "").strip()
        fecha = parse_date(row.get("FECHA_RECEPCION"))
        km_value = first_km_base(row, headers) if km_base_columns else None

        if not device_id:
            row_errors.append("sin ID_FLESPI")
        elif not device_id.isdigit():
            row_errors.append("ID_FLESPI debe ser numerico")
        else:
            ids.append(device_id)

        if not fecha:
            row_errors.append("FECHA_RECEPCION invalida")
        elif fecha.date() > today:
            row_errors.append("FECHA_RECEPCION no puede estar en el futuro")

        if km_base_columns and km_value is None:
            row_errors.append("KM_RECEPCION invalido")
        elif km_value is not None and km_value < 0:
            row_errors.append("KM_RECEPCION no puede ser negativo")

        if device_id and fecha:
            key_counter[(device_id, fecha.isoformat())] += 1

        if row_errors:
            errors.append(f"Fila {index}: {', '.join(row_errors)}")
        else:
            valid_rows += 1

        row_options.append({
            "index": index - 2,
            "excel_row": index,
            "can_update": not row_errors and not fatal_errors,
            "action": "ACTUALIZAR" if not row_errors and not fatal_errors else "REVISAR",
            "message": "Lista para calcular" if not row_errors and not fatal_errors else ", ".join(row_errors or fatal_errors),
            "plate": str(row.get("PLACA") or ""),
            "chassis": str(row.get("CHASIS") or ""),
            "model": str(row.get("MODELO") or ""),
            "date": fecha.strftime("%d/%m/%Y %H:%M:%S") if fecha else str(row.get("FECHA_RECEPCION") or ""),
            "km_base": km_value,
            "device_id": device_id,
        })

    duplicate_ids = [device_id for device_id, count in Counter(ids).items() if count > 1]
    duplicate_keys = [key for key, count in key_counter.items() if count > 1]
    if duplicate_ids:
        warnings.append(f"Hay {len(duplicate_ids)} ID_FLESPI repetidos. Se procesan, pero revisa si son mantenimientos distintos.")
    if duplicate_keys:
        warnings.append(f"Hay {len(duplicate_keys)} combinaciones ID_FLESPI + FECHA_RECEPCION repetidas.")

    return {
        "valid": not fatal_errors and not errors,
        "can_process": not fatal_errors and valid_rows > 0,
        "rows": len(rows),
        "valid_rows": valid_rows,
        "unique_ids": len(set(ids)),
        "km_columns": km_base_columns,
        "headers": headers,
        "errors": (fatal_errors + errors)[:100],
        "warnings": warnings[:50],
        "row_options": row_options,
    }


@app.get("/")
def index():
    return render_template_string(PAGE)


@app.post("/api/validate")
def validate_file():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "Sube un archivo .xlsx", "valid": False}), 400
    ext = Path(uploaded.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": "Solo se aceptan archivos .xlsx", "valid": False}), 400

    job_id = f"validacion_{uuid.uuid4().hex[:10]}"
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    input_path = job_dir / (secure_filename(uploaded.filename) or "MANTENIMIENTO.xlsx")
    uploaded.save(input_path)
    enriched_path = job_dir / "MANTENIMIENTO_CON_DISPOSITIVOS.xlsx"
    source_path, resolution = enrich_workbook_by_chassis(input_path, enriched_path)
    result = validate_workbook(source_path)
    result["chassis_resolution"] = resolution
    return jsonify(result), 200 if result["can_process"] else 400


@app.post("/api/jobs")
def create_job():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "Sube un archivo .xlsx"}), 400
    ext = Path(uploaded.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": "Solo se aceptan archivos .xlsx"}), 400

    calculation_mode = request.form.get("calculation_mode", "calculadora")
    # Compatibilidad con pestañas que quedaron abiertas antes del cambio a distancia diaria.
    if calculation_mode == "configuracion_dispositivo":
        calculation_mode = "reporte_diario"
    if calculation_mode not in {"calculadora", "reporte_diario"}:
        return jsonify({"error": "Metodo de calculo no valido"}), 400

    job_id = uuid.uuid4().hex[:10]
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    input_name = secure_filename(uploaded.filename) or "MANTENIMIENTO.xlsx"
    uploaded_path = job_dir / input_name
    enriched_path = job_dir / "MANTENIMIENTO_CON_DISPOSITIVOS.xlsx"
    input_path = job_dir / "SELECCION_ODOMETROS.xlsx"
    output_path = job_dir / OUTPUT_NAME
    log_path = job_dir / "progreso.log"
    err_path = job_dir / "errores.log"
    uploaded.save(uploaded_path)

    source_path, resolution = enrich_workbook_by_chassis(uploaded_path, enriched_path)
    validation = validate_workbook(source_path)
    validation["chassis_resolution"] = resolution
    if not validation["can_process"]:
        return jsonify({"error": "El archivo no paso el control previo", **validation}), 400

    try:
        requested_indices = json.loads(request.form.get("selected_rows", "[]"))
        requested_indices = sorted({int(value) for value in requested_indices})
    except (TypeError, ValueError, json.JSONDecodeError):
        return jsonify({"error": "La seleccion de filas no es valida"}), 400

    allowed_indices = {item["index"] for item in validation["row_options"] if item["can_update"]}
    if not requested_indices:
        return jsonify({"error": "Selecciona al menos un odometro para actualizar"}), 400
    if any(index not in allowed_indices for index in requested_indices):
        return jsonify({"error": "La seleccion contiene filas marcadas para REVISAR"}), 400

    headers, rows = read_xlsx_rows(source_path)
    write_selected_workbook(headers, rows, requested_indices, input_path)

    log_file = log_path.open("w", encoding="utf-8", buffering=1)
    err_file = err_path.open("w", encoding="utf-8", buffering=1)
    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    process = subprocess.Popen(
        [
            sys.executable,
            "-u",
            "-B",
            str(CALC_SCRIPT),
            str(input_path),
            str(output_path),
            calculation_mode,
        ],
        cwd=APP_DIR,
        stdout=log_file,
        stderr=err_file,
        creationflags=creationflags,
    )
    jobs[job_id] = {
        "process": process,
        "input": input_path,
        "output": output_path,
        "log": log_path,
        "err": err_path,
        "log_file": log_file,
        "err_file": err_file,
        "selected_count": len(requested_indices),
        "calculation_mode": calculation_mode,
    }
    return jsonify({"job_id": job_id, "selected_count": len(requested_indices), "calculation_mode": calculation_mode})


@app.get("/api/jobs/<job_id>")
def get_job(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Proceso no encontrado"}), 404

    process = job["process"]
    returncode = process.poll()
    if returncode is not None:
        job["log_file"].close()
        job["err_file"].close()

    log_text = read_text(job["log"])
    err_text = read_text(job["err"])
    progress = parse_progress(log_text, returncode)

    done = returncode is not None
    success = returncode == 0 and job["output"].exists()
    if success:
        actions = result_action_counts(job["output"])
        message = f"Proceso terminado: {actions['actualizar']} para ACTUALIZAR y {actions['revisar']} para REVISAR."
    elif returncode is None:
        message = "Procesando archivo..."
    else:
        message = "El proceso termino con error. Revisa el log."

    combined_log = log_text
    if err_text.strip():
        combined_log = f"{combined_log}\n\nERRORES:\n{err_text}"

    return jsonify(
        {
            **progress,
            "done": done,
            "success": success,
            "message": message,
            "log": combined_log[-12000:],
            "selected_count": job["selected_count"],
            "actions": result_action_counts(job["output"]) if success else {"actualizar": 0, "revisar": 0},
        }
    )


@app.get("/api/jobs/<job_id>/download")
def download(job_id):
    job = jobs.get(job_id)
    if not job or not job["output"].exists():
        return jsonify({"error": "Archivo no disponible"}), 404
    return send_file(job["output"], as_attachment=True, download_name=OUTPUT_NAME)


@app.get("/api/jobs/<job_id>/results")
def job_results(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Proceso no encontrado"}), 404
    if job["process"].poll() is None:
        return jsonify({"error": "El analisis todavia esta en proceso"}), 409
    if job["process"].returncode != 0 or not job["output"].exists():
        return jsonify({"error": "El analisis no genero un archivo valido"}), 409

    headers, rows = read_xlsx_rows(job["output"])
    result_rows = []
    for index, row in enumerate(rows):
        action = str(row.get("ACCION_RECOMENDADA") or "REVISAR")
        estimated = parse_float(row.get("KM_ACTUAL_ESTIMADO"))
        applied = parse_float(row.get("KM_FLESPI_ACTUAL"))
        difference = parse_float(row.get("DIFERENCIA_KM"))
        update_marker = str(row.get("ODOMETRO_ACTUALIZADO") or "").strip().upper()
        was_updated = update_marker in {"SI", "PENDIENTE"}
        calculation_status = str(row.get("ESTADO_CALCULO") or "")
        can_update = calculation_status == "OK" and estimated is not None
        result_rows.append({
            "index": index,
            "can_update": can_update,
            "action": action,
            "plate": str(row.get("PLACA") or ""),
            "chassis": str(row.get("CHASIS") or ""),
            "km_base": first_km_base(row, headers),
            "km_estimated": estimated,
            "km_applied": applied,
            "km_difference": difference,
            "difference_status": str(row.get("ESTADO_DIFERENCIA") or "SIN DATOS"),
            "applied": was_updated,
            "status": calculation_status,
            "message": str(row.get("OBSERVACION") or ("Listo para actualizar" if can_update else "Requiere revision")),
            "flespi_update_status": str(row.get("ESTADO_ACTUALIZACION_FLESPI") or ""),
        })
    return jsonify({"rows": result_rows})


@app.post("/api/jobs/<job_id>/apply")
def apply_job_updates(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Proceso no encontrado"}), 404
    if job["process"].poll() is None:
        return jsonify({"error": "Espera a que termine el analisis"}), 409
    if job["process"].returncode != 0 or not job["output"].exists():
        return jsonify({"error": "No existe un resultado valido para actualizar"}), 409

    payload = request.get_json(silent=True) or {}
    try:
        selected_indices = sorted({int(value) for value in payload.get("rows", [])})
    except (TypeError, ValueError):
        return jsonify({"error": "La seleccion de resultados no es valida"}), 400
    if not selected_indices:
        return jsonify({"error": "Selecciona al menos un resultado"}), 400

    headers, rows = read_xlsx_rows(job["output"])
    invalid = []
    for index in selected_indices:
        if index < 0 or index >= len(rows):
            invalid.append(index)
            continue
        row = rows[index]
        if row.get("ESTADO_CALCULO") != "OK" or parse_float(row.get("KM_ACTUAL_ESTIMADO")) is None:
            invalid.append(index)
    if invalid:
        return jsonify({"error": "Solo se pueden aplicar filas con un calculo valido", "invalid_rows": invalid}), 400

    existing_update = update_jobs.get(job_id)
    if existing_update and not existing_update["done"]:
        return jsonify({"error": "Ya existe una actualizacion de Flespi en proceso"}), 409

    with update_lock:
        update_jobs[job_id] = {
            "done": False,
            "done_count": 0,
            "total": len(selected_indices),
            "percent": 0,
            "success": 0,
            "pending": 0,
            "failed": 0,
            "message": "Preparando actualizacion en Flespi...",
            "log": [],
        }
    worker = threading.Thread(target=run_flespi_updates, args=(job_id, selected_indices), daemon=True)
    worker.start()
    return jsonify({"message": "Actualizacion iniciada", "total": len(selected_indices)}), 202


@app.get("/api/jobs/<job_id>/apply-status")
def apply_job_status(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Proceso no encontrado"}), 404
    with update_lock:
        state = update_jobs.get(job_id)
        if not state:
            return jsonify({"error": "No existe una actualizacion iniciada"}), 404
        return jsonify({**state, "log": state["log"][-500:]})


@app.get("/api/template/download")
def download_template():
    template_path = APP_DIR / TEMPLATE_NAME
    if not template_path.exists():
        return jsonify({"error": "Plantilla no disponible"}), 404
    return send_file(template_path, as_attachment=True, download_name=TEMPLATE_NAME)


if __name__ == "__main__":
    JOBS_DIR.mkdir(parents=True, exist_ok=True)
    app.run(host="127.0.0.1", port=8060, debug=False)

