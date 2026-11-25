from flask import Flask, render_template_string


app = Flask(__name__)

PAGE = """
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Menú de odómetros</title>
  <style>
    :root { --ink:#102a43; --blue:#155eef; --green:#087f5b; --line:#d9e2ec; }
    * { box-sizing:border-box; }
    body { margin:0; min-height:100vh; font-family:Arial,sans-serif; color:var(--ink); background:linear-gradient(135deg,#eef7ff,#f5fffb); }
    main { width:auto; max-width:1200px; margin:0 32px 0 280px; padding:58px 0; }
    .eyebrow { color:#087f5b; font-size:12px; letter-spacing:1.2px; font-weight:800; }
    h1 { font-size:36px; margin:8px 0 12px; }
    .intro { max-width:700px; color:#52677d; font-size:17px; line-height:1.5; }
    .cards { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:20px; margin-top:34px; }
    .card { background:white; border:1px solid var(--line); border-radius:18px; padding:26px; box-shadow:0 10px 30px rgba(26,68,105,.08); }
    .icon { width:48px; height:48px; display:grid; place-items:center; border-radius:14px; color:white; font-weight:800; font-size:20px; background:var(--blue); }
    .distance .icon { background:var(--green); }
    .real .icon { background:#7c3aed; }
    h2 { margin:18px 0 8px; font-size:22px; }
    p { color:#52677d; line-height:1.5; min-height:72px; }
    ul { padding-left:19px; color:#344b61; line-height:1.7; font-size:14px; }
    a { display:inline-flex; margin-top:18px; text-decoration:none; color:white; background:var(--blue); padding:12px 16px; border-radius:9px; font-weight:700; }
    .distance a { background:var(--green); }
    .real a { background:#7c3aed; }
    .note { margin-top:25px; padding:14px 16px; background:#fff; border:1px solid var(--line); border-radius:10px; color:#52677d; font-size:14px; }
    .sidebar { position:fixed; inset:0 auto 0 0; width:248px; padding:24px 16px; background:#0f2742; color:white; display:flex; flex-direction:column; z-index:10; box-shadow:10px 0 30px rgba(15,39,66,.12); }
    .brand { padding:4px 10px 22px; border-bottom:1px solid rgba(255,255,255,.14); }
    .brandMark { display:flex; align-items:center; gap:10px; font-weight:800; font-size:18px; }
    .brandIcon { width:38px; height:38px; display:grid; place-items:center; border-radius:11px; background:linear-gradient(135deg,#1c7ed6,#20c997); }
    .brand small { display:block; margin-top:7px; color:#a9bdd1; font-size:12px; letter-spacing:.8px; }
    .navLabel { margin:22px 10px 9px; color:#91a8bd; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:1px; }
    .sideNav { display:grid; gap:7px; }
    .sideNav a { display:flex; align-items:center; gap:11px; margin:0; padding:12px 13px; border-radius:10px; color:#d9e8f5; background:transparent; text-decoration:none; font-size:14px; font-weight:700; border:1px solid transparent; }
    .sideNav a:hover { background:rgba(255,255,255,.08); color:white; }
    .sideNav a.active { color:white; background:linear-gradient(135deg,#155eef,#087f5b); border-color:rgba(255,255,255,.2); box-shadow:0 8px 18px rgba(0,0,0,.18); }
    .navIcon { width:28px; height:28px; display:grid; place-items:center; border-radius:8px; background:rgba(255,255,255,.1); font-size:13px; }
    .sidebarFoot { margin-top:auto; padding:14px 10px 2px; border-top:1px solid rgba(255,255,255,.14); color:#a9bdd1; font-size:12px; line-height:1.5; }
    @media (max-width:900px) {
      .sidebar { position:static; width:auto; min-height:0; padding:14px; }
      .brand { padding:0 4px 12px; }
      .brand small, .navLabel, .sidebarFoot { display:none; }
      .sideNav { grid-template-columns:repeat(3,minmax(0,1fr)); }
      .sideNav a { justify-content:center; }
      main { width:min(1060px,calc(100% - 28px)); margin:0 auto; padding:34px 0; }
    }
    @media (max-width:720px) { h1 { font-size:29px; } .cards { grid-template-columns:1fr; } }
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
      <a class="active" href="http://127.0.0.1:8062"><span class="navIcon">⌂</span><span>Inicio</span></a>
      <a href="http://127.0.0.1:8060"><span class="navIcon">RC</span><span>Recepción</span></a>
      <a href="http://127.0.0.1:8061"><span class="navIcon">DT</span><span>Distancia</span></a>
      <a href="http://127.0.0.1:8063"><span class="navIcon">OR</span><span>Odómetro real</span></a>
    </nav>
    <div class="sidebarFoot">Centro de navegación<br><b style="color:white">Gestión de odómetros</b></div>
  </aside>
  <main>
    <div class="eyebrow">TITANIA · HERRAMIENTAS OPERATIVAS</div>
    <h1>Gestión de odómetros</h1>
    <div class="intro">Selecciona el módulo según el origen del kilometraje. Ambos funcionan de forma independiente.</div>
    <section class="cards">
      <article class="card">
        <div class="icon">01</div>
        <h2>Mantenimiento por recepción</h2>
        <p>Flujo original para trabajar con el kilometraje y fecha de recepción del archivo de mantenimiento.</p>
        <ul><li>Base: KM_RECEPCION</li><li>Fecha: FECHA_RECEPCION</li><li>Enlace local: puerto 8060</li></ul>
        <a href="http://127.0.0.1:8060">Abrir módulo de recepción</a>
      </article>
      <article class="card distance">
        <div class="icon">02</div>
        <h2>Mantenimiento por distancia</h2>
        <p>Nuevo flujo exclusivo para archivos simples. Calcula el kilometraje esperado usando distancia diaria Flespi.</p>
        <ul><li>Campos: CHASIS, FECHA_MANTENIMIENTO y KM_MANTENIMIENTO</li><li>Sin calculadora Flespi</li><li>Enlace local: puerto 8061</li></ul>
        <a href="http://127.0.0.1:8061">Abrir módulo por distancia</a>
      </article>
      <article class="card real">
        <div class="icon">03</div>
        <h2>Odómetro real aproximado</h2>
        <p>Análisis masivo de chasis para reconstruir el odómetro desde la creación y eliminar saltos anormales.</p>
        <ul><li>Entrada: CHASIS y KM_INICIAL opcional</li><li>Compara aproximado contra actual</li><li>Enlace local: puerto 8063</li></ul>
        <a href="http://127.0.0.1:8063">Abrir odómetro real</a>
      </article>
    </section>
    <div class="note">El módulo por distancia busca IMEI e ID Flespi por CHASIS y entrega la columna KM_DEBERIA_ESTAR_HOY.</div>
  </main>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(PAGE)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8062, debug=False)
