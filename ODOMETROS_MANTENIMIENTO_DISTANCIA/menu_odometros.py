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
    main { width:min(1060px,calc(100% - 36px)); margin:0 auto; padding:58px 0; }
    .eyebrow { color:#087f5b; font-size:12px; letter-spacing:1.2px; font-weight:800; }
    h1 { font-size:36px; margin:8px 0 12px; }
    .intro { max-width:700px; color:#52677d; font-size:17px; line-height:1.5; }
    .cards { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:20px; margin-top:34px; }
    .card { background:white; border:1px solid var(--line); border-radius:18px; padding:26px; box-shadow:0 10px 30px rgba(26,68,105,.08); }
    .icon { width:48px; height:48px; display:grid; place-items:center; border-radius:14px; color:white; font-weight:800; font-size:20px; background:var(--blue); }
    .distance .icon { background:var(--green); }
    h2 { margin:18px 0 8px; font-size:22px; }
    p { color:#52677d; line-height:1.5; min-height:72px; }
    ul { padding-left:19px; color:#344b61; line-height:1.7; font-size:14px; }
    a { display:inline-flex; margin-top:18px; text-decoration:none; color:white; background:var(--blue); padding:12px 16px; border-radius:9px; font-weight:700; }
    .distance a { background:var(--green); }
    .note { margin-top:25px; padding:14px 16px; background:#fff; border:1px solid var(--line); border-radius:10px; color:#52677d; font-size:14px; }
    @media (max-width:720px) { main { padding:34px 0; } h1 { font-size:29px; } .cards { grid-template-columns:1fr; } }
  </style>
</head>
<body>
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

