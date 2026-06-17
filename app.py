from flask import Flask, request, redirect, url_for, make_response
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import pytz
import qrcode
import os
import openpyxl
import io
from flask import send_file
from flask import session

LIMA = pytz.timezone("America/Lima")

def ahora_lima():
    return datetime.now(LIMA).replace(tzinfo=None)

from flask import request
app = Flask(__name__)
app.secret_key = "control_personal_2026"

DB_PATH = os.environ.get("DATABASE_URL", "sqlite:///almacen.db")
app.config['SQLALCHEMY_DATABASE_URI'] = DB_PATH
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# =========================
# TABLAS
# =========================

class Trabajador(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(20))
    nombre = db.Column(db.String(200))
    condicion = db.Column(db.String(50))
    area = db.Column(db.String(50))
    supervisor = db.Column(db.String(50))
    estado = db.Column(db.String(20))


class Asistencia(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(20))
    nombre = db.Column(db.String(200))
    fecha = db.Column(db.String(20))
    hora = db.Column(db.String(20))
    supervisor = db.Column(db.String(50))
    tipo = db.Column(db.String(30))
    escaneado_por = db.Column(db.String(50))  # NUEVO

class Movimiento(db.Model):

    id = db.Column(db.Integer, primary_key=True)

    fecha = db.Column(db.String(20))

    codigo = db.Column(db.String(20))

    nombre = db.Column(db.String(200))

    area_anterior = db.Column(db.String(50))

    area_nueva = db.Column(db.String(50))

class AsistenciaEspecial(db.Model):

    id = db.Column(db.Integer, primary_key=True)

    fecha = db.Column(db.String(20))

    codigo = db.Column(db.String(20))

    nombre = db.Column(db.String(200))

    tipo = db.Column(db.String(50))

    supervisor = db.Column(db.String(100))

class HorasExtras(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(20))
    nombre = db.Column(db.String(200))
    fecha = db.Column(db.String(20))
    horas = db.Column(db.Float)
    supervisor = db.Column(db.String(50))

class Incidencia(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(20))
    nombre = db.Column(db.String(200))
    tipo = db.Column(db.String(10))
    descripcion = db.Column(db.String(50))
    fecha_inicio = db.Column(db.String(20))
    fecha_fin = db.Column(db.String(20))
    activo = db.Column(db.Boolean, default=True)

# =========================
# LOGIN
# =========================

USUARIOS = {
    "JOSE":      {"password": "123",      "rol": "VISUALIZADOR"},
    "FRANCISCO": {"password": "1234",     "rol": "SUPERVISOR"},
    "JAROLD":    {"password": "12345",    "rol": "VISUALIZADOR"},
    "JEAN":      {"password": "123456",   "rol": "SUPERVISOR"},
    "OSCAR":     {"password": "44660096", "rol": "ADMIN"},
}

@app.route("/")
def login():
    return """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Control de Personal</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body {
    font-family:'Segoe UI',sans-serif;
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%);
    min-height:100vh;
    display:flex; align-items:center; justify-content:center;
    padding:20px;
}
.card {
    background:#1e293b;
    border-radius:20px;
    padding:36px 28px;
    width:100%; max-width:380px;
    border:1px solid rgba(255,255,255,0.08);
}
.logo {
    width:56px; height:56px; border-radius:14px;
    background: linear-gradient(135deg, #1a56db, #3b82f6);
    display:flex; align-items:center; justify-content:center;
    font-size:24px; font-weight:800; color:white;
    margin:0 auto 16px;
}
h1 {
    text-align:center; color:#fff;
    font-size:18px; font-weight:700;
    letter-spacing:1px; margin-bottom:4px;
}
.sub {
    text-align:center; color:#64748b;
    font-size:12px; margin-bottom:28px;
}
label {
    display:block; font-size:11px; font-weight:700;
    text-transform:uppercase; letter-spacing:1px;
    color:#64748b; margin-bottom:6px;
}
input {
    width:100%; padding:14px 16px;
    background:#0f172a;
    border:1px solid rgba(255,255,255,0.1);
    border-radius:10px; color:#fff;
    font-size:16px; margin-bottom:16px;
    outline:none;
}
input:focus { border-color:#3b82f6; }
button {
    width:100%; padding:15px;
    background: linear-gradient(90deg, #1a56db, #3b82f6);
    border:none; border-radius:10px;
    color:#fff; font-size:16px; font-weight:700;
    cursor:pointer; letter-spacing:1px;
    margin-top:4px;
}
button:active { opacity:0.9; transform:scale(0.99); }
.error {
    background:#7f1d1d; color:#fca5a5;
    border-radius:8px; padding:10px 14px;
    font-size:13px; margin-bottom:16px;
    text-align:center;
}
</style>
</head>
<body>
<div class="card">
  <div class="logo">CP</div>
  <h1>CONTROL DE PERSONAL</h1>
  <p class="sub">Almacén — Acceso al Sistema</p>

  <form action="/dashboard" method="post">
    <label>Usuario</label>
    <input type="text" name="usuario" placeholder="Tu usuario" autocomplete="off" autocapitalize="characters">

    <label>Contraseña</label>
    <input type="password" name="password" placeholder="Tu contraseña">

    <button type="submit">INGRESAR</button>
  </form>
</div>
</body>
</html>
"""


# =========================
# DASHBOARD
# =========================

@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    
    usuario = request.form.get("usuario", "").upper()
    password = request.form.get("password", "")

    if request.method == "POST":
        if usuario not in USUARIOS or USUARIOS[usuario]["password"] != password:
            return """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI',sans-serif; background:linear-gradient(135deg,#0f172a,#1e3a5f); min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }
.card { background:#1e293b; border-radius:20px; padding:36px 28px; width:100%; max-width:380px; border:1px solid rgba(255,255,255,0.08); }
.logo { width:56px; height:56px; border-radius:14px; background:linear-gradient(135deg,#1a56db,#3b82f6); display:flex; align-items:center; justify-content:center; font-size:24px; font-weight:800; color:white; margin:0 auto 16px; }
h1 { text-align:center; color:#fff; font-size:18px; font-weight:700; letter-spacing:1px; margin-bottom:4px; }
.sub { text-align:center; color:#64748b; font-size:12px; margin-bottom:28px; }
label { display:block; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#64748b; margin-bottom:6px; }
input { width:100%; padding:14px 16px; background:#0f172a; border:1px solid rgba(255,255,255,0.1); border-radius:10px; color:#fff; font-size:16px; margin-bottom:16px; outline:none; }
input:focus { border-color:#3b82f6; }
button { width:100%; padding:15px; background:linear-gradient(90deg,#1a56db,#3b82f6); border:none; border-radius:10px; color:#fff; font-size:16px; font-weight:700; cursor:pointer; letter-spacing:1px; margin-top:4px; }
.error { background:#7f1d1d; color:#fca5a5; border-radius:8px; padding:10px 14px; font-size:13px; margin-bottom:16px; text-align:center; }
</style>
</head>
<body>
<div class="card">
  <div class="logo">CP</div>
  <h1>CONTROL DE PERSONAL</h1>
  <p class="sub">Almacén — Acceso al Sistema</p>
  <div class="error">⚠️ Usuario o contraseña incorrectos</div>
  <form action="/dashboard" method="post">
    <label>Usuario</label>
    <input type="text" name="usuario" placeholder="Tu usuario" autocomplete="off" autocapitalize="characters">
    <label>Contraseña</label>
    <input type="password" name="password" placeholder="Tu contraseña">
    <button type="submit">INGRESAR</button>
  </form>
</div>
</body>
</html>
"""
        session["usuario"] = usuario
        session["rol"] = USUARIOS[usuario]["rol"]
    from datetime import date

    hoy = date.today().strftime("%d/%m/%Y")
    mes = date.today().strftime("%m/%Y")
    fecha_larga = date.today().strftime("%d/%m/%Y")

    total = Trabajador.query.count()
    activos = Trabajador.query.filter_by(estado="ACTIVO").count()
    cesados = Trabajador.query.filter_by(estado="CESADO").count()
    asistencias_hoy = Asistencia.query.filter_by(fecha=hoy).count()
    horas_hoy = db.session.query(db.func.sum(HorasExtras.horas)).filter_by(fecha=hoy).scalar() or 0
    movimientos_mes = Movimiento.query.filter(Movimiento.fecha.like(f"%{mes}")).count()
    especiales_hoy = AsistenciaEspecial.query.filter_by(fecha=hoy).count()
    incidencias_activas = Incidencia.query.filter_by(activo=True).count()
    faltantes = activos - asistencias_hoy
    porcentaje = round((asistencias_hoy / activos * 100)) if activos > 0 else 0

    porcentaje = round((asistencias_hoy / activos * 100)) if activos > 0 else 0

    user_agent = request.headers.get('User-Agent', '').lower()
    es_celular = any(x in user_agent for x in ['mobile', 'android', 'iphone', 'ipad'])

    if es_celular:
        return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Control Personal</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#0f172a; color:#fff; padding:12px; }}
.topbar {{ background:linear-gradient(90deg,#1e3a5f,#1a56db); border-radius:12px; padding:12px 16px; margin-bottom:12px; display:flex; align-items:center; justify-content:space-between; }}
.topbar-title {{ font-size:13px; font-weight:700; }}
.topbar-date {{ font-size:11px; background:rgba(255,255,255,0.15); padding:3px 10px; border-radius:20px; }}
.section {{ font-size:10px; font-weight:700; letter-spacing:2px; color:#64748b; text-transform:uppercase; margin:10px 0 6px; }}
.kpi-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:4px; }}
.kpi {{ background:#1e293b; border-radius:10px; padding:10px 12px; border-top:3px solid #3b82f6; }}
.kpi.red {{ border-top-color:#ef4444; }}
.kpi.orange {{ border-top-color:#f59e0b; }}
.kpi.purple {{ border-top-color:#8b5cf6; }}
.kpi.green {{ border-top-color:#10b981; }}
.kpi.slate {{ border-top-color:#64748b; }}
.kpi-label {{ font-size:9px; text-transform:uppercase; letter-spacing:1px; color:#64748b; margin-bottom:4px; }}
.kpi-value {{ font-size:24px; font-weight:800; color:#60a5fa; }}
.kpi.red .kpi-value {{ color:#f87171; }}
.kpi.orange .kpi-value {{ color:#fbbf24; }}
.kpi.purple .kpi-value {{ color:#a78bfa; }}
.kpi.green .kpi-value {{ color:#34d399; }}
.kpi.slate .kpi-value {{ color:#94a3b8; }}
.btn-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:8px; }}
.btn {{ display:flex; align-items:center; gap:8px; background:#1e293b; border:1px solid rgba(255,255,255,0.08); border-radius:10px; padding:12px; text-decoration:none; color:#e2e8f0; font-size:12px; font-weight:600; }}
.btn:active {{ background:#1a56db; }}
.btn-icon {{ font-size:18px; }}
.badge {{ margin-left:auto; background:#ef4444; color:#fff; font-size:10px; padding:2px 6px; border-radius:99px; }}
</style>
</head>
<body>

<div class="topbar">
  <div class="topbar-title">⚙ CONTROL DE PERSONAL</div>
  <div class="topbar-date">{fecha_larga}</div>
</div>

<div class="section">Indicadores del día</div>
<div class="kpi-grid">
  <div class="kpi blue">
    <div class="kpi-label">Asistencias</div>
    <div class="kpi-value">{asistencias_hoy}</div>
  </div>
  <div class="kpi red">
    <div class="kpi-label">Faltantes</div>
    <div class="kpi-value">{faltantes}</div>
  </div>
  <div class="kpi orange">
    <div class="kpi-label">Horas extras</div>
    <div class="kpi-value">{horas_hoy}</div>
  </div>
  <div class="kpi purple">
    <div class="kpi-label">Incidencias</div>
    <div class="kpi-value">{incidencias_activas}</div>
  </div>
</div>

<div class="section">Trabajadores</div>
<div class="kpi-grid">
  <div class="kpi green">
    <div class="kpi-label">Activos</div>
    <div class="kpi-value">{activos}</div>
  </div>
  <div class="kpi red">
    <div class="kpi-label">Cesados</div>
    <div class="kpi-value">{cesados}</div>
  </div>
</div>

<div class="section">Módulos</div>
<div class="btn-grid">
  <a href="/asistencia" class="btn"><span class="btn-icon">✅</span>Asistencia QR</a>
  <a href="/reporte_diario" class="btn"><span class="btn-icon">📅</span>Reporte Diario</a>
  <a href="/horas_extras" class="btn"><span class="btn-icon">⏰</span>Horas Extras</a>
  <a href="/reporte_horas" class="btn"><span class="btn-icon">📊</span>Rep. Horas</a>
  <a href="/asistencias_especiales" class="btn"><span class="btn-icon">⭐</span>Especiales</a>
  <a href="/reporte_asistencias_especiales" class="btn"><span class="btn-icon">📑</span>Rep. Especiales</a>
  <a href="/incidencias" class="btn"><span class="btn-icon">🏥</span>Incidencias{"<span class='badge'>" + str(incidencias_activas) + "</span>" if incidencias_activas > 0 else ""}</a>
  <a href="/reporte_incidencias" class="btn"><span class="btn-icon">📋</span>Rep. Incidencias</a>
  <a href="/trabajadores" class="btn"><span class="btn-icon">👥</span>Trabajadores</a>
  <a href="/reporte_movimientos" class="btn"><span class="btn-icon">🔄</span>Movimientos</a>
  <a href="/exportar_mensual_formato" class="btn"><div class="btn-icon">📊</div>Control Mensual Excel</a>
</div>

</body>
</html>
"""

    else:
        return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Control de Personal</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
html, body {{ height:100vh; overflow:hidden; font-family:'Segoe UI',sans-serif; background:#0f172a; color:#fff; }}
.topbar {{
    background: linear-gradient(90deg, #1e3a5f 0%, #1a56db 100%);
    padding: 0 24px; height: 52px;
    display: flex; align-items: center; justify-content: space-between;
}}
.topbar-brand {{ display:flex; align-items:center; gap:10px; }}
.topbar-logo {{ width:32px; height:32px; border-radius:8px; background:rgba(255,255,255,0.2); display:flex; align-items:center; justify-content:center; font-size:14px; font-weight:700; }}
.topbar-title {{ font-size:15px; font-weight:700; letter-spacing:1px; }}
.topbar-date {{ font-size:12px; background:rgba(255,255,255,0.15); padding:4px 12px; border-radius:20px; }}
.body {{ padding:12px 20px; height:calc(100vh - 52px); display:flex; flex-direction:column; gap:10px; }}
.section-label {{ font-size:10px; font-weight:700; letter-spacing:2px; text-transform:uppercase; color:#94a3b8; margin-bottom:6px; }}
.kpi-row {{ display:grid; gap:10px; }}
.kpi-row.cols-4 {{ grid-template-columns: repeat(4, 1fr); }}
.kpi-row.cols-3 {{ grid-template-columns: repeat(3, 1fr); }}
.kpi {{ background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border-radius:12px; padding:12px 16px; border:1px solid rgba(255,255,255,0.08); position:relative; overflow:hidden; }}
.kpi::before {{ content:''; position:absolute; top:0; left:0; right:0; height:3px; }}
.kpi.blue::before {{ background: linear-gradient(90deg,#3b82f6,#60a5fa); }}
.kpi.red::before {{ background: linear-gradient(90deg,#ef4444,#f87171); }}
.kpi.orange::before {{ background: linear-gradient(90deg,#f59e0b,#fbbf24); }}
.kpi.purple::before {{ background: linear-gradient(90deg,#8b5cf6,#a78bfa); }}
.kpi.green::before {{ background: linear-gradient(90deg,#10b981,#34d399); }}
.kpi.teal::before {{ background: linear-gradient(90deg,#06b6d4,#22d3ee); }}
.kpi.slate::before {{ background: linear-gradient(90deg,#64748b,#94a3b8); }}
.kpi-top {{ display:flex; align-items:center; justify-content:space-between; margin-bottom:8px; }}
.kpi-label {{ font-size:10px; text-transform:uppercase; letter-spacing:1px; color:#64748b; font-weight:600; }}
.kpi-icon {{ font-size:16px; }}
.kpi-value {{ font-size:28px; font-weight:800; line-height:1; }}
.kpi.blue .kpi-value {{ color:#60a5fa; }}
.kpi.red .kpi-value {{ color:#f87171; }}
.kpi.orange .kpi-value {{ color:#fbbf24; }}
.kpi.purple .kpi-value {{ color:#a78bfa; }}
.kpi.green .kpi-value {{ color:#34d399; }}
.kpi.teal .kpi-value {{ color:#22d3ee; }}
.kpi.slate .kpi-value {{ color:#94a3b8; }}
.progress {{ background:rgba(255,255,255,0.08); border-radius:99px; height:4px; margin-top:8px; }}
.progress-fill {{ height:4px; border-radius:99px; background:linear-gradient(90deg,#3b82f6,#60a5fa); width:{porcentaje}%; }}
.kpi-sub {{ font-size:10px; color:#475569; margin-top:4px; }}
.menu-row {{ display:grid; gap:8px; }}
.menu-row.cols-4 {{ grid-template-columns: repeat(4, 1fr); }}
.menu-row.cols-3 {{ grid-template-columns: repeat(3, 1fr); }}
.btn {{ display:flex; align-items:center; gap:8px; background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border:1px solid rgba(255,255,255,0.08); border-radius:10px; padding:10px 14px; text-decoration:none; color:#e2e8f0; font-size:12px; font-weight:600; transition:all 0.15s; }}
.btn:hover {{ background: linear-gradient(135deg, #1a56db 0%, #1e3a5f 100%); border-color:#3b82f6; color:#fff; transform:translateY(-1px); }}
.btn-icon {{ width:28px; height:28px; border-radius:7px; display:flex; align-items:center; justify-content:center; font-size:13px; flex-shrink:0; background:rgba(255,255,255,0.08); }}
.badge {{ margin-left:auto; background:#ef4444; color:#fff; font-size:10px; font-weight:700; padding:2px 7px; border-radius:99px; }}
</style>
</head>
<body>

<div class="topbar">
  <div class="topbar-brand">
    <div class="topbar-logo">CP</div>
    <div class="topbar-title">CONTROL DE PERSONAL — ALMACÉN | Global Sourcing S.A.C.</div>
  </div>
  <div class="topbar-date">📅 {fecha_larga}</div>
</div>

<div class="body">

  <div>
    <div class="section-label">Indicadores del día</div>
    <div class="kpi-row cols-4">
      <div class="kpi blue">
        <div class="kpi-top"><span class="kpi-label">Asistencias</span><span class="kpi-icon">✅</span></div>
        <div class="kpi-value">{asistencias_hoy}</div>
        <div class="progress"><div class="progress-fill"></div></div>
        <div class="kpi-sub">{porcentaje}% de {activos} activos</div>
      </div>
      <div class="kpi red">
        <div class="kpi-top"><span class="kpi-label">Faltantes</span><span class="kpi-icon">❌</span></div>
        <div class="kpi-value">{faltantes}</div>
        <div class="kpi-sub">Sin marcar hoy</div>
      </div>
      <div class="kpi orange">
        <div class="kpi-top"><span class="kpi-label">Horas extras</span><span class="kpi-icon">⏰</span></div>
        <div class="kpi-value">{horas_hoy}</div>
        <div class="kpi-sub">Registradas hoy</div>
      </div>
      <div class="kpi purple">
        <div class="kpi-top"><span class="kpi-label">Incidencias</span><span class="kpi-icon">🏥</span></div>
        <div class="kpi-value">{incidencias_activas}</div>
        <div class="kpi-sub">Activas</div>
      </div>
    </div>
  </div>

  <div>
    <div class="section-label">Trabajadores</div>
    <div class="kpi-row cols-3">
      <div class="kpi green">
        <div class="kpi-top"><span class="kpi-label">Activos</span><span class="kpi-icon">👥</span></div>
        <div class="kpi-value">{activos}</div>
      </div>
      <div class="kpi red">
        <div class="kpi-top"><span class="kpi-label">Cesados</span><span class="kpi-icon">🚪</span></div>
        <div class="kpi-value">{cesados}</div>
      </div>
      <div class="kpi slate">
        <div class="kpi-top"><span class="kpi-label">Total registrados</span><span class="kpi-icon">📋</span></div>
        <div class="kpi-value">{total}</div>
      </div>
    </div>
  </div>

  <div>
    <div class="section-label">Módulos</div>
    <div style="display:flex; flex-direction:column; gap:8px;">
      <div class="menu-row cols-4">
        <a href="/asistencia" class="btn"><div class="btn-icon">✅</div>Asistencia QR</a>
        <a href="/reporte_diario" class="btn"><div class="btn-icon">📅</div>Reporte Diario</a>
        <a href="/horas_extras" class="btn"><div class="btn-icon">⏰</div>Horas Extras</a>
        <a href="/reporte_horas" class="btn"><div class="btn-icon">📊</div>Reporte Horas Extras</a>
      </div>
      <div class="menu-row cols-4">
        <a href="/asistencias_especiales" class="btn"><div class="btn-icon">⭐</div>Asistencias Especiales</a>
        <a href="/reporte_asistencias_especiales" class="btn"><div class="btn-icon">📑</div>Reporte Especiales</a>
        <a href="/incidencias" class="btn"><div class="btn-icon">🏥</div>Incidencias{"<span class='badge'>" + str(incidencias_activas) + "</span>" if incidencias_activas > 0 else ""}</a>
        <a href="/reporte_incidencias" class="btn"><div class="btn-icon">📋</div>Reporte Incidencias</a>
      </div>
      <div class="menu-row cols-3">
        <a href="/trabajadores" class="btn"><div class="btn-icon">👥</div>Trabajadores</a>
        <a href="/reporte_movimientos" class="btn"><div class="btn-icon">🔄</div>Movimientos Personal</a>
        <a href="/exportar_mensual_formato" class="btn"><div class="btn-icon">📊</div>Control Mensual Excel</a>
      </div>
    </div>
  </div>

</div>
</body>
</html>
"""


# =========================
# ASISTENCIA
# =========================

@app.route("/asistencia")
def asistencia():
    if session.get("rol") == "VISUALIZADOR":
        return redirect("/dashboard")
    return """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Asistencia QR</title>
</head>
<body style="font-family:Segoe UI,sans-serif;text-align:center;padding:20px;background:#f0f2f5;">

<h2 style="color:#0f172a;">ASISTENCIA QR</h2>

<div id="reader" style="width:300px;margin:0 auto;"></div>

<br>
<p style="color:#64748b;font-size:13px;">Apunta la cámara al código QR</p>

<hr style="margin:20px 0;">

<p style="font-size:13px;color:#64748b;">O escribe el código manualmente:</p>
<form action="/registrar_asistencia" method="post">
    <input type="text"
           name="codigo"
           placeholder="Código trabajador"
           autocomplete="off"
           onchange="this.form.submit()"
           style="padding:10px;font-size:16px;width:220px;border:1px solid #ccc;border-radius:8px;">
</form>

<br>
<a href="/dashboard" style="color:#3b82f6;font-size:13px;">Volver al Dashboard</a>

<script src="https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js"></script>
<script>
let ultimoCodigo = null;

function beep() {
    const audio = new Audio("https://actions.google.com/sounds/v1/alarms/beep_short.ogg");
    audio.play();
}

function onScanSuccess(decodedText) {
    html5QrCode.stop().then(() => {
        var form = document.createElement("form");
        form.method = "POST";
        form.action = "/registrar_asistencia";
        var input = document.createElement("input");
        input.type = "hidden";
        input.name = "codigo";
        input.value = decodedText;
        form.appendChild(input);
        document.body.appendChild(form);
        form.submit();
    });
}

var html5QrCode = new Html5Qrcode("reader");
html5QrCode.start(
    { facingMode: "environment" },
    { fps: 10, qrbox: 250 },
    onScanSuccess
);
</script>
</body>
</html>
    """
@app.route("/registrar_asistencia", methods=["POST"])
def registrar_asistencia():

    codigo = request.form["codigo"]

    trabajador = Trabajador.query.filter_by(
        codigo=codigo
    ).first()

    if not trabajador:
        return """
        <h2>TRABAJADOR NO ENCONTRADO</h2>
        <a href="/asistencia">VOLVER</a>
        """

    ahora = ahora_lima()
    hoy = ahora.strftime("%d/%m/%Y")

    # Verificar si ya marcó hoy
    ya_registro = Asistencia.query.filter_by(
        codigo=trabajador.codigo,
        fecha=hoy
    ).first()

    if ya_registro:
        session["resultado"] = {
            "tipo": "ya_registrado",
            "nombre": trabajador.nombre,
            "hora": ya_registro.hora
        }
        return redirect("/resultado_asistencia")


    registro = Asistencia(
        codigo=trabajador.codigo,
        nombre=trabajador.nombre,
        fecha=hoy,
        hora=ahora.strftime("%H:%M:%S"),
        supervisor=trabajador.supervisor,
        tipo="ASISTENCIA",
        escaneado_por=session.get("usuario", "DESCONOCIDO")
    )
    db.session.add(registro)
    db.session.commit()

    session["resultado"] = {
        "tipo": "exitoso",
        "nombre": trabajador.nombre,
        "area": trabajador.area,
        "hora": ahora.strftime("%H:%M:%S")
    }
    return redirect("/resultado_asistencia")

@app.route("/resultado_asistencia")
def resultado_asistencia():
    resultado = session.pop("resultado", None)
    if not resultado:
        return redirect("/asistencia")

    if resultado["tipo"] == "ya_registrado":
        return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }}
.card {{ background:#1e293b; border-radius:16px; padding:32px 24px; width:100%; max-width:380px; text-align:center; }}
.icon {{ font-size:48px; margin-bottom:16px; }}
h2 {{ color:#fbbf24; font-size:20px; margin-bottom:10px; }}
p {{ color:#94a3b8; font-size:14px; margin-bottom:6px; }}
b {{ color:#fff; }}
.btn {{ display:block; margin-top:24px; padding:14px; background:linear-gradient(90deg,#1a56db,#3b82f6); border-radius:10px; color:#fff; font-size:16px; font-weight:700; text-decoration:none; }}
</style>
</head>
<body>
<div class="card">
  <div class="icon">⚠️</div>
  <h2>Ya registrado hoy</h2>
  <p><b>{resultado["nombre"]}</b></p>
  <p>Hora de registro: <b>{resultado["hora"]}</b></p>
  <a href="/asistencia" class="btn">SIGUIENTE</a>
</div>
</body>
</html>
        """

    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }}
.card {{ background:#1e293b; border-radius:16px; padding:32px 24px; width:100%; max-width:380px; text-align:center; }}
.icon {{ font-size:48px; margin-bottom:16px; }}
h2 {{ color:#34d399; font-size:20px; margin-bottom:10px; }}
p {{ color:#94a3b8; font-size:14px; margin-bottom:6px; }}
b {{ color:#fff; }}
.area {{ background:#0f172a; border-radius:8px; padding:8px 16px; display:inline-block; color:#60a5fa; font-size:13px; margin:8px 0; }}
.hora {{ font-size:28px; font-weight:800; color:#fff; margin:10px 0; }}
.btn {{ display:block; margin-top:24px; padding:14px; background:linear-gradient(90deg,#10b981,#34d399); border-radius:10px; color:#fff; font-size:16px; font-weight:700; text-decoration:none; }}
</style>
</head>
<body>
<div class="card">
  <div class="icon">✅</div>
  <h2>Asistencia Registrada</h2>
  <p><b>{resultado["nombre"]}</b></p>
  <div class="area">{resultado["area"]}</div>
  <div class="hora">{resultado["hora"]}</div>
  <a href="/asistencia" class="btn">SIGUIENTE</a>
</div>
</body>
</html>
    """

# =========================
# TRABAJADORES
# =========================

@app.route("/nuevo_trabajador")
def nuevo_trabajador():

    return """
    <h1>NUEVO TRABAJADOR</h1>

    <form action="/guardar_trabajador" method="post">

        Nombre Completo:<br>
        <input type="text" name="nombre">

        <br><br>

        Condición:<br>
        <select name="condicion">
            <option>FIJO</option>
            <option>DOTACION</option>
            <option>CAMPAÑA</option>
        </select>

        <br><br>

        Área:<br>
        <select name="area">
            <option>RECEPCION</option>
            <option>REPOSICION</option>
            <option>PICKING</option>
            <option>PACKING</option>
        </select>

        <br><br>

        Supervisor:<br>
        <input type="text" name="supervisor">

        <br><br>

        <button type="submit">
        GUARDAR
        </button>

    </form>
    """
@app.route("/guardar_trabajador", methods=["POST"])
def guardar_trabajador():

    nombre = request.form["nombre"]
    condicion = request.form["condicion"]
    area = request.form["area"]
    supervisor = request.form["supervisor"]

    ultimo = Trabajador.query.order_by(
        Trabajador.id.desc()
    ).first()

    if ultimo:

        numero = int(ultimo.codigo.replace("ALM-", "")) + 1

    else:

        numero = 1

    codigo = f"ALM-{numero:04d}"

    trabajador = Trabajador(
        codigo=codigo,
        nombre=nombre,
        condicion=condicion,
        area=area,
        supervisor=supervisor,
        estado="ACTIVO"
    )

    db.session.add(trabajador)
    db.session.commit()

    # Generar QR

    ruta_qr = f"static/qr/{codigo}.png"

    qr = qrcode.make(codigo)

    qr.save(ruta_qr)

    return f"""
    <h2>TRABAJADOR REGISTRADO</h2>

    <p>Código generado:</p>

    <h1>{codigo}</h1>

    <p>{nombre}</p>

    <a href='/trabajadores'>
    VOLVER
    </a>
    """
@app.route("/trabajadores")
def trabajadores():

    trabajadores = Trabajador.query.order_by(
        Trabajador.codigo
    ).all()

    html = """
<h1>TRABAJADORES</h1>

<a href="/nuevo_trabajador">
NUEVO TRABAJADOR
</a>

<br><br>

<table border="1" cellpadding="5">

<tr>
    <th>CODIGO</th>
    <th>NOMBRE</th>
    <th>AREA</th>
    <th>CONDICION</th>
    <th>ESTADO</th>
    <th>EDITAR</th>
    <th>MOVER</th>
    <th>CESAR</th>
</tr>
"""

    for t in trabajadores:

        html += f"""
        <tr>
    <td>{t.codigo}</td>
    <td>{t.nombre}</td>
    <td>{t.area}</td>
    <td>{t.condicion}</td>
    <td>{t.estado}</td>

    <td>
        <a href="/editar/{t.id}">
        EDITAR
        </a>
    </td>
    <td>
        <a href="/mover/{t.id}">
        MOVER
        </a>
    </td>

    <td>
        <a href="/cesar/{t.id}">
        CESAR
        </a>
    </td>
</tr>
        """

    html += "</table>"
    html += """
    <br><br>
    <a href="/exportar_trabajadores">📥 DESCARGAR EXCEL</a>
    <br><br>
    <a href="/dashboard">⬅️ VOLVER AL DASHBOARD</a>
    """
    return html

# =========================
# INICIO
# =========================
@app.route("/mover/<int:id>")
def mover(id):

    trabajador = Trabajador.query.get(id)

    return f"""
    <h1>MOVER TRABAJADOR</h1>

    <p><b>{trabajador.codigo}</b></p>
    <p>{trabajador.nombre}</p>

    <p>Área actual: {trabajador.area}</p>

    <form action="/guardar_movimiento/{id}" method="post">

        Nueva Área:<br>

        <select name="area">
            <option>RECEPCION</option>
            <option>REPOSICION</option>
            <option>PICKING</option>
            <option>PACKING</option>
            <option>INVENTARIOS</option>
            <option>OTROS</option>
        </select>

        <br><br>

        <button type="submit">
        GUARDAR
        </button>

    </form>
    """


@app.route("/guardar_movimiento/<int:id>", methods=["POST"])
def guardar_movimiento(id):

    trabajador = Trabajador.query.get(id)

    area_anterior = trabajador.area

    area_nueva = request.form["area"]

    movimiento = Movimiento(
        fecha=datetime.now().strftime("%d/%m/%Y"),
        codigo=trabajador.codigo,
        nombre=trabajador.nombre,
        area_anterior=area_anterior,
        area_nueva=area_nueva
    )

    trabajador.area = area_nueva

    db.session.add(movimiento)

    db.session.commit()

    return """
    <h2>MOVIMIENTO REGISTRADO</h2>
    <a href="/trabajadores">VOLVER A LISTA</a>
    &nbsp;&nbsp;
    <a href="/dashboard">DASHBOARD</a>
    """


@app.route("/cesar/<int:id>")
def cesar(id):

    trabajador = Trabajador.query.get(id)

    trabajador.estado = "CESADO"

    db.session.commit()

    return """
    <h2>TRABAJADOR CESADO CORRECTAMENTE</h2>
    <a href="/trabajadores">VOLVER A LISTA</a>
    &nbsp;&nbsp;
    <a href="/dashboard">DASHBOARD</a>
    """

@app.route("/horas_extras")
def horas_extras():
    return """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Horas Extras</title>
</head>
<body style="font-family:Segoe UI,sans-serif;text-align:center;padding:20px;background:#f0f2f5;">

<h2 style="color:#0f172a;">HORAS EXTRAS</h2>

<div id="reader" style="width:300px;margin:0 auto;"></div>

<br>
<p style="color:#64748b;font-size:13px;">Apunta la cámara al código QR</p>

<hr style="margin:20px 0;">

<p style="font-size:13px;color:#64748b;">O escribe el código manualmente:</p>
<form action="/buscar_trabajador_horas" method="post">
    <input type="text"
           name="codigo"
           placeholder="Código trabajador"
           autocomplete="off"
           onchange="this.form.submit()"
           style="padding:10px;font-size:16px;width:220px;border:1px solid #ccc;border-radius:8px;">
</form>

<br>
<a href="/dashboard" style="color:#3b82f6;font-size:13px;">Volver al Dashboard</a>

<script src="https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js"></script>
<script>
function onScanSuccess(decodedText) {
    html5QrCode.stop().then(() => {
        var form = document.createElement("form");
        form.method = "POST";
        form.action = "/buscar_trabajador_horas";
        var input = document.createElement("input");
        input.type = "hidden";
        input.name = "codigo";
        input.value = decodedText;
        form.appendChild(input);
        document.body.appendChild(form);
        form.submit();
    });
}

var html5QrCode = new Html5Qrcode("reader");
html5QrCode.start(
    { facingMode: "environment" },
    { fps: 10, qrbox: 250 },
    onScanSuccess
);
</script>
</body>
</html>
    """

@app.route("/buscar_trabajador_horas", methods=["POST"])
def buscar_trabajador_horas():

    codigo = request.form["codigo"]

    trabajador = Trabajador.query.filter_by(
        codigo=codigo
    ).first()

    if not trabajador:
        return """
        <h2>TRABAJADOR NO ENCONTRADO</h2>
        <a href="/horas_extras">VOLVER</a>
        """

    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }}
.card {{ background:#1e293b; border-radius:16px; padding:28px 24px; width:100%; max-width:380px; }}
.worker {{ text-align:center; margin-bottom:20px; }}
.worker-name {{ color:#fff; font-size:18px; font-weight:700; margin-bottom:4px; }}
.worker-area {{ background:#0f172a; border-radius:8px; padding:6px 14px; display:inline-block; color:#60a5fa; font-size:13px; }}
label {{ display:block; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#64748b; margin-bottom:6px; }}
select, input {{ width:100%; padding:12px 14px; background:#0f172a; border:1px solid rgba(255,255,255,0.1); border-radius:8px; color:#fff; font-size:15px; margin-bottom:16px; outline:none; }}
button {{ width:100%; padding:14px; background:linear-gradient(90deg,#f59e0b,#fbbf24); border:none; border-radius:8px; color:#0f172a; font-size:16px; font-weight:700; cursor:pointer; }}
a {{ display:block; text-align:center; color:#64748b; font-size:12px; margin-top:14px; text-decoration:none; }}
</style>
</head>
<body>
<div class="card">
  <div class="worker">
    <div class="worker-name">{trabajador.nombre}</div>
    <div class="worker-area">{trabajador.area}</div>
  </div>

  <form action="/guardar_horas_extras" method="post">
    <input type="hidden" name="codigo" value="{trabajador.codigo}">

    <label>Horas</label>
    <select name="horas">
        <option value="1">1 HORA</option>
        <option value="2">2 HORAS</option>
        <option value="3">3 HORAS</option>
        <option value="4">4 HORAS</option>
        <option value="5">5 HORAS</option>
    </select>

    <label>Supervisor</label>
    <input type="text" name="supervisor" value="{trabajador.supervisor}">

    <button type="submit">GUARDAR</button>
  </form>
  <a href="/horas_extras">← Volver</a>
</div>
</body>
</html>
    """

@app.route("/guardar_horas_extras", methods=["POST"])
def guardar_horas_extras():

    codigo = request.form["codigo"]

    trabajador = Trabajador.query.filter_by(
        codigo=codigo
    ).first()

    hoy = datetime.now().strftime("%d/%m/%Y")

    # Verificar si ya registró horas extras hoy
    ya_registro = HorasExtras.query.filter_by(
        codigo=trabajador.codigo,
        fecha=hoy
    ).first()

    if ya_registro:
        return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }}
.card {{ background:#1e293b; border-radius:16px; padding:32px 24px; width:100%; max-width:380px; text-align:center; }}
.icon {{ font-size:48px; margin-bottom:16px; }}
h2 {{ color:#fbbf24; font-size:20px; margin-bottom:10px; }}
p {{ color:#94a3b8; font-size:14px; margin-bottom:6px; }}
b {{ color:#fff; }}
.btn {{ display:block; margin-top:24px; padding:14px; background:linear-gradient(90deg,#1a56db,#3b82f6); border-radius:10px; color:#fff; font-size:16px; font-weight:700; text-decoration:none; }}
</style>
</head>
<body>
<div class="card">
  <div class="icon">⚠️</div>
  <h2>Ya registrado hoy</h2>
  <p><b>{trabajador.nombre}</b></p>
  <p>Ya tiene horas extras registradas hoy.</p>
  <a href="/horas_extras" class="btn">SIGUIENTE</a>
</div>
</body>
</html>
        """

    registro = HorasExtras(
        codigo=trabajador.codigo,
        nombre=trabajador.nombre,
        fecha=hoy,
        horas=float(request.form["horas"]),
        supervisor=request.form["supervisor"]
    )
    db.session.add(registro)
    db.session.commit()

    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }}
.card {{ background:#1e293b; border-radius:16px; padding:32px 24px; width:100%; max-width:380px; text-align:center; }}
.icon {{ font-size:48px; margin-bottom:16px; }}
h2 {{ color:#34d399; font-size:20px; margin-bottom:10px; }}
p {{ color:#94a3b8; font-size:14px; margin-bottom:6px; }}
b {{ color:#fff; }}
.horas {{ font-size:40px; font-weight:800; color:#fbbf24; margin:10px 0; }}
.label {{ color:#64748b; font-size:12px; }}
.btn {{ display:block; margin-top:24px; padding:14px; background:linear-gradient(90deg,#f59e0b,#fbbf24); border-radius:10px; color:#0f172a; font-size:16px; font-weight:700; text-decoration:none; }}
</style>
</head>
<body>
<div class="card">
  <div class="icon">⏰</div>
  <h2>Horas Extras Registradas</h2>
  <p><b>{trabajador.nombre}</b></p>
  <div class="horas">{request.form['horas']}</div>
  <div class="label">horas registradas</div>
  <a href="/horas_extras" class="btn">SIGUIENTE</a>
</div>
</body>
</html>
    """
@app.route("/reporte_horas")
def reporte_horas():
    return """
    <h1>REPORTE HORAS EXTRAS</h1>

    <form action="/filtrar_horas" method="post">
        Fecha inicio:<br>
        <input type="date" name="fecha_inicio">
        <br><br>
        Fecha fin:<br>
        <input type="date" name="fecha_fin">
        <br><br>
        <button type="submit">VER REPORTE</button>
    </form>

    <br>
    <a href="/dashboard">Volver</a>
    """

@app.route("/filtrar_horas", methods=["POST"])
def filtrar_horas():

    from datetime import datetime, timedelta

    fecha_inicio = request.form["fecha_inicio"]
    fecha_fin = request.form["fecha_fin"]

    fi = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    ff = datetime.strptime(fecha_fin, "%Y-%m-%d")

    fechas = []
    actual = fi
    while actual <= ff:
        fechas.append(actual.strftime("%d/%m/%Y"))
        actual += timedelta(days=1)

    registros = HorasExtras.query.filter(
        HorasExtras.fecha.in_(fechas)
    ).order_by(HorasExtras.fecha, HorasExtras.nombre).all()

    total = sum(r.horas for r in registros)

    html = f"""
    <h1>REPORTE HORAS EXTRAS</h1>
    <p>Del <b>{fi.strftime("%d/%m/%Y")}</b> al <b>{ff.strftime("%d/%m/%Y")}</b></p>
    <p>Total horas: <b>{total}</b></p>

    <table border="1" cellpadding="5">
    <tr>
        <th>FECHA</th><th>CODIGO</th><th>NOMBRE</th><th>HORAS</th><th>SUPERVISOR</th>
    </tr>
    """

    for r in registros:
        html += f"""
        <tr>
            <td>{r.fecha}</td><td>{r.codigo}</td><td>{r.nombre}</td>
            <td>{r.horas}</td><td>{r.supervisor}</td>
        </tr>
        """

    html += f"""
    </table>
    <br>
    <a href="/exportar_horas_filtrado?fi={fecha_inicio}&ff={fecha_fin}">📥 EXPORTAR EXCEL</a>
    &nbsp;&nbsp;
    <a href="/reporte_horas">NUEVO FILTRO</a>
    &nbsp;&nbsp;
    <a href="/dashboard">DASHBOARD</a>
    """

    return html

@app.route("/exportar_horas_filtrado")
def exportar_horas_filtrado():

    from datetime import datetime, timedelta

    fecha_inicio = request.args.get("fi")
    fecha_fin = request.args.get("ff")

    fi = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    ff = datetime.strptime(fecha_fin, "%Y-%m-%d")

    fechas = []
    actual = fi
    while actual <= ff:
        fechas.append(actual.strftime("%d/%m/%Y"))
        actual += timedelta(days=1)

    registros = HorasExtras.query.filter(
        HorasExtras.fecha.in_(fechas)
    ).order_by(HorasExtras.fecha).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Horas Extras"
    ws.append([f"REPORTE HORAS EXTRAS {fi.strftime('%d/%m/%Y')} AL {ff.strftime('%d/%m/%Y')}"])
    ws.append([])
    ws.append(["FECHA", "CODIGO", "NOMBRE", "HORAS", "SUPERVISOR"])
    for r in registros:
        ws.append([r.fecha, r.codigo, r.nombre, r.horas, r.supervisor])

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"horas_extras_{fecha_inicio}_{fecha_fin}.xlsx"
    )

@app.route("/reporte_asistencia")
def reporte_asistencia():

    registros = Asistencia.query.order_by(
        Asistencia.id.desc()
    ).all()

    html = """
    <h1>REPORTE DE ASISTENCIA</h1>

    <table border="1" cellpadding="5">

    <tr>
        <th>FECHA</th>
        <th>HORA</th>
        <th>CODIGO</th>
        <th>NOMBRE</th>
        <th>SUPERVISOR</th>
    </tr>
    """

    for r in registros:

        html += f"""
        <tr>
            <td>{r.fecha}</td>
            <td>{r.hora}</td>
            <td>{r.codigo}</td>
            <td>{r.nombre}</td>
            <td>{r.supervisor}</td>
        </tr>
        """

    html += "</table>"
    html += """
    <br><br>
    <a href="/exportar_asistencia">📥 DESCARGAR EXCEL</a>
    """
    return html

@app.route("/generar_qr_todos")
def generar_qr_todos():

    import qrcode
    import os

    os.makedirs("static/qr", exist_ok=True)

    trabajadores = Trabajador.query.all()

    total = 0

    for t in trabajadores:

        ruta = f"static/qr/{t.codigo}.png"

        qr = qrcode.make(t.codigo)

        qr.save(ruta)

        total += 1

    return f"""
    <h2>QR GENERADOS CORRECTAMENTE</h2>

    <p>Total generados: {total}</p>

    <a href="/trabajadores">
    VOLVER
    </a>
    """
@app.route("/reporte_movimientos")
def reporte_movimientos():

    movimientos = Movimiento.query.order_by(
        Movimiento.id.desc()
    ).all()

    html = """
    <h1>MOVIMIENTOS DE PERSONAL</h1>

    <table border="1" cellpadding="5">

    <tr>
        <th>FECHA</th>
        <th>CODIGO</th>
        <th>NOMBRE</th>
        <th>AREA ANTERIOR</th>
        <th>AREA NUEVA</th>
    </tr>
    """

    for m in movimientos:

        html += f"""
        <tr>
            <td>{m.fecha}</td>
            <td>{m.codigo}</td>
            <td>{m.nombre}</td>
            <td>{m.area_anterior}</td>
            <td>{m.area_nueva}</td>
        </tr>
        """

    html += "</table>"
    html += """
    <br><br>
    <a href="/exportar_asistencia">📥 DESCARGAR EXCEL</a>
    <br><br>
    <a href="/dashboard">⬅️ VOLVER AL DASHBOARD</a>
    """
    return html

@app.route("/asistencias_especiales")
def asistencias_especiales():
    return """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Asistencias Especiales</title>
</head>
<body style="font-family:Segoe UI,sans-serif;text-align:center;padding:20px;background:#f0f2f5;">

<h2 style="color:#0f172a;">ASISTENCIAS ESPECIALES</h2>

<div id="reader" style="width:300px;margin:0 auto;"></div>

<br>
<p style="color:#64748b;font-size:13px;">Apunta la cámara al código QR</p>

<hr style="margin:20px 0;">

<p style="font-size:13px;color:#64748b;">O escribe el código manualmente:</p>
<form action="/buscar_trabajador_especial" method="post">
    <input type="text"
           name="codigo"
           placeholder="Código trabajador"
           autocomplete="off"
           onchange="this.form.submit()"
           style="padding:10px;font-size:16px;width:220px;border:1px solid #ccc;border-radius:8px;">
</form>

<br>
<a href="/dashboard" style="color:#3b82f6;font-size:13px;">Volver al Dashboard</a>

<script src="https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js"></script>
<script>
function onScanSuccess(decodedText) {
    html5QrCode.stop().then(() => {
        var form = document.createElement("form");
        form.method = "POST";
        form.action = "/buscar_trabajador_especial";
        var input = document.createElement("input");
        input.type = "hidden";
        input.name = "codigo";
        input.value = decodedText;
        form.appendChild(input);
        document.body.appendChild(form);
        form.submit();
    });
}

var html5QrCode = new Html5Qrcode("reader");
html5QrCode.start(
    { facingMode: "environment" },
    { fps: 10, qrbox: 250 },
    onScanSuccess
);
</script>
</body>
</html>
    """

@app.route("/buscar_trabajador_especial", methods=["POST"])
def buscar_trabajador_especial():

    codigo = request.form["codigo"]

    trabajador = Trabajador.query.filter_by(
        codigo=codigo
    ).first()

    if not trabajador:
        return """
        <h2>TRABAJADOR NO ENCONTRADO</h2>
        <a href="/asistencias_especiales">VOLVER</a>
        """

    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }}
.card {{ background:#1e293b; border-radius:16px; padding:28px 24px; width:100%; max-width:380px; }}
.worker {{ text-align:center; margin-bottom:20px; }}
.worker-name {{ color:#fff; font-size:18px; font-weight:700; margin-bottom:4px; }}
.worker-area {{ background:#0f172a; border-radius:8px; padding:6px 14px; display:inline-block; color:#60a5fa; font-size:13px; }}
label {{ display:block; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#64748b; margin-bottom:6px; }}
select, input {{ width:100%; padding:12px 14px; background:#0f172a; border:1px solid rgba(255,255,255,0.1); border-radius:8px; color:#fff; font-size:15px; margin-bottom:16px; outline:none; }}
button {{ width:100%; padding:14px; background:linear-gradient(90deg,#8b5cf6,#a78bfa); border:none; border-radius:8px; color:#fff; font-size:16px; font-weight:700; cursor:pointer; }}
a {{ display:block; text-align:center; color:#64748b; font-size:12px; margin-top:14px; text-decoration:none; }}
</style>
</head>
<body>
<div class="card">
  <div class="worker">
    <div class="worker-name">{trabajador.nombre}</div>
    <div class="worker-area">{trabajador.area}</div>
  </div>

  <form action="/guardar_asistencia_especial" method="post">
    <input type="hidden" name="codigo" value="{trabajador.codigo}">

    <label>Tipo</label>
    <select name="tipo">
        <option>SABADO</option>
        <option>DOMINGO</option>
        <option>FERIADO</option>
        <option>INVENTARIO</option>
        <option>APOYO</option>
    </select>

    <label>Supervisor</label>
    <input type="text" name="supervisor" value="{trabajador.supervisor}">

    <button type="submit">GUARDAR</button>
  </form>
  <a href="/asistencias_especiales">← Volver</a>
</div>
</body>
</html>
    """
@app.route("/guardar_asistencia_especial", methods=["POST"])
def guardar_asistencia_especial():

    codigo = request.form["codigo"]

    trabajador = Trabajador.query.filter_by(
        codigo=codigo
    ).first()

    hoy = datetime.now().strftime("%d/%m/%Y")

    # Verificar si ya registró asistencia especial hoy
    ya_registro = AsistenciaEspecial.query.filter_by(
        codigo=trabajador.codigo,
        fecha=hoy
    ).first()

    if ya_registro:
        return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }}
.card {{ background:#1e293b; border-radius:16px; padding:32px 24px; width:100%; max-width:380px; text-align:center; }}
.icon {{ font-size:48px; margin-bottom:16px; }}
h2 {{ color:#fbbf24; font-size:20px; margin-bottom:10px; }}
p {{ color:#94a3b8; font-size:14px; margin-bottom:6px; }}
b {{ color:#fff; }}
.btn {{ display:block; margin-top:24px; padding:14px; background:linear-gradient(90deg,#8b5cf6,#a78bfa); border-radius:10px; color:#fff; font-size:16px; font-weight:700; text-decoration:none; }}
</style>
</head>
<body>
<div class="card">
  <div class="icon">⚠️</div>
  <h2>Ya registrado hoy</h2>
  <p><b>{trabajador.nombre}</b></p>
  <p>Ya tiene asistencia especial registrada hoy.</p>
  <a href="/asistencias_especiales" class="btn">SIGUIENTE</a>
</div>
</body>
</html>
        """

    registro = AsistenciaEspecial(
        fecha=hoy,
        codigo=trabajador.codigo,
        nombre=trabajador.nombre,
        tipo=request.form["tipo"],
        supervisor=request.form["supervisor"]
    )
    db.session.add(registro)
    db.session.commit()

    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }}
.card {{ background:#1e293b; border-radius:16px; padding:32px 24px; width:100%; max-width:380px; text-align:center; }}
.icon {{ font-size:48px; margin-bottom:16px; }}
h2 {{ color:#34d399; font-size:20px; margin-bottom:10px; }}
p {{ color:#94a3b8; font-size:14px; margin-bottom:6px; }}
b {{ color:#fff; }}
.tipo {{ background:#0f172a; border-radius:8px; padding:8px 16px; display:inline-block; color:#a78bfa; font-size:16px; font-weight:700; margin:10px 0; }}
.btn {{ display:block; margin-top:24px; padding:14px; background:linear-gradient(90deg,#8b5cf6,#a78bfa); border-radius:10px; color:#fff; font-size:16px; font-weight:700; text-decoration:none; }}
</style>
</head>
<body>
<div class="card">
  <div class="icon">⭐</div>
  <h2>Asistencia Especial Registrada</h2>
  <p><b>{trabajador.nombre}</b></p>
  <div class="tipo">{request.form['tipo']}</div>
  <a href="/asistencias_especiales" class="btn">SIGUIENTE</a>
</div>
</body>
</html>
    """
@app.route("/reporte_asistencias_especiales")
def reporte_asistencias_especiales():
    return """
    <h1>REPORTE ASISTENCIAS ESPECIALES</h1>

    <form action="/filtrar_especiales" method="post">
        Fecha inicio:<br>
        <input type="date" name="fecha_inicio">
        <br><br>
        Fecha fin:<br>
        <input type="date" name="fecha_fin">
        <br><br>
        <button type="submit">VER REPORTE</button>
    </form>

    <br>
    <a href="/dashboard">Volver</a>
    """

@app.route("/filtrar_especiales", methods=["POST"])
def filtrar_especiales():

    from datetime import datetime, timedelta

    fecha_inicio = request.form["fecha_inicio"]
    fecha_fin = request.form["fecha_fin"]

    fi = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    ff = datetime.strptime(fecha_fin, "%Y-%m-%d")

    fechas = []
    actual = fi
    while actual <= ff:
        fechas.append(actual.strftime("%d/%m/%Y"))
        actual += timedelta(days=1)

    registros = AsistenciaEspecial.query.filter(
        AsistenciaEspecial.fecha.in_(fechas)
    ).order_by(AsistenciaEspecial.fecha, AsistenciaEspecial.nombre).all()

    html = f"""
    <h1>REPORTE ASISTENCIAS ESPECIALES</h1>
    <p>Del <b>{fi.strftime("%d/%m/%Y")}</b> al <b>{ff.strftime("%d/%m/%Y")}</b></p>
    <p>Total registros: <b>{len(registros)}</b></p>

    <table border="1" cellpadding="5">
    <tr>
        <th>FECHA</th><th>CODIGO</th><th>NOMBRE</th><th>TIPO</th><th>SUPERVISOR</th>
    </tr>
    """

    for r in registros:
        html += f"""
        <tr>
            <td>{r.fecha}</td><td>{r.codigo}</td><td>{r.nombre}</td>
            <td>{r.tipo}</td><td>{r.supervisor}</td>
        </tr>
        """

    html += f"""
    </table>
    <br>
    <a href="/exportar_especiales_filtrado?fi={fecha_inicio}&ff={fecha_fin}">📥 EXPORTAR EXCEL</a>
    &nbsp;&nbsp;
    <a href="/reporte_asistencias_especiales">NUEVO FILTRO</a>
    &nbsp;&nbsp;
    <a href="/dashboard">DASHBOARD</a>
    """

    return html

@app.route("/exportar_especiales_filtrado")
def exportar_especiales_filtrado():

    from datetime import datetime, timedelta

    fecha_inicio = request.args.get("fi")
    fecha_fin = request.args.get("ff")

    fi = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    ff = datetime.strptime(fecha_fin, "%Y-%m-%d")

    fechas = []
    actual = fi
    while actual <= ff:
        fechas.append(actual.strftime("%d/%m/%Y"))
        actual += timedelta(days=1)

    registros = AsistenciaEspecial.query.filter(
        AsistenciaEspecial.fecha.in_(fechas)
    ).order_by(AsistenciaEspecial.fecha).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Asistencias Especiales"
    ws.append([f"REPORTE ASISTENCIAS ESPECIALES {fi.strftime('%d/%m/%Y')} AL {ff.strftime('%d/%m/%Y')}"])
    ws.append([])
    ws.append(["FECHA", "CODIGO", "NOMBRE", "TIPO", "SUPERVISOR"])
    for r in registros:
        ws.append([r.fecha, r.codigo, r.nombre, r.tipo, r.supervisor])

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"especiales_{fecha_inicio}_{fecha_fin}.xlsx"
    )

@app.route("/exportar_asistencia")
def exportar_asistencia():
    registros = Asistencia.query.order_by(Asistencia.id.desc()).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Asistencia"
    ws.append(["FECHA", "HORA", "CODIGO", "NOMBRE", "SUPERVISOR"])
    for r in registros:
        ws.append([r.fecha, r.hora, r.codigo, r.nombre, r.supervisor])
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="reporte_asistencia.xlsx")


@app.route("/exportar_horas_extras")
def exportar_horas_extras():
    registros = HorasExtras.query.order_by(HorasExtras.id.desc()).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Horas Extras"
    ws.append(["FECHA", "CODIGO", "NOMBRE", "HORAS", "SUPERVISOR"])
    for r in registros:
        ws.append([r.fecha, r.codigo, r.nombre, r.horas, r.supervisor])
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="reporte_horas_extras.xlsx")


@app.route("/exportar_movimientos")
def exportar_movimientos():
    registros = Movimiento.query.order_by(Movimiento.id.desc()).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Movimientos"
    ws.append(["FECHA", "CODIGO", "NOMBRE", "AREA ANTERIOR", "AREA NUEVA"])
    for m in registros:
        ws.append([m.fecha, m.codigo, m.nombre, m.area_anterior, m.area_nueva])
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="reporte_movimientos.xlsx")


@app.route("/exportar_asistencias_especiales")
def exportar_asistencias_especiales():
    registros = AsistenciaEspecial.query.order_by(AsistenciaEspecial.id.desc()).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Asistencias Especiales"
    ws.append(["FECHA", "CODIGO", "NOMBRE", "TIPO", "SUPERVISOR"])
    for r in registros:
        ws.append([r.fecha, r.codigo, r.nombre, r.tipo, r.supervisor])
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="reporte_asistencias_especiales.xlsx")

@app.route("/incidencias")
def incidencias():
    trabajadores = Trabajador.query.filter_by(
        estado="ACTIVO"
    ).order_by(Trabajador.nombre).all()

    html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Incidencias</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }
.card { background:#1e293b; border-radius:16px; padding:28px; width:100%; max-width:420px; border:1px solid rgba(255,255,255,0.08); }
h2 { color:#fff; font-size:16px; font-weight:700; margin-bottom:6px; }
.sub { color:#64748b; font-size:12px; margin-bottom:24px; }
label { display:block; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#64748b; margin-bottom:6px; }
select { width:100%; padding:12px 14px; background:#0f172a; border:1px solid rgba(255,255,255,0.1); border-radius:8px; color:#fff; font-size:14px; margin-bottom:16px; outline:none; }
input[type=date] { width:100%; padding:12px 14px; background:#0f172a; border:1px solid rgba(255,255,255,0.1); border-radius:8px; color:#fff; font-size:14px; margin-bottom:16px; outline:none; }
button { width:100%; padding:14px; background:linear-gradient(90deg,#1a56db,#3b82f6); border:none; border-radius:8px; color:#fff; font-size:15px; font-weight:700; cursor:pointer; }
a { display:block; text-align:center; color:#64748b; font-size:12px; margin-top:14px; text-decoration:none; }
</style>
</head>
<body>
<div class="card">
  <h2>🏥 INCIDENCIAS</h2>
  <p class="sub">Selecciona el trabajador y tipo de incidencia</p>

  <form action="/guardar_incidencia" method="post">

    <label>Trabajador</label>
    <select name="codigo">
"""

    for t in trabajadores:
        html += f'<option value="{t.codigo}">{t.nombre} — {t.area}</option>'

    html += """
    </select>

    <label>Tipo</label>
    <select name="tipo">
        <option value="V">V — VACACIONES</option>
        <option value="LSG">LSG — LICENCIA SIN GOCE</option>
        <option value="DM">DM — DESCANSO MEDICO</option>
    </select>

    <label>Fecha inicio</label>
    <input type="date" name="fecha_inicio">

    <label>Fecha fin</label>
    <input type="date" name="fecha_fin">

    <button type="submit">GUARDAR</button>

  </form>
  <a href="/dashboard">← Volver</a>
</div>
</body>
</html>
    """

    return html

@app.route("/buscar_trabajador_incidencia", methods=["POST"])
def buscar_trabajador_incidencia():

    codigo = request.form["codigo"]

    trabajador = Trabajador.query.filter_by(
        codigo=codigo
    ).first()

    if not trabajador:
        return """
        <h2>TRABAJADOR NO ENCONTRADO</h2>
        <a href="/incidencias">VOLVER</a>
        """

    return f"""
    <h1>INCIDENCIAS</h1>

    <p><b>{trabajador.nombre}</b></p>
    <p>Área: {trabajador.area}</p>

    <form action="/guardar_incidencia" method="post">

        <input type="hidden" name="codigo" value="{trabajador.codigo}">

        Tipo:<br>
        <select name="tipo">
            <option value="V">V — VACACIONES</option>
            <option value="LSG">LSG — LICENCIA SIN GOCE</option>
            <option value="DM">DM — DESCANSO MEDICO</option>
        </select>

        <br><br>

        Fecha inicio:<br>
        <input type="date" name="fecha_inicio">

        <br><br>

        Fecha fin:<br>
        <input type="date" name="fecha_fin">

        <br><br>

        <button type="submit">GUARDAR</button>

    </form>
    """

@app.route("/guardar_incidencia", methods=["POST"])
def guardar_incidencia():

    codigo = request.form["codigo"]

    trabajador = Trabajador.query.filter_by(
        codigo=codigo
    ).first()

    tipo = request.form["tipo"]

    descripciones = {
        "V": "VACACIONES",
        "LSG": "LICENCIA SIN GOCE",
        "DM": "DESCANSO MEDICO"
    }

    registro = Incidencia(
        codigo=trabajador.codigo,
        nombre=trabajador.nombre,
        tipo=tipo,
        descripcion=descripciones[tipo],
        fecha_inicio=request.form["fecha_inicio"],
        fecha_fin=request.form["fecha_fin"],
        activo=True
    )

    db.session.add(registro)
    db.session.commit()

    return f"""
    <h2>✅ INCIDENCIA REGISTRADA</h2>
    <p><b>{trabajador.nombre}</b></p>
    <p>Tipo: {tipo} — {descripciones[tipo]}</p>
    <p>Desde: {request.form['fecha_inicio']}</p>
    <p>Hasta: {request.form['fecha_fin']}</p>
    <br>
    <a href="/reporte_incidencias">VER REPORTE</a>
    &nbsp;&nbsp;
    <a href="/incidencias">NUEVA INCIDENCIA</a>
    &nbsp;&nbsp;
    <a href="/dashboard">DASHBOARD</a>
    """

@app.route("/reporte_incidencias")
def reporte_incidencias():
    return """
    <h1>REPORTE DE INCIDENCIAS</h1>

    <form action="/filtrar_incidencias" method="post">
        Fecha inicio:<br>
        <input type="date" name="fecha_inicio">
        <br><br>
        Fecha fin:<br>
        <input type="date" name="fecha_fin">
        <br><br>
        <button type="submit">VER REPORTE</button>
    </form>

    <br>
    <a href="/dashboard">Volver</a>
    """

@app.route("/filtrar_incidencias", methods=["POST"])
def filtrar_incidencias():

    fecha_inicio = request.form["fecha_inicio"]
    fecha_fin = request.form["fecha_fin"]

    from datetime import datetime
    fi = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    ff = datetime.strptime(fecha_fin, "%Y-%m-%d")

    registros = Incidencia.query.filter(
        Incidencia.fecha_inicio >= fecha_inicio,
        Incidencia.fecha_inicio <= fecha_fin
    ).order_by(Incidencia.fecha_inicio, Incidencia.nombre).all()

    html = f"""
    <h1>REPORTE DE INCIDENCIAS</h1>
    <p>Del <b>{fi.strftime("%d/%m/%Y")}</b> al <b>{ff.strftime("%d/%m/%Y")}</b></p>
    <p>Total registros: <b>{len(registros)}</b></p>

    <table border="1" cellpadding="5">
    <tr>
        <th>CODIGO</th><th>NOMBRE</th><th>TIPO</th><th>DESCRIPCION</th>
        <th>DESDE</th><th>HASTA</th><th>ESTADO</th><th>DESACTIVAR</th>
    </tr>
    """

    for r in registros:
        estado = "ACTIVA" if r.activo else "CERRADA"
        boton = f'<a href="/desactivar_incidencia/{r.id}">❌ QUITAR</a>' if r.activo else "—"
        html += f"""
        <tr>
            <td>{r.codigo}</td><td>{r.nombre}</td>
            <td><b>{r.tipo}</b></td><td>{r.descripcion}</td>
            <td>{r.fecha_inicio}</td><td>{r.fecha_fin}</td>
            <td>{estado}</td><td>{boton}</td>
        </tr>
        """

    html += f"""
    </table>
    <br>
    <a href="/exportar_incidencias_filtrado?fi={fecha_inicio}&ff={fecha_fin}">📥 EXPORTAR EXCEL</a>
    &nbsp;&nbsp;
    <a href="/reporte_incidencias">NUEVO FILTRO</a>
    &nbsp;&nbsp;
    <a href="/dashboard">DASHBOARD</a>
    """

    return html

@app.route("/exportar_incidencias_filtrado")
def exportar_incidencias_filtrado():

    fecha_inicio = request.args.get("fi")
    fecha_fin = request.args.get("ff")

    from datetime import datetime
    fi = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    ff = datetime.strptime(fecha_fin, "%Y-%m-%d")

    registros = Incidencia.query.filter(
        Incidencia.fecha_inicio >= fecha_inicio,
        Incidencia.fecha_inicio <= fecha_fin
    ).order_by(Incidencia.fecha_inicio).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Incidencias"
    ws.append([f"REPORTE INCIDENCIAS {fi.strftime('%d/%m/%Y')} AL {ff.strftime('%d/%m/%Y')}"])
    ws.append([])
    ws.append(["CODIGO", "NOMBRE", "TIPO", "DESCRIPCION", "DESDE", "HASTA", "ESTADO"])
    for r in registros:
        ws.append([r.codigo, r.nombre, r.tipo, r.descripcion, r.fecha_inicio, r.fecha_fin, "ACTIVA" if r.activo else "CERRADA"])

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"incidencias_{fecha_inicio}_{fecha_fin}.xlsx"
    )
@app.route("/desactivar_incidencia/<int:id>")
def desactivar_incidencia(id):

    incidencia = Incidencia.query.get(id)
    incidencia.activo = False
    db.session.commit()

    return """
    <h2>✅ INCIDENCIA DESACTIVADA</h2>
    <a href="/reporte_incidencias">VOLVER</a>
    """
@app.route("/reporte_mensual")
def reporte_mensual():
    return """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Reporte Mensual</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }
.card { background:#1e293b; border-radius:16px; padding:28px; width:100%; max-width:380px; border:1px solid rgba(255,255,255,0.08); }
h2 { color:#fff; font-size:16px; font-weight:700; margin-bottom:6px; }
.sub { color:#64748b; font-size:12px; margin-bottom:24px; }
label { display:block; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#64748b; margin-bottom:6px; }
select { width:100%; padding:12px 14px; background:#0f172a; border:1px solid rgba(255,255,255,0.1); border-radius:8px; color:#fff; font-size:15px; margin-bottom:16px; outline:none; }
.btns { display:flex; gap:10px; }
button { flex:1; padding:14px; border:none; border-radius:8px; color:#fff; font-size:14px; font-weight:700; cursor:pointer; }
.btn-ver { background:linear-gradient(90deg,#0d9488,#14b8a6); }
.btn-exp { background:linear-gradient(90deg,#1a56db,#3b82f6); }
a { display:block; text-align:center; color:#64748b; font-size:12px; margin-top:14px; text-decoration:none; }
</style>
</head>
<body>
<div class="card">
  <h2>📆 REPORTE MENSUAL</h2>
  <p class="sub">Selecciona el mes a reportar</p>

  <form action="/generar_reporte_mensual" method="post">
    <label>Mes</label>
    <select name="mes">
      <option value="1">Enero</option>
      <option value="2">Febrero</option>
      <option value="3">Marzo</option>
      <option value="4">Abril</option>
      <option value="5">Mayo</option>
      <option value="6" selected>Junio</option>
      <option value="7">Julio</option>
      <option value="8">Agosto</option>
      <option value="9">Septiembre</option>
      <option value="10">Octubre</option>
      <option value="11">Noviembre</option>
      <option value="12">Diciembre</option>
    </select>

    <label>Año</label>
    <select name="año">
      <option value="2025">2025</option>
      <option value="2026" selected>2026</option>
      <option value="2027">2027</option>
    </select>

    <div class="btns">
      <button type="submit" name="accion" value="ver" class="btn-ver">👁 VER</button>
      <button type="submit" name="accion" value="exportar" class="btn-exp">📥 EXCEL</button>
    </div>
  </form>
  <a href="/dashboard">← Volver</a>
</div>
</body>
</html>
    """

@app.route("/generar_reporte_mensual", methods=["POST"])
def generar_reporte_mensual():

    from datetime import datetime, timedelta
    import calendar

    accion = request.form.get("accion", "ver")

    if accion == "exportar":
        return exportar_reporte_mensual()

    mes = int(request.form.get("mes", datetime.now().month))
    año = int(request.form.get("año", datetime.now().year))
    dias_mes = calendar.monthrange(año, mes)[1]

    fi = datetime(año, mes, 1)
    ff = datetime(año, mes, dias_mes)

    fechas = []
    actual = fi
    while actual <= ff:
        fechas.append(actual.strftime("%d/%m/%Y"))
        actual += timedelta(days=1)

    fecha_inicio = fi.strftime("%Y-%m-%d")
    fecha_fin = ff.strftime("%Y-%m-%d")

    asistencias = Asistencia.query.filter(
        Asistencia.fecha.in_(fechas)
    ).order_by(Asistencia.fecha, Asistencia.nombre).all()

    horas = HorasExtras.query.filter(
        HorasExtras.fecha.in_(fechas)
    ).order_by(HorasExtras.fecha, HorasExtras.nombre).all()

    especiales = AsistenciaEspecial.query.filter(
        AsistenciaEspecial.fecha.in_(fechas)
    ).order_by(AsistenciaEspecial.fecha, AsistenciaEspecial.nombre).all()

    incidencias = Incidencia.query.filter(
        Incidencia.fecha_inicio >= fecha_inicio
    ).order_by(Incidencia.fecha_inicio, Incidencia.nombre).all()

    trabajadores_activos = Trabajador.query.filter_by(estado="ACTIVO").all()
    codigos_area = {t.codigo: t.area for t in trabajadores_activos}
    incidencias_activas = Incidencia.query.filter_by(activo=True).all()
    codigos_incidencia = {i.codigo: i for i in incidencias_activas}

    faltas = []
    for fecha in fechas:
        codigos_presentes = set(
            r.codigo for r in Asistencia.query.filter_by(fecha=fecha).all()
        )
        for t in trabajadores_activos:
            if t.codigo not in codigos_presentes:
                incidencia = codigos_incidencia.get(t.codigo)
                tipo = f"{incidencia.tipo} — {incidencia.descripcion}" if incidencia else "FALTA"
                faltas.append({
                    "fecha": fecha,
                    "codigo": t.codigo,
                    "nombre": t.nombre,
                    "area": t.area,
                    "tipo": tipo
                })

    total_horas = sum(r.horas for r in horas)

    html = f"""
    <h1>REPORTE MENSUAL</h1>
    <p>Del <b>{fi.strftime("%d/%m/%Y")}</b> al <b>{ff.strftime("%d/%m/%Y")}</b></p>
    <hr>

    <h2>RESUMEN</h2>
    <table border="1" cellpadding="5">
    <tr><td>Total asistencias</td><td><b>{len(asistencias)}</b></td></tr>
    <tr><td>Total horas extras</td><td><b>{total_horas}</b></td></tr>
    <tr><td>Total asistencias especiales</td><td><b>{len(especiales)}</b></td></tr>
    <tr><td>Total incidencias</td><td><b>{len(incidencias)}</b></td></tr>
    <tr><td>Total faltas/ausencias</td><td><b>{len(faltas)}</b></td></tr>
    </table><br>

    <h2>ASISTENCIAS ({len(asistencias)})</h2>
    <table border="1" cellpadding="5">
    <tr><th>FECHA</th><th>HORA</th><th>CODIGO</th><th>NOMBRE</th><th>AREA</th><th>SUPERVISOR</th></tr>
    """

    for r in asistencias:
        area = codigos_area.get(r.codigo, "")
        html += f"""
        <tr>
            <td>{r.fecha}</td><td>{r.hora}</td><td>{r.codigo}</td>
            <td>{r.nombre}</td><td>{area}</td><td>{r.supervisor}</td>
        </tr>
        """

    html += f"""
    </table><br>

    <h2>FALTAS Y AUSENCIAS ({len(faltas)})</h2>
    <table border="1" cellpadding="5">
    <tr><th>FECHA</th><th>CODIGO</th><th>NOMBRE</th><th>AREA</th><th>MOTIVO</th></tr>
    """

    for f in faltas:
        html += f"""
        <tr>
            <td>{f['fecha']}</td><td>{f['codigo']}</td><td>{f['nombre']}</td>
            <td>{f['area']}</td><td><b>{f['tipo']}</b></td>
        </tr>
        """

    html += f"""
    </table><br>

    <h2>HORAS EXTRAS ({len(horas)})</h2>
    <table border="1" cellpadding="5">
    <tr><th>FECHA</th><th>CODIGO</th><th>NOMBRE</th><th>HORAS</th><th>SUPERVISOR</th></tr>
    """

    for r in horas:
        html += f"""
        <tr>
            <td>{r.fecha}</td><td>{r.codigo}</td><td>{r.nombre}</td>
            <td>{r.horas}</td><td>{r.supervisor}</td>
        </tr>
        """

    html += f"""
    </table><br>

    <h2>ASISTENCIAS ESPECIALES ({len(especiales)})</h2>
    <table border="1" cellpadding="5">
    <tr><th>FECHA</th><th>CODIGO</th><th>NOMBRE</th><th>TIPO</th><th>SUPERVISOR</th></tr>
    """

    for r in especiales:
        html += f"""
        <tr>
            <td>{r.fecha}</td><td>{r.codigo}</td><td>{r.nombre}</td>
            <td>{r.tipo}</td><td>{r.supervisor}</td>
        </tr>
        """

    html += f"""
    </table><br>

    <h2>INCIDENCIAS ({len(incidencias)})</h2>
    <table border="1" cellpadding="5">
    <tr><th>CODIGO</th><th>NOMBRE</th><th>TIPO</th><th>DESDE</th><th>HASTA</th><th>ESTADO</th></tr>
    """

    for r in incidencias:
        estado = "ACTIVA" if r.activo else "CERRADA"
        html += f"""
        <tr>
            <td>{r.codigo}</td><td>{r.nombre}</td>
            <td>{r.tipo} — {r.descripcion}</td>
            <td>{r.fecha_inicio}</td><td>{r.fecha_fin}</td><td>{estado}</td>
        </tr>
        """

    html += """
    </table><br>
    <a href="/reporte_mensual">NUEVO REPORTE</a>
    &nbsp;&nbsp;
    <a href="/dashboard">DASHBOARD</a>
    """

    return html
@app.route("/exportar_reporte_mensual", methods=["POST"])
def exportar_reporte_mensual():

    from datetime import datetime, timedelta

    fecha_inicio = request.form["fecha_inicio"]
    fecha_fin = request.form["fecha_fin"]

    fi = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    ff = datetime.strptime(fecha_fin, "%Y-%m-%d")

    fechas = []
    actual = fi
    while actual <= ff:
        fechas.append(actual.strftime("%d/%m/%Y"))
        actual += timedelta(days=1)

    asistencias = Asistencia.query.filter(Asistencia.fecha.in_(fechas)).order_by(Asistencia.fecha).all()
    horas = HorasExtras.query.filter(HorasExtras.fecha.in_(fechas)).order_by(HorasExtras.fecha).all()
    especiales = AsistenciaEspecial.query.filter(AsistenciaEspecial.fecha.in_(fechas)).order_by(AsistenciaEspecial.fecha).all()
    incidencias = Incidencia.query.filter(Incidencia.fecha_inicio >= fecha_inicio).order_by(Incidencia.fecha_inicio).all()

    # Calcular faltas
    trabajadores_activos = Trabajador.query.filter_by(estado="ACTIVO").all()
    codigos_area = {t.codigo: t.area for t in trabajadores_activos}
    incidencias_activas = Incidencia.query.filter_by(activo=True).all()
    codigos_incidencia = {i.codigo: i for i in incidencias_activas}

    faltas = []
    for fecha in fechas:
        codigos_presentes = set(
            r.codigo for r in Asistencia.query.filter_by(fecha=fecha).all()
        )
        for t in trabajadores_activos:
            if t.codigo not in codigos_presentes:
                incidencia = codigos_incidencia.get(t.codigo)
                tipo = f"{incidencia.tipo} - {incidencia.descripcion}" if incidencia else "FALTA"
                faltas.append([fecha, t.codigo, t.nombre, t.area, tipo])

    wb = openpyxl.Workbook()

    # Hoja Asistencias
    ws1 = wb.active
    ws1.title = "Asistencias"
    ws1.append([f"REPORTE MENSUAL {fecha_inicio} AL {fecha_fin}"])
    ws1.append([])
    ws1.append(["FECHA", "HORA", "CODIGO", "NOMBRE", "AREA", "SUPERVISOR"])
    for r in asistencias:
        area = codigos_area.get(r.codigo, "")
        ws1.append([r.fecha, r.hora, r.codigo, r.nombre, area, r.supervisor])

    # Hoja Faltas
    ws2 = wb.create_sheet("Faltas y Ausencias")
    ws2.append([f"REPORTE MENSUAL {fecha_inicio} AL {fecha_fin}"])
    ws2.append([])
    ws2.append(["FECHA", "CODIGO", "NOMBRE", "AREA", "MOTIVO"])
    for f in faltas:
        ws2.append(f)

    # Hoja Horas Extras
    ws3 = wb.create_sheet("Horas Extras")
    ws3.append([f"REPORTE MENSUAL {fecha_inicio} AL {fecha_fin}"])
    ws3.append([])
    ws3.append(["FECHA", "CODIGO", "NOMBRE", "HORAS", "SUPERVISOR"])
    for r in horas:
        ws3.append([r.fecha, r.codigo, r.nombre, r.horas, r.supervisor])

    # Hoja Especiales
    ws4 = wb.create_sheet("Especiales")
    ws4.append([f"REPORTE MENSUAL {fecha_inicio} AL {fecha_fin}"])
    ws4.append([])
    ws4.append(["FECHA", "CODIGO", "NOMBRE", "TIPO", "SUPERVISOR"])
    for r in especiales:
        ws4.append([r.fecha, r.codigo, r.nombre, r.tipo, r.supervisor])

    # Hoja Incidencias
    ws5 = wb.create_sheet("Incidencias")
    ws5.append([f"REPORTE MENSUAL {fecha_inicio} AL {fecha_fin}"])
    ws5.append([])
    ws5.append(["CODIGO", "NOMBRE", "TIPO", "DESCRIPCION", "DESDE", "HASTA", "ESTADO"])
    for r in incidencias:
        ws5.append([r.codigo, r.nombre, r.tipo, r.descripcion, r.fecha_inicio, r.fecha_fin, "ACTIVA" if r.activo else "CERRADA"])

    # Ajustar anchos
    for ws in [ws1, ws2, ws3, ws4, ws5]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 4

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"reporte_{fecha_inicio}_{fecha_fin}.xlsx"
    )
@app.route("/reporte_diario")
def reporte_diario():
    from datetime import date
    hoy = date.today().strftime("%d/%m/%Y")

    orden_area = {"RECEPCION": 1, "REPOSICION": 2, "PICKING": 3, "PACKING": 4}
    trabajadores_activos = sorted(
        Trabajador.query.filter_by(estado="ACTIVO").all(),
        key=lambda t: (orden_area.get(t.area, 99), t.nombre)
    )

    presentes = Asistencia.query.filter_by(fecha=hoy).all()
    codigos_presentes = {r.codigo: r for r in presentes}

    incidencias_activas = Incidencia.query.filter_by(activo=True).all()
    codigos_incidencia = {i.codigo: i for i in incidencias_activas}

    total = len(trabajadores_activos)
    total_presentes = len(presentes)
    total_ausentes = total - total_presentes

    filas = ""
    for t in trabajadores_activos:
        if t.codigo in codigos_presentes:
            registro = codigos_presentes[t.codigo]
            estado = "P"
            hora = registro.hora
            escaneado = registro.escaneado_por or "—"
            color = "#16a34a"
            bg = "#f0fdf4"
        else:
            incidencia = codigos_incidencia.get(t.codigo)
            if incidencia:
                estado = incidencia.tipo
            else:
                estado = "F"
            escaneado = "—"
            color = "#dc2626" if estado == "F" else "#d97706"
            bg = "#fef2f2" if estado == "F" else "#fffbeb"

        filas += f"""
        <tr style="background:{bg};">
            <td style="padding:8px 10px;border:1px solid #e2e8f0;font-size:13px;font-weight:600;">{t.nombre}</td>
            <td style="padding:8px 10px;border:1px solid #e2e8f0;font-size:12px;text-align:center;">{t.area}</td>
            <td style="padding:8px 10px;border:1px solid #e2e8f0;font-size:12px;text-align:center;">{t.supervisor}</td>
            <td style="padding:8px 10px;border:1px solid #e2e8f0;font-size:12px;text-align:center;color:{color};font-weight:800;">{estado}</td>
            <td style="padding:8px 10px;border:1px solid #e2e8f0;font-size:12px;text-align:center;color:#64748b;">{hora if estado == "P" else "—"}</td>
            <td style="padding:8px 10px;border:1px solid #e2e8f0;font-size:12px;text-align:center;color:#64748b;">{escaneado}</td>
        </tr>
        """

    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Reporte Diario</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',sans-serif; background:#f1f5f9; padding:16px; }}
.header {{ background:linear-gradient(90deg,#1e3a5f,#1a56db); border-radius:12px; padding:16px 20px; margin-bottom:16px; color:#fff; display:flex; justify-content:space-between; align-items:center; }}
.header h1 {{ font-size:16px; font-weight:700; letter-spacing:1px; }}
.header .fecha {{ font-size:13px; background:rgba(255,255,255,0.2); padding:4px 12px; border-radius:20px; }}
.kpis {{ display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-bottom:16px; }}
.kpi {{ background:#fff; border-radius:10px; padding:12px 16px; text-align:center; border-top:3px solid #3b82f6; }}
.kpi.green {{ border-top-color:#16a34a; }}
.kpi.red {{ border-top-color:#dc2626; }}
.kpi-value {{ font-size:28px; font-weight:800; color:#1e293b; }}
.kpi.green .kpi-value {{ color:#16a34a; }}
.kpi.red .kpi-value {{ color:#dc2626; }}
.kpi-label {{ font-size:10px; text-transform:uppercase; letter-spacing:1px; color:#94a3b8; margin-top:4px; }}
.leyenda {{ background:#fff; border-radius:10px; padding:10px 16px; margin-bottom:16px; font-size:11px; color:#64748b; }}
table {{ width:100%; border-collapse:collapse; background:#fff; border-radius:10px; overflow:hidden; }}
thead {{ background:#1e293b; }}
thead th {{ padding:10px; color:#fff; font-size:11px; text-transform:uppercase; letter-spacing:1px; text-align:center; }}
thead th:nth-child(2) {{ text-align:left; }}
.btns {{ display:flex; gap:10px; margin-top:16px; }}
.btn {{ padding:12px 20px; border-radius:8px; font-size:13px; font-weight:700; text-decoration:none; color:#fff; }}
.btn-excel {{ background:linear-gradient(90deg,#16a34a,#22c55e); }}
.btn-dash {{ background:linear-gradient(90deg,#1a56db,#3b82f6); }}
</style>
</head>
<body>

<div class="header">
  <h1>📅 REPORTE DIARIO</h1>
  <div class="fecha">{hoy}</div>
</div>

<div class="kpis">
  <div class="kpi blue">
    <div class="kpi-value">{total}</div>
    <div class="kpi-label">Total Activos</div>
  </div>
  <div class="kpi green">
    <div class="kpi-value">{total_presentes}</div>
    <div class="kpi-label">Presentes</div>
  </div>
  <div class="kpi red">
    <div class="kpi-value">{total_ausentes}</div>
    <div class="kpi-label">Ausentes</div>
  </div>
</div>

<div class="leyenda">
  P = Presente &nbsp;|&nbsp; F = Falta &nbsp;|&nbsp; DM = Descanso Médico &nbsp;|&nbsp; V = Vacaciones &nbsp;|&nbsp; LSG = Licencia Sin Goce
</div>

<table>
  <thead>
    <tr>
      <th style="text-align:left;">Nombre</th>
      <th>Área</th>
      <th>Supervisor</th>
      <th>Asist.</th>
      <th>Hora Ingreso</th>
      <th>Escaneado por</th>
    </tr>
  </thead>
  <tbody>
    {filas}
  </tbody>
</table>

<div class="btns">
  <a href="/exportar_reporte_diario" class="btn btn-excel">📥 EXPORTAR EXCEL</a>
  <a href="/dashboard" class="btn btn-dash">⬅ DASHBOARD</a>
</div>

</body>
</html>
"""

@app.route("/exportar_reporte_diario")
def exportar_reporte_diario():
    from datetime import date
    hoy = date.today().strftime("%d/%m/%Y")

    orden_area = {"RECEPCION": 1, "REPOSICION": 2, "PICKING": 3, "PACKING": 4}
    trabajadores_activos = sorted(
        Trabajador.query.filter_by(estado="ACTIVO").all(),
        key=lambda t: (orden_area.get(t.area, 99), t.nombre)
    )

    presentes = Asistencia.query.filter_by(fecha=hoy).all()
    codigos_presentes = {r.codigo: r for r in presentes}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte Diario"
    ws.append([f"REPORTE DIARIO — {hoy}"])
    ws.append([])
    incidencias_activas = Incidencia.query.filter_by(activo=True).all()
    codigos_incidencia = {i.codigo: i for i in incidencias_activas}

    ws.append(["HORA", "NOMBRE", "AREA", "ASIST."])

    for t in trabajadores_activos:
        if t.codigo in codigos_presentes:
            registro = codigos_presentes[t.codigo]
            ws.append([registro.hora, t.nombre, t.area, "P"])
        else:
            incidencia = codigos_incidencia.get(t.codigo)
            estado = incidencia.tipo if incidencia else "F"
            ws.append(["", t.nombre, t.area, estado])

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"reporte_diario_{hoy.replace('/','_')}.xlsx"
    )
@app.route("/exportar_mensual_formato")
def exportar_mensual_formato():
    return """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Control Mensual</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI',sans-serif; background:#0f172a; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }
.card { background:#1e293b; border-radius:16px; padding:28px; width:100%; max-width:380px; border:1px solid rgba(255,255,255,0.08); }
h2 { color:#fff; font-size:16px; font-weight:700; margin-bottom:6px; }
.sub { color:#64748b; font-size:12px; margin-bottom:24px; }
label { display:block; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#64748b; margin-bottom:6px; }
input { width:100%; padding:12px 14px; background:#0f172a; border:1px solid rgba(255,255,255,0.1); border-radius:8px; color:#fff; font-size:15px; margin-bottom:16px; outline:none; }
button { width:100%; padding:14px; background:linear-gradient(90deg,#1a56db,#3b82f6); border:none; border-radius:8px; color:#fff; font-size:15px; font-weight:700; cursor:pointer; }
a { display:block; text-align:center; color:#64748b; font-size:12px; margin-top:14px; text-decoration:none; }
</style>
</head>
<body>
<div class="card">
  <h2>📊 CONTROL MENSUAL EXCEL</h2>
  <p class="sub">Selecciona el rango de fechas</p>

  <form action="/generar_mensual_formato" method="post">
    <label>Fecha inicio</label>
    <input type="date" name="fecha_inicio">

    <label>Fecha fin</label>
    <input type="date" name="fecha_fin">

    <button type="submit">📥 DESCARGAR EXCEL</button>
  </form>
  <a href="/dashboard">← Volver</a>
</div>
</body>
</html>
    """

@app.route("/generar_mensual_formato", methods=["POST"])
def generar_mensual_formato():
    from datetime import date
    import calendar

    from datetime import datetime, timedelta

    fecha_inicio = request.form["fecha_inicio"]
    fecha_fin = request.form["fecha_fin"]

    fi = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    ff = datetime.strptime(fecha_fin, "%Y-%m-%d")

    mes_nombre = fi.strftime("%b").upper()
    año = fi.year

    fechas = []
    fechas_iso = []
    actual = fi
    while actual <= ff:
        fechas.append(actual.strftime("%d/%m/%Y"))
        fechas_iso.append(actual.strftime("%Y-%m-%d"))
        actual += timedelta(days=1)

    # Trabajadores ordenados por condición laboral luego nombre
    orden_condicion = {"FIJO": 1, "DOTACION": 2, "CAMPAÑA": 3}
    trabajadores = Trabajador.query.filter_by(estado="ACTIVO").order_by(Trabajador.condicion, Trabajador.nombre).all()

    # Registros del mes
    asistencias = set()
    for r in Asistencia.query.filter(Asistencia.fecha.in_(fechas)).all():
        asistencias.add((r.codigo, r.fecha))

    incidencias_list = Incidencia.query.all()

    # Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    dias_mes = len(fechas)
    ws.title = f"Control {mes_nombre} {año}"

    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # Colores
    COLOR_P   = PatternFill("solid", fgColor="FFFFFF")  # Blanco
    COLOR_F   = PatternFill("solid", fgColor="FF0000")  # Rojo
    COLOR_DM  = PatternFill("solid", fgColor="FFFF00")  # Amarillo
    COLOR_C   = PatternFill("solid", fgColor="7B3F00")  # Marrón
    COLOR_LSG = PatternFill("solid", fgColor="7030A0")  # Morado
    COLOR_V   = PatternFill("solid", fgColor="00B050")  # Verde
    COLOR_HDR  = PatternFill("solid", fgColor="1F3864")
    COLOR_HDR2 = PatternFill("solid", fgColor="2F5496")
    borde = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    centro = Alignment(horizontal="center", vertical="center", wrap_text=False)
    izq = Alignment(horizontal="left", vertical="center")

    # FILA 1 - Título
    total_cols = 4 + dias_mes
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    ws["A1"] = f"CONTROL DE ASISTENCIA — {mes_nombre} {año}"
    ws["A1"].fill = COLOR_HDR
    ws["A1"].font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    ws["A1"].alignment = centro

    # FILA 2 - Leyenda
    ws.merge_cells(f"A2:{get_column_letter(total_cols)}2")
    ws["A2"] = "P=Presente  |  F=Falta  |  DM=Descanso Médico  |  V=Vacaciones  |  LSG=Licencia Sin Goce  |  C=Cesado"
    ws["A2"].font = Font(name="Calibri", size=8, italic=True)
    ws["A2"].alignment = centro

    # FILA 3 - Encabezados fijos
    for col, h in enumerate(["Nro", "Apellidos y Nombres", "Condición", "Área"], 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.fill = COLOR_HDR2
        cell.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        cell.alignment = centro if col != 2 else izq
        cell.border = borde

    # FILA 3 - Encabezados de fechas formato "8-Jun"
    for d in range(1, dias_mes + 1):
        col = 4 + d
        fecha_obj = fi + timedelta(days=d-1)
        label = f"{fecha_obj.day}-{fecha_obj.strftime('%b').upper()}"
        cell = ws.cell(row=3, column=col, value=label)
        cell.fill = COLOR_HDR2
        cell.font = Font(name="Calibri", size=8, bold=True, color="FFFFFF")
        cell.alignment = centro
        cell.border = borde

    # Anchos
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 11
    ws.column_dimensions["D"].width = 12
    for d in range(1, dias_mes + 1):
        ws.column_dimensions[get_column_letter(4 + d)].width = 6

    # Alturas
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 13
    ws.row_dimensions[3].height = 16

    # DATOS
    for i, t in enumerate(trabajadores, 1):
        row = 3 + i
        ws.row_dimensions[row].height = 14

        ws.cell(row=row, column=1, value=i).alignment = centro
        ws.cell(row=row, column=1).border = borde
        ws.cell(row=row, column=1).font = Font(name="Calibri", size=9)

        ws.cell(row=row, column=2, value=t.nombre).alignment = izq
        ws.cell(row=row, column=2).border = borde
        ws.cell(row=row, column=2).font = Font(name="Calibri", size=9)

        ws.cell(row=row, column=3, value=t.condicion).alignment = centro
        ws.cell(row=row, column=3).border = borde
        ws.cell(row=row, column=3).font = Font(name="Calibri", size=9)

        ws.cell(row=row, column=4, value=t.area).alignment = centro
        ws.cell(row=row, column=4).border = borde
        ws.cell(row=row, column=4).font = Font(name="Calibri", size=9)

        # Incidencia activa del trabajador
        inc = None
        for incd in incidencias_list:
            if incd.codigo == t.codigo and incd.activo:
                inc = incd
                break

        for d, (fecha, fecha_iso) in enumerate(zip(fechas, fechas_iso), 1):
            col = 4 + d
            cell = ws.cell(row=row, column=col)
            cell.alignment = centro
            cell.border = borde
            cell.font = Font(name="Calibri", size=9, color="000000")

            # Estado
            if t.estado == "CESADO":
                estado = "C"
            elif inc and inc.tipo == "V" and inc.fecha_inicio <= fecha_iso <= inc.fecha_fin:
                estado = "V"
            elif inc and inc.tipo == "LSG" and inc.fecha_inicio <= fecha_iso <= inc.fecha_fin:
                estado = "LSG"
            elif inc and inc.tipo == "DM" and inc.fecha_inicio <= fecha_iso <= inc.fecha_fin:
                estado = "DM"
            elif (t.codigo, fecha) in asistencias:
                estado = "P"
            else:
                estado = "F"

            cell.value = estado

            if estado == "P":
                cell.fill = COLOR_P
            elif estado == "F":
                cell.fill = COLOR_F
            elif estado == "DM":
                cell.fill = COLOR_DM
            elif estado == "C":
                cell.fill = COLOR_C
            elif estado == "LSG":
                cell.fill = COLOR_LSG
                cell.font = Font(name="Calibri", size=8, color="000000")
            elif estado == "V":
                cell.fill = COLOR_V
                cell.font = Font(name="Calibri", size=9, color="FFFFFF")

    # Congelar paneles en E4
    ws.freeze_panes = "E4"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"control_asistencia_{mes_nombre}_{año}.xlsx"
    )

with app.app_context():
    db.create_all()
    try:
        with db.engine.connect() as conn:
            conn.execute(db.text("ALTER TABLE asistencia ADD COLUMN escaneado_por VARCHAR(50)"))
            conn.commit()
    except:
        pass

@app.route("/cargar_trabajadores_masivo")
def cargar_trabajadores_masivo():
    trabajadores = [
        ("ALM-0001", "MANUEL IGNACIO VALENCIA MANRIQUE", "FIJO", "PACKING", "JOSE"),
        ("ALM-0002", "ANDRÉS SANTOS JULCA ZURITA", "FIJO", "PACKING", "JOSE"),
        ("ALM-0003", "JOSUE RUBÉN CHIRINOS DE LA CRUZ", "FIJO", "PACKING", "JOSE"),
        ("ALM-0004", "JOSÉ ALBERTO BALTAZAR NAPÁN", "FIJO", "PACKING", "JOSE"),
        ("ALM-0005", "DAVID ISAAC CORNEJO OROPEZA", "CAMPAÑA", "PACKING", "JOSE"),
        ("ALM-0006", "RAÚL SNAIDES SANGANÍA RAMÍREZ", "CAMPAÑA", "PACKING", "JOSE"),
        ("ALM-0007", "ÁNGEL DE JESÚS NAVARRETE JAVIER", "CAMPAÑA", "PACKING", "JOSE"),
        ("ALM-0008", "EDWIN MARCOS VENTURA CAYTUERO", "FIJO", "PICKING", "JEAN"),
        ("ALM-0009", "FLAZ COTAQUISPE QUISPE", "FIJO", "PICKING", "JEAN"),
        ("ALM-0010", "JOAQUÍN SEBASTIÁN MAZA JIMÉNEZ", "FIJO", "PICKING", "JEAN"),
        ("ALM-0011", "JOHAN SINCAS GUTIÉRREZ", "FIJO", "PICKING", "JEAN"),
        ("ALM-0012", "JHON MARK CAJUSOL INOÑAN", "CAMPAÑA", "PICKING", "JEAN"),
        ("ALM-0013", "JOSÉ ÁNGEL IPANAQUÉ SILVA", "CAMPAÑA", "PICKING", "JEAN"),
        ("ALM-0014", "JEAN CARLO VEGA CALLAN", "CAMPAÑA", "PICKING", "JEAN"),
        ("ALM-0015", "FRANCISCO CRUZ", "FIJO", "RECEPCION", "FRANCISCO"),
        ("ALM-0016", "VICTOR GUSTAVO TORRES YATACO", "FIJO", "RECEPCION", "FRANCISCO"),
        ("ALM-0017", "GABRIEL IVAN ASTO SALAZAR", "FIJO", "RECEPCION", "FRANCISCO"),
        ("ALM-0018", "PIERO ALESSANDRO ASENCIO LOZANO", "FIJO", "RECEPCION", "FRANCISCO"),
        ("ALM-0019", "ROBERTO CARLOS SALAZAR HUARACA", "FIJO", "RECEPCION", "FRANCISCO"),
        ("ALM-0020", "CHRISTIAN ORDINOLA QUISPE", "FIJO", "RECEPCION", "FRANCISCO"),
        ("ALM-0021", "RONALD RODRIGUEZ VENTURA", "FIJO", "RECEPCION", "FRANCISCO"),
        ("ALM-0022", "EUDES CERRON QUISPEAYALA", "CAMPAÑA", "RECEPCION", "FRANCISCO"),
        ("ALM-0023", "DENIS MANUEL MANCO MANCO", "CAMPAÑA", "RECEPCION", "FRANCISCO"),
        ("ALM-0024", "WILLIANS SÁNCHEZ LÓPEZ", "CAMPAÑA", "RECEPCION", "FRANCISCO"),
        ("ALM-0025", "MARIO DANIEL GONZALO ZÁRATE HIDALGO", "CAMPAÑA", "RECEPCION", "FRANCISCO"),
        ("ALM-0026", "CARLOS ISMAEL LÓPEZ VERA", "CAMPAÑA", "RECEPCION", "FRANCISCO"),
        ("ALM-0027", "LUIS ALBERTO SANDOVAL DURAND", "CAMPAÑA", "RECEPCION", "FRANCISCO"),
        ("ALM-0028", "LUIS REYES PRADO", "CAMPAÑA", "RECEPCION", "FRANCISCO"),
        ("ALM-0029", "THELMA GOMEZ CHUQUIHUAMANI", "CAMPAÑA", "RECEPCION", "FRANCISCO"),
        ("ALM-0030", "ROCIO MADRID ESCOBAR", "CAMPAÑA", "RECEPCION", "FRANCISCO"),
        ("ALM-0031", "JAROLD GUEVARA LLUSEMA", "CAMPAÑA", "REPOSICION", "JAROLD"),
        ("ALM-0032", "JORDAN LANFRANCO LUCAS", "FIJO", "REPOSICION", "JAROLD"),
        ("ALM-0033", "JEREMI RENATO RAMOS VILLARREAL", "FIJO", "REPOSICION", "JAROLD"),
        ("ALM-0034", "JEFFERSON JESUS CRUZ CABANILLAS", "FIJO", "REPOSICION", "JAROLD"),
        ("ALM-0035", "ANDRÉS FELIPE CÓRDOVA SECLÉN", "CAMPAÑA", "REPOSICION", "JAROLD"),
        ("ALM-0036", "FERNANDO MANUEL MIRANDA PÉREZ", "CAMPAÑA", "REPOSICION", "JAROLD"),
    ]
    total = 0
    for codigo, nombre, condicion, area, supervisor in trabajadores:
        existe = Trabajador.query.filter_by(codigo=codigo).first()
        if not existe:
            t = Trabajador(codigo=codigo, nombre=nombre, condicion=condicion,
                          area=area, supervisor=supervisor, estado="ACTIVO")
            db.session.add(t)
            total += 1
    db.session.commit()
    return f"<h2>✅ {total} trabajadores cargados correctamente</h2><a href='/trabajadores'>VER LISTA</a>"

@app.route("/exportar_trabajadores")
def exportar_trabajadores():
    trabajadores = Trabajador.query.order_by(Trabajador.codigo).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trabajadores"
    ws.append(["CODIGO", "NOMBRE", "CONDICION", "AREA", "SUPERVISOR", "ESTADO"])
    for t in trabajadores:
        ws.append([t.codigo, t.nombre, t.condicion, t.area, t.supervisor, t.estado])
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="trabajadores.xlsx")
@app.route("/editar/<int:id>", methods=["GET", "POST"])
def editar(id):
    t = Trabajador.query.get(id)
    if request.method == "POST":
        t.nombre = request.form["nombre"]
        t.condicion = request.form["condicion"]
        t.area = request.form["area"]
        t.supervisor = request.form["supervisor"]
        db.session.commit()
        return "<h2>✅ Actualizado</h2><a href='/trabajadores'>VOLVER A LISTA</a> &nbsp;&nbsp; <a href='/dashboard'>DASHBOARD</a>"
    return f"""
<h2>EDITAR TRABAJADOR</h2>
<form action="/editar/{t.id}" method="post">
    Nombre:<br><input type="text" name="nombre" value="{t.nombre}"><br><br>
    Condición:<br>
    <select name="condicion">
        <option {'selected' if t.condicion=='FIJO' else ''}>FIJO</option>
        <option {'selected' if t.condicion=='DOTACION' else ''}>DOTACION</option>
        <option {'selected' if t.condicion=='CAMPAÑA' else ''}>CAMPAÑA</option>
    </select><br><br>
    Área:<br>
    <select name="area">
        <option {'selected' if t.area=='RECEPCION' else ''}>RECEPCION</option>
        <option {'selected' if t.area=='REPOSICION' else ''}>REPOSICION</option>
        <option {'selected' if t.area=='PICKING' else ''}>PICKING</option>
        <option {'selected' if t.area=='PACKING' else ''}>PACKING</option>
    </select><br><br>
    Supervisor:<br><input type="text" name="supervisor" value="{t.supervisor}"><br><br>
    <button type="submit">GUARDAR</button>
</form>
<a href="/trabajadores">VOLVER</a>
"""

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )