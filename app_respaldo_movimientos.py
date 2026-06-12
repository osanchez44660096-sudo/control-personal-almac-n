from flask import Flask, request
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import qrcode
import os

app = Flask(__name__)

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///almacen.db'
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


# =========================
# LOGIN
# =========================

@app.route("/")
def login():

    return """
    <h1>CONTROL DE PERSONAL DE ALMACEN</h1>

    <form action="/dashboard" method="post">

        Usuario:<br>
        <input type="text" name="usuario">

        <br><br>

        Contraseña:<br>
        <input type="password" name="password">

        <br><br>

        <button type="submit">
        Ingresar
        </button>

    </form>
    """


# =========================
# DASHBOARD
# =========================

@app.route("/dashboard", methods=["POST"])
def dashboard():

    total = Trabajador.query.count()

    return f"""
    <h1>DASHBOARD</h1>

    <p>Trabajadores registrados: {total}</p>

    <hr>

    <ul>
        <li><a href="/asistencia">ASISTENCIA QR</a></li>
        <li>HORAS EXTRAS</li>
        <li>SÁBADO</li>
        <li>DOMINGO</li>
        <li><a href="/trabajadores">TRABAJADORES</a></li>
        <li>REPORTES</li>
        <li>MOVIMIENTOS PERSONAL</li>
    </ul>
    """


# =========================
# ASISTENCIA
# =========================

@app.route("/asistencia")
def asistencia():

    return """
    <h1>ASISTENCIA QR</h1>

    <form action="/registrar_asistencia" method="post">

        Código trabajador:<br>

        <input type="text"
               name="codigo"
               autofocus>

        <br><br>

        <button type="submit">
        REGISTRAR
        </button>

    </form>

    <br>

    <a href="/">
    Volver
    </a>
    """


@app.route("/registrar_asistencia", methods=["POST"])
def registrar_asistencia():

    codigo = request.form["codigo"]

    trabajador = Trabajador.query.filter_by(
        codigo=codigo
    ).first()

    if trabajador:

        ahora = datetime.now()

        registro = Asistencia(
            codigo=trabajador.codigo,
            nombre=trabajador.nombre,
            fecha=ahora.strftime("%d/%m/%Y"),
            hora=ahora.strftime("%H:%M:%S"),
            supervisor="SUPERVISOR",
            tipo="ASISTENCIA"
        )

        db.session.add(registro)
        db.session.commit()

        return f"""
<h2>TRABAJADOR REGISTRADO</h2>

<p>Código generado:</p>

<h1>{codigo}</h1>

<p>{nombre}</p>

<p>QR generado correctamente</p>

<a href='/static/qr/{codigo}.png' target='_blank'>
VER QR
</a>

<br><br>

<a href='/trabajadores'>
VOLVER
</a>
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
    ➕ NUEVO TRABAJADOR
    </a>

    <br><br>

    <table border="1" cellpadding="5">

    <tr>
        <th>CODIGO</th>
        <th>NOMBRE</th>
        <th>AREA</th>
        <th>CONDICION</th>
        <th>ESTADO</th>
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

    trabajador.area = request.form["area"]

    db.session.commit()

    return """
    <h2>ÁREA ACTUALIZADA CORRECTAMENTE</h2>

    <a href="/trabajadores">
    VOLVER
    </a>
    """


@app.route("/cesar/<int:id>")
def cesar(id):

    trabajador = Trabajador.query.get(id)

    trabajador.estado = "CESADO"

    db.session.commit()

    return """
    <h2>TRABAJADOR CESADO CORRECTAMENTE</h2>

    <a href="/trabajadores">
    VOLVER
    </a>
    """
if __name__ == "__main__":

    with app.app_context():
        db.create_all()

    app.run(debug=True)