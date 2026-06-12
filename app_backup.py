from flask import Flask, request
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///almacen.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

class Trabajador(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(20))
    nombre = db.Column(db.String(200))
    condicion = db.Column(db.String(50))
    area = db.Column(db.String(50))
    supervisor = db.Column(db.String(50))
    estado = db.Column(db.String(20))

@app.route("/")
def login():
    return """
    <h1>CONTROL DE PERSONAL DE ALMACEN</h1>

    <form action="/dashboard" method="post">
        Usuario:<br>
        <input type="text" name="usuario"><br><br>

        Contraseña:<br>
        <input type="password" name="password"><br><br>

        <button type="submit">Ingresar</button>
    </form>
    """

@app.route("/dashboard", methods=["POST"])
def dashboard():
    total = Trabajador.query.count()

    return f"""
    <h1>Dashboard</h1>

    <p>Trabajadores registrados: {total}</p>

    <ul>
        <li>Asistencia QR</li>
        <li>Horas Extras</li>
        <li>Sábados</li>
        <li>Domingos</li>
        <li>Trabajadores</li>
        <li>Reportes</li>
        <li>Movimientos de Personal</li>
    </ul>
    """

if __name__ == "__main__":

    with app.app_context():
        db.create_all()

        if Trabajador.query.count() == 0:

            import pandas as pd

            df = pd.read_excel("trabajadores.xlsx.xlsx")

            for _, fila in df.iterrows():

                trabajador = Trabajador(
                    codigo=str(fila["CODIGO"]),
                    nombre=str(fila["Nombre Completo"]),
                    condicion=str(fila["Condición Laboral"]),
                    area=str(fila["AREA"]),
                    supervisor=str(fila["SUPERVISOR"]),
                    estado=str(fila["ESTADO"])
                )

                db.session.add(trabajador)

            db.session.commit()

            print("Trabajadores importados correctamente")

    app.run(debug=True)