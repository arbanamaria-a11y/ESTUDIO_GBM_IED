import io
from flask import Flask, render_template, request, redirect, url_for, send_file, session
from flask_mysqldb import MySQL
import MySQLdb.cursors
from openpyxl import Workbook
from conexion import Config

app = Flask(__name__)

# Clave secreta requerida para manejar sesiones de usuario (session)
app.secret_key = "clave_secreta_estudio_gbm"

# ==========================
# Configuración de MySQL
# ==========================
app.config['MYSQL_HOST'] = Config.MYSQL_HOST
app.config['MYSQL_USER'] = Config.MYSQL_USER
app.config['MYSQL_PASSWORD'] = Config.MYSQL_PASSWORD
app.config['MYSQL_DB'] = Config.MYSQL_DB
app.config['MYSQL_PORT'] = int(getattr(Config, 'MYSQL_PORT', 24725))

# Configuración SSL obligatoria para Aiven
app.config['MYSQL_SSL'] = {'ssl_mode': 'REQUIRED'}

# Inicialización de la extensión MySQL (AQUÍ ESTABA LA LÍNEA QUE FALTABA)
mysql = MySQL(app)


# ==========================
# PÁGINAS PRINCIPALES
# ==========================

@app.route("/")
def inicio():
    return render_template("index.html")

@app.route("/misionyvision")
def mision():
    return render_template("misionyvision.html")

@app.route("/quienessomos")
def quienes_somos():
    return render_template("quienessomos.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        tipo_usuario = request.form.get("tipo_usuario")
        usuario_ingresado = request.form.get("usuario")
        contrasena_ingresada = request.form.get("contrasena")

       
        cursor = mysql.connection.cursor()
        
        cursor.execute("""
            SELECT * FROM usuario
            WHERE nombre_usuario = %s
            AND contrasena = %s
            AND roles = %s
        """, (usuario_ingresado, contrasena_ingresada, tipo_usuario))

        cuenta = cursor.fetchone()
        cursor.close()


        if cuenta:
            # Guardamos los datos principales en la sesión del usuario
            session["id_usuario"] = cuenta[0]
            session["usuario"] = cuenta[1]

            # Redirige según el rol seleccionado
            if tipo_usuario == "administrador":
                return redirect(url_for("administrador"))
            elif tipo_usuario == "estudiante":
                return redirect(url_for("estudiante"))
            elif tipo_usuario == "docente":
                return redirect(url_for("docente"))
            elif tipo_usuario == "directivos":
                return redirect(url_for("directivos"))
            else:
                return redirect(url_for("inicio_usuario"))
        else:
            return render_template("login.html", error="Datos incorrectos")

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ==========================
# PANELS Y MÓDULOS
# ==========================

@app.route("/inicio_usuario")
def inicio_usuario():
    if "id_usuario" not in session:
        return redirect(url_for("login"))
    return render_template("inicio_usuario.html", nombre=session.get("usuario"), id_usuario=session.get("id_usuario"))

@app.route("/administrador")
def administrador():
    return render_template("administrador.html")

@app.route("/cursos")
def cursos():
    return render_template("cursos.html")

@app.route("/notas", methods=["GET", "POST"])
def notas():
    if request.method == "POST":
        id_usuario = request.form["id_usuario"]
        asignatura = request.form["asignatura"]
        nota = request.form["nota"]
        fecha = request.form["fecha"]

        cursor = mysql.connection.cursor()
        cursor.execute("""
            INSERT INTO notas (id_usuario, asignatura, nota, fecha)
            VALUES (%s, %s, %s, %s)
        """, (id_usuario, asignatura, nota, fecha))
        mysql.connection.commit()
        cursor.close()

        return redirect(url_for("notas"))

    cursor = mysql.connection.cursor()
    cursor.execute("SELECT * FROM notas")
    datos = cursor.fetchall()
    cursor.close()
    return render_template("notas.html", notas=datos)

@app.route("/editar_nota/<int:id_nota>", methods=["GET", "POST"])
def editar_nota(id_nota):
    if request.method == "POST":
        asignatura = request.form["asignatura"]
        nota = request.form["nota"]
        fecha = request.form["fecha"]

        cursor = mysql.connection.cursor()
        cursor.execute("""
            UPDATE notas
            SET asignatura = %s, nota = %s, fecha = %s
            WHERE id_nota = %s
        """, (asignatura, nota, fecha, id_nota))
        mysql.connection.commit()
        cursor.close()

        return redirect(url_for("notas"))

    cursor = mysql.connection.cursor()
    cursor.execute("SELECT * FROM notas WHERE id_nota = %s", (id_nota,))
    dato = cursor.fetchone()
    cursor.close()
    return render_template("editar_nota.html", nota=dato)

@app.route("/eliminar_nota/<int:id_nota>")
def eliminar_nota(id_nota):
    cursor = mysql.connection.cursor()
    cursor.execute("DELETE FROM notas WHERE id_nota = %s", (id_nota,))
    mysql.connection.commit()
    cursor.close()
    return redirect(url_for("notas"))

@app.route("/registrousuario", methods=["GET", "POST"])
def registro():
    return render_template("registrousuario.html")

@app.route("/estudiante")
def estudiante():
    return render_template("estudiante.html")

@app.route("/registroestudiante")
def registro_estudiante():
    return render_template("registroestudiante.html")

@app.route("/docente")
def docente():
    return render_template("docente.html")

@app.route("/directivos")
def directivos():
    return render_template("directivos.html")

@app.route("/consultar_usuarios")
def consultar_usuarios():
    return render_template("registrousuario.html")

@app.route("/editar_usuario/<int:id>")
def editar_usuario(id):
    return render_template("registrousuario.html")

@app.route("/eliminar_usuarios")
def eliminar_usuarios():
    return render_template("registrousuario.html")


# ======================================
# REPORTES Y EXPORTAR A EXCEL (CORREGIDO)
# ======================================

@app.route("/reporte")
def reporte():
    if "id_usuario" not in session:
        return redirect(url_for("login"))

    fecha = request.args.get("fecha")
    cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)

    if fecha:
        cursor.execute("""
            SELECT
                r.id_reserva,
                u.id_usuario,
                u.nombre_usuario AS nombre,
                u.correo,
                r.fecha_reserva,
                r.hora_reserva,
                r.cantidad_personas,
                r.estado,
                r.observaciones
            FROM reservas r
            INNER JOIN usuario u ON r.id_usuario = u.id_usuario
            WHERE r.fecha_reserva = %s
            ORDER BY r.fecha_reserva DESC, r.hora_reserva DESC
        """, (fecha,))
    else:
        cursor.execute("""
            SELECT
                r.id_reserva,
                u.id_usuario,
                u.nombre_usuario AS nombre,
                u.correo,
                r.fecha_reserva,
                r.hora_reserva,
                r.cantidad_personas,
                r.estado,
                r.observaciones
            FROM reservas r
            INNER JOIN usuario u ON r.id_usuario = u.id_usuario
            ORDER BY r.fecha_reserva DESC, r.hora_reserva DESC
        """)

    reservas = cursor.fetchall()
    cursor.close()

    return render_template("reporte.html", reservas=reservas, fecha=fecha)


@app.route("/reporte_excel")
def reporte_excel():
    if "id_usuario" not in session:
        return redirect(url_for("login"))

    fecha = request.args.get("fecha")
    cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)

    if fecha:
        cursor.execute("""
            SELECT
                r.id_reserva,
                u.id_usuario,
                u.nombre_usuario AS nombre,
                u.correo,
                r.fecha_reserva,
                r.hora_reserva,
                r.cantidad_personas,
                r.estado,
                r.observaciones
            FROM reservas r
            INNER JOIN usuario u ON r.id_usuario = u.id_usuario
            WHERE r.fecha_reserva = %s
            ORDER BY r.fecha_reserva DESC
        """, (fecha,))
    else:
        cursor.execute("""
            SELECT
                r.id_reserva,
                u.id_usuario,
                u.nombre_usuario AS nombre,
                u.correo,
                r.fecha_reserva,
                r.hora_reserva,
                r.cantidad_personas,
                r.estado,
                r.observaciones
            FROM reservas r
            INNER JOIN usuario u ON r.id_usuario = u.id_usuario
            ORDER BY r.fecha_reserva DESC
        """)

    datos = cursor.fetchall()
    cursor.close()

    libro = Workbook()
    hoja = libro.active
    hoja.title = "Reporte Reservas"

    hoja.append([
        "ID Reserva",
        "ID Usuario",
        "Nombre",
        "Correo",
        "Fecha Reserva",
        "Hora Reserva",
        "Cantidad Personas",
        "Estado",
        "Observaciones"
    ])

    for fila in datos:
        hoja.append([
            fila["id_reserva"],
            fila["id_usuario"],
            fila["nombre"],
            fila["correo"],
            str(fila["fecha_reserva"]),
            str(fila["hora_reserva"]),
            fila["cantidad_personas"],
            fila["estado"],
            fila["observaciones"]
        ])

    archivo = io.BytesIO()
    libro.save(archivo)
    archivo.seek(0)

    return send_file(
        archivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        download_name="reporte_reservas.xlsx",
        as_attachment=True
    )


# ==========================
# VERIFICACIÓN DE CONEXIÓN
# ==========================

@app.route("/usuarios")
def usuarios():
    try:
        cursor = mysql.connection.cursor()
        cursor.execute("SELECT * FROM usuario")
        datos = cursor.fetchall()
        cursor.close()
        return {
            "mensaje": "Conexión exitosa con la base de datos MySQL",
            "cantidad_registros": len(datos),
            "usuarios": datos
        }
    except Exception as error:
        return {
            "mensaje": "Error al conectar con la base de datos",
            "detalle": str(error)
        }, 500


if __name__ == "__main__":
    app.run(debug=True)