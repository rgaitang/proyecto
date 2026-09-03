import os
import calendar
from datetime import date, datetime, timedelta
from functools import wraps

from flask import (Flask, render_template, request, redirect, url_for,
                   flash, session, jsonify, send_file, abort)
from werkzeug.utils import secure_filename

from models import db, Sucursal, Empleado, Usuario, Turno, RegistroHoras, Novedad
from excel_generator import generar_archivo_siigo

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INSTANCE_DIR = os.path.join(BASE_DIR, 'instance')
OUTPUT_DIR = os.environ.get('OUTPUT_DIR', os.path.join(BASE_DIR, 'output'))
os.makedirs(INSTANCE_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'cambia-esta-clave-en-produccion')
DATABASE_URL = os.environ.get('DATABASE_URL')
if not DATABASE_URL:
    DATABASE_URL = 'sqlite:///' + os.path.join(INSTANCE_DIR, 'horarios.db')
if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = DATABASE_URL
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db.init_app(app)


# ---------------- Decoradores de permisos ----------------
def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            flash('Debes iniciar sesión', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrapper


def rol_required(*roles):
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if 'user_id' not in session:
                flash('Debes iniciar sesión', 'warning')
                return redirect(url_for('login'))
            usuario = db.session.get(Usuario, session['user_id'])
            if usuario.rol not in roles:
                abort(403)
            return f(*args, **kwargs)
        return wrapper
    return decorator


def get_usuario_actual():
    if 'user_id' not in session:
        return None
    return db.session.get(Usuario, session['user_id'])


# ---------------- Utilidades ----------------
SPANISH_DAYS = ['LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO']


def dias_del_periodo(mes, anio, quincena):
    """quincena = 1 (dias 1-15) o 2 (dias 16-fin)"""
    if quincena == 1:
        inicio, fin = 1, 15
    else:
        inicio = 16
        fin = calendar.monthrange(anio, mes)[1]
    return [date(anio, mes, d) for d in range(inicio, fin + 1)]


# ---------------- Rutas de autenticación ----------------
@app.route('/')
def index():
    return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        usuario = Usuario.query.filter_by(username=username).first()
        if usuario and usuario.check_password(password):
            session['user_id'] = usuario.id
            session['username'] = usuario.username
            session['rol'] = usuario.rol
            flash('Bienvenido', 'success')
            if usuario.rol == 'admin_global':
                return redirect(url_for('panel_global'))
            if usuario.rol == 'admin_local':
                return redirect(url_for('panel_local'))
            return redirect(url_for('dashboard'))
        flash('Usuario o contraseña incorrectos', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    flash('Sesión cerrada', 'info')
    return redirect(url_for('login'))


# ---------------- Dashboard Empleado ----------------
@app.route('/dashboard')
@login_required
def dashboard():
    usuario = get_usuario_actual()
    emp = Empleado.query.filter_by(user_id=usuario.id).first()
    if not emp:
        flash('Tu usuario no está vinculado a un empleado. Contacta a RRHH.', 'danger')
        return redirect(url_for('logout'))

    hoy = date.today()
    mes_act = request.args.get('mes', type=int, default=hoy.month)
    anio_act = request.args.get('anio', type=int, default=hoy.year)

    # Mostrar período actual (quincena vigente)
    quincena_act = 1 if hoy.day <= 15 else 2
    dias = dias_del_periodo(mes_act, anio_act, quincena_act)

    # Registros existentes
    registros = {r.fecha: r for r in RegistroHoras.query.filter_by(
        empleado_id=emp.id).filter(RegistroHoras.fecha >= dias[0],
                                   RegistroHoras.fecha <= dias[-1]).all()}

    turnos = Turno.query.order_by(Turno.codigo).all()

    # Validación: solo puede editar propio
    return render_template('dashboard.html', emp=emp, dias=dias, registros=registros,
                           turnos=turnos, mes=mes_act, anio=anio_act,
                           quincena=quincena_act, hoy=hoy,
                           spanish_days=SPANISH_DAYS)


@app.route('/guardar_turno', methods=['POST'])
@login_required
def guardar_turno():
    usuario = get_usuario_actual()
    emp = Empleado.query.filter_by(user_id=usuario.id).first()
    if not emp:
        flash('Error de vinculación', 'danger')
        return redirect(url_for('dashboard'))

    fecha_str = request.form.get('fecha')
    turno = request.form.get('turno').strip()
    observacion = request.form.get('observacion', '').strip()

    fecha = datetime.strptime(fecha_str, '%Y-%m-%d').date()

    # Regla: no puede editar fechas de períodos cerrados
    if fecha < date.today() - timedelta(days=15):
        flash('El período anterior ya está cerrado, contacta a tu administrador', 'warning')
        return redirect(url_for('dashboard'))

    if turno == '':
        flash('Selecciona un turno o deja vacío para no reportar', 'warning')
        return redirect(url_for('dashboard'))

    reg = RegistroHoras.query.filter_by(empleado_id=emp.id, fecha=fecha).first()
    if reg:
        reg.turno_codigo = turno
        reg.observacion = observacion
        reg.estado = 'pendiente'
    else:
        reg = RegistroHoras(empleado_id=emp.id, fecha=fecha, turno_codigo=turno,
                            observacion=observacion, estado='pendiente',
                            creado_por=usuario.id)
        db.session.add(reg)
    db.session.commit()
    flash('Turno guardado', 'success')
    return redirect(url_for('dashboard'))


@app.route('/borrar_turno', methods=['POST'])
@login_required
def borrar_turno():
    usuario = get_usuario_actual()
    emp = Empleado.query.filter_by(user_id=usuario.id).first()
    fecha_str = request.form.get('fecha')
    fecha = datetime.strptime(fecha_str, '%Y-%m-%d').date()
    reg = RegistroHoras.query.filter_by(empleado_id=emp.id, fecha=fecha).first()
    if reg:
        db.session.delete(reg)
        db.session.commit()
        flash('Turno eliminado', 'success')
    return redirect(url_for('dashboard'))


# ---------------- Panel Admin Local ----------------
@app.route('/panel-local')
@rol_required('admin_local', 'admin_global')
def panel_local():
    usuario = get_usuario_actual()
    emp_admin = Empleado.query.filter_by(user_id=usuario.id).first()

    if usuario.rol == 'admin_local':
        # El admin local solo administra su propia sucursal
        if not emp_admin or not emp_admin.sucursal_id:
            flash('Tu sucursal no está configurada', 'danger')
            return redirect(url_for('dashboard'))
        sucursal_id = emp_admin.sucursal_id
    else:
        # admin global puede elegir sucursal
        sucursal_id = request.args.get('sucursal', type=int)
        if not sucursal_id:
            sucursal_id = 1

    sucursal = db.session.get(Sucursal, sucursal_id)
    empleados = Empleado.query.filter_by(sucursal_id=sucursal_id).order_by(Empleado.nombre).all()

    hoy = date.today()
    mes_act = request.args.get('mes', type=int, default=hoy.month)
    anio_act = request.args.get('anio', type=int, default=hoy.year)
    quincena_act = 1 if hoy.day <= 15 else 2

    # Todas las quincenas disponibles
    quincenas = [(1, dias_del_periodo(mes_act, anio_act, 1)),
                 (2, dias_del_periodo(mes_act, anio_act, 2))]
    quincena_sel = request.args.get('quincena', type=int, default=quincena_act)
    dias = dias_del_periodo(mes_act, anio_act, quincena_sel)

    registros = {r.fecha: r for r in RegistroHoras.query
                 .filter(RegistroHoras.empleado_id.in_([e.id for e in empleados]),
                         RegistroHoras.fecha >= dias[0],
                         RegistroHoras.fecha <= dias[-1]).all()}

    turnos = Turno.query.order_by(Turno.codigo).all()
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()

    return render_template('panel_local.html', empleados=empleados, dias=dias,
                           registros=registros, turnos=turnos, mes=mes_act, anio=anio_act,
                           quincena=quincena_sel, quincenas=quincenas,
                           sucursal=sucursal, sucursales=sucursales,
                           es_global=(usuario.rol == 'admin_global'),
                           spanish_days=SPANISH_DAYS)


@app.route('/admin_editar_turno', methods=['POST'])
@rol_required('admin_local', 'admin_global')
def admin_editar_turno():
    usuario = get_usuario_actual()
    emp_admin = Empleado.query.filter_by(user_id=usuario.id).first()

    empleado_id = request.form.get('empleado_id', type=int)
    fecha_str = request.form.get('fecha')
    turno = request.form.get('turno').strip()
    observacion = request.form.get('observacion', '').strip()
    fecha = datetime.strptime(fecha_str, '%Y-%m-%d').date()

    target = db.session.get(Empleado, empleado_id)
    if not target:
        abort(404)

    # Control: admin local solo su sucursal
    if usuario.rol == 'admin_local':
        if target.sucursal_id != emp_admin.sucursal_id:
            abort(403)

    reg = RegistroHoras.query.filter_by(empleado_id=empleado_id, fecha=fecha).first()
    if turno == '':
        if reg:
            db.session.delete(reg)
            db.session.commit()
        flash('Turno eliminado', 'success')
    else:
        if reg:
            reg.turno_codigo = turno
            reg.observacion = observacion
            reg.estado = 'pendiente'
        else:
            reg = RegistroHoras(empleado_id=empleado_id, fecha=fecha, turno_codigo=turno,
                                observacion=observacion, estado='pendiente', creado_por=usuario.id)
            db.session.add(reg)
        db.session.commit()
        flash('Turno actualizado', 'success')
    return redirect(request.referrer or url_for('panel_local'))


# ---------------- Panel Admin Global (Carolina) ----------------
@app.route('/panel-global')
@rol_required('admin_global')
def panel_global():
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()
    hoy = date.today()

    # Resumen por sucursal
    resumen = []
    for s in sucursales:
        emp_count = Empleado.query.filter_by(sucursal_id=s.id).count()
        pendientes = 0
        empleados = Empleado.query.filter_by(sucursal_id=s.id).all()
        if empleados:
            pendientes = RegistroHoras.query.filter(
                RegistroHoras.empleado_id.in_([e.id for e in empleados])
            ).filter(RegistroHoras.fecha >= date(hoy.year, hoy.month, 1)).count()
        resumen.append({'sucursal': s, 'empleados': emp_count, 'pendientes': pendientes})

    return render_template('panel_global.html', resumen=resumen, hoy=hoy)


@app.route('/admin/empleados')
@rol_required('admin_global')
def admin_empleados():
    empleados = Empleado.query.order_by(Empleado.nombre).all()
    return render_template('admin_empleados.html', empleados=empleados)


@app.route('/admin/usuarios')
@rol_required('admin_global')
def admin_usuarios():
    usuarios = Usuario.query.all()
    return render_template('admin_usuarios.html', usuarios=usuarios)


@app.route('/admin/crear_empleado', methods=['GET', 'POST'])
@rol_required('admin_global')
def crear_empleado():
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()
    if request.method == 'POST':
        cedula = request.form.get('cedula').strip()
        nombre = request.form.get('nombre').strip()
        cargo = request.form.get('cargo', '').strip()
        sucursal_id = request.form.get('sucursal_id', type=int)
        username = request.form.get('username').strip()
        password = request.form.get('password')

        if Empleado.query.filter_by(cedula=cedula).first():
            flash('La cédula ya existe', 'danger')
            return redirect(url_for('crear_empleado'))

        emp = Empleado(cedula=cedula, nombre=nombre, cargo=cargo, sucursal_id=sucursal_id)
        db.session.add(emp)
        db.session.flush()

        usuario = Usuario(username=username, rol='empleado')
        usuario.set_password(password)
        db.session.add(usuario)
        db.session.flush()
        emp.user_id = usuario.id

        db.session.commit()
        flash('Empleado creado', 'success')
        return redirect(url_for('admin_empleados'))
    return render_template('crear_empleado.html', sucursales=sucursales)


@app.route('/admin/editar_empleado/<int:empleado_id>', methods=['GET', 'POST'])
@rol_required('admin_global')
def editar_empleado(empleado_id):
    emp = db.session.get(Empleado, empleado_id)
    if not emp:
        flash('Empleado no encontrado', 'danger')
        return redirect(url_for('admin_empleados'))
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()
    if request.method == 'POST':
        nombre = request.form.get('nombre', '').strip()
        cargo = request.form.get('cargo', '').strip()
        sucursal_id = request.form.get('sucursal_id', type=int)
        cedula = request.form.get('cedula', '').strip()

        if nombre:
            emp.nombre = nombre
        emp.cargo = cargo
        emp.cedula = cedula or emp.cedula
        emp.sucursal_id = sucursal_id or None

        db.session.commit()
        flash('Empleado actualizado', 'success')
        return redirect(url_for('admin_empleados'))
    return render_template('editar_empleado.html', emp=emp, sucursales=sucursales)


@app.route('/admin/crear_usuario', methods=['GET', 'POST'])
@rol_required('admin_global')
def crear_usuario():
    if request.method == 'POST':
        username = request.form.get('username').strip()
        password = request.form.get('password')
        rol = request.form.get('rol')
        if Usuario.query.filter_by(username=username).first():
            flash('El usuario ya existe', 'danger')
            return redirect(url_for('crear_usuario'))
        u = Usuario(username=username, rol=rol)
        u.set_password(password)
        db.session.add(u)
        db.session.commit()
        flash('Usuario creado', 'success')
        return redirect(url_for('admin_usuarios'))
    return render_template('crear_usuario.html')


# ---------------- Generación de Excel para SIIGO ----------------
@app.route('/generar_excel', methods=['POST'])
@rol_required('admin_local', 'admin_global')
def generar_excel():
    usuario = get_usuario_actual()
    emp_admin = Empleado.query.filter_by(user_id=usuario.id).first()

    mes = request.form.get('mes', type=int)
    anio = request.form.get('anio', type=int)

    if not mes or not anio:
        flash('Mes y año obligatorios', 'danger')
        return redirect(request.referrer or url_for('panel_global'))

    if usuario.rol == 'admin_local':
        if not emp_admin or not emp_admin.sucursal_id:
            abort(403)
        sucursal_id = emp_admin.sucursal_id
    else:
        sucursal_id = request.form.get('sucursal_id', type=int)
        if not sucursal_id:
            sucursal_id = 1

    nombre_mes = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO',
                  'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'][mes - 1]

    sucursal = db.session.get(Sucursal, sucursal_id)
    empleados = Empleado.query.filter_by(sucursal_id=sucursal_id).order_by(Empleado.nombre).all()

    # Obtener registros del mes
    registros = {r.empleado_id: {} for r in []}
    for e in empleados:
        regs = RegistroHoras.query.filter_by(empleado_id=e.id).filter(
            RegistroHoras.fecha >= date(anio, mes, 1),
            RegistroHoras.fecha <= date(anio, mes, calendar.monthrange(anio, mes)[1])
        ).all()
        registros[e.id] = {r.fecha.day: r for r in regs}

    novedades = Novedad.query.filter(
        Novedad.empleado_id.in_([e.id for e in empleados]),
        Novedad.fecha_inicio >= date(anio, mes, 1),
        Novedad.fecha_inicio <= date(anio, mes, calendar.monthrange(anio, mes)[1])
    ).all() if empleados else []

    # Mapa turno->descripcion
    turnos_map = {t.codigo: t.descripcion for t in Turno.query.all()}

    archivo = generar_archivo_siigo(OUTPUT_DIR, nombre_mes, anio, mes, sucursal.nombre,
                                    empleados, registros, novedades, turnos_map)

    flash('Archivo Excel generado', 'success')
    return send_file(archivo, as_attachment=True)


# ---------------- Carga inicial (script) ----------------
def init_db():
    db.create_all()
    if Sucursal.query.count() == 0:
        for i in range(1, 7):
            db.session.add(Sucursal(nombre=f'SOLOFARMA {i}'))
        db.session.commit()
        print('Sucursales creadas')


# ---------------- Inicializacion al arrancar (para gunicorn app:app) ----------------
# Crea las tablas y llena datos base (sucursales, turnos, admin, empleados) si la
# base esta vacia. Esto garantiza el funcionamiento aunque Render use 'gunicorn app:app'
# en lugar del start.sh. Es idempotente: solo actua cuando la BD esta vacia.
with app.app_context():
    db.create_all()
    if Sucursal.query.count() == 0:
        print('Base de datos vacia: ejecutando seed inicial...')
        from seed import main as seed_main
        seed_main()

if __name__ == '__main__':
    with app.app_context():
        init_db()
    app.run(debug=True)
