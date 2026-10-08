import os
import calendar
from datetime import date, datetime, timedelta
from functools import wraps

from flask import (Flask, render_template, request, redirect, url_for,
                   flash, session, jsonify, send_file, abort)
from werkzeug.utils import secure_filename
from sqlalchemy import inspect as sa_inspect, text as sa_text

from models import (db, Sucursal, Empleado, Usuario, Turno, RegistroHoras,
                    Novedad, Actividad, Notificacion, ExcelGenerado)
from excel_generator import generar_archivo_siigo, generar_informe_global_siigo

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


# ---------------- Utilidades de trazabilidad y notificaciones ----------------
def quincena_de(fecha):
    return 1 if fecha.day <= 15 else 2


def sucursal_de_usuario(usuario):
    """Sucursal que administra el usuario (admin_local) o None."""
    if not usuario:
        return None
    if usuario.sucursal_id:
        return usuario.sucursal_id
    emp = Empleado.query.filter_by(user_id=usuario.id).first()
    if emp:
        return emp.sucursal_id
    return None


def admin_local_de_sucursal(sucursal_id):
    """Devuelve el usuario admin_local de una sucursal (por user.sucursal_id
    o por su vínculo con un Empleado de esa sucursal)."""
    u = Usuario.query.filter_by(rol='admin_local', sucursal_id=sucursal_id).first()
    if u:
        return u
    for e in Empleado.query.filter_by(sucursal_id=sucursal_id).all():
        if e.user_id:
            usr = db.session.get(Usuario, e.user_id)
            if usr and usr.rol == 'admin_local':
                return usr
    return None


def registrar_actividad(tipo, accion, detalle='', mes=None, anio=None, quincena=None,
                        empleado_id=None, empleado_nombre=None,
                        sucursal_id=None, sucursal_nombre=None):
    """Registra una entrada en la bitácora de trazabilidad."""
    usuario = get_usuario_actual()
    if not usuario:
        return
    if sucursal_id is None:
        sucursal_id = sucursal_de_usuario(usuario)
    if sucursal_nombre is None and sucursal_id:
        s = db.session.get(Sucursal, sucursal_id)
        sucursal_nombre = s.nombre if s else ''
    act = Actividad(fecha=datetime.now(), usuario_id=usuario.id,
                    usuario_nombre=usuario.username, rol=usuario.rol,
                    sucursal_id=sucursal_id, sucursal_nombre=sucursal_nombre,
                    tipo=tipo, accion=accion, detalle=detalle, mes=mes, anio=anio,
                    quincena=quincena, empleado_id=empleado_id,
                    empleado_nombre=empleado_nombre)
    db.session.add(act)
    db.session.commit()


def crear_notificacion(sucursal_id, mensaje, usuario, mes, anio):
    """Crea un aviso hacia RRHH cuando un admin local ingresa una novedad."""
    s = db.session.get(Sucursal, sucursal_id) if sucursal_id else None
    n = Notificacion(creada=datetime.now(), usuario_id=usuario.id,
                     usuario_nombre=usuario.username,
                     sucursal_id=sucursal_id,
                     sucursal_nombre=s.nombre if s else '',
                     tipo='novedad', mensaje=mensaje, leida=False,
                     mes=mes, anio=anio)
    db.session.add(n)
    db.session.commit()


def admin_local_bloqueado(usuario):
    return usuario and usuario.rol == 'admin_local' and usuario.bloqueado


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
            usuario.ultimo_login = datetime.now()
            db.session.commit()
            registrar_actividad('login', 'login',
                                detalle=f'{usuario.username} ingresó al sistema')
            if usuario.bloqueado:
                flash('Tu cuenta está BLOQUEADA: no puedes modificar horarios ni novedades.', 'warning')
            else:
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
        sucursal_id = sucursal_de_usuario(usuario)
        if not sucursal_id:
            flash('Tu sucursal no está configurada', 'danger')
            return redirect(url_for('dashboard'))
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

    # Novedades de la sucursal en el mes seleccionado
    novedades = []
    if empleados:
        novedades = Novedad.query.filter(
            Novedad.empleado_id.in_([e.id for e in empleados]),
            Novedad.fecha_inicio >= date(anio_act, mes_act, 1),
            Novedad.fecha_inicio <= date(anio_act, mes_act, calendar.monthrange(anio_act, mes_act)[1])
        ).order_by(Novedad.fecha_inicio.desc()).all()

    turnos = Turno.query.order_by(Turno.codigo).all()
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()

    return render_template('panel_local.html', empleados=empleados, dias=dias,
                           registros=registros, turnos=turnos, mes=mes_act, anio=anio_act,
                           quincena=quincena_sel, quincenas=quincenas,
                           sucursal=sucursal, sucursales=sucursales,
                           novedades=novedades,
                           admin_bloqueado=admin_local_bloqueado(usuario),
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
        sucursal_admin = sucursal_de_usuario(usuario)
        if target.sucursal_id != sucursal_admin:
            abort(403)
        # Control de bloqueo de RRHH
        if usuario.bloqueado:
            flash('Estás bloqueado por RRHH. No puedes modificar horarios.', 'danger')
            return redirect(request.referrer or url_for('panel_local'))

    reg = RegistroHoras.query.filter_by(empleado_id=empleado_id, fecha=fecha).first()
    detalle = (f'{target.nombre} | {fecha.isoformat()}'
               f' | antes: {reg.turno_codigo if reg else "sin turno"}'
               f' -> después: {turno if turno else "sin turno"}')
    if turno == '':
        if reg:
            db.session.delete(reg)
            db.session.commit()
            registrar_actividad('horario', 'eliminar', detalle=detalle,
                                mes=fecha.month, anio=fecha.year,
                                quincena=quincena_de(fecha),
                                empleado_id=target.id, empleado_nombre=target.nombre)
            flash('Turno eliminado', 'success')
    else:
        if reg:
            reg.turno_codigo = turno
            reg.observacion = observacion
            reg.estado = 'pendiente'
            accion = 'editar'
        else:
            reg = RegistroHoras(empleado_id=empleado_id, fecha=fecha, turno_codigo=turno,
                                observacion=observacion, estado='pendiente', creado_por=usuario.id)
            db.session.add(reg)
            accion = 'crear'
        db.session.commit()
        registrar_actividad('horario', accion, detalle=detalle,
                            mes=fecha.month, anio=fecha.year,
                            quincena=quincena_de(fecha),
                            empleado_id=target.id, empleado_nombre=target.nombre)
        flash('Turno actualizado', 'success')
    return redirect(request.referrer or url_for('panel_local'))


# ---------------- Novedades (admin local / global) ----------------
@app.route('/admin_novedad', methods=['POST'])
@rol_required('admin_local', 'admin_global')
def admin_novedad():
    usuario = get_usuario_actual()

    if admin_local_bloqueado(usuario):
        flash('Estás bloqueado por RRHH. No puedes ingresar novedades.', 'danger')
        return redirect(request.referrer or url_for('panel_local'))

    tipo = request.form.get('tipo', '').strip().upper()
    empleado_id = request.form.get('empleado_id', type=int)
    fecha_str = request.form.get('fecha_inicio', '')
    fecha_fin_str = request.form.get('fecha_fin', '').strip()
    descripcion = request.form.get('descripcion', '').strip()
    reporta = request.form.get('reporta', '').strip()

    emp = db.session.get(Empleado, empleado_id)
    if not emp or not tipo or not fecha_str:
        flash('Completa todos los campos obligatorios de la novedad', 'danger')
        return redirect(request.referrer or url_for('panel_local'))

    try:
        fecha_inicio = datetime.strptime(fecha_str, '%Y-%m-%d').date()
        fecha_fin = datetime.strptime(fecha_fin_str, '%Y-%m-%d').date() if fecha_fin_str else None
    except ValueError:
        flash('Fecha inválida', 'danger')
        return redirect(request.referrer or url_for('panel_local'))

    sucursal_id = emp.sucursal_id
    if usuario.rol == 'admin_local':
        if emp.sucursal_id != sucursal_de_usuario(usuario):
            abort(403)

    nov = Novedad(empleado_id=emp.id, tipo=tipo, fecha_inicio=fecha_inicio,
                  fecha_fin=fecha_fin, descripcion=descripcion, reporta=reporta)
    db.session.add(nov)
    db.session.commit()

    sucursal = db.session.get(Sucursal, sucursal_id)
    registrar_actividad('novedad', 'crear',
                        detalle=f'{emp.nombre} | {tipo} | {fecha_inicio.isoformat()}'
                                f'{" a " + fecha_fin.isoformat() if fecha_fin else ""}'
                                f' | {descripcion}',
                        mes=fecha_inicio.month, anio=fecha_inicio.year,
                        quincena=quincena_de(fecha_inicio),
                        empleado_id=emp.id, empleado_nombre=emp.nombre,
                        sucursal_id=sucursal_id,
                        sucursal_nombre=sucursal.nombre if sucursal else '')
    crear_notificacion(
        sucursal_id,
        f'El admin {usuario.username} ingresó la novedad "{tipo}" de {emp.nombre} '
        f'para el {fecha_inicio.isoformat()}.',
        usuario, fecha_inicio.month, fecha_inicio.year)
    flash('Novedad ingresada', 'success')
    return redirect(request.referrer or url_for('panel_local'))


@app.route('/admin_novedad_borrar', methods=['POST'])
@rol_required('admin_local', 'admin_global')
def admin_novedad_borrar():
    usuario = get_usuario_actual()

    if admin_local_bloqueado(usuario):
        flash('Estás bloqueado por RRHH. No puedes eliminar novedades.', 'danger')
        return redirect(request.referrer or url_for('panel_local'))

    nov = db.session.get(Novedad, request.form.get('novedad_id', type=int))
    if not nov:
        abort(404)
    emp = db.session.get(Empleado, nov.empleado_id)
    if usuario.rol == 'admin_local' and emp and emp.sucursal_id != sucursal_de_usuario(usuario):
        abort(403)

    detalle = f'{emp.nombre if emp else nov.empleado_id} | {nov.tipo} | {nov.fecha_inicio}'
    db.session.delete(nov)
    db.session.commit()
    registrar_actividad('novedad', 'eliminar', detalle=detalle,
                        mes=nov.fecha_inicio.month, anio=nov.fecha_inicio.year,
                        quincena=quincena_de(nov.fecha_inicio),
                        empleado_id=nov.empleado_id,
                        empleado_nombre=emp.nombre if emp else '')
    flash('Novedad eliminada', 'success')
    return redirect(request.referrer or url_for('panel_local'))
@app.route('/panel-global')
@rol_required('admin_global')
def panel_global():
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()
    hoy = date.today()

    # Resumen por sucursal + estado de bloqueo + novedades del mes
    resumen = []
    informe_mes = request.args.get('informe_mes', type=int, default=hoy.month)
    informe_anio = request.args.get('informe_anio', type=int, default=hoy.year)
    for s in sucursales:
        emp_count = Empleado.query.filter_by(sucursal_id=s.id).count()
        empleados = Empleado.query.filter_by(sucursal_id=s.id).all()
        pendientes = 0
        if empleados:
            pendientes = RegistroHoras.query.filter(
                RegistroHoras.empleado_id.in_([e.id for e in empleados])
            ).filter(RegistroHoras.fecha >= date(informe_anio, informe_mes, 1)).count()
        # Novedades del mes
        nov_count = 0
        if empleados:
            nov_count = Novedad.query.filter(
                Novedad.empleado_id.in_([e.id for e in empleados]),
                Novedad.fecha_inicio >= date(informe_anio, informe_mes, 1),
                Novedad.fecha_inicio <= date(informe_anio, informe_mes, calendar.monthrange(informe_anio, informe_mes)[1])
            ).count()
        admin = admin_local_de_sucursal(s.id)
        resumen.append({'sucursal': s, 'empleados': emp_count, 'pendientes': pendientes,
                        'novedades': nov_count, 'admin': admin,
                        'bloqueado': bool(admin and admin.bloqueado)})

    notificaciones = Notificacion.query.order_by(Notificacion.creada.desc()).limit(40).all()
    no_leidas = Notificacion.query.filter_by(leida=False).count()
    sucursales_con_novedad = sum(1 for r in resumen if r['novedades'] > 0)

    return render_template('panel_global.html', resumen=resumen, hoy=hoy,
                           notificaciones=notificaciones, no_leidas=no_leidas,
                           informe_mes=informe_mes, informe_anio=informe_anio,
                           sucursales_con_novedad=sucursales_con_novedad)


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


@app.route('/admin/crear_usuario', methods=['GET', 'POST'])
@rol_required('admin_global')
def crear_usuario():
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()
    if request.method == 'POST':
        username = request.form.get('username').strip()
        password = request.form.get('password')
        rol = request.form.get('rol')
        sucursal_id = request.form.get('sucursal_id', type=int)
        if Usuario.query.filter_by(username=username).first():
            flash('El usuario ya existe', 'danger')
            return redirect(url_for('crear_usuario'))
        u = Usuario(username=username, rol=rol)
        if rol == 'admin_local':
            u.sucursal_id = sucursal_id
        u.set_password(password)
        db.session.add(u)
        db.session.commit()
        flash('Usuario creado', 'success')
        return redirect(url_for('admin_usuarios'))
    return render_template('crear_usuario.html', sucursales=sucursales)


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
        sucursal_id = sucursal_de_usuario(usuario)
        if not sucursal_id:
            abort(403)
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

    # Histórico + trazabilidad
    db.session.add(ExcelGenerado(fecha=datetime.now(), usuario_id=usuario.id,
                                 usuario_nombre=usuario.username,
                                 sucursal_id=sucursal.id, sucursal_nombre=sucursal.nombre,
                                 nombre_mes=nombre_mes, mes=mes, anio=anio,
                                 nombre_archivo=os.path.basename(archivo), es_global=False))
    db.session.commit()
    registrar_actividad('excel', 'generar',
                        detalle=f'Excel SIIGO de {sucursal.nombre} - {nombre_mes} {anio}',
                        mes=mes, anio=anio,
                        sucursal_id=sucursal.id, sucursal_nombre=sucursal.nombre)

    flash('Archivo Excel generado', 'success')
    return send_file(archivo, as_attachment=True)


@app.route('/generar_excel_global', methods=['POST'])
@rol_required('admin_global')
def generar_excel_global():
    mes = request.form.get('mes', type=int)
    anio = request.form.get('anio', type=int)
    if not mes or not anio:
        flash('Mes y año obligatorios', 'danger')
        return redirect(url_for('panel_global'))

    nombre_mes = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO',
                  'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'][mes - 1]

    # Verificar que las 6 sucursales ya ingresaron novedades
    faltantes = []
    sucursales_data = []
    for s in Sucursal.query.order_by(Sucursal.nombre).all():
        empleados = Empleado.query.filter_by(sucursal_id=s.id).all()
        novedades = []
        if empleados:
            novedades = Novedad.query.filter(
                Novedad.empleado_id.in_([e.id for e in empleados]),
                Novedad.fecha_inicio >= date(anio, mes, 1),
                Novedad.fecha_inicio <= date(anio, mes, calendar.monthrange(anio, mes)[1])
            ).all()
        if not novedades:
            faltantes.append(s.nombre)
        registros = {}
        for e in empleados:
            registros[e.id] = {
                r.fecha.day: r for r in RegistroHoras.query.filter_by(empleado_id=e.id).filter(
                    RegistroHoras.fecha >= date(anio, mes, 1),
                    RegistroHoras.fecha <= date(anio, mes, calendar.monthrange(anio, mes)[1])
                ).all()}
        sucursales_data.append({'sucursal': s, 'empleados': empleados,
                                'registros': registros, 'novedades': novedades})

    if faltantes:
        flash(f'Faltan novedades de: {", ".join(faltantes)}. El informe global no puede generarse aún.', 'danger')
        return redirect(url_for('panel_global'))

    turnos_map = {t.codigo: t.descripcion for t in Turno.query.all()}
    archivo = generar_informe_global_siigo(OUTPUT_DIR, nombre_mes, anio, mes,
                                           sucursales_data, turnos_map)

    usuario = get_usuario_actual()
    db.session.add(ExcelGenerado(fecha=datetime.now(), usuario_id=usuario.id,
                                 usuario_nombre=usuario.username,
                                 sucursal_id=None, sucursal_nombre='TODAS LAS SUCURSALES',
                                 nombre_mes=nombre_mes, mes=mes, anio=anio,
                                 nombre_archivo=os.path.basename(archivo), es_global=True))
    db.session.commit()
    registrar_actividad('excel', 'generar',
                        detalle=f'INFORME GLOBAL de novedades - {nombre_mes} {anio}',
                        mes=mes, anio=anio,
                        sucursal_nombre='TODAS LAS SUCURSALES')

    flash('Informe global de novedades generado', 'success')
    return send_file(archivo, as_attachment=True)


# ---------------- Notificaciones y bloqueo de admin local ----------------
@app.route('/notificacion/leer/<int:nid>', methods=['POST'])
@rol_required('admin_global')
def notificacion_leer(nid):
    n = db.session.get(Notificacion, nid)
    if n:
        n.leida = True
        db.session.commit()
    return redirect(url_for('panel_global'))


@app.route('/notificaciones/leer_todas', methods=['POST'])
@rol_required('admin_global')
def notificaciones_leer_todas():
    for n in Notificacion.query.filter_by(leida=False).all():
        n.leida = True
    db.session.commit()
    flash('Notificaciones marcadas como leídas', 'info')
    return redirect(url_for('panel_global'))


@app.route('/admin/bloquear_sucursal/<int:sucursal_id>', methods=['POST'])
@rol_required('admin_global')
def bloquear_sucursal(sucursal_id):
    admin = admin_local_de_sucursal(sucursal_id)
    if not admin:
        flash('No hay admin local configurado para esta sucursal', 'warning')
        return redirect(url_for('panel_global'))
    admin.bloqueado = True
    admin.motivo_bloqueo = request.form.get('motivo', '')
    db.session.commit()
    s = db.session.get(Sucursal, sucursal_id)
    registrar_actividad('bloqueo', 'bloquear',
                        detalle=f'Bloqueado el admin {admin.username} '
                                f'({request.form.get("motivo", "") or "sin motivo"})',
                        sucursal_id=sucursal_id,
                        sucursal_nombre=s.nombre if s else '')
    flash(f'Admin {admin.username} bloqueado', 'success')
    return redirect(url_for('panel_global'))


@app.route('/admin/desbloquear_sucursal/<int:sucursal_id>', methods=['POST'])
@rol_required('admin_global')
def desbloquear_sucursal(sucursal_id):
    admin = admin_local_de_sucursal(sucursal_id)
    if not admin:
        flash('No hay admin local configurado para esta sucursal', 'warning')
        return redirect(url_for('panel_global'))
    admin.bloqueado = False
    admin.motivo_bloqueo = ''
    db.session.commit()
    s = db.session.get(Sucursal, sucursal_id)
    registrar_actividad('bloqueo', 'desbloquear',
                        detalle=f'Desbloqueado el admin {admin.username}',
                        sucursal_id=sucursal_id,
                        sucursal_nombre=s.nombre if s else '')
    flash(f'Admin {admin.username} desbloqueado', 'success')
    return redirect(url_for('panel_global'))


# ---------------- Trazabilidad e histórico de Excel ----------------
@app.route('/trazabilidad')
@rol_required('admin_global')
def trazabilidad():
    f_sucursal = request.args.get('sucursal', type=int)
    f_tipo = request.args.get('tipo', type=str)
    f_mes = request.args.get('mes', type=int)
    f_anio = request.args.get('anio', type=int)

    q = Actividad.query
    if f_sucursal:
        q = q.filter_by(sucursal_id=f_sucursal)
    if f_tipo:
        q = q.filter_by(tipo=f_tipo)
    if f_mes:
        q = q.filter_by(mes=f_mes)
    if f_anio:
        q = q.filter_by(anio=f_anio)
    actividades = q.order_by(Actividad.fecha.desc()).limit(800).all()

    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()
    return render_template('trazabilidad.html', actividades=actividades,
                           sucursales=sucursales, f_sucursal=f_sucursal,
                           f_tipo=f_tipo, f_mes=f_mes, f_anio=f_anio)


@app.route('/historico_excel')
@rol_required('admin_global')
def historico_excel():
    logs = ExcelGenerado.query.order_by(ExcelGenerado.fecha.desc()).all()
    return render_template('historico_excel.html', logs=logs)


@app.route('/descargar_excel/<int:log_id>')
@rol_required('admin_global')
def descargar_excel(log_id):
    log = db.session.get(ExcelGenerado, log_id)
    if not log:
        abort(404)
    ruta = os.path.join(OUTPUT_DIR, log.nombre_archivo)
    if not os.path.exists(ruta):
        flash('El archivo ya no existe en el servidor', 'danger')
        return redirect(url_for('historico_excel'))
    return send_file(ruta, as_attachment=True)


# ---------------- Carga inicial (script) ----------------
def migrar_bd():
    """Migraciones ligeras: agrega columnas nuevas a tablas existentes
    sin perder los datos actuales. Compatible con SQLite y Postgres."""
    try:
        insp = sa_inspect(db.engine)
        tablas = insp.get_table_names()
        if 'usuario' not in tablas:
            return
        dialect = db.engine.dialect.name
        tipo_dt = 'TIMESTAMP' if dialect == 'postgresql' else 'DATETIME'
        cols = {c['name'] for c in insp.get_columns('usuario')}
        adiciones = {
            'bloqueado': 'BOOLEAN',
            'motivo_bloqueo': 'VARCHAR(200)',
            'ultimo_login': tipo_dt,
            'sucursal_id': 'INTEGER',
        }
        with db.engine.begin() as conn:
            for col, ddl in adiciones.items():
                if col not in cols:
                    conn.execute(sa_text(f'ALTER TABLE usuario ADD COLUMN {col} {ddl}'))
    except Exception as e:
        print(f'[migrar_bd] No se pudo ajustar la tabla usuario: {e}')


def init_db():
    migrar_bd()
    db.create_all()
    if Sucursal.query.count() == 0:
        for i in range(1, 7):
            db.session.add(Sucursal(nombre=f'SOLOFARMA {i}'))
        db.session.commit()
        print('Sucursales creadas')


if __name__ == '__main__':
    with app.app_context():
        init_db()
    app.run(debug=True)
