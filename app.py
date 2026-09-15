import os
import calendar
from datetime import date, datetime, timedelta
from functools import wraps

from flask import (Flask, render_template, request, redirect, url_for,
                   flash, session, jsonify, send_file, abort)
from werkzeug.utils import secure_filename
from sqlalchemy import inspect as sa_inspect, text as sa_text, func as sa_func, cast as sa_cast, BigInteger as sa_BigInteger

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


# ---------------- Utilidades de trazabilidad, notificaciones y bloqueo ----------------
def quincena_de(fecha):
    return 1 if fecha.day <= 15 else 2


def sucursal_de_usuario(usuario):
    """Sucursal que administra el usuario (admin_local) o None."""
    if not usuario:
        return None
    emp = usuario.empleado
    if emp:
        return emp.sucursal_id
    return None


def admin_local_de_sucursal(sucursal_id):
    """Devuelve el usuario admin_local de una sucursal (vinculado a un Empleado)."""
    for e in Empleado.query.filter_by(sucursal_id=sucursal_id).all():
        if e.usuario and e.usuario.rol == 'admin_local':
            return e.usuario
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
    return bool(usuario and usuario.rol == 'admin_local' and usuario.bloqueado)


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
            # Los empleados NO pueden iniciar sesión; solo admins (sucursal y global).
            if usuario.rol == 'empleado':
                flash('Los empleados no pueden iniciar sesión. Contacta a la administración.', 'danger')
                return render_template('login.html')
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

    turnos = Turno.query.order_by(Turno.codigo).all()
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()

    return render_template('panel_local.html', empleados=empleados, dias=dias,
                           registros=registros, turnos=turnos, mes=mes_act, anio=anio_act,
                           quincena=quincena_sel, quincenas=quincenas,
                           sucursal=sucursal, sucursales=sucursales,
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


# ---------------- Novedades (solo admin local / global) ----------------
TIPOS_NOVEDAD = ['INCAPACIDAD', 'NO VINO', 'S-POSITIVA', 'INGRESO', 'RETIRO', 'CAMBIO_TURNO', 'VACACIONES', 'OTRO']


@app.route('/nuevas_novedades', methods=['GET', 'POST'])
@rol_required('admin_local', 'admin_global')
def nuevas_novedades():
    usuario = get_usuario_actual()
    emp_admin = Empleado.query.filter_by(user_id=usuario.id).first()

    def _sucursal_de_admin():
        if usuario.rol == 'admin_local':
            if not emp_admin or not emp_admin.sucursal_id:
                return None
            return emp_admin.sucursal_id
        # admin global: elegir sucursal
        return request.args.get('sucursal', type=int) or request.form.get('sucursal_id', type=int)

    sucursal_id = _sucursal_de_admin()
    if not sucursal_id:
        flash('Tu sucursal no está configurada', 'danger')
        return redirect(url_for('panel_local' if usuario.rol == 'admin_local' else 'panel_global'))

    if request.method == 'POST':
        if admin_local_bloqueado(usuario):
            flash('Estás bloqueado por RRHH. No puedes ingresar novedades.', 'danger')
            return redirect(url_for('nuevas_novedades', sucursal=sucursal_id))

        empleado_id = request.form.get('empleado_id', type=int)
        tipo = request.form.get('tipo')
        codigo = request.form.get('codigo', '').strip()
        fecha_inicio = request.form.get('fecha_inicio')
        fecha_fin = request.form.get('fecha_fin') or None
        descripcion = request.form.get('descripcion', '')
        reporta = request.form.get('reporta', '')

        target = db.session.get(Empleado, empleado_id)
        if not target:
            flash('Empleado no encontrado', 'danger')
            return redirect(url_for('nuevas_novedades', sucursal=sucursal_id))

        # Control: admin local solo su sucursal
        if usuario.rol == 'admin_local' and target.sucursal_id != sucursal_id:
            abort(403)

        try:
            f_inicio = datetime.strptime(fecha_inicio, '%Y-%m-%d').date()
        except Exception:
            flash('Fecha de inicio inválida', 'danger')
            return redirect(url_for('nuevas_novedades', sucursal=sucursal_id))

        f_fin = None
        if fecha_fin:
            try:
                f_fin = datetime.strptime(fecha_fin, '%Y-%m-%d').date()
            except Exception:
                f_fin = None

        n = Novedad(empleado_id=empleado_id, tipo=tipo, codigo=codigo, fecha_inicio=f_inicio,
                    fecha_fin=f_fin, descripcion=descripcion, reporta=reporta,
                    estado='pendiente')
        db.session.add(n)
        db.session.commit()

        suc = db.session.get(Sucursal, sucursal_id)
        registrar_actividad('novedad', 'crear',
                            detalle=f'{target.nombre} | {tipo} | {f_inicio.isoformat()}'
                                    f'{" a " + f_fin.isoformat() if f_fin else ""}'
                                    f' | {descripcion}',
                            mes=f_inicio.month, anio=f_inicio.year,
                            quincena=quincena_de(f_inicio),
                            empleado_id=target.id, empleado_nombre=target.nombre,
                            sucursal_id=sucursal_id,
                            sucursal_nombre=suc.nombre if suc else '')
        if usuario.rol == 'admin_local':
            crear_notificacion(sucursal_id,
                               f'El admin {usuario.username} ingresó la novedad "{tipo}" '
                               f'de {target.nombre} para el {f_inicio.isoformat()}.',
                               usuario, f_inicio.month, f_inicio.year)
        flash('Novedad registrada', 'success')
        return redirect(url_for('nuevas_novedades', sucursal=sucursal_id))

    # GET: mostrar formulario + listado de novedades de la sucursal
    sucursal = db.session.get(Sucursal, sucursal_id)
    empleados = Empleado.query.filter_by(sucursal_id=sucursal_id).order_by(Empleado.nombre).all()
    novedades = Novedad.query.filter(Novedad.empleado_id.in_([e.id for e in empleados])) \
        .order_by(Novedad.fecha_inicio.desc()).all() if empleados else []
    emp_ids = {e.id for e in empleados}

    return render_template('nuevas_novedades.html', sucursal=sucursal, empleados=empleados,
                           novedades=novedades, tipos=TIPOS_NOVEDAD, emp_ids=emp_ids,
                           es_global=(usuario.rol == 'admin_global'),
                           admin_bloqueado=admin_local_bloqueado(usuario),
                           sucursales=Sucursal.query.order_by(Sucursal.nombre).all(),
                           hoy=date.today())


@app.route('/eliminar_novedad/<int:novedad_id>', methods=['POST'])
@rol_required('admin_local', 'admin_global')
def eliminar_novedad(novedad_id):
    usuario = get_usuario_actual()
    emp_admin = Empleado.query.filter_by(user_id=usuario.id).first()
    n = db.session.get(Novedad, novedad_id)
    if not n:
        flash('Novedad no encontrada', 'danger')
        return redirect(request.referrer or url_for('panel_local'))
    target = db.session.get(Empleado, n.empleado_id)
    if usuario.rol == 'admin_local':
        if not emp_admin or not target or target.sucursal_id != emp_admin.sucursal_id:
            abort(403)
        if usuario.bloqueado:
            flash('Estás bloqueado por RRHH. No puedes eliminar novedades.', 'danger')
            return redirect(request.referrer or url_for('panel_local'))
    detalle = f'{target.nombre if target else n.empleado_id} | {n.tipo} | {n.fecha_inicio}'
    db.session.delete(n)
    db.session.commit()
    registrar_actividad('novedad', 'eliminar', detalle=detalle,
                        mes=n.fecha_inicio.month, anio=n.fecha_inicio.year,
                        quincena=quincena_de(n.fecha_inicio),
                        empleado_id=n.empleado_id,
                        empleado_nombre=target.nombre if target else '')
    flash('Novedad eliminada', 'success')
    return redirect(request.referrer or url_for('panel_local'))


# ---------------- Panel Admin Global (Carolina) ----------------
@app.route('/panel-global')
@rol_required('admin_global')
def panel_global():
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()
    hoy = date.today()

    informe_mes = request.args.get('informe_mes', type=int, default=hoy.month)
    informe_anio = request.args.get('informe_anio', type=int, default=hoy.year)

    # Resumen por sucursal + estado de bloqueo + novedades del mes
    resumen = []
    for s in sucursales:
        emp_count = Empleado.query.filter_by(sucursal_id=s.id).count()
        empleados = Empleado.query.filter_by(sucursal_id=s.id).all()
        pendientes = 0
        nov_count = 0
        if empleados:
            pendientes = RegistroHoras.query.filter(
                RegistroHoras.empleado_id.in_([e.id for e in empleados])
            ).filter(RegistroHoras.fecha >= date(informe_anio, informe_mes, 1)).count()
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

    # Consolidado de novedades de todas las sucursales para el periodo elegido
    ini = date(informe_anio, informe_mes, 1)
    fin = date(informe_anio, informe_mes, calendar.monthrange(informe_anio, informe_mes)[1])
    filas = db.session.query(Novedad, Empleado).join(
        Empleado, Novedad.empleado_id == Empleado.id
    ).filter(Novedad.fecha_inicio >= ini, Novedad.fecha_inicio <= fin) \
        .order_by(Novedad.fecha_inicio, Empleado.sucursal_id).all()
    sucursal_cache = {}
    novedades_global = []
    for n, e in filas:
        s = sucursal_cache.get(e.sucursal_id) or db.session.get(Sucursal, e.sucursal_id)
        sucursal_cache[e.sucursal_id] = s
        novedades_global.append({'novedad': n, 'empleado': e,
                                 'sucursal_nombre': s.nombre if s else '?'})

    return render_template('panel_global.html', resumen=resumen, hoy=hoy,
                           notificaciones=notificaciones, no_leidas=no_leidas,
                           informe_mes=informe_mes, informe_anio=informe_anio,
                           sucursales_con_novedad=sucursales_con_novedad,
                           novedades_global=novedades_global)


@app.route('/admin/empleados')
@rol_required('admin_global')
def admin_empleados():
    sort = request.args.get('sort', 'nombre')
    direction = request.args.get('dir', 'asc')
    SORTABLE = {'cedula', 'nombre', 'cargo', 'sucursal', 'usuario'}
    if sort not in SORTABLE:
        sort = 'nombre'
    if direction not in ('asc', 'desc'):
        direction = 'asc'

    q = Empleado.query
    if sort == 'cedula':
        col = sa_cast(sa_func.coalesce(Empleado.cedula_real, Empleado.cedula), sa_BigInteger)
    elif sort == 'nombre':
        col = sa_func.coalesce(Empleado.nombre_real, Empleado.nombre)
    elif sort == 'cargo':
        col = Empleado.cargo
    elif sort == 'sucursal':
        q = q.outerjoin(Empleado.sucursal_ref)
        col = Sucursal.nombre
    else:  # 'usuario'
        q = q.outerjoin(Empleado.usuario)
        col = Usuario.username

    col = col.desc() if direction == 'desc' else col.asc()
    col = col.nullslast()
    empleados = q.order_by(col, Empleado.nombre).all()
    return render_template('admin_empleados.html', empleados=empleados,
                           sort=sort, dir=direction)


@app.route('/admin/usuarios')
@rol_required('admin_global')
def admin_usuarios():
    sort = request.args.get('sort', 'usuario')
    direction = request.args.get('dir', 'asc')
    SORTABLE = {'usuario', 'rol', 'empleado', 'cedula', 'sucursal'}
    if sort not in SORTABLE:
        sort = 'usuario'
    if direction not in ('asc', 'desc'):
        direction = 'asc'

    q = Usuario.query
    if sort == 'usuario':
        col = Usuario.username
    elif sort == 'rol':
        col = Usuario.rol
    elif sort == 'empleado':
        q = q.outerjoin(Usuario.empleado)
        col = sa_func.coalesce(Empleado.nombre_real, Empleado.nombre)
    elif sort == 'cedula':
        q = q.outerjoin(Usuario.empleado)
        col = sa_cast(sa_func.coalesce(Empleado.cedula_real, Empleado.cedula), sa_BigInteger)
    else:  # 'sucursal'
        q = q.outerjoin(Usuario.empleado).outerjoin(Empleado.sucursal_ref)
        col = Sucursal.nombre

    col = col.desc() if direction == 'desc' else col.asc()
    col = col.nullslast()
    usuarios = q.order_by(col, Usuario.username).all()
    return render_template('admin_usuarios.html', usuarios=usuarios,
                           sort=sort, dir=direction)


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
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()
    if request.method == 'POST':
        username = request.form.get('username').strip()
        password = request.form.get('password')
        rol = request.form.get('rol')
        sucursal_id = request.form.get('sucursal_id', type=int)
        empleado_id = request.form.get('empleado_id', type=int)

        if Usuario.query.filter_by(username=username).first():
            flash('El usuario ya existe', 'danger')
            return redirect(url_for('crear_usuario'))

        # Para admin_local y empleado, debe estar vinculado a una persona (empleado).
        # El empleado define la sucursal a la que pertenece el usuario.
        emp = None
        if rol in ('admin_local', 'empleado'):
            if not empleado_id:
                flash('Selecciona la sede y el empleado para este usuario', 'danger')
                return redirect(url_for('crear_usuario'))
            emp = db.session.get(Empleado, empleado_id)
            if not emp:
                flash('Empleado no encontrado', 'danger')
                return redirect(url_for('crear_usuario'))

        u = Usuario(username=username, rol=rol)
        u.set_password(password)
        db.session.add(u)
        db.session.flush()

        if emp:
            emp.user_id = u.id
            # admin_local: el empleado vinculado debe tener sucursal asignada
            if rol == 'admin_local' and not emp.sucursal_id and sucursal_id:
                emp.sucursal_id = sucursal_id

        db.session.commit()
        flash('Usuario creado', 'success')
        return redirect(url_for('admin_usuarios'))
    return render_template('crear_usuario.html', sucursales=sucursales)


@app.route('/admin/editar_usuario/<int:usuario_id>', methods=['GET', 'POST'])
@rol_required('admin_global')
def editar_usuario(usuario_id):
    u = db.session.get(Usuario, usuario_id)
    if not u:
        flash('Usuario no encontrado', 'danger')
        return redirect(url_for('admin_usuarios'))
    sucursales = Sucursal.query.order_by(Sucursal.nombre).all()

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password')
        rol = request.form.get('rol')
        sucursal_id = request.form.get('sucursal_id', type=int)
        empleado_id = request.form.get('empleado_id', type=int)

        # Validar username unico (si se cambio)
        if username and username != u.username:
            if Usuario.query.filter_by(username=username).first():
                flash('Ese nombre de usuario ya existe', 'danger')
                return redirect(url_for('editar_usuario', usuario_id=usuario_id))
            u.username = username

        if rol:
            u.rol = rol

        if password:
            if len(password) < 4:
                flash('La contraseña debe tener al menos 4 caracteres', 'danger')
                return redirect(url_for('editar_usuario', usuario_id=usuario_id))
            u.set_password(password)

        # Manejar vinculo con empleado
        if rol in ('admin_local', 'empleado'):
            if empleado_id:
                emp = db.session.get(Empleado, empleado_id)
                if emp:
                    emp.user_id = u.id
                    if rol == 'admin_local' and not emp.sucursal_id and sucursal_id:
                        emp.sucursal_id = sucursal_id
                    db.session.commit()
                    flash('Usuario actualizado', 'success')
                    return redirect(url_for('admin_usuarios'))
            flash('Selecciona la sede y el empleado para este usuario', 'danger')
            return redirect(url_for('editar_usuario', usuario_id=usuario_id))
        else:
            # admin_global: quitar vinculo a empleado si lo tenia (pierde la sede)
            if u.empleado:
                u.empleado.user_id = None

        db.session.commit()
        flash('Usuario actualizado', 'success')
        return redirect(url_for('admin_usuarios'))

    # GET: preparar datos para el formulario
    emp_actual = u.empleado if u.empleado else None
    sucursal_sel = emp_actual.sucursal_id if emp_actual else None
    return render_template('editar_usuario.html', u=u, sucursales=sucursales,
                           emp_actual=emp_actual, sucursal_sel=sucursal_sel,
                           todos_tipos=['admin_global', 'admin_local', 'empleado'])


@app.route('/admin/eliminar_empleado/<int:empleado_id>', methods=['POST'])
@rol_required('admin_global')
def eliminar_empleado(empleado_id):
    emp = db.session.get(Empleado, empleado_id)
    if not emp:
        flash('Empleado no encontrado', 'danger')
        return redirect(url_for('admin_empleados'))
    if emp.usuario:
        flash('No se puede eliminar: el empleado tiene un usuario de login asociado', 'danger')
        return redirect(url_for('admin_empleados'))
    nombre = emp.nombre
    RegistroHoras.query.filter_by(empleado_id=emp.id).delete(synchronize_session=False)
    Novedad.query.filter_by(empleado_id=emp.id).delete(synchronize_session=False)
    db.session.delete(emp)
    db.session.commit()
    flash(f'Empleado "{nombre}" eliminado', 'success')
    return redirect(url_for('admin_empleados'))


@app.route('/admin/eliminar_usuario/<int:usuario_id>', methods=['POST'])
@rol_required('admin_global')
def eliminar_usuario(usuario_id):
    u = db.session.get(Usuario, usuario_id)
    if not u:
        flash('Usuario no encontrado', 'danger')
        return redirect(url_for('admin_usuarios'))
    if u.username == 'carolina':
        flash('No se puede eliminar la cuenta principal (carolina)', 'danger')
        return redirect(url_for('admin_usuarios'))
    username = u.username
    if u.empleado:
        u.empleado.user_id = None
    db.session.delete(u)
    db.session.commit()
    flash(f'Usuario "{username}" eliminado', 'success')
    return redirect(url_for('admin_usuarios'))


@app.route('/cambiar_password', methods=['POST'])
@login_required
def cambiar_password():
    usuario = get_usuario_actual()
    actual = request.form.get('actual', '')
    nueva = request.form.get('nueva', '')
    confirmar = request.form.get('confirmar', '')

    if not usuario.check_password(actual):
        flash('La contraseña actual es incorrecta', 'danger')
        return redirect(request.referrer or url_for('panel_local'))
    if len(nueva) < 4:
        flash('La nueva contraseña debe tener al menos 4 caracteres', 'danger')
        return redirect(request.referrer or url_for('panel_local'))
    if nueva != confirmar:
        flash('Las contraseñas no coinciden', 'danger')
        return redirect(request.referrer or url_for('panel_local'))

    usuario.set_password(nueva)
    db.session.commit()
    flash('Contraseña actualizada', 'success')
    return redirect(request.referrer or url_for('panel_local'))


@app.route('/admin/empleados_por_sucursal')
@rol_required('admin_global')
def empleados_por_sucursal():
    sucursal_id = request.args.get('sucursal_id', type=int)
    if not sucursal_id:
        return jsonify([])
    empleados = Empleado.query.filter_by(sucursal_id=sucursal_id).order_by(Empleado.nombre).all()
    return jsonify([{'id': e.id, 'nombre': e.nombre, 'cedula': e.cedula} for e in empleados])


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


# ---------------- Informe global de novedades ----------------
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

    # Verificar que todas las sucursales ya ingresaron novedades
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
        sucursales_data.append({'sucursal': s, 'empleados': empleados, 'novedades': novedades})

    if faltantes:
        flash(f'Faltan novedades de: {", ".join(faltantes)}. El informe global no puede generarse aún.', 'danger')
        return redirect(url_for('panel_global'))

    archivo = generar_informe_global_siigo(OUTPUT_DIR, nombre_mes, anio, mes, sucursales_data)

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


# ---------------- Migración (SQLite y Postgres) ----------------
def migrar_bd():
    """Añade columnas nuevas a tablas existentes según la BD usada."""
    insp = sa_inspect(db.engine)
    if 'usuario' in insp.get_table_names():
        cols = {c['name'] for c in insp.get_columns('usuario')}
        with db.engine.begin() as conn:
            if 'bloqueado' not in cols:
                t = 'BOOLEAN DEFAULT FALSE' if db.engine.dialect.name == 'postgresql' else 'BOOLEAN DEFAULT 0'
                conn.execute(sa_text(f'ALTER TABLE usuario ADD COLUMN bloqueado {t}'))
            if 'motivo_bloqueo' not in cols:
                t = 'VARCHAR(200)' if db.engine.dialect.name == 'postgresql' else 'VARCHAR(200)'
                conn.execute(sa_text(f'ALTER TABLE usuario ADD COLUMN motivo_bloqueo {t}'))
            if 'ultimo_login' not in cols:
                t = 'TIMESTAMP' if db.engine.dialect.name == 'postgresql' else 'DATETIME'
                conn.execute(sa_text(f'ALTER TABLE usuario ADD COLUMN ultimo_login {t}'))


def init_db():
    db.create_all()
    if Sucursal.query.count() == 0:
        from seed import NOMBRES_SUCURSALES
        for nombre in NOMBRES_SUCURSALES:
            db.session.add(Sucursal(nombre=nombre))
        db.session.commit()
        print('Sucursales creadas')


# ---------------- Inicializacion al arrancar (para gunicorn app:app) ----------------
# Migra, crea las tablas y llena datos base (sucursales, turnos, admin, empleados) si la
# base esta vacia. Esto garantiza el funcionamiento aunque Render use 'gunicorn app:app'
# en lugar del start.sh. Es idempotente: solo actua cuando la BD esta vacia.
with app.app_context():
    migrar_bd()
    db.create_all()
    from seed import main as seed_main
    seed_main()

if __name__ == '__main__':
    with app.app_context():
        init_db()
    app.run(debug=True)
