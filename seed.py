"""Script para poblar la base de datos (local o en producción).

Uso:
    - Local (SQLite):        python seed.py
    - Produccion (Postgres):  set DATABASE_URL=postgresql://...  y luego python seed.py
                              (o ejecútalo dentro del servicio de Render con un comando shell)

Crea: 6 sucursales, tabla de turnos completa (~45 turnos), usuarios admin demo.
"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app, db
from models import Sucursal, Empleado, Usuario, Turno

# Nombres de las 6 sucursales, en orden (1 a 6)
NOMBRES_SUCURSALES = ['SOLOFARMA NIQUIA', 'SOLOFARMA PRADO', 'SOLOFARMA GUANTEROS',
                      'SOLOFARMA POSITIVA', 'SOLOFARMA BICENTENARIO', 'SOLOFARMA CASTILLA']

# Tabla de turnos (codigo -> (descripcion, es_domingo))
# Sincronizada con la hoja 'TURNO MODIFICADO' / 'TURNOS' del excel
# 'HORARIOS PUNTOS DE VENTA SOLOFARMA.xlsx'
TURNOS = [
    ('A', '7 AM - 2 PM- BICENTENARIO', False),
    ('B', '1 PM - 8 PM DOMINGO', True),
    ('C', '2 PM - 9 PM BICENTENARIO', False),
    ('D', '8 AM - 3 PM- BICENTENARIO', False),
    ('E', '8 AM - 3 PM- BICENTENARIO DOMINGO', True),
    ('F', '1 PM - 8 PM BICENTENARIO', False),
    ('G', '8 AM - 4 PM', False),
    ('H', '8 AM - 5 PM', False),
    ('I', '8 AM - 3 PM', False),
    ('J', '8 AM - 8 PM BICENTENARIO DOMINGO (S.O.S)', True),
    ('K', '8 AM - 7 PM', False),
    ('L', '8 AM - 1 PM/ 5 PM - 10 PM', False),
    ('M', 'NO TENER EN CUENTA', False),
    ('N', '8 AM - 1 PM/ 6 PM - 10 PM', False),
    ('O', '8 AM - 1 PM/ 7 PM - 10 PM', False),
    ('P', '9 AM - 5 PM', False),
    ('Q', '9 AM - 4 PM', False),
    ('R', '9 AM - 10 PM', False),
    ('S', '10 AM - 6 PM', False),
    ('T', '10 AM - 10 PM', False),
    ('U', '1 PM - 9 PM', False),
    ('V', '9 AM - 5 PM DOMINGO', True),
    ('W', '11 AM - 7 PM', False),
    ('A1', '1 PM - 10 PM', False),
    ('F1', '3 PM - 10 PM', False),
    ('I1', '7 AM - 9 PM - BICENTENARIO', False),
    ('J1', '9 AM - 1 PM/ 6 PM - 10 PM DOMINGO', True),
    ('K1', '9 AM - 4 PM DOMINGO', True),
    ('L1', '9 AM -2 PM/ 5 PM-10 PM DOMINGO', True),
    ('M1', '9 AM - 1 PM/ 5 PM - 10 PM DOMINGO', True),
    ('N1', '9 AM - 2 PM/ 7 PM - 10 PM DOMINGO', True),
    ('O1', '9 AM - 10 PM - DOMINGO', True),
    ('P1', '9 AM - 1 PM / 4 PM - 10 PM DOMINGO', True),
    ('Q1', '9 AM - 1 PM / 7 PM - 10 PM DOMINGO', True),
    ('R1', '10 AM - 10 PM DOMINGO', True),
    ('S1', '10 AM - 6 PM DOMINGO', True),
    ('T1', '11 AM - 7 PM DOMINGO', True),
    ('U1', '1 PM - 10 PM DOMINGO', True),
    ('V1', '9 AM - 1 PM / 6 PM - 10 PM', False),
    ('W1', '2 PM - 10 PM - DOMINGO', True),
    ('A2', '3 PM - 10 PM - DOMINGO', True),
    ('B2', '1 PM - 9 PM - DOMINGO', True),
    ('Z', 'HORARIO MARGARITA', False),
    # Novedades (codigos de letra para novedades en el Excel SIIGO)
    ('INCAPACIDAD', 'Incapacidad', False),
    ('NO VINO', 'No vino', False),
    ('S-POSITIVA', 'S-Positiva (permiso/cita)', False),
    ('VACACIONES', 'Vacaciones', False),
]


def cargar_empleados_json():
    """Intenta cargar empleados desde un JSON exportado (data/empleados.json).
    Si no existe, retorna lista vacia (solo se crearan sucursales/turnos/usuario admin)."""
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'empleados.json')
    if os.path.exists(ruta):
        with open(ruta, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def cargar_nomina():
    """Carga la nómina real (data/nomina.json): lista de dicts cedula/nombre/cargo.
    Generado a partir de Trabajadores.xls. Lista vacia si el archivo no existe."""
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'nomina.json')
    if os.path.exists(ruta):
        with open(ruta, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def _normalizar_tokens(s):
    """Conjunto de palabras de un nombre normalizado (sin acentos, minusculas)."""
    import unicodedata
    s = unicodedata.normalize('NFD', s.lower())
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return set(p for p in s.replace('ñ', 'n').split() if p)


def _migrar_esquema():
    """Agrega columnas faltantes a tablas existentes (compatible SQLite y PostgreSQL)."""
    from sqlalchemy import inspect, text
    insp = inspect(db.engine)
    if 'novedad' in insp.get_table_names():
        cols = {c['name'] for c in insp.get_columns('novedad')}
        if 'codigo' not in cols:
            db.session.execute(text('ALTER TABLE novedad ADD COLUMN codigo VARCHAR(20) DEFAULT \'\''))
            db.session.commit()
            print('Migracion: columna codigo agregada a novedad')
    if 'usuario' in insp.get_table_names():
        cols = {c['name'] for c in insp.get_columns('usuario')}
        if 'rol' not in cols:
            db.session.execute(text('ALTER TABLE usuario ADD COLUMN rol VARCHAR(20) DEFAULT \'empleado\''))
            db.session.commit()
            print('Migracion: columna rol agregada a usuario')
    if 'empleado' in insp.get_table_names():
        cols = {c['name'] for c in insp.get_columns('empleado')}
        if 'cedula_real' not in cols:
            db.session.execute(text('ALTER TABLE empleado ADD COLUMN cedula_real VARCHAR(20)'))
            db.session.commit()
            print('Migracion: columna cedula_real agregada a empleado')
        if 'nombre_real' not in cols:
            db.session.execute(text('ALTER TABLE empleado ADD COLUMN nombre_real VARCHAR(120)'))
            db.session.commit()
            print('Migracion: columna nombre_real agregada a empleado')


def _cargar_cedulas_reales():
    """Asigna cedula_real / nombre_real a los empleados del sistema usando la
    lista real de la hoja DATOS de la plantilla. Idempotente: solo actualiza
    cuando el valor difiere. Los empleados sin match quedan sin cedula real."""
    import unicodedata

    def normalizar(s):
        s = unicodedata.normalize('NFD', s.lower())
        s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
        s = s.replace('ñ', 'n')
        return ' '.join(p for p in s.split() if p)

    mapeo = {
        # matches automaticos (2+ palabras)
        'andres felipe murillo': ('1054919363', 'Murillo Vasquez Andres Felipe'),
        'bryan andres montoya': ('1000761939', 'Montoya Higuita Bryan Andres'),
        'caterine lopez': ('1037607749', 'Lopez Sotelo Caterine'),
        'deicy natalia acevedo': ('1035433568', 'Acevedo Echeverri Deicy Natalia'),
        'denis alejandra quintero': ('1152703555', 'Quintero Lujan Denis Alejandra'),
        'diego arturo henao': ('98498026', 'Henao Munoz Diego Arturo'),
        'edith florinda garcia': ('43692133', 'Garcia Palencia Edith Florinda'),
        'edwin jose frias': ('1080021724', 'Frias Florez Edwin Jose'),
        'erica vannesa posada': ('1038358213', 'Posada Mesa Erica Vannesa'),
        'erika juliana perez': ('1036618798', 'Perez Bedoya Erika Juliana'),
        'esneider de jesus vasquez': ('70551493', 'Vasquez Giron Esneider de Jesus'),
        'francisco jose correa': ('71717188', 'Correa Vasquez Francisco Jose'),
        'idania esther osorio': ('1066508855', 'Osorio Oviedo Idania'),
        'ingrid jhonna cueto': ('22705114', 'Cueto Ramos Ingrid Jhonna'),
        'jennifer natalia laverde': ('1000456347', 'Laverde Torres Jenifer Natalia'),
        'john alexander duque': ('98696981', 'Duque Garcia John Alexander'),
        'jose de jesus orrego': ('1001724926', 'Orrego Franco Jose de Jesus'),
        'jose manuel jaramillo': ('1000446110', 'Jaramillo Londoño Jose Manuel'),
        'juvenal sanchez': ('70569992', 'Sanchez Maya Juvenal'),
        'leonel arturo barrera': ('1102825707', 'Barrera Beltran Leonel Arturo'),
        'lina patricia taborda': ('43584878', 'Taborda Londoño Lina Patricia'),
        'lisveth karina munera': ('1041410064', 'Munera Monsalve Lisveth Karina'),
        'maira alejandra londono': ('1120354380', 'Londoño Avila Maira Alejandra'),
        'mario alberto urrea': ('98577161', 'Urrea Giraldo Mario Alberto'),
        'melany guerra': ('1001248624', 'Guerra Acevedo Melany'),
        'monica maria garcia': ('43904796', 'Munera Garcia Monica Maria'),
        'osnaider andres moreno': ('1127609039', 'Moreno Payares Osnaider Andres'),
        'paula andrea restrepo': ('1037671519', 'Restrepo Riaza Paula Andrea'),
        'samuel restrepo': ('1017922576', 'Restrepo Valderrama Samuel'),
        'sebastian ganan': ('1000194657', 'Ganan Montoya Sebastian'),
        'sulman yurley muneton': ('1007328591', 'Muneton ochoa Sulman Yurley'),
        'viviana isabel silva': ('1038109472', 'Silva Piñerez Viviana Isabel'),
        'yeison stiven jimenez': ('1193122499', 'Jimenez Loaiza Yeison Stiven'),
        'yennifer andrea muneton': ('1007283803', 'Muneton Ochoa Yennifer Andrea'),
        # nombres cortos de NIQUIA (confirmados por el usuario)
        'adriana': ('56098624', 'Campo Lopez Adriana Maria'),
        'andrea': ('56098624', 'Campo Lopez Adriana Maria'),
        'edwin': ('1080021724', 'Frias Florez Edwin Jose'),
        'monica': ('43904796', 'Munera Garcia Monica Maria'),
        'melany': ('1001248624', 'Guerra Acevedo Melany'),
        'deicy': ('1035433568', 'Acevedo Echeverri Deicy Natalia'),
        'dayana': ('1001738439', 'Pino Avendano Dayana Alejandra'),
        'angie': ('1001753481', 'Angie Niyered Agudelo'),
        'yennifer': ('1007283803', 'Muneton Ochoa Yennifer Andrea'),
        'karina': ('1041410064', 'Munera Monsalve Lisveth Karina'),
        'edith': ('43692133', 'Garcia Palencia Edith Florinda'),
    }
    mapeo = {normalizar(k): v for k, v in mapeo.items()}
    cambios = 0
    for emp in Empleado.query.all():
        clave = normalizar(emp.nombre)
        if clave in mapeo:
            ced, nom = mapeo[clave]
            if emp.cedula_real != ced or emp.nombre_real != nom:
                emp.cedula_real = ced
                emp.nombre_real = nom
                cambios += 1
    if cambios:
        db.session.commit()
        print(f'Cedulas reales actualizadas: {cambios} empleados')


def _limpiar_empleados_duplicados():
    """Elimina empleados que quedaron sin cedula_real y sin estar vinculados a un
    usuario de login (los nombres cortos/duplicados que no matchean a la plantilla).
    Conserva siempre a los 6 administradores de sucursal y a Carolina. Idempotente:
    si no hay candidatos, no hace nada."""
    import unicodedata

    def normalizar(s):
        s = unicodedata.normalize('NFD', s.lower())
        s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
        s = s.replace('ñ', 'n')
        return ' '.join(p for p in s.split() if p)

    # Nombres completos que ya tienen match con cédula real (no se deben tocar)
    nombres_completos = {
        normalizar(n) for n in [
            'Andres Felipe Murillo', 'Bryan Andres Montoya', 'Caterine Lopez',
            'Deicy Natalia Acevedo', 'Denis Alejandra Quintero', 'Diego Arturo Henao',
            'Edith Florinda Garcia', 'Edwin Jose Frias', 'Erica Vannesa Posada',
            'Erika Juliana Perez', 'Esneider De Jesus Vasquez', 'Francisco Jose Correa',
            'Idania Esther Osorio', 'Ingrid Jhonna Cueto', 'Jennifer Natalia Laverde',
            'John Alexander Duque', 'Jose De Jesus Orrego', 'Jose Manuel Jaramillo',
            'Juvenal Sanchez', 'Leonel Arturo Barrera', 'Lina Patricia Taborda',
            'Lisveth Karina Munera', 'Maira Alejandra Londono', 'Mario Alberto Urrea',
            'Melany Guerra', 'Monica Maria Garcia', 'Osnaider Andres Moreno',
            'Paula Andrea Restrepo', 'Samuel Restrepo', 'Sebastian Ganan',
            'Sulman Yurley Muneton', 'Viviana Isabel Silva', 'Yeison Stiven Jimenez',
            'Yennifer Andrea Muneton',
        ]
    }

    candidatos = []
    for emp in Empleado.query.all():
        # Conservar empleados de login (admin) y a los que ya tienen cedula real
        if emp.user_id is not None:
            continue
        if emp.cedula_real:
            continue
        # Conservar empleados cuyo nombre es un 'nombre completo' (aunque no matcheen)
        if normalizar(emp.nombre) in nombres_completos:
            continue
        # Candidato a eliminar: nombre corto/duplicado sin cedula real
        candidatos.append(emp)

    if not candidatos:
        print('Limpieza: no hay empleados duplicados que eliminar')
        return

    # Eliminar dependencias primero (novedades y registros) para no violar FK
    for emp in candidatos:
        try:
            from models import Novedad, RegistroHoras
            Novedad.query.filter_by(empleado_id=emp.id).delete(synchronize_session=False)
            RegistroHoras.query.filter_by(empleado_id=emp.id).delete(synchronize_session=False)
            db.session.delete(emp)
        except Exception as e:
            db.session.rollback()
            print(f'  [skip] {emp.nombre}: {e}')
    db.session.commit()
    print(f'Limpieza: {len(candidatos)} empleados duplicados eliminados: ' +
          ', '.join(e.nombre for e in candidatos))


def _cargar_empleados_nomina():
    """Sube los empleados de la nómina real (data/nomina.json) a la BD.
    - Crea los empleados que no existen (sin crear usuarios de login).
    - Actualiza nombre/cédula/cargo de los existentes, detectándolos por
      cédula real o por coincidencia de nombre (tokens).
    - NO toca user_id: Carolina y los regentes conservan su vinculación.
    - NO asigna sucursal: queda a cargo del admin global desde la aplicación.
    Idempotente: se puede ejecutar en cada arranque sin duplicar."""
    nomina = cargar_nomina()
    if not nomina:
        print('Nomina: no hay data/nomina.json')
        return

    empleados = Empleado.query.all()
    creados = []
    actualizados = []

    for row in nomina:
        ced = str(row.get('cedula', '')).strip()
        nombre = str(row.get('nombre', '')).strip()
        cargo = str(row.get('cargo', '')).strip()
        if not ced or not nombre:
            continue
        toks = _normalizar_tokens(nombre)

        # 1) Match por cédula (real o interna)
        emp = next((e for e in empleados
                    if (e.cedula_real or '') == ced or e.cedula == ced), None)
        # 2) Match por nombre (tokens: el conjunto más corto es subconjunto del otro)
        if emp is None:
            for e in empleados:
                et = _normalizar_tokens(e.nombre or '') or _normalizar_tokens(e.nombre_real or '')
                if et and toks and (toks <= et or et <= toks):
                    emp = e
                    break

        if emp is None:
            db.session.add(Empleado(cedula=ced, nombre=nombre, cedula_real=ced,
                                    nombre_real=nombre, cargo=cargo))
            creados.append(nombre)
        else:
            if (emp.cedula_real or '') != ced:
                emp.cedula_real = ced
            if emp.cedula != ced:
                emp.cedula = ced
            if (emp.nombre_real or '') != nombre:
                emp.nombre_real = nombre
            if emp.nombre != nombre:
                emp.nombre = nombre
            if not emp.cargo:
                emp.cargo = cargo
            actualizados.append(nombre)

    db.session.commit()
    print(f'Nomina: {len(creados)} empleados creados, {len(actualizados)} actualizados')
    if creados:
        print('  Creados: ' + ', '.join(creados[:15]))


def _fusionar_duplicados_nomina():
    """Fusiona registros duplicados de la misma persona (misma cedula_real),
    conservando el vinculado a un usuario de login y reasignando sus
    registros de horas/novedades al registro principal. Idempotente."""
    from models import Novedad, RegistroHoras
    por_cedula = {}
    for e in Empleado.query.all():
        if e.cedula_real:
            por_cedula.setdefault(e.cedula_real, []).append(e)

    fusionados = 0
    for ced, lista in por_cedula.items():
        if len(lista) < 2:
            continue
        primario = next((e for e in lista if e.user_id is not None), None)
        if primario is None:
            primario = max(lista, key=lambda e: len(e.nombre))
        for e in lista:
            if e.id == primario.id or e.user_id is not None:
                continue
            # reasignar registros de horas
            for reg in RegistroHoras.query.filter_by(empleado_id=e.id).all():
                if RegistroHoras.query.filter_by(empleado_id=primario.id, fecha=reg.fecha).first():
                    db.session.delete(reg)
                else:
                    reg.empleado_id = primario.id
            # reasignar novedades
            for n in Novedad.query.filter_by(empleado_id=e.id).all():
                n.empleado_id = primario.id
            db.session.delete(e)
            fusionados += 1

    if fusionados:
        db.session.commit()
        print(f'Nomina: {fusionados} duplicados fusionados')


def _depurar_solo_nomina():
    """Deja únicamente los empleados de la nómina real (data/nomina.json).
    Elimina cualquier otro empleado (datos semilla/plantilla) junto con sus
    registros y novedades. NUNCA elimina usuarios de login: si se elimina el
    empleado vinculado a un admin, el usuario se conserva para reasignarlo.
    Idempotente."""
    from models import Novedad, RegistroHoras
    nomina = cargar_nomina()
    if not nomina:
        return
    mantenidos = set()
    for row in nomina:
        ced = str(row.get('cedula', '')).strip()
        if ced:
            mantenidos.add(ced)

    eliminados = []
    for emp in Empleado.query.all():
        if (emp.cedula_real or emp.cedula) in mantenidos:
            continue
        Novedad.query.filter_by(empleado_id=emp.id).delete(synchronize_session=False)
        RegistroHoras.query.filter_by(empleado_id=emp.id).delete(synchronize_session=False)
        db.session.delete(emp)
        eliminados.append(emp.nombre)

    if eliminados:
        db.session.commit()
        print(f'Nomina: {len(eliminados)} empleados fuera de la nómina eliminados: ' +
              ', '.join(eliminados))
    else:
        print('Nomina: no hay empleados fuera de la nómina que eliminar')


def _vincular_carolina_y_regente():
    """Vinculaciones SOLO iniciales y no destructivas:
    - carolina (admin_global) -> 'Sepulveda Vides Ana Carolina' si está libre.
    - admin1 (admin_local NIQUIA) -> regente 'Orrego Franco Jose De Jesus'.
    NUNCA sobrescribe asignaciones existentes: si el usuario ya tiene empleado
    vinculado o el empleado ya está ocupado, no toca nada (el admin global
    reasigna manualmente al admin de cada sucursal)."""
    cambios = []

    carolina = Usuario.query.filter_by(username='carolina').first()
    admin1 = Usuario.query.filter_by(username='admin1').first()

    if carolina:
        ana = (Empleado.query.filter_by(nombre='Sepulveda Vides Ana Carolina').first()
               or Empleado.query.filter_by(nombre='Ana Carolina Sepulveda').first())
        if ana:
            if ana.user_id and ana.user_id != carolina.id:
                otro = db.session.get(Usuario, ana.user_id)
                ana.user_id = None
                cambios.append(f"'Sepulveda Vides Ana Carolina' desvinculada de "
                               f"{otro.username if otro else ana.user_id}")
        if carolina.empleado is None and ana and ana.user_id is None:
            ana.user_id = carolina.id
            cambios.append("'carolina' vinculada a 'Sepulveda Vides Ana Carolina'")

    if admin1 and admin1.empleado is None:
        regente = (Empleado.query.filter_by(nombre='Orrego Franco Jose De Jesus').first()
                   or Empleado.query.filter_by(nombre='Jose De Jesus Orrego').first())
        if regente and regente.user_id is None:
            regente.user_id = admin1.id
            cambios.append("admin1 vinculado al regente 'Orrego Franco Jose De Jesus'")

    if cambios:
        db.session.commit()
        print('Vinculaciones iniciales:')
        for c in cambios:
            print('  -', c)
    else:
        print('Vinculaciones ya correctas')


def main():
    with app.app_context():
        db.create_all()

        # ---- Migracion ligera de esquema (idempotente) ----
        _migrar_esquema()

        # Sucursales: crear las faltantes y renombrar las existentes en orden (1 a 6)
        if Sucursal.query.count() == 0:
            for nombre in NOMBRES_SUCURSALES:
                db.session.add(Sucursal(nombre=nombre))
            db.session.commit()
            print('Sucursales creadas')
        else:
            existing = Sucursal.query.order_by(Sucursal.id).limit(len(NOMBRES_SUCURSALES)).all()
            for i, s in enumerate(existing):
                nuevo = NOMBRES_SUCURSALES[i]
                if s.nombre != nuevo:
                    s.nombre = nuevo
            db.session.commit()
            print('Sucursales renombradas')

        # Turnos: crear los faltantes y actualizar las descripciones de los existentes
        for codigo, desc, es_dom in TURNOS:
            t = Turno.query.filter_by(codigo=codigo).first()
            if t:
                if t.descripcion != desc or t.es_domingo != es_dom:
                    t.descripcion = desc
                    t.es_domingo = es_dom
            else:
                db.session.add(Turno(codigo=codigo, descripcion=desc, es_domingo=es_dom))
        db.session.commit()
        print(f'{len(TURNOS)} turnos garantizados')

        # Usuario admin global (si no existe)
        if not Usuario.query.filter_by(username='carolina').first():
            carolina = Usuario(username='carolina', rol='admin_global')
            carolina.set_password(os.environ.get('ADMIN_PASS', 'admin123'))
            db.session.add(carolina)
            db.session.commit()
            print('Usuario admin global creado: carolina')
        else:
            print('Usuario carolina ya existia')

        # Empleados desde JSON (si hay)
        empleados = cargar_empleados_json()
        if empleados and Empleado.query.count() == 0:
            for i, emp in enumerate(empleados):
                e = Empleado(cedula=emp.get('cedula', f'100{i:06d}'),
                             nombre=emp.get('nombre', f'EMP {i}'),
                             cargo=emp.get('cargo', ''),
                             sucursal_id=emp.get('sucursal_id', (i % 6) + 1))
                db.session.add(e)
            db.session.commit()
            print(f'{len(empleados)} empleados cargados desde JSON')

        # Administradores de sucursal (admin1..admin6), uno por sede.
        # Cada uno queda vinculado al primer empleado de su sucursal.
        for n in range(1, 7):
            username = f'admin{n}'
            if not Usuario.query.filter_by(username=username).first():
                empleado_sede = Empleado.query.filter_by(sucursal_id=n).first()
                if empleado_sede:
                    u = Usuario(username=username, rol='admin_local')
                    u.set_password(os.environ.get(f'ADMIN{n}_PASS', 'admin123'))
                    db.session.add(u)
                    db.session.flush()
                    empleado_sede.user_id = u.id
                    print(f'Admin de sucursal creado: {username} -> {empleado_sede.nombre}')

        db.session.commit()

        print('Seed completado.')

        # Asignar cedulas/ nombres reales desde la plantilla (idempotente)
        _cargar_cedulas_reales()

        # Limpiar empleados duplicados/nombres cortos sin cedula real (idempotente)
        _limpiar_empleados_duplicados()

        # Corregir vinculaciones: carolina -> Ana Carolina Sepulveda, admin1 -> regente
        _vincular_carolina_y_regente()

        # Subir la nómina real desde data/nomina.json
        _cargar_empleados_nomina()

        # Fusionar duplicados de la misma persona (nombres cortos del seed antiguo)
        _fusionar_duplicados_nomina()

        # Dejar SOLO los empleados de la nómina (elimina los que vienen de la plantilla)
        _depurar_solo_nomina()


if __name__ == '__main__':
    main()
