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


if __name__ == '__main__':
    main()
