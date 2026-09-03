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


if __name__ == '__main__':
    main()
