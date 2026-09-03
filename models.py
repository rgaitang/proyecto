from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class Sucursal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), unique=True, nullable=False)
    empleados = db.relationship('Empleado', backref='sucursal_ref', lazy=True)


class Empleado(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cedula = db.Column(db.String(20), unique=True, nullable=False)
    nombre = db.Column(db.String(120), nullable=False)
    cargo = db.Column(db.String(80), default='')
    sucursal_id = db.Column(db.Integer, db.ForeignKey('sucursal.id'))

    # Relación con usuario de login (uno a uno)
    user_id = db.Column(db.Integer, db.ForeignKey('usuario.id'))
    usuario = db.relationship('Usuario', uselist=False, back_populates='empleado')

    registros = db.relationship('RegistroHoras', backref='empleado_ref', lazy=True, cascade="all, delete-orphan")


class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    # Roles: 'admin_global' (Carolina), 'admin_local', 'empleado'
    rol = db.Column(db.String(20), nullable=False, default='empleado')
    empleado = db.relationship('Empleado', uselist=False, back_populates='usuario')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Turno(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(10), unique=True, nullable=False)
    descripcion = db.Column(db.String(200), default='')
    es_domingo = db.Column(db.Boolean, default=False)


class RegistroHoras(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    empleado_id = db.Column(db.Integer, db.ForeignKey('empleado.id'), nullable=False)
    fecha = db.Column(db.Date, nullable=False)
    turno_codigo = db.Column(db.String(10), nullable=False)
    observacion = db.Column(db.String(300), default='')
    # Estados: 'pendiente', 'aprobado_local', 'aprobado', 'rechazado'
    estado = db.Column(db.String(20), default='pendiente')
    creado_por = db.Column(db.Integer, db.ForeignKey('usuario.id'))

    __table_args__ = (db.UniqueConstraint('empleado_id', 'fecha', name='uniq_emp_fecha'),)


class Novedad(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    empleado_id = db.Column(db.Integer, db.ForeignKey('empleado.id'), nullable=False)
    tipo = db.Column(db.String(50), nullable=False)  # INCAPACIDAD, VACACIONES, CAMBIO_TURNO, RETIRO, INGRESO, OTRO
    fecha_inicio = db.Column(db.Date, nullable=False)
    fecha_fin = db.Column(db.Date)
    descripcion = db.Column(db.String(300), default='')
    reporta = db.Column(db.String(120), default='')
    estado = db.Column(db.String(20), default='pendiente')
    empleado_ref = db.relationship('Empleado', foreign_keys=[empleado_id])
