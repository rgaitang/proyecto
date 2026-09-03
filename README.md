# Sistema de Horarios SOLOFARMA

Aplicación web en Python (Flask) para que los empleados reporten sus turnos de forma individual, **sin poder modificar los horarios de otros**, y que **genera automáticamente el archivo Excel en el formato que se reporta a SIIGO**.

## Problema resuelto

Antes: un archivo Excel en Google Drive donde cualquier empleado con acceso podía editar los horarios de todos.

Ahora: cada empleado solo ve y edita **sus propios turnos**. Los administradores locales aprueban por sucursal y la coordinadora de RRHH genera el Excel final.

## Roles

| Rol | Quién | Permisos |
|-----|-------|----------|
| `admin_global` | Carolina Sepúlveda (RRHH) | Todo: ver todas las sucursales, crear empleados/usuarios, generar Excel de cualquier sucursal |
| `admin_local` | 1 por sucursal (6) | Ve y edita solo SU sucursal, genera Excel de su sucursal |
| `empleado` | Cada trabajador | Solo ve/reporta sus propios turnos |

## Estructura

```
sistema_horarios/
├── app.py              # Rutas, lógica y control de permisos
├── models.py           # Modelos: Sucursal, Empleado, Usuario, Turno, RegistroHoras, Novedad
├── excel_generator.py  # Genera el Excel con formato SIIGO (plantillas por quincena + novedades)
├── seed.py             # Puebla la BD con sucursales, turnos y empleados
├── run.py              # Arranca la aplicación
├── templates/          # Vistas HTML
├── output/             # Excels generados
└── instance/           # Base de datos SQLite (horarios.db)
```

## Instalación y arranque

```powershell
# 1. Instalar dependencias
pip install flask flask-sqlalchemy openpyxl werkzeug

# 2. Inicializar la base de datos (con los datos de los Excel)
python seed.py

# 3. Arrancar
python run.py
# Abre http://127.0.0.1:5000
```

## Usuario inicial

- **Usuario:** `carolina`
- **Contraseña:** `admin123`
- **Rol:** admin_global (RRHH)

> Cambia esta contraseña apenas entres. Desde el panel global puedes crear usuarios para empleados y administradores locales.

## Flujo de trabajo

1. **Empleado** inicia sesión → ve su calendario mensual → selecciona el código de su turno de cada día (dropdown con todos los turnos válidos) → solo edita sus propias celdas.
2. **Admin local** revisa y ajusta los turnos de los empleados de su sucursal.
3. **Admin global (Carolina)** selecciona sucursal/mes y pulsa **"Generar Excel SIIGO"**.
4. El archivo se descarga con el mismo formato de hojas (`MES 1-15`, `MES 16-31`, `NOVEDADES MES`) que usan para reportar a SIIGO.

## Seguridad

- Contraseñas con hash (werkzeug `generate_password_hash`).
- Permisos por decorador: un empleado jamás accede a `/panel-local`, `/panel-global` o `/generar_excel` (403).
- Un admin local solo edita los registros de **su** sucursal (validación en servidor).
- Un empleado solo puede guardar turnos en fechas del período vigente (no en períodos cerrados antiguos).

## Nota sobre los datos semilla

El `seed.py` crea los 6 puntos de venta y la tabla de turnos completa con sus descripciones (extraídas del archivo `HORARIOS PUNTOS DE VENTA SOLOFARMA.xlsx`). La asignación de cada empleado a su sucursal se hace de forma rotativa/estimada en el seed; **debes revisar y asignar correctamente cada empleado a su sucursal** desde la base de datos o ajustando el seed con los datos reales de tu nómina.
