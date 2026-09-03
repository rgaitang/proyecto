import calendar
from datetime import date, datetime, time
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


def _estilizar_hoja_horario(ws, n_cols):
    bold = Font(bold=True, size=11)
    header_fill = PatternFill('solid', fgColor='D9E1F2')
    thin = Side(style='thin', color='999999')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r in range(1, 4):
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            if r == 2:
                cell.font = bold
                cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')


def generar_archivo_siigo(OUTPUT_DIR, nombre_mes, anio, mes, sucursal_nombre,
                          empleados, registros, novedades, turnos_map):
    """
    Genera un libro Excel con el formato usado para reportar novedades a SIIGO.

    - registros: dict {empleado_id: {dia_int: RegistroHoras}}
    - turnos_map: {codigo: descripcion}
    """
    wb = Workbook()
    wb.remove(wb.active)

    quincenas = [(1, 1, 15), (2, 16, calendar.monthrange(anio, mes)[1])]

    for quincena, d_inicio, d_fin in quincenas:
        # Si la quincena no tiene días del mes (ej. siempre tiene al menos 15) se crea
        fechas = [date(anio, mes, d) for d in range(d_inicio, d_fin + 1)]
        if not fechas:
            continue

        nombre_hoja = f'{nombre_mes} {d_inicio}-{d_fin}'
        ws = wb.create_sheet(nombre_hoja)

        # Fila 1: título
        ws.cell(row=1, column=1, value=f'HORARIO {sucursal_nombre}')
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fechas) + 1)

        # Fila 2: FECHA + fechas
        ws.cell(row=2, column=1, value='FECHA')
        for i, f in enumerate(fechas):
            ws.cell(row=2, column=2 + i, value=f)

        # Fila 3: EMPLEADOS + días de la semana
        ws.cell(row=3, column=1, value='EMPLEADOS')
        dias_semana = ['LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO']
        for i, f in enumerate(fechas):
            ws.cell(row=3, column=2 + i, value=dias_semana[f.weekday()])

        # Filas de empleados
        fila = 4
        for emp in empleados:
            ws.cell(row=fila, column=1, value=emp.nombre)
            regs_emp = registros.get(emp.id, {})
            for dia in range(d_inicio, d_fin + 1):
                reg = regs_emp.get(dia)
                if reg:
                    ws.cell(row=fila, column=1 + dia).value = reg.turno_codigo
                    if reg.observacion:
                        ws.cell(row=fila, column=1 + dia).value = reg.turno_codigo
            fila += 1

        # Nota de días especiales
        ws.cell(row=fila + 3, column=1, value='Dias Dominicales y festivos')
        ws.cell(row=fila + 4, column=1, value='Dias ordinarios')
        ws.cell(row=fila + 5, column=1, value='Descansos')
        ws.cell(row=fila + 6, column=1, value='Novedad')

        _estilizar_hoja_horario(ws, len(fechas) + 1)

        # Ancho de columnas
        ws.column_dimensions['A'].width = 20
        for i in range(1, len(fechas) + 1):
            ws.column_dimensions[get_column_letter(1 + i)].width = 14

    # Hoja de Novedades
    if novedades:
        ws_nov = wb.create_sheet(f'NOVEDADES {nombre_mes}')
        ws_nov.cell(row=1, column=1, value=f'NOVEDADES DE {nombre_mes} {anio}')
        ws_nov.cell(row=3, column=1, value='INGRESO')
        headers_ingreso = ['CEDULA', 'NOMBRE', 'FECHA', 'CARGO', 'REPORTA', 'OBSERVACIONES']
        for i, h in enumerate(headers_ingreso):
            ws_nov.cell(row=4, column=2 + i, value=h)

        fila = 5
        num_ingreso = 1
        for n in [x for x in novedades if x.tipo in ('INGRESO', 'CAMBIO_TURNO')]:
            emp = next((e for e in empleados if e.id == n.empleado_id), None)
            if not emp:
                continue
            ws_nov.cell(row=fila, column=1, value=float(num_ingreso))
            ws_nov.cell(row=fila, column=2, value=emp.cedula)
            ws_nov.cell(row=fila, column=3, value=emp.nombre)
            ws_nov.cell(row=fila, column=4, value=n.fecha_inicio)
            ws_nov.cell(row=fila, column=5, value=emp.cargo)
            ws_nov.cell(row=fila, column=6, value=n.reporta or '')
            ws_nov.cell(row=fila, column=7, value=n.descripcion)
            fila += 1
            num_ingreso += 1

        ws_nov.cell(row=fila + 1, column=1, value='RETIROS')
        headers_retiro = ['CEDULA', 'NOMBRE', 'EPS', 'PENSION', 'FECHA RETIRO', 'DEPENDENCIA', 'CARGO']
        for i, h in enumerate(headers_retiro):
            ws_nov.cell(row=fila + 2, column=1 + i, value=h)

        fila_retiro = fila + 3
        num_retiro = 1
        for n in [x for x in novedades if x.tipo == 'RETIRO']:
            emp = next((e for e in empleados if e.id == n.empleado_id), None)
            if not emp:
                continue
            ws_nov.cell(row=fila_retiro, column=1, value=float(num_retiro))
            ws_nov.cell(row=fila_retiro, column=2, value=emp.cedula)
            ws_nov.cell(row=fila_retiro, column=3, value=emp.nombre)
            ws_nov.cell(row=fila_retiro, column=6, value=n.fecha_inicio)
            ws_nov.cell(row=fila_retiro, column=7, value='NIQUIA' if 'Niquia' in sucursal_nombre else sucursal_nombre)
            ws_nov.cell(row=fila_retiro, column=8, value=emp.cargo)
            fila_retiro += 1
            num_retiro += 1

    nombre_archivo = f'HORARIOS_{sucursal_nombre.replace(" ", "_")}_{nombre_mes}_{anio}.xlsx'
    ruta = f'{OUTPUT_DIR}/{nombre_archivo}'
    wb.save(ruta)
    return ruta
