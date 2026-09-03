"""Exporta los empleados de los Excel a un archivo JSON (data/empleados.json)
para poder cargarlos en la base de datos de produccion sin depender de los Excel.

Uso:  python exportar_empleados.py
"""
import os, json, openpyxl

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
EXCEL_SOLOFARMA = r'C:\Users\ronal\OneDrive\Escritorio\Documentos\Default Project\horarios nomina\HORARIOS PUNTOS DE VENTA SOLOFARMA.xlsx'
EXCEL_NIQUIA = r'C:\Users\ronal\OneDrive\Escritorio\Documentos\Default Project\horarios nomina\horarios niquia.xlsx'

os.makedirs(BASE, exist_ok=True)


def extraer_solofarma():
    wb = openpyxl.load_workbook(EXCEL_SOLOFARMA, data_only=True)
    ws = wb['DATOS']
    empleados = []
    nombres_vistos = set()
    for row in ws.iter_rows(min_row=4, max_row=60):
        nombre = row[2].value  # C
        if not nombre:
            continue
        nombre = str(nombre).strip().upper()
        if not nombre or nombre in ('TOTAL', 'EMPLEADO') or nombre in nombres_vistos:
            continue
        nombres_vistos.add(nombre)
        empleados.append({'nombre': nombre.title(), 'cargo': '', 'sucursal_id': None})
    return empleados


def extraer_niquia():
    wb = openpyxl.load_workbook(EXCEL_NIQUIA)
    ws = wb['JULIO 1-15']
    empleados = []
    nombres_vistos = set()
    for row in ws.iter_rows(min_row=5, max_row=25, min_col=1, max_col=1):
        val = row[0].value
        if not val:
            continue
        nombre = str(val).strip().upper()
        if not nombre or nombre in ('TOTAL', 'EMPLEADOS', 'FECHA') or nombre in nombres_vistos:
            continue
        nombres_vistos.add(nombre)
        empleados.append({'nombre': nombre.title(), 'cargo': '', 'sucursal_id': 1})
    return empleados


def main():
    lista = extraer_solofarma() + extraer_niquia()

    # Quitar duplicados entre ambos
    nombres = set()
    unicos = []
    for e in lista:
        if e['nombre'] not in nombres:
            nombres.add(e['nombre'])
            unicos.append(e)

    # Asignar sucursal de forma rotativa a los que no la tienen (supuesto, revisar!)
    idx = 0
    for e in unicos:
        if e['sucursal_id'] is None:
            e['sucursal_id'] = (idx % 6) + 1
            idx += 1

    ruta = os.path.join(BASE, 'empleados.json')
    with open(ruta, 'w', encoding='utf-8') as f:
        json.dump(unicos, f, ensure_ascii=False, indent=2)

    print(f'Exportados {len(unicos)} empleados a {ruta}')


if __name__ == '__main__':
    main()
