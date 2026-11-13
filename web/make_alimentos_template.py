# -*- coding: utf-8 -*-
"""
Arma la plantilla del ACTA DE VALIDACION DE ALIMENTOS
(plantilla/ACTA_VALIDACION_ALIMENTOS.xlsx) a partir del acta real que envio
el usuario (plantilla/ACTA_ALIMENTOS_EJEMPLO.xlsx).

El acta real es una hoja de trabajo: lo que se imprime es J1:V33, pero a los
costados y debajo tiene datos de otros expedientes. Aqui se deja solo el acta
(con su logo, letras, bordes y alturas) y sin los datos del ejemplo; la app la
rellena en el navegador (generarActaAlimentos en web/shell.html).

Ejecutar si cambia el acta de ejemplo, y luego:  python web/build_html.py
"""

import re
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill

BASE = Path(__file__).resolve().parent.parent
ORIGEN = BASE / "plantilla" / "ACTA_ALIMENTOS_EJEMPLO.xlsx"
SALIDA = BASE / "plantilla" / "ACTA_VALIDACION_ALIMENTOS.xlsx"

PRIMERA_COL, ULTIMA_COL = 10, 22      # J .. V: el acta impresa
ULTIMA_FILA = 34


def limpiar(cell):
    cell.value = None
    cell.font = Font()
    cell.fill = PatternFill()
    cell.border = Border()
    cell.alignment = Alignment()
    cell.number_format = "General"


def main():
    wb = load_workbook(ORIGEN, rich_text=True)
    ws = wb.worksheets[0]
    ws.title = "ACTA"

    # fuera del acta: combinaciones, valores y formatos de otros expedientes
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col < PRIMERA_COL or rng.min_col > ULTIMA_COL:
            ws.unmerge_cells(str(rng))
    if ws.max_row > ULTIMA_FILA:
        ws.delete_rows(ULTIMA_FILA + 1, ws.max_row - ULTIMA_FILA)
    # delete_rows no quita las alturas: sin esto la hoja "llega" hasta la fila 3572
    for r in [r for r in ws.row_dimensions if r > ULTIMA_FILA]:
        del ws.row_dimensions[r]
    ws.print_title_rows = None            # eran filas de otro expediente
    for row in ws.iter_rows(min_row=1, max_row=ULTIMA_FILA):
        for c in row:
            if c.column < PRIMERA_COL or c.column > ULTIMA_COL:
                limpiar(c)
    for letra in "ABCDEFGHI":
        ws.column_dimensions[letra].hidden = True
    ws.conditional_formatting = type(ws.conditional_formatting)()

    # datos del ejemplo dentro del acta: los pone la app
    for ref in ("L4", "T4", "T7", "J16", "S19", "U19", "L28"):
        ws[ref].value = None
    for r in (21, 22):
        for col in range(PRIMERA_COL, ULTIMA_COL + 1):
            ws.cell(r, col).value = None

    # textos del acta con errores de tipeo del ejemplo
    ws["J8"].value = ("en calidad de representante del área usuaria, procedió a la "
                      "revisión de las cotizaciones para verificar el cumplimiento de las ")
    ws["J25"].value = "según la cotización)"
    # resaltados amarillos de trabajo
    for ref in ("J24", "J26"):
        ws[ref].fill = PatternFill()

    wb.save(SALIDA)
    arreglar_logo()
    print("Generado %s" % SALIDA)


def arreglar_logo():
    """openpyxl guarda el logo con rutas absolutas y sin prefijos, y ExcelJS
    (con el que la app abre la plantilla) no lo reconoce. Se deja el dibujo tal
    como venía en el acta original, con rutas relativas."""
    src = zipfile.ZipFile(ORIGEN)
    tmp = SALIDA.with_suffix(".tmp")
    with zipfile.ZipFile(SALIDA) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in ("xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels"):
                data = src.read(item.filename)
            elif item.filename == "xl/worksheets/_rels/sheet1.xml.rels":
                data = data.replace(b'Target="/xl/drawings/', b'Target="../drawings/')
            elif item.filename == "xl/worksheets/sheet1.xml":
                # openpyxl siempre pone <outlinePr>; ExcelJS lo reescribe después de
                # <pageSetUpPr> y Excel abre el acta como archivo dañado
                data = re.sub(rb"<outlinePr[^>]*/>", b"", data)
            zout.writestr(item, data)
    tmp.replace(SALIDA)


if __name__ == "__main__":
    main()
