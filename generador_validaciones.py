# -*- coding: utf-8 -*-
"""
Generador de Cuadros de Validacion de Especificaciones Tecnicas (Bienes).

A partir de la plantilla base (plantilla/FORMATO_VALIDACIONES_DE_BIENES.xlsx)
rellena automaticamente los datos que ingresa el usuario y ajusta las columnas
de postores segun cuantos hayan presentado su cotizacion, SIN alterar la
estructura ni el formato del modelo (fuentes, colores, bordes y logo).

Uso interactivo:
    python generador_validaciones.py

Uso como libreria:
    from generador_validaciones import generar
    generar(datos, salida="mi_cuadro.xlsx")
"""

from copy import copy
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# --------------------------------------------------------------------------- #
#  Configuracion / constantes de la plantilla
# --------------------------------------------------------------------------- #

BASE_DIR = Path(__file__).resolve().parent
PLANTILLA = BASE_DIR / "plantilla" / "FORMATO_VALIDACIONES_DE_BIENES.xlsx"

# Filas clave del cuadro
FILA_TITULO = 5          # "VALIDACION ... PARA LA ADQUISICION DE ..."
FILA_ACTA = 8            # "Siendo las ___ horas del ___ ..."
FILA_HDR_CUMPLE = 11     # "CUMPLE CON LAS ESPECIFICACIONES TECNICAS"
FILA_HDR_POSTOR = 12     # "POSTOR N"
FILA_HDR_EMPRESA = 13    # razon social de la empresa
FILA_HDR_MARCA = 14      # "MARCA" / "(colocar SI o NO)"
FILA_DATOS = 15          # datos del bien + marca/cumple por postor

# Columnas fijas (la parte izquierda no cambia)
COL_NUM = 2              # B  -> Nº
COL_CODIGO = 3          # C  -> CODIGO SIGA
COL_DENOM = 4           # D  -> DENOMINACION
COL_UNIDAD = 5          # E  -> UNIDAD DE MEDIDA
COL_CANT = 6            # F  -> CANT
COL_PRIMER_POSTOR = 7   # G  -> primera columna de postores

# Anchos uniformes para las columnas de cada postor
ANCHO_MARCA = 50.0
ANCHO_SINO = 20.0

TITULO_BASE = "VALIDACION DE ESPECIFICACIONES TECNICAS PARA LA ADQUISICION DE "

# Rangos combinados que dependen del ancho de la tabla (se rehacen segun postores)
_MERGES_ANCHO_COMPLETO = [
    (4, 4), (5, 5), (8, 8), (10, 10),
    (18, 18), (19, 19), (20, 20), (22, 22), (23, 23), (24, 24), (29, 29),
]
# Estas lineas de firma se extienden una columna mas (como en el modelo original)
_MERGES_ANCHO_MAS_UNO = [(21, 21)]

# Rangos combinados de la zona de postores que se rehacen en cada generacion
_MERGES_POSTOR_BASE = [
    "G11:J11", "G12:H12", "I12:J12", "G13:H13", "I13:J13",
]


# --------------------------------------------------------------------------- #
#  Utilidades
# --------------------------------------------------------------------------- #

def _unmerge_si_existe(ws, rango):
    if rango in {str(r) for r in ws.merged_cells.ranges}:
        ws.unmerge_cells(rango)


def _limpiar_celda(cell):
    """Deja una celda sin valor, sin relleno y sin bordes."""
    cell.value = None
    cell.fill = PatternFill(fill_type=None)
    cell.border = Border()


def _aplicar_estilo(ws, fila, col, estilo, valor=None):
    c = ws.cell(row=fila, column=col)
    c._style = copy(estilo)
    if valor is not None:
        c.value = valor
    return c


# --------------------------------------------------------------------------- #
#  Motor principal
# --------------------------------------------------------------------------- #

def generar(datos, salida, plantilla=PLANTILLA):
    """Genera un cuadro de validacion a partir de ``datos``.

    ``datos`` es un dict con las claves:
        adquisicion   (str)  -> texto que va despues de "... ADQUISICION DE"
        codigo_siga   (str)
        denominacion  (str)
        unidad        (str)  -> por defecto "UND"
        cantidad      (int/str)
        postores      (list de dicts) cada uno con:
                          nombre (str)  -> razon social
                          marca  (str)
                          cumple (str)  -> "SI" / "NO" / "" (opcional)

    Devuelve la ruta del archivo generado.
    """
    postores = datos.get("postores") or []
    n = len(postores)
    if n < 1:
        raise ValueError("Se requiere al menos un postor.")

    wb = load_workbook(plantilla)
    ws = wb.active

    ultima_col = COL_PRIMER_POSTOR + 2 * n - 1  # ultima columna de postores

    # 1) Titulo -----------------------------------------------------------
    ws.cell(row=FILA_TITULO, column=COL_NUM).value = (
        TITULO_BASE + str(datos.get("adquisicion", "")).strip().upper()
    )

    # 1b) Acta (hora / fecha / lugar) opcional ---------------------------
    acta = datos.get("acta")
    if acta:
        ws.cell(row=FILA_ACTA, column=COL_NUM).value = acta

    # 2) Datos del bien ---------------------------------------------------
    ws.cell(row=FILA_DATOS, column=COL_CODIGO).value = str(datos.get("codigo_siga", "")).strip()
    ws.cell(row=FILA_DATOS, column=COL_DENOM).value = str(datos.get("denominacion", "")).strip()
    ws.cell(row=FILA_DATOS, column=COL_UNIDAD).value = (
        str(datos.get("unidad", "UND")).strip().upper() or "UND"
    )
    cant = datos.get("cantidad", "")
    try:
        cant = int(str(cant).strip())
    except (TypeError, ValueError):
        pass
    ws.cell(row=FILA_DATOS, column=COL_CANT).value = cant

    # 3) Capturar estilos prototipo de la zona de postores ----------------
    #    (columnas I/J = "postor generico" del modelo base)
    est = {
        "cumple": copy(ws.cell(row=FILA_HDR_CUMPLE, column=7)._style),   # G11
        "postor_izq": copy(ws.cell(row=FILA_HDR_POSTOR, column=9)._style),   # I12
        "postor_der": copy(ws.cell(row=FILA_HDR_POSTOR, column=10)._style),  # J12
        "empresa_izq": copy(ws.cell(row=FILA_HDR_EMPRESA, column=9)._style),  # I13
        "empresa_der": copy(ws.cell(row=FILA_HDR_EMPRESA, column=10)._style), # J13
        "marca_hdr": copy(ws.cell(row=FILA_HDR_MARCA, column=9)._style),   # I14 MARCA
        "sino_hdr": copy(ws.cell(row=FILA_HDR_MARCA, column=10)._style),   # J14 (SI o NO)
        "marca_dato": copy(ws.cell(row=FILA_DATOS, column=9)._style),      # I15
        "sino_dato": copy(ws.cell(row=FILA_DATOS, column=10)._style),      # J15
    }

    # 4) Desarmar combinaciones que vamos a rehacer -----------------------
    for rango in _MERGES_POSTOR_BASE:
        _unmerge_si_existe(ws, rango)
    for fila, _ in _MERGES_ANCHO_COMPLETO + _MERGES_ANCHO_MAS_UNO:
        # se desarman por su definicion original en la plantilla base
        pass
    # combinaciones de ancho completo del modelo base (2 postores)
    for rango in ["B4:J4", "B5:J5", "B8:H8", "B10:J10", "B18:J18",
                  "B19:H19", "B20:H20", "B21:K21", "B22:J22", "B23:J23",
                  "B24:J24", "B29:H29"]:
        _unmerge_si_existe(ws, rango)

    # 5) Limpiar toda la zona de postores (por si sobran columnas) --------
    MAX_LIMPIAR = 40
    for fila in range(FILA_HDR_CUMPLE, FILA_DATOS + 1):
        for col in range(COL_PRIMER_POSTOR, MAX_LIMPIAR + 1):
            _limpiar_celda(ws.cell(row=fila, column=col))
    # quitar anchos sobrantes de columnas de postores que ya no existen
    for col in range(ultima_col + 1, MAX_LIMPIAR + 1):
        letra = get_column_letter(col)
        if letra in ws.column_dimensions:
            del ws.column_dimensions[letra]

    # 6) Reconstruir encabezado "CUMPLE ..." ------------------------------
    for col in range(COL_PRIMER_POSTOR, ultima_col + 1):
        _aplicar_estilo(ws, FILA_HDR_CUMPLE, col, est["cumple"])
    ws.cell(row=FILA_HDR_CUMPLE, column=COL_PRIMER_POSTOR).value = \
        "CUMPLE CON LAS ESPECIFICACIONES TECNICAS"
    ws.merge_cells(start_row=FILA_HDR_CUMPLE, start_column=COL_PRIMER_POSTOR,
                   end_row=FILA_HDR_CUMPLE, end_column=ultima_col)

    # 7) Reconstruir cada postor -----------------------------------------
    for i, postor in enumerate(postores):
        c_marca = COL_PRIMER_POSTOR + 2 * i
        c_sino = c_marca + 1

        # Fila 12: "POSTOR N"
        _aplicar_estilo(ws, FILA_HDR_POSTOR, c_marca, est["postor_izq"],
                        valor="POSTOR %d" % (i + 1))
        _aplicar_estilo(ws, FILA_HDR_POSTOR, c_sino, est["postor_der"])
        ws.merge_cells(start_row=FILA_HDR_POSTOR, start_column=c_marca,
                       end_row=FILA_HDR_POSTOR, end_column=c_sino)

        # Fila 13: razon social
        _aplicar_estilo(ws, FILA_HDR_EMPRESA, c_marca, est["empresa_izq"],
                        valor=str(postor.get("nombre", "")).strip())
        _aplicar_estilo(ws, FILA_HDR_EMPRESA, c_sino, est["empresa_der"])
        ws.merge_cells(start_row=FILA_HDR_EMPRESA, start_column=c_marca,
                       end_row=FILA_HDR_EMPRESA, end_column=c_sino)

        # Fila 14: "MARCA" / "(colocar SI o NO)"
        _aplicar_estilo(ws, FILA_HDR_MARCA, c_marca, est["marca_hdr"], valor="MARCA")
        _aplicar_estilo(ws, FILA_HDR_MARCA, c_sino, est["sino_hdr"],
                        valor="(colocar SI o NO)")

        # Fila 15: marca y cumple
        _aplicar_estilo(ws, FILA_DATOS, c_marca, est["marca_dato"],
                        valor=str(postor.get("marca", "")).strip())
        cumple = str(postor.get("cumple", "")).strip().upper()
        _aplicar_estilo(ws, FILA_DATOS, c_sino, est["sino_dato"],
                        valor=cumple if cumple else None)

        # Anchos de columna
        ws.column_dimensions[get_column_letter(c_marca)].width = ANCHO_MARCA
        ws.column_dimensions[get_column_letter(c_sino)].width = ANCHO_SINO

    # 8) Rehacer combinaciones de ancho completo (textos y firmas) --------
    for fila, _ in _MERGES_ANCHO_COMPLETO:
        ws.merge_cells(start_row=fila, start_column=COL_NUM,
                       end_row=fila, end_column=ultima_col)
    for fila, _ in _MERGES_ANCHO_MAS_UNO:
        ws.merge_cells(start_row=fila, start_column=COL_NUM,
                       end_row=fila, end_column=ultima_col + 1)

    # 9) Ajuste de impresion: que siempre entre a lo ancho ----------------
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.scale = None
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    salida = Path(salida)
    wb.save(salida)
    return salida


# --------------------------------------------------------------------------- #
#  Interfaz interactiva (preguntas por consola)
# --------------------------------------------------------------------------- #

def _pedir(texto, obligatorio=True, defecto=None):
    while True:
        sufijo = " [%s]" % defecto if defecto is not None else ""
        val = input("%s%s: " % (texto, sufijo)).strip()
        if not val and defecto is not None:
            return defecto
        if val or not obligatorio:
            return val
        print("  -> Este dato es obligatorio, intenta de nuevo.")


def _pedir_entero(texto, defecto=None):
    while True:
        val = _pedir(texto, obligatorio=defecto is None, defecto=defecto)
        try:
            return int(str(val).strip())
        except (TypeError, ValueError):
            print("  -> Ingresa un numero valido.")


def _nombre_archivo(adquisicion):
    base = "".join(ch if ch.isalnum() or ch in " -_" else "_"
                   for ch in str(adquisicion).upper()).strip().replace(" ", "_")
    base = base[:60] or "CUADRO"
    return "VALIDACION_%s.xlsx" % base


def main():
    print("=" * 60)
    print("  GENERADOR DE CUADROS DE VALIDACION DE BIENES")
    print("=" * 60)
    print("Responde las preguntas. (Enter para aceptar el valor por defecto)\n")

    adquisicion = _pedir("Adquisicion (va despues de '...ADQUISICION DE')")
    codigo_siga = _pedir("Codigo SIGA")
    denominacion = _pedir("Denominacion del bien")
    unidad = _pedir("Unidad de medida", defecto="UND")
    cantidad = _pedir_entero("Cantidad")

    n = _pedir_entero("Cuantos postores presentaron cotizacion")
    postores = []
    for i in range(1, n + 1):
        print("\n--- POSTOR %d ---" % i)
        nombre = _pedir("  Razon social / nombre del postor")
        marca = _pedir("  Marca ofertada", obligatorio=False, defecto="")
        cumple = _pedir("  Cumple? (SI/NO, Enter = dejar en blanco)",
                        obligatorio=False, defecto="")
        postores.append({"nombre": nombre, "marca": marca, "cumple": cumple})

    datos = {
        "adquisicion": adquisicion,
        "codigo_siga": codigo_siga,
        "denominacion": denominacion,
        "unidad": unidad,
        "cantidad": cantidad,
        "postores": postores,
    }

    sugerido = _nombre_archivo(adquisicion)
    salida = _pedir("\nNombre del archivo de salida", defecto=sugerido)
    if not salida.lower().endswith(".xlsx"):
        salida += ".xlsx"

    ruta = generar(datos, salida)
    print("\n[OK] Cuadro generado: %s" % Path(ruta).resolve())


if __name__ == "__main__":
    main()
