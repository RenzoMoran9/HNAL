# -*- coding: utf-8 -*-
"""
Generador de Cuadros de Validacion de Especificaciones Tecnicas.

A partir de la plantilla base (plantilla/FORMATO_VALIDACIONES_DE_BIENES.xlsx)
rellena automaticamente los datos que ingresa el usuario, SIN alterar la
estructura ni el formato del modelo (fuentes, colores, bordes y logo).

Soporta:
  - Bienes ("...ADQUISICION DE ...") y servicios ("...TERMINO DE REFERENCIA
    PARA EL SERVICIO DE ...").
  - Varios postores (agrega/quita columnas segun cuantos cotizaron).
  - Varias filas / items (agrega/quita filas), cada uno con la marca de cada
    postor. La columna SI/NO se deja en blanco (la llena el area usuario).

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

# --------------------------------------------------------------------------- #
#  Configuracion / constantes de la plantilla
# --------------------------------------------------------------------------- #

BASE_DIR = Path(__file__).resolve().parent
PLANTILLA = BASE_DIR / "plantilla" / "FORMATO_VALIDACIONES_DE_BIENES.xlsx"

FILA_TITULO = 5
FILA_ITEM = 15           # primera fila de datos (item)
FOOT_TOP = 16            # primera fila del pie
FOOT_BOT = 46            # ultima fila usada en la plantilla

COL_PRIMER_POSTOR = 7    # G

ANCHO_MARCA = 50.0
ANCHO_SINO = 63.0

TITULOS = {
    "bien": "VALIDACION DE ESPECIFICACIONES TECNICAS PARA LA ADQUISICION DE ",
    "servicio": "VALIDACION DEL TERMINO DE REFERENCIA PARA EL SERVICIO DE ",
}


# --------------------------------------------------------------------------- #
#  Utilidades
# --------------------------------------------------------------------------- #

def _limpiar_celda(cell):
    cell.value = None
    cell.fill = PatternFill(fill_type=None)
    cell.border = Border()


def _merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _cant(valor):
    txt = str("" if valor is None else valor).strip()
    try:
        n = int(txt)
        if str(n) == txt:
            return n
    except (TypeError, ValueError):
        pass
    return txt or None


def _normalizar_items(datos, n_postores):
    """Devuelve la lista de items. Acepta el modelo nuevo (``items``) o el
    modelo antiguo (un solo item con datos en el nivel superior)."""
    items = datos.get("items")
    if items:
        norm = []
        for it in items:
            marcas = list(it.get("marcas") or [])
            marcas += [""] * (n_postores - len(marcas))
            norm.append({
                "codigo_siga": str(it.get("codigo_siga", "")).strip(),
                "denominacion": str(it.get("denominacion", "")).strip(),
                "unidad": (str(it.get("unidad", "UND")).strip().upper() or "UND"),
                "cantidad": it.get("cantidad", ""),
                "marcas": [str(m or "").strip() for m in marcas[:n_postores]],
            })
        return norm
    # modelo antiguo: 1 item, marca en cada postor
    marcas = [str((p or {}).get("marca", "") or "").strip()
              for p in (datos.get("postores") or [])]
    return [{
        "codigo_siga": str(datos.get("codigo_siga", "")).strip(),
        "denominacion": str(datos.get("denominacion", "")).strip(),
        "unidad": (str(datos.get("unidad", "UND")).strip().upper() or "UND"),
        "cantidad": datos.get("cantidad", ""),
        "marcas": marcas,
    }]


# --------------------------------------------------------------------------- #
#  Motor principal
# --------------------------------------------------------------------------- #

def generar(datos, salida, plantilla=PLANTILLA):
    """Genera un cuadro de validacion a partir de ``datos``.

    ``datos`` admite:
        tipo          "bien" | "servicio"  (por defecto "bien")
        adquisicion   texto que va despues del prefijo del titulo
        postores      lista de dicts con clave ``nombre``
        items         lista de dicts con: codigo_siga, denominacion, unidad,
                      cantidad, marcas (lista, una marca por postor).
    """
    postores = [p for p in (datos.get("postores") or [])
                if str(p.get("nombre", "")).strip()]
    n_p = len(postores)
    if n_p < 1:
        raise ValueError("Se requiere al menos un postor.")

    items = _normalizar_items(datos, n_p)
    n_i = len(items)

    first = COL_PRIMER_POSTOR
    ultima = 6 + 2 * n_p
    off = n_i - 1
    last_item = 14 + n_i
    wmax = max(ultima + 2, 12)

    def F(r):
        return r + off

    wb = load_workbook(plantilla)
    ws = wb.active

    # Prototipos de estilo
    st = {
        "cumple": copy(ws.cell(11, 7)._style),
        "postor_izq": copy(ws.cell(12, 9)._style),
        "postor_der": copy(ws.cell(12, 10)._style),
        "emp_izq": copy(ws.cell(13, 9)._style),
        "emp_der": copy(ws.cell(13, 10)._style),
        "marca_hdr": copy(ws.cell(14, 9)._style),
        "sino_hdr": copy(ws.cell(14, 10)._style),
        "marca_dato": copy(ws.cell(15, 9)._style),
        "sino_dato": copy(ws.cell(15, 10)._style),
        "num": copy(ws.cell(15, 2)._style),
        "codigo": copy(ws.cell(15, 3)._style),
        "denom": copy(ws.cell(15, 4)._style),
        "unidad": copy(ws.cell(15, 5)._style),
        "cant": copy(ws.cell(15, 6)._style),
    }
    altura_item = ws.row_dimensions[15].height

    # Capturar el PIE (filas 16..46)
    foot = []
    for r in range(FOOT_TOP, FOOT_BOT + 1):
        cells = []
        for c in range(1, wmax + 1):
            cell = ws.cell(r, c)
            cells.append((c, cell.value, copy(cell._style)))
        foot.append((r, ws.row_dimensions[r].height, cells))

    # Desarmar todas las combinaciones
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))

    # Limpiar valores de la zona dinamica
    for r in range(15, FOOT_BOT + off + 1):
        for c in range(1, wmax + 1):
            ws.cell(r, c).value = None

    # Reubicar el pie desplazado +off
    for (r, h, cells) in foot:
        nr = r + off
        if h:
            ws.row_dimensions[nr].height = h
        for (c, val, style) in cells:
            cell = ws.cell(nr, c)
            cell._style = copy(style)
            cell.value = val

    # Titulo
    tipo = "servicio" if datos.get("tipo") == "servicio" else "bien"
    ws.cell(FILA_TITULO, 2).value = (
        TITULOS[tipo] + str(datos.get("adquisicion", "")).strip().upper()
    )

    # Limpiar zona de postores (encabezado + filas de items)
    for f in range(11, last_item + 1):
        for c in range(7, 41):
            cell = ws.cell(f, c)
            if c > ultima:
                _limpiar_celda(cell)
            else:
                cell.value = None

    # Filas de items (columnas B..F)
    for k, it in enumerate(items):
        r = 15 + k
        if altura_item:
            ws.row_dimensions[r].height = altura_item
        ws.cell(r, 2)._style = copy(st["num"]); ws.cell(r, 2).value = k + 1
        ws.cell(r, 3)._style = copy(st["codigo"]); ws.cell(r, 3).value = it["codigo_siga"]
        ws.cell(r, 4)._style = copy(st["denom"]); ws.cell(r, 4).value = it["denominacion"]
        ws.cell(r, 5)._style = copy(st["unidad"]); ws.cell(r, 5).value = it["unidad"]
        ws.cell(r, 6)._style = copy(st["cant"]); ws.cell(r, 6).value = _cant(it["cantidad"])

    # Encabezado CUMPLE
    for c in range(first, ultima + 1):
        ws.cell(11, c)._style = copy(st["cumple"])
    ws.cell(11, first).value = "CUMPLE CON LAS ESPECIFICACIONES TECNICAS"
    _merge(ws, 11, first, 11, ultima)

    # Postores (encabezado) + marca/sino por item
    for i, p in enumerate(postores):
        cm = first + 2 * i
        cs = cm + 1
        ws.cell(12, cm)._style = copy(st["postor_izq"]); ws.cell(12, cs)._style = copy(st["postor_der"])
        ws.cell(12, cm).value = "POSTOR %d" % (i + 1)
        _merge(ws, 12, cm, 12, cs)
        ws.cell(13, cm)._style = copy(st["emp_izq"]); ws.cell(13, cs)._style = copy(st["emp_der"])
        ws.cell(13, cm).value = str(p.get("nombre", "")).strip()
        _merge(ws, 13, cm, 13, cs)
        ws.cell(14, cm)._style = copy(st["marca_hdr"]); ws.cell(14, cm).value = "MARCA"
        ws.cell(14, cs)._style = copy(st["sino_hdr"]); ws.cell(14, cs).value = "(colocar SI o NO)"
        ws.column_dimensions[get_column_letter(cm)].width = ANCHO_MARCA
        ws.column_dimensions[get_column_letter(cs)].width = ANCHO_SINO
        for k, it in enumerate(items):
            r = 15 + k
            ws.cell(r, cm)._style = copy(st["marca_dato"])
            ws.cell(r, cm).value = it["marcas"][i] if i < len(it["marcas"]) else ""
            ws.cell(r, cs)._style = copy(st["sino_dato"])

    # Combinaciones del encabezado
    _merge(ws, 6, 2, 6, 6); _merge(ws, 7, 2, 7, 6)
    for c in range(2, 7):
        _merge(ws, 11, c, 14, c)
    for f in (4, 5, 8, 10):
        _merge(ws, f, 2, f, ultima)

    # Combinaciones del pie (desplazadas)
    _merge(ws, F(16), 2, F(17), 3)
    _merge(ws, F(16), 4, F(17), 8)
    for f in (18, 19, 20, 22, 23, 24, 29):
        _merge(ws, F(f), 2, F(f), ultima)
    _merge(ws, F(21), 2, F(21), ultima + 1)
    _merge(ws, F(25), 2, F(25), 6)

    # Impresion: que entre a lo ancho
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.scale = None
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    if isinstance(salida, (str, Path)):
        salida = Path(salida)
        wb.save(salida)
    else:
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


def _si_no(texto):
    return _pedir(texto + " (s/n)", defecto="n").strip().lower().startswith("s")


def _nombre_archivo(adquisicion):
    base = "".join(ch if ch.isalnum() or ch in " -_" else "_"
                   for ch in str(adquisicion).upper()).strip().replace(" ", "_")
    base = base[:60] or "CUADRO"
    return "VALIDACION_%s.xlsx" % base


def main():
    print("=" * 60)
    print("  GENERADOR DE CUADROS DE VALIDACION (BIENES Y SERVICIOS)")
    print("=" * 60)

    tipo = "servicio" if _pedir("Tipo: 1) Bien  2) Servicio", defecto="1").strip() == "2" else "bien"
    ref = "SERVICIO DE" if tipo == "servicio" else "ADQUISICION DE"
    adquisicion = _pedir("Texto que va despues de '...%s'" % ref)

    n = _pedir_entero("Cuantos postores presentaron cotizacion")
    postores = []
    for i in range(1, n + 1):
        postores.append({"nombre": _pedir("  Razon social del POSTOR %d" % i)})

    items = []
    while True:
        idx = len(items) + 1
        print("\n--- BIEN/SERVICIO %d ---" % idx)
        codigo = _pedir("  Codigo SIGA", obligatorio=False, defecto="")
        denom = _pedir("  Denominacion")
        unidad = _pedir("  Unidad de medida", defecto="UND")
        cantidad = _pedir_entero("  Cantidad")
        marcas = []
        for i, p in enumerate(postores):
            marcas.append(_pedir("  Marca de %s" % p["nombre"], obligatorio=False, defecto=""))
        items.append({"codigo_siga": codigo, "denominacion": denom,
                      "unidad": unidad, "cantidad": cantidad, "marcas": marcas})
        if not _si_no("\n¿Agregar otro bien/servicio?"):
            break

    datos = {"tipo": tipo, "adquisicion": adquisicion,
             "postores": postores, "items": items}

    sugerido = _nombre_archivo(adquisicion)
    salida = _pedir("\nNombre del archivo de salida", defecto=sugerido)
    if not salida.lower().endswith(".xlsx"):
        salida += ".xlsx"

    ruta = generar(datos, salida)
    print("\n[OK] Cuadro generado: %s" % Path(ruta).resolve())


if __name__ == "__main__":
    main()
