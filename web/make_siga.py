# -*- coding: utf-8 -*-
"""
Arma el catalogo SIGA del hospital (web/siga.json) a partir de los
"REPORTE DE OC" que se bajan del SIGA (una hoja, una fila por item comprado).

    python web/make_siga.py                      # lee todos los .xlsx de datos/
    python web/make_siga.py otro_reporte.xlsx    # o los archivos que se indiquen

No es el catalogo nacional completo: son los items que el hospital realmente
compro, con su codigo SIGA, su nombre oficial, su unidad de medida y el precio
de cada compra. Eso es lo que sirve para autocompletar los cuadros y para
responder "a cuanto lo compramos la vez pasada".

Cuando salga el reporte de un ano nuevo, se deja en datos/ y se vuelve a correr
esto y despues:  python web/build_html.py
"""

import json
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
DATOS = BASE / "datos"
SALIDA = BASE / "web" / "siga.json"

# Columnas del reporte (el SIGA siempre las exporta en este orden)
C_FECHA, C_PROV, C_CODIGO, C_NOMBRE = 1, 2, 9, 10
C_CANT, C_PRECIO, C_UNIDAD = 12, 13, 18
CABECERA = ("nro_orden", "fecha_orden", "nombre_prov")


def limpiar(v):
    """Quita espacios, comillas sueltas y dobles espacios."""
    t = re.sub(r"\s+", " ", str(v or "")).strip()
    return t.strip('"').strip("'").strip()


def sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def fecha_num(v):
    """'20/08/2026 00:00:00' o un datetime -> 20260820 (entero, ordenable)."""
    if isinstance(v, datetime):
        return v.year * 10000 + v.month * 100 + v.day
    m = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})", str(v or ""))
    if not m:
        return 0
    d, mes, a = (int(x) for x in m.groups())
    return a * 10000 + mes * 100 + d


def numero(v):
    try:
        return round(float(v), 6)
    except (TypeError, ValueError):
        return 0.0


def leer(rutas):
    items = {}          # codigo -> {nombre, unidad, compras}
    provs, idx_prov = [], {}
    saltadas = 0

    for ruta in rutas:
        wb = load_workbook(ruta, read_only=True, data_only=True)
        ws = wb.active
        filas = ws.iter_rows(min_row=1, values_only=True)
        cab = next(filas, None)
        if not cab or tuple(str(c or "").strip() for c in cab[:3]) != CABECERA:
            print("  ! %s no parece un REPORTE DE OC, se omite" % Path(ruta).name)
            wb.close()
            continue

        n = 0
        for r in filas:
            codigo = limpiar(r[C_CODIGO] if len(r) > C_CODIGO else "")
            nombre = limpiar(r[C_NOMBRE] if len(r) > C_NOMBRE else "")
            if not re.fullmatch(r"\d{12}", codigo) or not nombre:
                saltadas += 1
                continue

            it = items.setdefault(codigo, {"nombre": nombre,
                                           "unidad": limpiar(r[C_UNIDAD]),
                                           "compras": []})
            # el nombre mas largo suele ser el completo (los cortos vienen recortados)
            if len(nombre) > len(it["nombre"]):
                it["nombre"] = nombre

            prov = limpiar(r[C_PROV])
            if prov not in idx_prov:
                idx_prov[prov] = len(provs)
                provs.append(prov)

            it["compras"].append([fecha_num(r[C_FECHA]), numero(r[C_PRECIO]),
                                  idx_prov[prov], numero(r[C_CANT])])
            n += 1
        wb.close()
        print("  %-28s %5d compras" % (Path(ruta).name, n))

    if saltadas:
        print("  (%d filas sin codigo SIGA valido, ignoradas)" % saltadas)
    return items, provs


def main():
    rutas = [Path(a) for a in sys.argv[1:]]
    if not rutas:
        rutas = sorted(DATOS.glob("*.xlsx")) + sorted(DATOS.glob("*.XLSX"))
    if not rutas:
        print("No hay reportes. Deja los .xlsx del SIGA en la carpeta datos/")
        return 1

    items, provs = leer(rutas)
    if not items:
        print("No se pudo leer ningun item.")
        return 1

    lista = []
    for codigo in sorted(items):
        it = items[codigo]
        # la compra mas reciente primero: es la que la app muestra como referencia
        compras = sorted(it["compras"], key=lambda c: -c[0])
        # se busca sin tildes, asi que el nombre ya va normalizado para buscar
        lista.append([codigo, it["nombre"], it["unidad"], compras])

    data = {
        "v": 1,
        "gen": datetime.now().strftime("%Y-%m-%d"),
        "fuente": [Path(r).name for r in rutas],
        "prov": provs,
        "items": lista,
    }
    SALIDA.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")),
                      encoding="utf-8")

    compras = sum(len(i[3]) for i in lista)
    kb = SALIDA.stat().st_size / 1024
    print("\n%d items, %d compras, %d proveedores -> %s (%.0f KB)"
          % (len(lista), compras, len(provs), SALIDA.name, kb))
    print("Ahora corre:  python web/build_html.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
