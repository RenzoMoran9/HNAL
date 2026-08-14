# -*- coding: utf-8 -*-
"""
Lee memos / notas ya emitidos (.docx) y saca de ellos el titulo profesional
(Q.F., ING., DR., LIC., C.P.C., ...) con el que se nombra a cada funcionario,
para guardarlo en web/directorio.json (campo "t" de cada persona).

    python web/aplicar_titulos.py documento1.docx documento2.docx ...

Solo toca a las personas que YA estan en el directorio: si alguien aparece en
un memo viejo pero ya no trabaja en el hospital, se informa y se ignora.
Cuando una persona aparece con dos titulos distintos gana el mas frecuente.

Despues de correrlo hay que reconstruir la app:  python web/build_html.py
"""

import collections
import json
import re
import sys
import unicodedata
from pathlib import Path

from docx import Document

BASE = Path(__file__).resolve().parent.parent
DIRECTORIO = BASE / "web" / "directorio.json"

# Los titulos se escriben de mil formas ("Q. F. SONIA", "ING.JORGE", "Lic. ARTURO"),
# asi que el punto tambien sirve de separador. Exigir punto o espacio despues del
# titulo evita confundir un nombre como "DRAGO" con "DR".
PATRON = (r"(Q\.?\s?F|C\.?P\.?C|ABOG|ING|LIC|DRA|DR|MG"
          r"|SRA|SR|M\.?C|OBST|ARQ|ECON|BLGO|T\.?M|ENF)(?:\.\s*|\s+)")
CANON = {
    "QF": "Q.F.", "CPC": "C.P.C.", "ABOG": "ABOG.", "ING": "ING.", "LIC": "LIC.",
    "DRA": "DRA.", "DR": "DR.", "MG": "MG.", "SRA": "SRA.", "SR": "SR.",
    "MC": "M.C.", "OBST": "OBST.", "ARQ": "ARQ.", "ECON": "ECON.",
    "BLGO": "BLGO.", "TM": "T.M.", "ENF": "ENF.",
}


def sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def palabras(nombre):
    """Nombre -> lista de palabras significativas, sin tildes ni iniciales sueltas."""
    t = re.sub(r"[^A-Z ]", " ", sin_tildes(nombre).upper()).split()
    return [w for w in t if len(w) > 2]


def claves(nombre):
    """(nombre + 2 apellidos) y (2 apellidos), para cotejar escrituras distintas."""
    p = palabras(nombre)
    if len(p) < 2:
        return None, None
    apellidos = (p[-2], p[-1]) if len(p) > 2 else (p[-1], p[-1])
    return (p[0],) + apellidos, apellidos


def leer_titulos(rutas):
    """Recorre los encabezados 'A:' / 'DE:' de cada .docx y cuenta los titulos."""
    cuenta = collections.defaultdict(collections.Counter)
    for ruta in rutas:
        doc = Document(ruta)
        for tabla in doc.tables:
            for fila in tabla.rows:
                celdas = [c.text.strip() for c in fila.cells]
                if len(celdas) < 3 or celdas[0].upper().strip() not in ("A", "DE", "PARA"):
                    continue
                partes = [p.strip() for p in celdas[2].split("\n") if p.strip()]
                if not partes:
                    continue
                crudo = sin_tildes(partes[0])
                m = re.match(r"^\s*" + PATRON, crudo, re.I)
                if not m:
                    continue
                titulo = CANON.get(re.sub(r"[^A-Za-z]", "", m.group(1)).upper())
                if not titulo:
                    continue
                cuenta[crudo[m.end():].strip()][titulo] += 1
    return cuenta


def main():
    rutas = sys.argv[1:]
    if not rutas:
        print(__doc__)
        return 1

    crudos = leer_titulos(rutas)
    por_completo = collections.defaultdict(collections.Counter)
    por_apellidos = collections.defaultdict(collections.Counter)
    for nombre, cont in crudos.items():
        completo, apellidos = claves(nombre)
        if not completo:
            continue
        por_completo[completo].update(cont)
        por_apellidos[apellidos].update(cont)

    personas = json.loads(DIRECTORIO.read_text(encoding="utf-8"))
    # Los apellidos solo sirven de respaldo si identifican a una sola persona.
    repetidos = collections.Counter(claves(p["n"])[1] for p in personas)

    puestos, encontrados = [], set()
    for persona in personas:
        completo, apellidos = claves(persona["n"])
        if not completo:
            continue
        cont = por_completo.get(completo)
        if not cont and repetidos[apellidos] == 1:
            cont = por_apellidos.get(apellidos)
        if not cont:
            continue
        titulo = cont.most_common(1)[0][0]
        encontrados.add(apellidos)
        if persona.get("t") != titulo:
            puestos.append((persona["n"], persona.get("t"), titulo))
        persona["t"] = titulo

    sobrantes = [(n, dict(c)) for n, c in crudos.items()
                 if claves(n)[1] and claves(n)[1] not in encontrados]

    for nombre, antes, ahora in puestos:
        print("  %-42s %s%s" % (nombre[:42], ahora, "" if antes is None else "  (antes %s)" % antes))
    if sobrantes:
        print("\nAparecen en los documentos pero no estan en el directorio "
              "(seguramente ya no trabajan aqui):")
        for nombre, cont in sobrantes:
            print("  %-42s %s" % (nombre[:42], cont))

    con_titulo = sum(1 for p in personas if p.get("t"))
    DIRECTORIO.write_text(json.dumps(personas, ensure_ascii=False, indent=1),
                          encoding="utf-8")
    print("\n%d cambios. %d de %d personas tienen titulo." %
          (len(puestos), con_titulo, len(personas)))
    print("Ahora corre:  python web/build_html.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
