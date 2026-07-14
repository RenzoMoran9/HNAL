# -*- coding: utf-8 -*-
"""
Construye el archivo HTML autonomo (docs/index.html) incrustando:
  - la libreria ExcelJS (web/vendor/exceljs.min.js)
  - la plantilla base en base64 (plantilla/FORMATO_VALIDACIONES_DE_BIENES.xlsx)
dentro de la plantilla web/shell.html.

Ejecutar cada vez que cambie la plantilla o el formulario:
    python web/build_html.py
"""

import base64
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
SHELL = BASE / "web" / "shell.html"
EXCELJS = BASE / "web" / "vendor" / "exceljs.min.js"
PLANTILLA = BASE / "plantilla" / "FORMATO_VALIDACIONES_DE_BIENES.xlsx"
PLANTILLA_CC = BASE / "plantilla" / "CUADRO_COMPARATIVO.xlsx"
SALIDA = BASE / "docs" / "index.html"


def main():
    shell = SHELL.read_text(encoding="utf-8")
    exceljs = EXCELJS.read_text(encoding="utf-8")
    b64 = base64.b64encode(PLANTILLA.read_bytes()).decode("ascii")
    b64cc = base64.b64encode(PLANTILLA_CC.read_bytes()).decode("ascii")

    html = shell.replace("/*__EXCELJS__*/", exceljs)
    html = html.replace("__TEMPLATE_B64__", b64)
    html = html.replace("__TEMPLATE_CC_B64__", b64cc)

    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    SALIDA.write_text(html, encoding="utf-8")
    kb = len(html.encode("utf-8")) / 1024
    print("Generado %s (%.0f KB)" % (SALIDA, kb))


if __name__ == "__main__":
    main()
