# -*- coding: utf-8 -*-
"""
Genera la plantilla de la CARTA AL PROVEEDOR (plantilla/CARTA_URGENTE.docx) a
partir de la carta real. Conserva todas las partes del .docx (encabezado con el
logo MINSA/HNAL, pie con la direccion del hospital, estilos, media) y solo
reemplaza word/document.xml por un cuerpo con marcadores de docxtemplater
({campo}), que se rellenan en el navegador.

Es la carta que se manda a la empresa ganadora para que adelante el producto
cuando la adquisicion es muy urgente y el area usuaria ya valido.

Ejecutar cuando cambie la estructura:
    python web/make_carta_template.py
"""

import re
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
EJEMPLO = BASE / "plantilla" / "CARTA_EJEMPLO.docx"
SALIDA = BASE / "plantilla" / "CARTA_URGENTE.docx"

XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
AZUL_CAB = "BDD6EE"          # sombreado de la cabecera de la tabla de items
ANCHO = 8504                 # ancho util de la pagina (A4 menos margenes)


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def run(text, font="Arial", sz=24, bold=False, underline=False, preserve=True):
    rpr = "<w:rPr>"
    rpr += '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (font, font, font)
    if bold:
        rpr += "<w:b/>"
    if underline:
        rpr += '<w:u w:val="single"/>'
    rpr += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    rpr += "</w:rPr>"
    sp = ' xml:space="preserve"' if preserve else ""
    return "<w:r>%s<w:t%s>%s</w:t></w:r>" % (rpr, sp, esc(text))


def salto():
    """Salto de linea dentro del mismo parrafo (RUC / direccion)."""
    return "<w:r><w:br/></w:r>"


def para(runs_xml, jc="both", font="Arial", sz=24, bold=False,
         spacing_after=0, spacing_before=0):
    ppr = "<w:pPr>"
    ppr += ('<w:spacing w:before="%d" w:after="%d" w:line="259" '
            'w:lineRule="auto"/>' % (spacing_before, spacing_after))
    if jc:
        ppr += '<w:jc w:val="%s"/>' % jc
    ppr += '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (font, font, font)
    if bold:
        ppr += "<w:b/>"
    ppr += '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:pPr>' % (sz, sz)
    return "<w:p>%s%s</w:p>" % (ppr, runs_xml)


def blank(sz=18):
    return para("", jc=None, sz=sz)


def tcell(width, paras_xml, valign="center", fill=None, gridspan=None):
    tcpr = '<w:tcPr><w:tcW w:w="%d" w:type="dxa"/>' % width
    if gridspan:
        tcpr += '<w:gridSpan w:val="%d"/>' % gridspan
    if fill:
        tcpr += '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % fill
    tcpr += '<w:vAlign w:val="%s"/></w:tcPr>' % valign
    return "<w:tc>%s%s</w:tc>" % (tcpr, paras_xml)


def bordes(visible):
    lado = ('<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            if visible else
            '<w:%s w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
    return "<w:tblBorders>%s</w:tblBorders>" % "".join(
        lado % l for l in ("top", "left", "bottom", "right", "insideH", "insideV"))


# ---------------- BLOQUE ASUNTO / REFERENCIA / FECHA ----------------

def tabla_asunto():
    W0, W1, W2 = 1507, 438, 6559

    def fila(label, valor_xml):
        return "<w:tr>%s%s%s</w:tr>" % (
            tcell(W0, para(run(label), jc="left"), valign="top"),
            tcell(W1, para(run(":"), jc="left"), valign="top"),
            tcell(W2, valor_xml, valign="top"))

    tblpr = ('<w:tblPr><w:tblW w:w="%d" w:type="dxa"/>%s</w:tblPr>'
             % (W0 + W1 + W2, bordes(False)))
    grid = ('<w:tblGrid><w:gridCol w:w="%d"/><w:gridCol w:w="%d"/>'
            '<w:gridCol w:w="%d"/></w:tblGrid>' % (W0, W1, W2))
    filas = (fila("Asunto", para(run("{asunto}"), jc="both"))
             + fila("Referencia", para(run("{referencia}"), jc="both"))
             + fila("Fecha", para(run("{fecha}"), jc="both")))
    return "<w:tbl>%s%s%s</w:tbl>" % (tblpr, grid, filas)


# ---------------- TABLA DE ITEMS ----------------

def tabla_items():
    # proporciones de la carta original, ajustadas para que entren en la hoja
    cols = [1689, 2071, 1231, 1006, 1253, 1254]
    hdrs = ["CÓDIGO SIGA", "DESCRIPCIÓN", "CANTIDAD", "UNIDAD DE MEDIDA",
            "PRECIO UNITARIO", "MONTO TOTAL"]
    campos = ["{codigo}", "{descripcion}", "{cantidad}", "{unidad}",
              "{precio}", "{monto}"]

    def celda(w, txt, sz, bold, fill=None, span=None, jc="center"):
        return tcell(w, para(run(txt, sz=sz, bold=bold), jc=jc, sz=sz, bold=bold),
                     fill=fill, gridspan=span)

    cab = "<w:tr>%s</w:tr>" % "".join(
        celda(cols[i], hdrs[i], 18, True, fill=AZUL_CAB) for i in range(6))
    # una fila por item (bucle de docxtemplater)
    loop = list(campos)
    loop[0] = "{#filas}" + campos[0]
    loop[5] = campos[5] + "{/filas}"
    datos = "<w:tr>%s</w:tr>" % "".join(
        celda(cols[i], loop[i], 16, False, jc="left" if i == 1 else "center")
        for i in range(6))
    # fila TOTAL: solo aparece cuando hay mas de un item
    total = "<w:tr>%s%s</w:tr>" % (
        celda(sum(cols[:5]), "{#hayTotal}TOTAL", 18, True, span=5),
        celda(cols[5], "{montoTotal}{/hayTotal}", 18, True))

    tblpr = ('<w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/>%s'
             '</w:tblPr>' % (sum(cols), bordes(True)))
    grid = "<w:tblGrid>%s</w:tblGrid>" % "".join(
        '<w:gridCol w:w="%d"/>' % w for w in cols)
    return "<w:tbl>%s%s%s%s%s</w:tbl>" % (tblpr, grid, cab, datos, total)


# ---------------- TEXTOS FIJOS ----------------

LEGAL = ("Mediante la presente le hago llegar mi saludo y a su vez, en atención "
         "a los documentos de referencia, siendo la mejor oferta validada por "
         "el área usuaria, y en virtud a la Constitución Política del Perú – "
         "Artículo 7°, la Ley General de Salud – Artículo 37°, el DL N° 1439 "
         "del Sistema Nacional de Abastecimiento - Artículo 2° numeral 4°, y el "
         "Principio de Eficacia y Eficiencia establecido en el artículo 2° "
         "literal f) del TUO de la Ley General de Contrataciones Públicas, se "
         "le solicita con carácter ")
COMPROMISO_1 = ("Asimismo, por lo expuesto nos comprometemos a atender lo "
                "requerido mediante ")
COMPROMISO_2 = (", la cual será enviada a su correo electrónico a la brevedad "
                "posible, siendo indispensable contar con el producto "
                "solicitado.")


def build_body():
    p = []
    # Titulo: el numero lo pone la secretaria a mano
    p.append(para(run("CARTA N°           - HNAL/OL-", sz=24, bold=True, underline=True)
                  + run("{anio}", sz=24, bold=True, underline=True)
                  + run(".", sz=24, bold=True, underline=True),
                  jc="both", sz=24, bold=True))
    p.append(blank())
    # Destinatario
    p.append(para(run("Señores:"), jc="left"))
    p.append(para(run("{empresa}", bold=True), jc="left", bold=True))
    p.append(para(run("RUC: ") + run("{ruc}") + salto() + run("{direccion}"),
                  jc="left"))
    p.append(para(run("{email}"), jc="left"))
    p.append(blank())
    p.append(para(run("Presente", bold=True) + run(". –"), jc="left"))
    p.append(blank())
    p.append(tabla_asunto())
    p.append(blank())
    p.append(para(run("_" * 63, sz=24), jc="both", sz=24))
    p.append(blank())
    p.append(para(run("De mi consideración:", sz=18), jc="both", sz=18))
    p.append(blank())
    p.append(para(run(LEGAL, sz=18)
                  + run("MUY URGENTE", sz=18, bold=True)
                  + run(":", sz=18), jc="both", sz=18))
    p.append(blank())
    p.append(para(run("Brindar en forma ", sz=18)
                  + run("INMEDIATA “Solicitud de adquisición “", sz=18, bold=True)
                  + run("{denominacion}", sz=18)
                  + run("”, conforme el siguiente detalle:", sz=18),
                  jc="both", sz=18))
    p.append(blank())
    p.append(tabla_items())
    p.append(blank())
    p.append(para(run(COMPROMISO_1, sz=18)
                  + run("orden de compra", sz=18, bold=True)
                  + run(COMPROMISO_2, sz=18), jc="both", sz=18))
    p.append(blank())
    p.append(para(run("Sin otro particular, quedo de usted.", sz=18),
                  jc="both", sz=18))
    p.append(para(run("Atentamente,", sz=18), jc="both", sz=18))
    # espacio para la firma y el sello
    for _ in range(4):
        p.append(blank())
    p.append(para(run("ELABORADO POR: ", sz=16, bold=True)
                  + run("{elaborado}", sz=16, bold=True), jc="both", sz=16, bold=True))
    p.append(para(run("Se adjunta copia de cotización", sz=16, bold=True),
                  jc="both", sz=16, bold=True))
    return "".join(p)


def main():
    src = zipfile.ZipFile(EJEMPLO, "r")
    orig = src.read("word/document.xml").decode("utf-8")
    ini = orig.find("<w:document")
    root_open = orig[ini:orig.find(">", ini) + 1]
    sectpr = re.search(r"<w:sectPr.*?</w:sectPr>", orig, re.S).group(0)
    doc = XMLDECL + root_open + "<w:body>" + build_body() + sectpr + "</w:body></w:document>"

    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(SALIDA, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "word/document.xml":
                data = doc.encode("utf-8")
            out.writestr(item, data)
    src.close()
    print("Generado", SALIDA, "(%d bytes)" % SALIDA.stat().st_size)


if __name__ == "__main__":
    main()
