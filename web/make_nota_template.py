# -*- coding: utf-8 -*-
"""
Genera la plantilla de la NOTA INFORMATIVA (plantilla/NOTA_INFORMATIVA.docx) a
partir del ejemplo real. Conserva todas las partes del .docx (encabezado con el
logo MINSA/HNAL, pie, estilos, media) y solo reemplaza word/document.xml por un
cuerpo con marcadores de docxtemplater ({campo}), que se rellenan en el
navegador.

Ejecutar cuando cambie la estructura:
    python web/make_nota_template.py
"""

import re
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
EJEMPLO = BASE / "plantilla" / "NOTA_INFORMATIVA_EJEMPLO.docx"
SALIDA = BASE / "plantilla" / "NOTA_INFORMATIVA.docx"

XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
AZUL = "0000FF"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def run(text, font="Arial Narrow", sz=21, bold=False, italic=False,
        underline=False, color=None, preserve=True):
    rpr = "<w:rPr>"
    if font:
        rpr += '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (font, font, font)
    if bold:
        rpr += "<w:b/>"
    if italic:
        rpr += "<w:i/>"
    if color:
        rpr += '<w:color w:val="%s"/>' % color
    if underline:
        rpr += '<w:u w:val="single"/>'
    rpr += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    rpr += "</w:rPr>"
    sp = ' xml:space="preserve"' if preserve else ""
    return "<w:r>%s<w:t%s>%s</w:t></w:r>" % (rpr, sp, esc(text))


FRAME_BOTTOM = ('<w:framePr w:w="9000" w:h="1200" w:hRule="atLeast" '
                'w:wrap="around" w:vAnchor="margin" w:hAnchor="margin" '
                'w:x="0" w:yAlign="bottom"/>')


def para(runs_xml, jc="both", font="Arial Narrow", sz=21, bold=False,
         border=False, spacing_after=None, frame=None):
    ppr = "<w:pPr>"
    if frame:
        ppr += frame
    if border:
        ppr += ('<w:pBdr><w:bottom w:val="single" w:sz="18" w:space="1" '
                'w:color="auto"/></w:pBdr>')
    after = 0 if spacing_after is None else spacing_after
    ppr += '<w:spacing w:after="%d" w:line="276" w:lineRule="auto"/>' % after
    if jc:
        ppr += '<w:jc w:val="%s"/>' % jc
    ppr += "<w:rPr>"
    if font:
        ppr += '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (font, font, font)
    if bold:
        ppr += "<w:b/>"
    ppr += '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:pPr>' % (sz, sz)
    return "<w:p>%s%s</w:p>" % (ppr, runs_xml)


def blank(sz=18):
    return para("", jc=None, sz=sz)


def tcell(width, paras_xml, gridspan=None, valign="center"):
    tcpr = '<w:tcPr><w:tcW w:w="%d" w:type="dxa"/>' % width
    if gridspan:
        tcpr += '<w:gridSpan w:val="%d"/>' % gridspan
    tcpr += '<w:vAlign w:val="%s"/></w:tcPr>' % valign
    return "<w:tc>%s%s</w:tc>" % (tcpr, paras_xml)


# ---------------- BLOQUE DESTINATARIO ----------------

def tabla_destinatario():
    W0, W1, W2 = 1651, 298, 7118
    GAP = 150

    def fila(label, cel2_paras):
        c0 = tcell(W0, para(run(label, sz=21), jc="left", sz=21))
        c1 = tcell(W1, para(run(":", sz=21), jc="left", sz=21))
        c2 = tcell(W2, cel2_paras)
        return "<w:tr>%s%s%s</w:tr>" % (c0, c1, c2)

    dest = (para(run("{destNombre}", sz=21, bold=True), jc="left", sz=21,
                 bold=True)
            + para(run("{destCargo}", sz=21), jc="left", sz=21,
                   spacing_after=GAP))
    rem = (para(run("{remNombre}", sz=21, bold=True), jc="left", sz=21,
                bold=True)
           + para(run("{remCargo}", sz=21), jc="left", sz=21,
                  spacing_after=GAP))
    asunto = para(
        run("Solicito Aprobación de Certificación de Crédito Presupuestario "
            "N°{asuntoNum}.", sz=21), jc="left", sz=21, spacing_after=GAP)
    ref = para(run("{referencia}", sz=21, color=AZUL), jc="left", sz=21,
               spacing_after=GAP)
    fecha = para(run("{fecha}", sz=21), jc="left", sz=21)

    tblpr = ('<w:tblPr><w:tblW w:w="9067" w:type="dxa"/><w:jc w:val="center"/>'
             '<w:tblBorders>'
             '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '</w:tblBorders></w:tblPr>')
    grid = ('<w:tblGrid><w:gridCol w:w="%d"/><w:gridCol w:w="%d"/>'
            '<w:gridCol w:w="%d"/></w:tblGrid>' % (W0, W1, W2))
    filas = (fila("A", dest) + fila("DE", rem) + fila("ASUNTO", asunto)
             + fila("REFERENCIA", ref) + fila("FECHA", fecha))
    return "<w:tbl>%s%s%s</w:tbl>" % (tblpr, grid, filas)


# ---------------- TABLA DE DATOS (CCP) ----------------

def tabla_datos():
    cols = [900, 820, 2990, 680, 1097, 709, 1588]
    hdrs = ["CCMN", "CCP SIAF", "Descripción", "Meta", "Especifica de Gasto",
            "F.F.", "Monto"]
    campos = ["{ccmn}", "{ccpSiaf}", "{descripcion}", "{meta}",
              "{especifica}", "{ff}", "{monto}"]

    def celda(w, txt, sz=20, span=None):
        return tcell(w, para(run(txt, sz=sz, bold=True), jc="center", sz=sz,
                             bold=True), gridspan=span)

    header = "<w:tr>%s</w:tr>" % "".join(
        celda(cols[i], hdrs[i]) for i in range(7))
    # Fila de datos: bucle de docxtemplater (una fila por item)
    campos_loop = list(campos)
    campos_loop[0] = "{#filas}" + campos[0]
    campos_loop[6] = campos[6] + "{/filas}"
    datos = "<w:tr>%s</w:tr>" % "".join(
        celda(cols[i], campos_loop[i]) for i in range(7))
    # Fila TOTAL: "TOTAL" abarca las 6 primeras columnas + Monto total
    total = "<w:tr>%s%s</w:tr>" % (
        celda(sum(cols[:6]), "TOTAL", sz=18, span=6),
        celda(cols[6], "{montoTotal}"))

    tblpr = ('<w:tblPr><w:tblW w:w="8784" w:type="dxa"/><w:jc w:val="center"/>'
             '<w:tblBorders>'
             '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '</w:tblBorders></w:tblPr>')
    grid = "<w:tblGrid>%s</w:tblGrid>" % "".join(
        '<w:gridCol w:w="%d"/>' % w for w in cols)
    return "<w:tbl>%s%s%s%s%s</w:tbl>" % (tblpr, grid, header, datos, total)


CTX1 = ("Tengo el agrado de dirigirme a usted para saludarlo cordialmente y a "
        "través del presente, solicitarle la aprobación de Certificación de "
        "Crédito Presupuestal, a fin de proseguir con el trámite "
        "correspondiente, según el siguiente detalle:")
CTX2 = ("En tal sentido, solicito a su despacho la aprobación de la "
        "Certificación de Crédito Presupuestario, en concordancia con los "
        "numerales 41.1 y 41.2 del artículo 41 del DECRETO LEGISLATIVO Nº 1440 "
        "DEL SISTEMA NACIONAL DE PRESUPUESTO PÚBLICO, se adjunta el expediente "
        "completo, para su atención correspondiente dentro de sus "
        "competencias.")


def build_body():
    parts = []
    # Titulo (sin numero: se agrega a mano)
    parts.append(para(
        run("INFORMATIVA Nº            - OL-J-HNAL-", font=None, sz=24,
            bold=True, underline=True)
        + run("{anio}", font=None, sz=24, bold=True, underline=True),
        jc="center", font=None, sz=24, bold=True))
    parts.append(blank())
    parts.append(tabla_destinatario())
    parts.append(blank())
    parts.append(para("", jc=None, border=True))
    parts.append(blank())
    parts.append(para(run(CTX1, sz=21), jc="both", sz=21))
    parts.append(blank())
    parts.append(tabla_datos())
    parts.append(blank())
    parts.append(para(run(CTX2, sz=21), jc="both", sz=21))
    parts.append(blank())
    parts.append(para(run("Sin otro particular, quedo de Usted.", sz=21),
                      jc="both", sz=21))
    parts.append(blank())
    parts.append(para(run("Atentamente,", sz=21), jc="left", sz=21))
    parts.append(blank())
    # Pie anclado al fondo (Arial 8, etiquetas en negrita) como el memo
    def pie_para(runs_xml):
        return para(runs_xml, jc="left", font="Arial", sz=16, frame=FRAME_BOTTOM)

    parts.append(pie_para(run("EXP. DIR.: {expDir}", font="Arial", sz=16,
                              bold=True)))
    parts.append(pie_para(run("EXP. LOG. {expLog}", font="Arial", sz=16,
                              bold=True)))
    parts.append(pie_para(
        run("Folios: ", font="Arial", sz=16)
        + run("(        {folios}        )", font="Arial", sz=16, bold=True)
        + run(" original", font="Arial", sz=16)))
    parts.append(pie_para(run("C.c.  Archivo", font="Arial", sz=16)))
    parts.append(pie_para(
        run("Elaborado por: ", font="Arial", sz=16, bold=True)
        + run("{elaborado}", font="Arial", sz=16)))
    parts.append(pie_para(run("{iniciales}", font="Arial", sz=16)))
    return "".join(parts)


def main():
    src = zipfile.ZipFile(EJEMPLO, "r")
    orig = src.read("word/document.xml").decode("utf-8")
    root_open = orig[orig.find("<w:document"):orig.find(">", orig.find("<w:document")) + 1]
    sectpr = re.search(r"<w:sectPr.*?</w:sectPr>", orig, re.S).group(0)
    body = "<w:body>" + build_body() + sectpr + "</w:body>"
    doc = XMLDECL + root_open + body + "</w:document>"

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
