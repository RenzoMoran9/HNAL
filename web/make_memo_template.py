# -*- coding: utf-8 -*-
"""
Genera la plantilla del MEMO (plantilla/MEMO_VALIDACION.docx) a partir del
ejemplo real que envio el usuario. Conserva TODAS las partes del .docx original
(encabezado con el logo MINSA/HNAL, estilos, fuentes, relaciones, media) y solo
reemplaza word/document.xml por un cuerpo de UN memo con marcadores de
docxtemplater ({campo}). En tiempo de ejecucion, docxtemplater rellena esos
marcadores en el navegador.

Ejecutar cuando cambie la estructura del memo:
    python web/make_memo_template.py
"""

import re
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
EJEMPLO = BASE / "plantilla" / "EJEMPLO_DE_MEMOS_DE_VALIDACION.docx"
SALIDA = BASE / "plantilla" / "MEMO_VALIDACION.docx"

# Tamanos en medios-puntos (Word): 7pt=14, 8pt=16, 10pt=20, 12pt=24, 16pt=32
XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'


def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def run(text, font="Arial", sz=14, bold=False, underline=False, preserve=True):
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


# Marco que ancla el bloque al fondo del area de texto (justo sobre el pie de
# pagina). Todos los parrafos que lo comparten se agrupan en un solo marco.
FRAME_BOTTOM = ('<w:framePr w:w="9000" w:h="1400" w:hRule="atLeast" '
                'w:wrap="around" w:vAnchor="margin" w:hAnchor="margin" '
                'w:x="0" w:yAlign="bottom"/>')


def para(runs_xml, jc="both", font="Arial", sz=14, bold=False, border=False,
         spacing_after=None, frame=None):
    ppr = "<w:pPr>"
    if frame:
        ppr += frame
    if border:
        ppr += ('<w:pBdr><w:bottom w:val="single" w:sz="18" w:space="1" '
                'w:color="auto"/></w:pBdr>')
    # Sin espaciado extra (como el estilo "Sin espaciado" del original), para
    # que el memo mantenga la densidad y quepa en una pagina.
    after = 0 if spacing_after is None else spacing_after
    ppr += '<w:spacing w:after="%d" w:line="240" w:lineRule="auto"/>' % after
    if jc:
        ppr += '<w:jc w:val="%s"/>' % jc
    ppr += "<w:rPr>"
    ppr += '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (font, font, font)
    if bold:
        ppr += "<w:b/>"
    ppr += '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:pPr>' % (sz, sz)
    return "<w:p>%s%s</w:p>" % (ppr, runs_xml)


def blank(sz=14):
    return para("", jc=None, sz=sz)


def tcell(width, paras_xml, shd=None, valign="center"):
    tcpr = '<w:tcPr><w:tcW w:w="%d" w:type="dxa"/>' % width
    if shd:
        tcpr += '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shd
    tcpr += '<w:vAlign w:val="%s"/></w:tcPr>' % valign
    return "<w:tc>%s%s</w:tc>" % (tcpr, paras_xml)


# ----------------- CUERPO DEL MEMO -----------------

def cell_para(text_or_runs, jc="left", font="Arial", sz=20, bold=False,
              frame=None):
    if isinstance(text_or_runs, str):
        rx = run(text_or_runs, font=font, sz=sz, bold=bold)
    else:
        rx = text_or_runs
    return para(rx, jc=jc, font=font, sz=sz, bold=bold, frame=frame)


def tabla_destinatario():
    W0, W1, W2 = 1124, 418, 6962

    def fila(label, cel2_paras):
        c0 = tcell(W0, cell_para(label, font="Arial Narrow", sz=20))
        c1 = tcell(W1, cell_para(":", font="Arial Narrow", sz=20))
        c2 = tcell(W2, cel2_paras)
        return "<w:tr>%s%s%s</w:tr>" % (c0, c1, c2)

    dest = (para(run("{destNombre}", font="Arial Narrow", sz=24, bold=True),
                 jc="left", font="Arial Narrow", sz=24, bold=True)
            + para(run("{destCargo}", font="Arial Narrow", sz=20),
                   jc="left", font="Arial Narrow", sz=20))
    rem = (para(run("{remNombre}", font="Arial", sz=20, bold=True),
                jc="left", font="Arial", sz=20, bold=True)
           + para(run("{remCargo}", font="Arial", sz=20),
                  jc="left", font="Arial", sz=20))
    asunto = para(
        run("Solicitud de Revisión y Evaluación de Cumplimiento de ",
            font="Arial Narrow", sz=20)
        + run("{asunto}", font="Arial Narrow", sz=20, bold=True),
        jc="left", font="Arial Narrow", sz=20)
    ref = para(run("{referencia}", font="Arial Narrow", sz=20, bold=True),
               jc="left", font="Arial Narrow", sz=20)
    fecha = cell_para("{fecha}", font="Arial Narrow", sz=20)

    tblpr = ('<w:tblPr><w:tblW w:w="0" w:type="auto"/>'
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
    filas = (fila("A", dest) + fila("DE", rem) + fila("Asunto", asunto)
             + fila("Referencia", ref) + fila("Fecha", fecha))
    return "<w:tbl>%s%s%s</w:tbl>" % (tblpr, grid, filas)


def tabla_postores():
    cols = [601, 3204, 1807, 3324]

    def hcell(w, txt):
        return tcell(w, cell_para(txt, jc="center", font="Arial", sz=16,
                                  bold=True), shd="A6A6A6")

    def dcell(w, txt):
        return tcell(w, cell_para(txt, jc="center", font="Arial", sz=16))

    header = "<w:tr>%s%s%s%s</w:tr>" % (
        hcell(cols[0], "N°"), hcell(cols[1], "EMPRESA"),
        hcell(cols[2], "RUC"), hcell(cols[3], "OBSERVACION"))
    # Fila plantilla con bucle de docxtemplater sobre {#filas}...{/filas}
    datos = "<w:tr>%s%s%s%s</w:tr>" % (
        dcell(cols[0], "{#filas}{num}"), dcell(cols[1], "{empresa}"),
        dcell(cols[2], "{ruc}"), dcell(cols[3], "{obs}{/filas}"))

    tblpr = ('<w:tblPr><w:tblW w:w="0" w:type="auto"/>'
             '<w:jc w:val="center"/>'
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
    return "<w:tbl>%s%s%s%s</w:tbl>" % (tblpr, grid, header, datos)


CONTEXTO_PRE = ("Por medio del presente me dirijo a usted, para saludarla "
                "cordialmente, y en atención al documento de la referencia, "
                "mediante el cual se solicita el ")
CONTEXTO_MID = " al rubro en materia de convocatoria, obteniendo "
CONTEXTO_MID2 = " cotizaciones a solicitud de "
CONTEXTO_END = ", según detalle:"

SOBRE = ("Sobre el particular se hace de conocimiento, que se ha enviado a "
         "diferentes empresas de los cuales se obtuvo cotizaciones, se solicita "
         "si está de acuerdo con las especificaciones técnicas")
ACUERDO = ("De acuerdo al Numeral 44.6 del Artículo 44° del Reglamento de la "
           "Ley General de Contrataciones públicas, señala lo siguiente: “El "
           "requerimiento no incluye exigencias desproporcionadas e innecesarias "
           "que limiten la concurrencia o favorezcan a determinado proveedor ni "
           "hace referencia a procedencia , fabricante , marca, patente, origen "
           "o tipos de producción, ni descripción que oriente la contratación "
           "hacía ellos, salvo que la autoridad de la gestión administrativa "
           "haya aprobado el correspondiente proceso de compatibilización del "
           "requerimiento, conforme a las disposiciones que establezca la DGA "
           "mediante directiva”")
ASIMISMO = ("Así mismo, el área usuaria tiene la responsabilidad técnica en la "
            "elaboración del requerimiento, precisos y objetivos mediante una "
            "participación activa en la evaluación de las cotizaciones asegurando "
            "que estas cumplan con las especificaciones técnicas, conforme al "
            "numeral 14.2 del Artículo 14º y Articulo 20 del Reglamento de la Ley "
            "de Contrataciones Públicas.")
DELOANTERIOR = ("De lo anterior, en caso la cotización no cumpla con los "
                "términos de referencia y/o Especificaciones Técnicas requeridos, "
                "deberá indicar el motivo de manera clara y objetiva considerando "
                "lo consignado en el Numeral 46.4 del Artículo 46 de la LGCP, así "
                "como los principios que rigen la contratación pública contenidos "
                "en el Artículo 5 de la Ley General de las Contrataciones "
                "Públicas.")
ENTALSENTIDO = ("En tal sentido, se corre el traslado del expediente en original "
                "de la referencia con la finalidad de que, a través de su "
                "despacho, se sirva evaluar de acuerdo con las Especificaciones "
                "Técnicas, a fin de continuar con los trámites correspondientes y "
                "ejecución del ejercicio presupuestal de acuerdo a Ley.")


def build_body():
    parts = []
    # Titulo
    parts.append(para(
        run("MEMORANDO N° ", font="Cambria", sz=32, bold=True, underline=True)
        + run("{memoNum}", font="Cambria", sz=32, bold=True, underline=True)
        + run(" -OL-J- H.N.A.L.-", font="Cambria", sz=32, bold=True,
              underline=True)
        + run("{anio}", font="Cambria", sz=32, bold=True, underline=True),
        jc="center", font="Cambria", sz=32, bold=True))
    # Tabla destinatario
    parts.append(tabla_destinatario())
    parts.append(blank())
    # Linea separadora
    parts.append(para("", jc=None, border=True))
    parts.append(blank())
    # Contexto
    ctx = (run(CONTEXTO_PRE, sz=16) + run("“", sz=16)
           + run("{titulo}", sz=16, bold=True) + run("”", sz=16)
           + run(CONTEXTO_MID, sz=16)
           + run("{cantidad}", sz=16, bold=True)
           + run(CONTEXTO_MID2, sz=16)
           + run("{solicitud}", sz=16)
           + run(CONTEXTO_END, sz=16))
    parts.append(para(ctx, jc="both", sz=16))
    parts.append(blank())
    # Tabla de postores
    parts.append(tabla_postores())
    parts.append(blank())
    # Parrafos legales
    for txt in (SOBRE, ACUERDO, ASIMISMO, DELOANTERIOR, ENTALSENTIDO):
        parts.append(para(run(txt, sz=14), jc="both", sz=14))
        parts.append(blank())
    # Atentamente + espacio para sello/firma
    parts.append(para(run("Atentamente,", sz=14), jc="left", sz=14))
    for _ in range(3):
        parts.append(blank())
    # Pie: bloque anclado al fondo de la pagina (framePr compartido). Siempre
    # sale pegado al pie, sin importar cuanto texto haya arriba.
    pie = [
        "EXP. DIR.: {expDir}",
        "EXP. LOG. {expLog}",
        "Folios: ( {folios} ) original",
        "C.c.  Archivo",
        "Elaborado por: {elaborado}",
        "{iniciales}",
    ]
    for linea in pie:
        parts.append(cell_para(linea, font="Arial Narrow", sz=20,
                               frame=FRAME_BOTTOM))
    return "".join(parts)


def main():
    src = zipfile.ZipFile(EJEMPLO, "r")
    orig = src.read("word/document.xml").decode("utf-8")

    # Raiz <w:document ...> con todos los namespaces del original
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
