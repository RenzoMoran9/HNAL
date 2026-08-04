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


def run(text, font="Arial", sz=14, bold=False, italic=False, underline=False,
        color=None, preserve=True):
    rpr = "<w:rPr>"
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


AZUL = "0000FF"


# Marco que ancla el bloque al fondo del area de texto (justo sobre el pie de
# pagina). Todos los parrafos que lo comparten se agrupan en un solo marco.
FRAME_BOTTOM = ('<w:framePr w:w="9000" w:h="1200" w:hRule="atLeast" '
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
    # Interlineado 1.15x (line=276) como el memo hecho a mano, para que no
    # salga "pegado". El espacio despues (after) se usa entre las filas del
    # bloque destinatario.
    after = 0 if spacing_after is None else spacing_after
    ppr += '<w:spacing w:after="%d" w:line="276" w:lineRule="auto"/>' % after
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
    # valign=None: no se escribe la etiqueta (Word alinea arriba por defecto)
    if valign:
        tcpr += '<w:vAlign w:val="%s"/>' % valign
    tcpr += '</w:tcPr>'
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
    # Anchos del memo manual: la columna de etiquetas debe caber "Referencia"
    # en Arial sin partirse en dos lineas.
    W0, W1, W2 = 1440, 280, 6780

    # Todo el encabezado va en Arial y alineado arriba, como el memo manual:
    # asi la etiqueta (A / DE) queda pareja con la primera linea del valor.
    def fila(label, cel2_paras):
        c0 = tcell(W0, cell_para(label, font="Arial", sz=20), valign=None)
        c1 = tcell(W1, cell_para(":", font="Arial", sz=20), valign=None)
        c2 = tcell(W2, cel2_paras, valign=None)
        return "<w:tr>%s%s%s</w:tr>" % (c0, c1, c2)

    GAP = 150   # espacio despues de cada fila (como el memo manual)
    # El CARGO va sin negrita (solo el nombre es negrita), igual en "A" y en "DE"
    dest = (para(run("{destNombre}", font="Arial", sz=21, bold=True),
                 jc="left", font="Arial", sz=21, bold=True)
            + para(run("{destCargo}", font="Arial", sz=21),
                   jc="left", font="Arial", sz=21,
                   spacing_after=GAP))
    rem = (para(run("{remNombre}", font="Arial", sz=21, bold=True),
                jc="left", font="Arial", sz=21, bold=True)
           + para(run("{remCargo}", font="Arial", sz=20),
                  jc="left", font="Arial", sz=20, spacing_after=GAP))
    asunto = para(
        run("Solicitud de Revisión y Evaluación de Cumplimiento de "
            "{asunto}", font="Arial", sz=20),
        jc="left", font="Arial", sz=20, spacing_after=GAP)
    ref = para(run("{referencia}", font="Arial", sz=20, bold=True),
               jc="left", font="Arial", sz=20, spacing_after=GAP)
    fecha = cell_para("{fecha}", font="Arial", sz=20)

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
        return tcell(w, cell_para(txt, jc="center", font="Arial", sz=16,
                                  bold=True))

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

# Partes con formato mixto (negrita / cursiva / subrayado) como el memo manual
ACUERDO_CITA1 = ("El requerimiento no incluye exigencias desproporcionadas e "
                 "innecesarias que limiten la concurrencia o favorezcan a "
                 "determinado proveedor ni hace referencia a procedencia , "
                 "fabricante , marca, patente, origen o tipos de producción, ni "
                 "descripción que oriente la contratación hacía ellos,")
ACUERDO_CITA2 = (" salvo que la autoridad de la gestión administrativa haya "
                 "aprobado el correspondiente proceso de compatibilización del "
                 "requerimiento, conforme a las disposiciones que establezca la "
                 "DGA mediante directiva”")
DELO_1 = ("De lo anterior, en caso la cotización no cumpla con los términos de "
          "referencia y/o Especificaciones Técnicas requeridos, ")
DELO_3 = (" considerando lo consignado en el Numeral 46.4 del Artículo 46 de la "
          "LGCP, así como los principios que rigen la contratación pública "
          "contenidos en el Artículo 5 de la Ley General de las Contrataciones "
          "Públicas.")


def parrafo_acuerdo():
    return (
        run("De acuerdo al", sz=14)
        + run(" Numeral 44.6 del Artículo 44°", sz=14, bold=True)
        + run(" del Reglamento de la Ley General de Contrataciones públicas, ",
              sz=14)
        + run("señala lo siguiente", sz=14, bold=True)
        + run(": ", sz=14)
        + run("“", sz=14, italic=True)
        + run(ACUERDO_CITA1, sz=14, bold=True, italic=True, underline=True)
        + run(ACUERDO_CITA2, sz=14, bold=True, italic=True)
    )


def parrafo_deloanterior():
    return (
        run(DELO_1, sz=14, italic=True)
        + run("deberá indicar el motivo de manera clara y objetiva",
              sz=14, bold=True, italic=True)
        + run(DELO_3, sz=14, italic=True)
    )


def build_body():
    parts = []
    # Titulo (sin numero: se agrega a mano; se deja el espacio subrayado)
    parts.append(para(
        run("MEMORANDO N°            -OL-J- H.N.A.L.-", font="Cambria",
            sz=32, bold=True, underline=True)
        + run("{anio}", font="Cambria", sz=32, bold=True, underline=True),
        jc="center", font="Cambria", sz=32, bold=True))
    # Tabla destinatario
    parts.append(tabla_destinatario())
    parts.append(blank())
    # Linea separadora
    parts.append(para("", jc=None, border=True))
    parts.append(blank())
    # Contexto: el objeto y el numero de cotizaciones van en AZUL + negrita
    ctx = (run(CONTEXTO_PRE, sz=16) + run("“", sz=16)
           + run("{titulo}", sz=16, bold=True, color=AZUL) + run("”", sz=16)
           + run(CONTEXTO_MID, sz=16)
           + run("{cantidad} cotizaciones", sz=16, bold=True, color=AZUL)
           + run(" a solicitud de ", sz=16)
           + run("{solicitud}", sz=16)
           + run(CONTEXTO_END, sz=16))
    parts.append(para(ctx, jc="both", sz=16))
    parts.append(blank())
    # Tabla de postores
    parts.append(tabla_postores())
    parts.append(blank())
    # Parrafos legales (con el formato del memo hecho a mano)
    parts.append(para(run(SOBRE, sz=16), jc="both", sz=16))
    parts.append(blank())
    parts.append(para(parrafo_acuerdo(), jc="both", sz=14))
    parts.append(blank())
    parts.append(para(run(ASIMISMO, sz=14, bold=True, italic=True),
                      jc="both", sz=14))
    parts.append(blank())
    parts.append(para(parrafo_deloanterior(), jc="both", sz=14))
    parts.append(blank())
    parts.append(para(run(ENTALSENTIDO, sz=14), jc="both", sz=14))
    parts.append(blank())
    # Atentamente (el espacio para sello/firma lo da el bloque EXP anclado al
    # pie, que flota; por eso aqui basta una linea en blanco)
    parts.append(para(run("Atentamente,", sz=14), jc="left", sz=14))
    parts.append(blank())
    # Pie: bloque anclado al fondo de la pagina (framePr compartido). Siempre
    # sale pegado al pie. Arial 8pt, etiquetas en negrita (como el manual).
    def pie_para(runs_xml):
        return para(runs_xml, jc="left", font="Arial", sz=16,
                    frame=FRAME_BOTTOM)

    parts.append(pie_para(run("EXP. DIR.: {expDir}", sz=16, bold=True)))
    parts.append(pie_para(run("EXP. LOG. {expLog}", sz=16, bold=True)))
    parts.append(pie_para(
        run("Folios: ", sz=16)
        + run("(        {folios}        )", sz=16, bold=True)
        + run(" original", sz=16)))
    parts.append(pie_para(run("C.c.  Archivo", sz=16)))
    parts.append(pie_para(
        run("Elaborado por: ", sz=16, bold=True) + run("{elaborado}", sz=16)))
    parts.append(pie_para(run("{iniciales}", sz=16)))
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
