# Generador de Formatos

App de una sola página que autocompleta los documentos de contratación del
área de Logística.

La interfaz **no lleva marca de la entidad** (sin logo, sin nombre del
hospital): eso va solo dentro de los documentos que genera, que sí llevan su
membrete oficial. Al tocar la portada, la barra lateral o los textos de ayuda,
mantenerla neutral.

## Cómo responderme

- **Sé conciso.** Ve al grano, sin resúmenes largos ni repetir lo que ya dije.
- Nada de tablas ni listas de "lo que verifiqué" salvo que las pida.
- Si el cambio salió bien, dilo en una o dos líneas. Si algo falló o quedó
  incompleto, eso sí explícamelo.
- Escríbeme en español.

## Cómo trabajar en este repo

- Se edita `web/shell.html`; **nunca** `docs/index.html`, que se genera.
- Después de cualquier cambio: `python web/build_html.py`.
- Si cambia una plantilla Word: `python web/make_memo_template.py`,
  `make_nota_template.py` o `make_carta_template.py`, y luego el build.
- `python web/aplicar_titulos.py <memos.docx>` saca los grados profesionales
  (Dr., Ing., Q.F.) de documentos ya emitidos y los guarda en
  `web/directorio.json`.
- Cuando salga el "REPORTE DE OC" de un año nuevo: se deja el .xlsx en `datos/`
  y se corre `python web/make_siga.py`, que rehace `web/siga.json`. El usuario
  también puede cargar reportes nuevos desde la pantalla *Buscar*: la app los
  lee en el navegador y los guarda ahí (IndexedDB, store `siga`), y esa copia
  manda mientras sea igual o más nueva que la que trae la app.
- Commit y push a la rama `claude/excel-validation-model-ogwy3i`.

## Estructura

- `web/shell.html` — toda la app (HTML + CSS + JS en un archivo).
- `web/build_html.py` — incrusta ExcelJS, docxtemplater, las plantillas en
  base64, el directorio y las fuentes → `docs/index.html` (GitHub Pages).
- `plantilla/` — los .xlsx y .docx oficiales que se rellenan.
- `web/directorio.json` — 151 personas con su área y su grado.
- `datos/` — los "REPORTE DE OC" bajados del SIGA (una fila por ítem comprado).
- `web/siga.json` — catálogo armado con esos reportes: 4100 ítems con su código
  SIGA, su nombre oficial, su unidad y el precio de cada compra.

## Documentos que genera

Validación (Excel) → Memo (Word) → Comparativo (Excel) → Nota Informativa
(Word). Aparte: Carta al proveedor (Word), para pedir adelanto por urgencia.
