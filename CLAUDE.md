# HNAL · Generador de Formatos

App de una sola página que autocompleta los documentos de Logística del
Hospital Nacional Arzobispo Loayza.

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
- Commit y push a la rama `claude/excel-validation-model-ogwy3i`.

## Estructura

- `web/shell.html` — toda la app (HTML + CSS + JS en un archivo).
- `web/build_html.py` — incrusta ExcelJS, docxtemplater, las plantillas en
  base64, el directorio y las fuentes → `docs/index.html` (GitHub Pages).
- `plantilla/` — los .xlsx y .docx oficiales que se rellenan.
- `web/directorio.json` — 151 personas del hospital con su área y su grado.

## Documentos que genera

Validación (Excel) → Memo (Word) → Comparativo (Excel) → Nota Informativa
(Word). Aparte: Carta al proveedor (Word), para pedir adelanto por urgencia.
