# Generador de Cuadros de Validación de Bienes y Servicios

Sistema para **rellenar automáticamente** el formato de *Validación de
Especificaciones Técnicas*. En lugar de llenar el Excel a mano, llenas un
formulario y te devuelve el **mismo modelo, con el mismo formato y logo**, pero
ya con los datos completos.

Características:

- **Bienes y servicios:** elige el tipo y el título del cuadro se arma solo
  (*"…ESPECIFICACIONES TÉCNICAS PARA LA ADQUISICIÓN DE…"* para bienes, o
  *"…TÉRMINO DE REFERENCIA PARA EL SERVICIO DE…"* para servicios).
- **Varios postores:** agrega o quita columnas de postores automáticamente.
- **Varias filas (items):** valida varios bienes/servicios en un mismo cuadro,
  cada uno con la marca de cada postor.
- La columna **SI / NO** se deja en blanco para el área usuario.

---

## Qué te pregunta

1. **Adquisición** — el texto que va después de *"…PARA LA ADQUISICIÓN DE"*
   (arma el título del cuadro).
2. **Código SIGA**
3. **Denominación** del bien
4. **Unidad de medida** (por defecto `UND`)
5. **Cantidad**
6. **Cuántos postores** presentaron cotización, y por cada uno:
   - Razón social / nombre
   - Marca ofertada

> **La columna `SI / NO` (cumple) se deja en blanco a propósito**: la llena el
> **área usuario** después de revisar. El sistema no la pregunta.

Lo que **no** cambia (encabezado institucional, logo, textos legales, bordes,
colores, fuentes) se conserva idéntico al modelo original.

---

## Requisitos (una sola vez)

1. Instalar **Python 3** — https://www.python.org/downloads/
   (en Windows, marca la casilla *"Add Python to PATH"* al instalar).
2. Instalar las librerías necesarias. Abre una terminal en esta carpeta y ejecuta:

   ```bash
   pip install -r requirements.txt
   ```

---

## Cómo usarlo

### Opción A — Un solo archivo HTML (la más fácil) ⭐

El archivo **[`docs/index.html`](docs/index.html)** es la app completa en un
solo archivo, **sin instalar nada** (ni Python, ni servidor). Todo funciona
dentro del navegador.

- **Doble clic** en `docs/index.html` → se abre, llenas el formulario y descargas
  el Excel.
- **¿Quieres un link público** (para usarlo desde el celular o compartirlo)? Sigue
  **[docs/COMO_PUBLICAR_LINK.md](docs/COMO_PUBLICAR_LINK.md)** para activarlo gratis
  con GitHub Pages.

> El HTML se genera con `python web/build_html.py` (incrusta la librería ExcelJS y
> la plantilla). Solo hace falta regenerarlo si cambias el formulario o la
> plantilla base.

### Opción B — App web con Python (Flask) 🌐

Un servidor local con el mismo formulario:

- **Windows:** doble clic en **`iniciar_app.bat`**
- **Mac / Linux:** ejecuta `bash iniciar_app.sh`
- **O manualmente:** `python app.py` → se abre en `http://localhost:5000`

Para publicarla en internet con un servidor, mira **[DESPLIEGUE.md](DESPLIEGUE.md)**.

### Opción C — Por consola (preguntas en la terminal)

```bash
python generador_validaciones.py
```

Responde las preguntas y, al final, se genera un archivo
`VALIDACION_<nombre>.xlsx` listo para abrir, imprimir o enviar.

---

## Uso avanzado (desde código)

También puedes generar cuadros sin la parte interactiva:

```python
from generador_validaciones import generar

datos = {
    "adquisicion": "TOKEN DE SEGURIDAD CON CONECTOR USB",
    "codigo_siga": "805200110001",
    "denominacion": "TOKEN DE SEGURIDAD CON CONECTOR USB",
    "unidad": "UND",
    "cantidad": 150,
    "postores": [
        {"nombre": "THIMMOTHE ROITTS PERUANA SAC", "marca": "LONGMAI", "cumple": "SI"},
        {"nombre": "INTERNATIONAL BUSINES TRACK EIRL", "marca": "LONGMAI", "cumple": "SI"},
        {"nombre": "BIT4ID PERU SAC", "marca": "BIT4ID", "cumple": "NO"},
    ],
}

generar(datos, "MI_CUADRO.xlsx")
```

---

## Estructura del proyecto

```
.
├── docs/
│   ├── index.html              # ⭐ App en UN SOLO archivo (abrir con doble clic)
│   └── COMO_PUBLICAR_LINK.md   # Guía para el link público (GitHub Pages)
├── web/
│   ├── shell.html              # Plantilla del HTML (formulario + lógica)
│   ├── build_html.py           # Genera docs/index.html
│   └── vendor/exceljs.min.js   # Librería para generar Excel en el navegador
├── app.py                      # App web con Python (servidor Flask)
├── templates/index.html        # Formulario de la app Flask
├── iniciar_app.bat             # Iniciar la app Flask en Windows (doble clic)
├── iniciar_app.sh              # Iniciar la app Flask en Mac / Linux
├── generador_validaciones.py   # Motor + preguntas por consola
├── plantilla/
│   └── FORMATO_VALIDACIONES_DE_BIENES.xlsx   # Modelo base (NO borrar)
├── ejemplos/
│   └── EJEMPLO_CON_MAS_POSTORES.xlsx         # Referencia de llenado
├── requirements.txt
└── README.md
```

> La carpeta `plantilla/` contiene el modelo base sobre el que se construye cada
> cuadro. **No lo borres ni cambies de nombre.** Si necesitas modificar el
> formato base, edita ese archivo y el sistema respetará los cambios.
