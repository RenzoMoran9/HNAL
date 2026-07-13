# Generador de Cuadros de Validación de Bienes

Sistema para **rellenar automáticamente** el formato de *Validación de
Especificaciones Técnicas para la Adquisición de Bienes*. En lugar de llenar el
Excel a mano, el programa te hace las preguntas necesarias y te devuelve el
**mismo modelo, con el mismo formato y logo**, pero ya con los datos completos.

El sistema también **agrega o quita las columnas de postores** automáticamente,
según cuántos hayan presentado su cotización.

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

Desde una terminal, en esta carpeta:

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
├── generador_validaciones.py   # Motor + preguntas interactivas
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
