# 🌐 Cómo tener un link público (GitHub Pages) — GRATIS y para siempre

El archivo **`docs/index.html`** es la app completa en un solo archivo. Funciona
de dos maneras:

1. **Sin internet:** descárgalo y haz **doble clic** → se abre en tu navegador.
2. **Con link público:** actívalo con GitHub Pages (pasos abajo) y tendrás una
   dirección web para usarlo desde cualquier lado, incluso el **celular**.

> No se "duerme" ni caduca, y es **gratis** (a diferencia de un servidor).

---

## Activar el link público (una sola vez, ~2 minutos)

1. Entra a tu repositorio **`renzomoran9/hnal`** en GitHub.
2. Ve a **Settings** (Configuración) → en el menú izquierdo, **Pages**.
3. En **Build and deployment → Source**, elige **Deploy from a branch**.
4. En **Branch**, selecciona:
   - la rama **`claude/excel-validation-model-ogwy3i`** (o `main` si ya la fusionaste)
   - y la carpeta **`/docs`**
5. Presiona **Save**.
6. Espera ~1 minuto y recarga. Arriba aparecerá tu link, algo como:
   **`https://renzomoran9.github.io/hnal/`**

¡Listo! Abre ese link y usa la app. 🎉

---

## Para actualizarla después

- Solo reemplaza el archivo **`docs/index.html`** en el repositorio (súbelo o
  pégalo) y GitHub Pages se actualiza **solo** en ~1 minuto.
- Si cambias el **formulario** o la **plantilla base**, edita `web/shell.html` o
  `plantilla/FORMATO_VALIDACIONES_DE_BIENES.xlsx` y vuelve a generar el HTML con:
  ```bash
  python web/build_html.py
  ```
  (eso regenera `docs/index.html`).

---

## Importante

- GitHub Pages gratis funciona con repositorios **públicos**. Si tu repo es
  privado, necesitarías GitHub Pro, o simplemente usa el archivo con **doble
  clic** (funciona igual, sin link).
- Todo ocurre en tu navegador: los datos que escribes **no se envían a ningún
  servidor**.
