# 🌐 Cómo subir la app a internet (gratis)

Con esto tendrás un **link** para usar la app desde cualquier computadora o
celular, **sin instalar nada**. Usaremos **Render** (tiene plan gratuito).

> Solo necesitas hacer esto **una vez**. Después, cada cambio que se suba al
> repositorio se publica **solo** (queda siempre actualizado).

---

## Paso a paso (unos 5 minutos)

### 1. Crea una cuenta en Render
- Entra a **https://render.com** y presiona **Get Started**.
- Regístrate con tu cuenta de **GitHub** (la más fácil).

### 2. Crea el servicio web
- En el panel, presiona **New +** → **Web Service**.
- Conecta tu repositorio **`renzomoran9/hnal`**
  (si te pide autorizar Render en GitHub, acepta).

### 3. Configura estos campos
| Campo | Valor |
|-------|-------|
| **Branch** (rama) | `claude/excel-validation-model-ogwy3i` |
| **Runtime / Language** | Python 3 *(lo detecta solo)* |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `gunicorn app:app` |
| **Instance Type** | **Free** |

> El nombre del servicio puede ser el que quieras, por ejemplo
> `cuadros-validacion`. Ese nombre saldrá en el link.

### 4. Crea el servicio
- Presiona **Create Web Service** y espera 2–4 minutos mientras se instala.
- Cuando diga **"Live"**, arriba verás tu link, algo como:
  **`https://cuadros-validacion.onrender.com`**

### 5. ¡Listo!
Abre ese link en cualquier navegador (compu o celular), llena el formulario y
descarga tu Excel. 🎉

---

## Alternativa aún más rápida (Blueprint)

El proyecto ya trae un archivo `render.yaml`, así que también puedes:
- **New +** → **Blueprint** → conectar `renzomoran9/hnal` → **Apply**.

Render lee la configuración solo y crea todo automáticamente.

---

## Cosas que debes saber

- **Se "duerme" tras 15 min sin uso** (plan gratuito). La primera vez que
  entres después de un rato, puede tardar ~30–60 segundos en despertar. Luego va
  rápido. *(Si quieres que nunca se duerma, se puede pasar a un plan pago o usar
  otro hosting; avísame.)*
- **Se actualiza solo:** cada vez que se suba un cambio al repositorio en la rama
  configurada, Render vuelve a publicar la app automáticamente.
- Si más adelante fusionas la rama a `main`, cambia el **Branch** en la
  configuración de Render (o en `render.yaml`) a `main`.

---

## ¿Problemas?
Si el despliegue falla o el link no abre, mándame el mensaje de error que
aparece en Render (pestaña **Logs**) y lo resolvemos.
