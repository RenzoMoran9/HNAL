# -*- coding: utf-8 -*-
"""
App web para generar Cuadros de Validacion de Bienes.

Levanta un formulario en el navegador donde se ingresan los datos
(adquisicion, codigo SIGA, denominacion, cantidad y los postores) y
descarga el Excel ya rellenado, identico al modelo (formato + logo).

Ejecutar:
    python app.py
Luego abrir en el navegador:
    http://localhost:5000
"""

import io
import os
import re
import webbrowser
import threading
from flask import Flask, request, render_template, send_file, jsonify

from generador_validaciones import generar

app = Flask(__name__)


def _nombre_archivo(adquisicion):
    base = re.sub(r"[^A-Za-z0-9]+", "_", str(adquisicion).upper()).strip("_")
    base = base[:50] or "CUADRO"
    return "VALIDACION_%s.xlsx" % base


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generar", methods=["POST"])
def generar_endpoint():
    data = request.get_json(force=True, silent=True) or {}

    adquisicion = (data.get("adquisicion") or "").strip()
    tipo = "servicio" if data.get("tipo") == "servicio" else "bien"

    postores = [
        {"nombre": (p.get("nombre") or "").strip()}
        for p in (data.get("postores") or [])
        if (p.get("nombre") or "").strip()
    ]

    items = [
        {
            "codigo_siga": (it.get("codigo_siga") or "").strip(),
            "denominacion": (it.get("denominacion") or "").strip(),
            "unidad": (it.get("unidad") or "UND").strip() or "UND",
            "cantidad": it.get("cantidad", ""),
            "marcas": [str(m or "").strip() for m in (it.get("marcas") or [])],
        }
        for it in (data.get("items") or [])
        if (it.get("denominacion") or "").strip() or (it.get("codigo_siga") or "").strip()
        or str(it.get("cantidad") or "").strip()
    ]

    if not adquisicion:
        return jsonify(error="Falta la adquisicion / servicio."), 400
    if not postores:
        return jsonify(error="Agrega al menos un postor con nombre."), 400
    if not items:
        return jsonify(error="Agrega al menos un bien/servicio."), 400

    datos = {
        "tipo": tipo,
        "adquisicion": adquisicion,
        "postores": postores,
        "items": items,
    }

    # Generar en memoria y devolver como descarga
    buffer = io.BytesIO()
    generar(datos, buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=_nombre_archivo(adquisicion),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _abrir_navegador():
    webbrowser.open("http://localhost:5000")


if __name__ == "__main__":
    # En un hosting (Render, etc.) la variable PORT viene definida.
    # En tu computadora no existe, y ahi abrimos el navegador solo.
    puerto = int(os.environ.get("PORT", "5000"))
    es_local = "PORT" not in os.environ
    if es_local:
        print("=" * 55)
        print("  APP DE CUADROS DE VALIDACION")
        print("  Abre tu navegador en:  http://localhost:5000")
        print("  (para cerrar la app, presiona Ctrl + C aqui)")
        print("=" * 55)
        threading.Timer(1.2, _abrir_navegador).start()
    app.run(host="0.0.0.0", port=puerto, debug=False)
