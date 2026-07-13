#!/usr/bin/env bash
# ============================================================
#  Iniciar la App de Cuadros de Validacion (Mac / Linux)
# ============================================================
cd "$(dirname "$0")" || exit 1

echo "Instalando/verificando dependencias (solo la primera vez)..."
python3 -m pip install -r requirements.txt

echo
echo "Iniciando la aplicacion..."
echo "Abre tu navegador en http://localhost:5000"
echo "Para cerrar la app, presiona Ctrl + C."
echo
python3 app.py
