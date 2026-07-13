@echo off
REM ============================================================
REM  Iniciar la App de Cuadros de Validacion (Windows)
REM  Doble clic en este archivo para abrir la aplicacion.
REM ============================================================
cd /d "%~dp0"

echo Instalando/verificando dependencias (solo la primera vez)...
python -m pip install -r requirements.txt

echo.
echo Iniciando la aplicacion...
echo Se abrira tu navegador en http://localhost:5000
echo Para cerrar la app, cierra esta ventana negra.
echo.
python app.py

pause
