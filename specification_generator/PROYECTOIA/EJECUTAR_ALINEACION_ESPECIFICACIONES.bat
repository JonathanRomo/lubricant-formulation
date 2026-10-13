@echo off
setlocal
cd /d "%~dp0"
title Alinear especificaciones

set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    where py >nul 2>&1
    if errorlevel 1 (
        echo ERROR: No se encontro Python Launcher. Instala Python 3.11 y el comando py.
        exit /b 1
    )
    echo Creando entorno virtual con Python 3.11...
    py -3.11 -m venv "%~dp0.venv"
    if errorlevel 1 (
        echo ERROR: No se pudo crear el entorno virtual de Python 3.11.
        exit /b 1
    )
)

"%PYTHON%" -c "import sys; raise SystemExit(sys.version_info[:2] != (3, 11))" >nul 2>&1
if errorlevel 1 (
    echo ERROR: El entorno .venv debe usar Python 3.11.
    echo Crea un entorno nuevo con: py -3.11 -m venv .venv
    exit /b 1
)

"%PYTHON%" -c "import openpyxl, pandas, rapidfuzz" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias necesarias para la alineacion...
    "%PYTHON%" -m pip install "openpyxl==3.1.5" "pandas>=2.2,<3" "rapidfuzz>=3.9,<4"
    if errorlevel 1 (
        echo ERROR: No se pudieron instalar las dependencias del alineador.
        exit /b 1
    )
)

echo.
echo Alineando pruebas y metodos...
"%PYTHON%" "%~dp0alinear_especificaciones_json.py" --base "%~dp0."
if errorlevel 1 (
    echo.
    echo ERROR: La alineacion termino con errores.
    exit /b 1
)

echo.
echo Alineacion terminada correctamente.
echo Revisa la carpeta ESPECIFICACIONES_JSON_ALINEADAS.
endlocal
