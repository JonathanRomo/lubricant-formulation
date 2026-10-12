@echo off
setlocal
cd /d "%~dp0"
title Pipeline de especificaciones

set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    where py >nul 2>&1
    if errorlevel 1 (
        echo ERROR: No se encontro Python Launcher. Instala Python 3.11 y el comando py.
        exit /b 1
    )
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

"%PYTHON%" -c "import docx, openpyxl, pandas, rapidfuzz" >nul 2>&1
if errorlevel 1 (
    echo Instalando las dependencias del flujo...
    "%PYTHON%" -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 (
        echo ERROR: No se pudieron instalar las dependencias.
        exit /b 1
    )
)

echo.
echo [1/3] Extrayendo especificaciones a JSON...
"%PYTHON%" "%~dp0gestor_especificaciones.py"
if errorlevel 1 (
    echo ERROR: Fallo la extraccion. No se ejecutaran las etapas siguientes.
    exit /b 1
)

echo.
echo [2/3] Alineando pruebas y metodos con la lista maestra...
"%PYTHON%" "%~dp0alinear_especificaciones_json.py" --base "%~dp0"
if errorlevel 1 (
    echo ERROR: Fallo la alineacion. No se ejecutara el generador.
    exit /b 1
)

if not exist "%~dp0FORMATOS\" (
    echo ERROR: No existe la carpeta de plantillas "%~dp0FORMATOS".
    echo JSON alineados disponibles en "%~dp0ESPECIFICACIONES_JSON_ALINEADAS".
    exit /b 1
)

echo.
echo [3/3] Generando documentos desde los JSON alineados...
"%PYTHON%" "%~dp0generador_especs_hii.py" ^
    --formatos "%~dp0FORMATOS" ^
    --json-dir "%~dp0ESPECIFICACIONES_JSON_ALINEADAS" ^
    --salida "%~dp0HISTORICO_ESPECIFICACIONES"
if errorlevel 1 (
    echo ERROR: El generador termino con errores.
    exit /b 1
)

echo.
echo Pipeline terminado correctamente.
endlocal
