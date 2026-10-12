@echo off
setlocal
cd /d "%~dp0"

set "PYTHON=%~dp0.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
    where py >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Se requiere Python Launcher y Python 3.11 para crear .venv.
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
    exit /b 1
)

"%PYTHON%" -c "import docx, openpyxl" >nul 2>&1
if errorlevel 1 (
    "%PYTHON%" -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 (
        echo ERROR: No se pudieron instalar las dependencias del gestor.
        exit /b 1
    )
)

"%PYTHON%" "%~dp0gestor_especificaciones.py" %*
exit /b %errorlevel%