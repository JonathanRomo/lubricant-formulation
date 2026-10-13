@echo off
setlocal
cd /d "%~dp0"
title Unificar Metodos Lab

echo ======================================================
echo  UNIFICAR METODOS LAB
echo ======================================================
echo.

set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    echo Creando entorno virtual con Python 3.11...
    py -3.11 -m venv "%~dp0.venv"
    if errorlevel 1 (
        echo ERROR: No se pudo crear el entorno virtual.
        echo Verifica que Python 3.11 y el comando py esten instalados.
        pause
        exit /b 1
    )
)

echo Verificando openpyxl...
"%PYTHON%" -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias en el entorno virtual del proyecto...
    "%PYTHON%" -m pip install -r "%~dp0requirements_metodos_lab.txt"
    if errorlevel 1 (
        echo.
        echo ERROR: No se pudo instalar openpyxl.
        echo Revisa la conexion a Internet y vuelve a ejecutar este archivo.
        pause
        exit /b 1
    )
)

echo.
echo Ejecutando unificacion...
"%PYTHON%" "%~dp0unificar_metodos_lab_v3.py" --base "%~dp0"
if errorlevel 1 (
    echo.
    echo La ejecucion termino con errores.
    pause
    exit /b 1
)

echo.
echo Proceso terminado correctamente.
echo Revisa la carpeta METODOS_LAB.
pause
endlocal
