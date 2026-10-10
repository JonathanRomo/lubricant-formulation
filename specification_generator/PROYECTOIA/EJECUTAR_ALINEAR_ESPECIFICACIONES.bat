@echo off
setlocal
cd /d "%~dp0"
title Alinear Especificaciones JSON

echo ======================================================
echo  ALINEAR ESPECIFICACIONES JSON CON METODOS LAB
echo ======================================================
echo.

where py >nul 2>&1
if errorlevel 1 (
  echo ERROR: No se encontro Python Launcher ^(py^).
  pause
  exit /b 1
)

py -3.11 --version >nul 2>&1
if errorlevel 1 (
  echo ERROR: No se encontro Python 3.11.
  pause
  exit /b 1
)

py -3.11 "%~dp0alinear_especificaciones_json.py" --base "%~dp0"
if errorlevel 1 (
  echo.
  echo El proceso termino con errores.
  pause
  exit /b 1
)

echo.
echo Proceso terminado. Revisa ESPECIFICACIONES_JSON_ALINEADAS.
pause
endlocal
