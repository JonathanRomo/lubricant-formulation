@echo off
setlocal
cd /d "%~dp0"
if not exist ".mplconfig" mkdir ".mplconfig"
set "MPLCONFIGDIR=%~dp0.mplconfig"
if not exist ".venv\Scripts\python.exe" (
  echo No se encontro el entorno Python local.
  echo Ejecuta la instalacion de dependencias antes de iniciar la aplicacion.
  pause
  exit /b 1
)
.venv\Scripts\python.exe app.py
if errorlevel 1 pause
