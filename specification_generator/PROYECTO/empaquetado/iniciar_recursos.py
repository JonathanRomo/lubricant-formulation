"""Instala recursos editables junto al ejecutable sin sobrescribir cambios."""
import shutil
import sys
from pathlib import Path

origen = Path(sys._MEIPASS) / 'recursos_iniciales'
destino = Path(sys.executable).resolve().parent
for recurso in origen.iterdir():
    salida = destino / recurso.name
    if not salida.exists():
        shutil.copy2(recurso, salida)
