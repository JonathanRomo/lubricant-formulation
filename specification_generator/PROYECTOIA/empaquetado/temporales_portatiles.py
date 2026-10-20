"""Mantiene los temporales de cada proceso en la carpeta portatil."""
import os
import sys
import tempfile
from pathlib import Path

# Una carpeta por proceso evita borrar temporales de otra instancia activa.
# TemporaryDirectory registra su limpieza al finalizar normalmente el proceso.
_raiz_temporales = Path(sys.executable).resolve().parent / 'TEMPORAL'
_raiz_temporales.mkdir(exist_ok=True)
_sesion_temporal = tempfile.TemporaryDirectory(
    prefix='sesion_', dir=str(_raiz_temporales)
)
tempfile.tempdir = _sesion_temporal.name
for _variable in ('TEMP', 'TMP', 'TMPDIR'):
    os.environ[_variable] = _sesion_temporal.name
