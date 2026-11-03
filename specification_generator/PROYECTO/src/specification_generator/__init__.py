"""Gestor de especificaciones tecnicas bilingues."""
from __future__ import annotations

import sys
from pathlib import Path

__version__ = "1.1.4"

# Compatibilidad temporal para los modulos historicos, que todavia utilizan
# importaciones planas entre si. Permite adoptar el layout src sin cambiar su
# identidad publica de golpe; la migracion a imports relativos sera gradual.
_PACKAGE_DIR = str(Path(__file__).resolve().parent)
if _PACKAGE_DIR not in sys.path:
    sys.path.insert(0, _PACKAGE_DIR)

__all__ = ["__version__"]
