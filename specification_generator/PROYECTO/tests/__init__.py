"""Configuracion compartida de las pruebas del proyecto."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "src"
PACKAGE = SOURCE / "specification_generator"

for directory in (str(SOURCE), str(PACKAGE)):
    if directory not in sys.path:
        sys.path.insert(0, directory)
