"""Conservar Tkinter; el spec recoge y valida Tcl/Tk explícitamente si hace falta."""
from pathlib import Path
import sys


def pre_find_module_path(hook_api):
    if not (Path(sys.base_prefix) / "Lib" / "tkinter" / "__init__.py").is_file():
        raise SystemExit("Faltan los módulos de Tkinter en la instalación de Python.")
    # No vaciar search_dirs cuando el entorno de análisis no puede iniciar Tcl.
