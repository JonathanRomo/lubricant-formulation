"""Punto de entrada estable de la aplicacion."""
from __future__ import annotations

def main() -> None:
    """Inicia la interfaz grafica principal."""
    # La interfaz y sus colaboradores aun usan importaciones planas. Cargarla
    # con esa misma identidad evita duplicar modulos durante la transicion.
    from interfaz_aplicacion_especificaciones import main as ejecutar_interfaz

    ejecutar_interfaz()
