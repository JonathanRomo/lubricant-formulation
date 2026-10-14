"""Ruta compartida para las listas y sus registros, con configuración en caché."""
from rutas_proyecto import BASE, rutas_configuradas


def carpeta_listas(base=BASE):
    return rutas_configuradas(base)["metodos_lab"]


def ruta_maestro(base=BASE):
    return carpeta_listas(base) / "METODOS_LAB_UNIFICADO.xlsx"
