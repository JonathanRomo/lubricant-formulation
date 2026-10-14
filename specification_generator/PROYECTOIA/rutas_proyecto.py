"""Rutas del proyecto: independientes del directorio de trabajo y portátiles al guardar."""
from __future__ import annotations

import json
import sys
from functools import lru_cache
from pathlib import Path


def base_programa() -> Path:
    origen = sys.executable if getattr(sys, "frozen", False) else __file__
    return Path(origen).resolve().parent


BASE = base_programa()
CARPETAS = {
    "especificaciones_viejas": "ESPECIFICACIONES_VIEJAS",
    "json_anteriores": "ESPECIFICACIONES_JSON",
    "json_alineadas": "ESPECIFICACIONES_JSON_ALINEADAS",
    "formatos": "FORMATOS",
    "historico": "HISTORICO_ESPECIFICACIONES",
    "metodos_lab": "METODOS_LAB",
}


def resolver_ruta(ruta, base=BASE) -> Path:
    """Las rutas relativas se interpretan respecto al programa, nunca respecto a cwd."""
    base = Path(base).expanduser()
    if not base.is_absolute():
        base = BASE / base
    ruta = Path(ruta).expanduser()
    return (ruta if ruta.is_absolute() else base / ruta).resolve()


def ruta_portatil(ruta, base=BASE) -> str:
    """Guarda recursos internos como relativos; conserva ubicaciones externas explícitas."""
    base = resolver_ruta(base)
    ruta = resolver_ruta(ruta, base)
    try:
        return ruta.relative_to(base).as_posix()
    except ValueError:
        return str(ruta)


def _leer_configuracion(archivo, firma):
    try:
        datos = json.loads(Path(archivo).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    if not isinstance(datos, dict):
        raise ValueError("La configuración de carpetas debe ser un objeto JSON.")
    return datos


@lru_cache(maxsize=16)
def _rutas_resueltas(base, archivo, firma):
    datos = _leer_configuracion(archivo, firma)
    base = Path(base)
    return {
        clave: reubicar_ruta_heredada(datos[clave], base)
        if isinstance(datos.get(clave), str) and datos[clave].strip()
        else base / carpeta
        for clave, carpeta in CARPETAS.items()
    }


def rutas_configuradas(base=BASE, *, archivo=None, estricto=False) -> dict[str, Path]:
    base = resolver_ruta(base)
    archivo = resolver_ruta(archivo, base) if archivo is not None else base / "configuracion_interfaz.json"
    try:
        info = archivo.stat()
        firma = (info.st_mtime_ns, info.st_size)
    except OSError:
        firma = None
    try:
        return dict(_rutas_resueltas(str(base), str(archivo), firma))
    except (OSError, ValueError):
        if estricto:
            raise
        return {clave: base / carpeta for clave, carpeta in CARPETAS.items()}


def reubicar_ruta_heredada(ruta, base=BASE) -> Path:
    """Recupera referencias antiguas a recursos existentes tras mover la carpeta madre."""
    base = resolver_ruta(base)
    original = resolver_ruta(ruta, base)
    if not Path(ruta).is_absolute() or original.exists():
        return original
    # Solo carpetas conocidas y archivos presentes: no adivinar rutas externas.
    carpetas = {nombre.casefold() for nombre in CARPETAS.values()}
    carpetas.update({"especificaciones_viejas_eng", "revisar"})
    for indice, parte in enumerate(original.parts):
        if parte.casefold() in carpetas:
            candidata = base.joinpath(*original.parts[indice:])
            if candidata.exists():
                return candidata.resolve()
    if original.name == "CODIGOS_200.xlsx" and (base / original.name).is_file():
        return base / original.name
    return original
