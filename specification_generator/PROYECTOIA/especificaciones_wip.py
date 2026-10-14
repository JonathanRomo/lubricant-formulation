"""Estado de trabajo fuera del esquema de las especificaciones JSON."""
import copy
import json
import os
from pathlib import Path
from rutas_proyecto import resolver_ruta


def ruta_estado(ruta):
    ruta = resolver_ruta(ruta)
    return ruta.with_name("." + ruta.name + ".wip")


def es_wip(ruta):
    return ruta_estado(ruta).is_file()


def serializar_estado(modelo, rutas, ruta_actual):
    ruta_actual = resolver_ruta(ruta_actual)
    return {
        "version": 1,
        "borradores": copy.deepcopy(modelo.borradores),
        "pendientes": {i: sorted(modelo.pendientes[i]) for i in ("es", "en")},
        "automaticas": {i: sorted(modelo.automaticas[i]) for i in ("es", "en")},
        "revisiones": {i: sorted(modelo.revisiones[i]) for i in ("es", "en")},
        "manuales": {i: sorted(modelo.manuales[i]) for i in ("es", "en")},
        "rutas": {i: Path(os.path.relpath(resolver_ruta(r), ruta_actual.parent)).as_posix() for i, r in rutas.items()},
    }


def restaurar_estado(modelo, ruta):
    ruta = resolver_ruta(ruta)
    datos = json.loads(ruta_estado(ruta).read_text(encoding="utf-8"))
    if datos.get("version") != 1 or set(datos.get("borradores", {})) != {"es", "en"}:
        raise ValueError("Estado WIP inválido o incompatible.")
    modelo.borradores = datos["borradores"]
    for campo in ("pendientes", "automaticas", "revisiones"):
        setattr(modelo, campo, {i: {tuple(c) for c in datos[campo][i]} for i in ("es", "en")})
    modelo.manuales = {i: {tuple(c) for c in datos.get("manuales", {}).get(i, [])} for i in ("es", "en")}
    return {i: (Path(ruta).parent / r).resolve() for i, r in datos["rutas"].items()}
