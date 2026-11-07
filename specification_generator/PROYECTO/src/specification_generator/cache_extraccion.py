"""Actualización de datos derivados sin crear revisiones documentales."""
from __future__ import annotations

import copy
import json
import os
import tempfile
from pathlib import Path

CAMPO_FUENTE = ("producto", "documento", "control_documental", "pruebas", "notas",
                "datos_institucionales", "estructura_detectada")
_AUSENTE = object()
REVISION_GESTIONADA = ("revision", "revision_fuente", "revision_recibida",
                      "historial_revisiones_recibidas")


def conservar_revision(resultado, anterior):
    """La revisión asignada y su procedencia no son campos del lector Word."""
    control = resultado.setdefault("control_documental", {})
    previo = anterior.get("control_documental", {})
    for campo in REVISION_GESTIONADA:
        if campo in previo:
            control[campo] = copy.deepcopy(previo[campo])
        else:
            control.pop(campo, None)


def base_extraida(datos):
    return {campo: copy.deepcopy(datos[campo]) for campo in CAMPO_FUENTE if campo in datos}


def fusionar(actual, nueva, base=None):
    """Tres vías; sin base heredada solo se rellenan datos ausentes o vacíos."""
    protegidos = set()

    def combinar(viejo, fresco, previo, ruta):
        if previo is not _AUSENTE and viejo == previo:
            return copy.deepcopy(fresco) if fresco is not _AUSENTE else _AUSENTE
        if viejo is _AUSENTE:
            if previo is not _AUSENTE:
                protegidos.add(ruta)  # Eliminación manual.
                return _AUSENTE
            return copy.deepcopy(fresco) if fresco is not _AUSENTE else _AUSENTE
        if viejo == fresco:
            return copy.deepcopy(viejo)
        if isinstance(viejo, dict) and isinstance(fresco, dict):
            anterior = previo if isinstance(previo, dict) else {}
            resultado = {}
            for campo in dict.fromkeys([*viejo, *fresco, *anterior]):
                valor = combinar(viejo.get(campo, _AUSENTE), fresco.get(campo, _AUSENTE),
                                 anterior.get(campo, _AUSENTE), ruta + "." + campo)
                if valor is not _AUSENTE:
                    resultado[campo] = valor
            return resultado
        if previo is _AUSENTE:
            if viejo in (None, "", [], {}):
                return copy.deepcopy(fresco) if fresco is not _AUSENTE else copy.deepcopy(viejo)
            if (ruta == "pruebas" and isinstance(viejo, list) and isinstance(fresco, list)
                    and len(viejo) == len(fresco)
                    and all(isinstance(a, dict) and isinstance(b, dict)
                            and a.get("fila_origen") == b.get("fila_origen") for a, b in zip(viejo, fresco))):
                return [combinar(a, b, _AUSENTE, f"pruebas.{i}") for i, (a, b) in enumerate(zip(viejo, fresco))]
        protegidos.add(ruta)
        return copy.deepcopy(viejo)

    resultado = copy.deepcopy(actual)
    for campo in CAMPO_FUENTE:
        valor = combinar(actual.get(campo, _AUSENTE), nueva.get(campo, _AUSENTE),
                         base.get(campo, _AUSENTE) if base is not None else _AUSENTE, campo)
        if valor is _AUSENTE:
            resultado.pop(campo, None)
        else:
            resultado[campo] = valor
    conservar_revision(resultado, actual)
    if not protegidos:
        resultado["validacion_codigos_200"] = copy.deepcopy(nueva["validacion_codigos_200"])
    return resultado, protegidos


def escribir_atomico(ruta: Path, contenido: bytes):
    temporal = None
    try:
        with tempfile.NamedTemporaryFile(dir=ruta.parent, prefix=".extraccion_", suffix=".tmp", delete=False) as archivo:
            temporal = Path(archivo.name)
            archivo.write(contenido)
        os.replace(temporal, ruta)
    finally:
        if temporal is not None:
            temporal.unlink(missing_ok=True)


def actualizar(con, registro, salida, nueva, version_lector, fecha):
    ruta = salida / registro["archivo_json"]
    original = ruta.read_bytes() if ruta.exists() else None
    actual = json.loads(original.decode("utf-8-sig")) if original is not None else {}
    if not isinstance(actual, dict):
        raise ValueError("El JSON existente no es un objeto; se conserva para revisión.")
    base = json.loads(registro["base_extraccion_json"]) if registro["base_extraccion_json"] else None
    if base is not None and not isinstance(base, dict):
        raise ValueError("La base de extracción existente es inválida; no se sobrescribe.")
    resultado, protegidos = fusionar(actual, nueva, base) if original is not None else (copy.deepcopy(nueva), set())
    if original is None:
        revision_fuente = copy.deepcopy(nueva.get("control_documental", {}).get("revision", ""))
        if base:
            conservar_revision(resultado, base)
        resultado.setdefault("control_documental", {}).setdefault("revision_fuente", revision_fuente)
        if "revision_especificacion" in registro.keys():
            resultado.setdefault("control_documental", {})["revision"] = registro["revision_especificacion"]
    base_actualizada = base_extraida(nueva)
    conservar_revision(base_actualizada, resultado)
    for campo in ("hash_sha256_origen", "hash_sha256_catalogo", "extractor_version", "version_lector"):
        resultado[campo] = nueva[campo]
    resultado.setdefault("archivo_origen", nueva["archivo_origen"])
    resultado["archivo_json"] = registro["archivo_json"]
    resultado.setdefault("revision_json", f"{registro['revision_json']:02d}")
    resultado.setdefault("fecha_creacion_json", registro["fecha_creacion"])
    contenido = json.dumps(resultado, ensure_ascii=False, indent=2).encode("utf-8")
    try:
        con.execute("""UPDATE versiones SET version_lector=?, base_extraccion_json=?,
                    campos_protegidos_json=?, hash_catalogo=?, revision_especificacion=? WHERE id=?""",
                    (version_lector, json.dumps(base_actualizada, ensure_ascii=False),
                     json.dumps(sorted(protegidos), ensure_ascii=False),
                     nueva["hash_sha256_catalogo"], resultado.get("control_documental", {}).get("revision"), registro["version_id"]))
        con.execute("""UPDATE archivos SET hash_catalogo=?, ultima_revision_procesada=?,
                    revision_especificacion=CASE WHEN ultimo_json=? THEN ? ELSE revision_especificacion END,
                    estado='OK', ultimo_error=NULL WHERE id=?""",
                    (nueva["hash_sha256_catalogo"], fecha, registro["archivo_json"],
                     resultado.get("control_documental", {}).get("revision"), registro["id"]))
        escribir_atomico(ruta, contenido)
        con.commit()
    except BaseException:
        con.rollback()
        if original is not None:
            escribir_atomico(ruta, original)
        else:
            ruta.unlink(missing_ok=True)
        raise
    return protegidos
