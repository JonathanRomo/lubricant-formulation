"""Retira índices reconstruibles y compacta SQLite, conservando un respaldo."""
from __future__ import annotations

import argparse
import json
import os
import sqlite3
import tempfile
from contextlib import closing
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from rutas_proyecto import BASE, reubicar_ruta_heredada, resolver_ruta


TABLAS_AUXILIARES = (
    "hii_enlaces_automaticos", "hii_traducciones", "hii_generadas",
    "hii_estado", "hii_especificaciones", "recuperacion_documentos",
)


def _relativa(archivo, carpeta):
    try:
        return Path(os.path.relpath(archivo, carpeta)).as_posix()
    except ValueError:
        return archivo.as_posix()


def _escribir_atomico(ruta, contenido):
    temporal = None
    try:
        with tempfile.NamedTemporaryFile(dir=ruta.parent, prefix=".migracion_", suffix=".tmp", delete=False) as archivo:
            temporal = Path(archivo.name)
            archivo.write(contenido)
            archivo.flush()
            os.fsync(archivo.fileno())
        os.replace(temporal, ruta)
    finally:
        if temporal is not None:
            temporal.unlink(missing_ok=True)


def _referencias(datos, origen):
    relaciones = datos.get("versiones_idioma", {})
    if isinstance(relaciones, dict):
        for referencias in relaciones.values():
            if isinstance(referencias, dict):
                referencias = [referencias]
            if isinstance(referencias, list):
                for referencia in referencias:
                    if isinstance(referencia, dict) and isinstance(referencia.get("archivo_json"), str):
                        yield resolver_ruta(referencia["archivo_json"], origen.parent)


def _rescatar_hii(con, tablas, base, resumen):
    """Prepara solo los vínculos o registros que aún existen exclusivamente en SQLite."""
    from generador_especs_hii import clave_texto, leer_idioma_datos
    documentos = {}
    def leer(ruta):
        if ruta not in documentos:
            contenido = ruta.read_bytes()
            datos = json.loads(contenido.decode("utf-8-sig"))
            if not isinstance(datos, dict):
                raise ValueError(f"El JSON debe contener un objeto: {ruta}")
            documentos[ruta] = (contenido, datos)
        return documentos[ruta][1]

    def identidad(datos):
        control = datos.get("control_documental", {})
        return tuple(clave_texto(str(control.get(c) or "")) for c in ("codigo_unico_producto", "nombre_cliente"))

    if "hii_traducciones" in tablas:
        automaticas = set()
        if "hii_enlaces_automaticos" in tablas:
            automaticas = {frozenset(f) for f in con.execute("SELECT origen,destino FROM main.hii_enlaces_automaticos")}
        for a, b in con.execute("SELECT origen,destino FROM main.hii_traducciones"):
            if frozenset((a, b)) in automaticas:
                continue
            origen, destino = (reubicar_ruta_heredada(r, base) for r in (a, b))
            if not origen.is_file() or not destino.is_file():
                resumen["advertencias"].append(f"Vínculo con archivo ausente conservado en el respaldo: {a} / {b}")
                continue
            datos_a, datos_b = leer(origen), leer(destino)
            if destino in set(_referencias(datos_a, origen)) or origen in set(_referencias(datos_b, destino)):
                continue
            if (not identidad(datos_a)[0] or identidad(datos_a) != identidad(datos_b)
                    or leer_idioma_datos(datos_a) == leer_idioma_datos(datos_b)):
                resumen["advertencias"].append(f"Vínculo incompatible conservado en el respaldo: {a} / {b}")
                continue
            for ruta, datos, otra, traducida in ((origen, datos_a, destino, datos_b), (destino, datos_b, origen, datos_a)):
                relaciones = datos.setdefault("versiones_idioma", {})
                if not isinstance(relaciones, dict):
                    raise ValueError(f"Relaciones de idiomas inválidas: {ruta}")
                idioma = leer_idioma_datos(traducida)
                referencias = relaciones.setdefault(idioma, [])
                if isinstance(referencias, dict):
                    referencias = relaciones[idioma] = [referencias]
                if not isinstance(referencias, list):
                    raise ValueError(f"Referencias de idiomas inválidas: {ruta}")
                control = traducida.get("control_documental", {})
                referencias.append({"archivo_json": _relativa(otra, ruta.parent),
                    "codigo": clave_texto(str(control.get("codigo_unico_producto") or "")),
                    "revision": clave_texto(str(control.get("revision") or "")),
                    "cliente": identidad(traducida)[1]})
            resumen["vinculos_rescatados"] += 1

    if "hii_generadas" in tablas:
        for archivo, especificacion, idioma, fecha in con.execute("SELECT ruta,especificacion,idioma,fecha FROM main.hii_generadas"):
            origen = reubicar_ruta_heredada(especificacion, base)
            destino = reubicar_ruta_heredada(archivo, base)
            if not origen.is_file():
                resumen["advertencias"].append(f"Registro HII sin JSON conservado en el respaldo: {archivo}")
                continue
            datos = leer(origen)
            historial = datos.get("hii_generadas", [])
            heredados = datos.get("hii_registros_heredados", [])
            if not isinstance(historial, list) or not isinstance(heredados, list):
                raise ValueError(f"Historial HII inválido: {origen}")
            if any(isinstance(r, dict) and isinstance(r.get("archivo"), str)
                   and resolver_ruta(r["archivo"], origen.parent) == destino for r in [*historial, *heredados]):
                continue
            # La tabla antigua no conocía ID, huella ni hora: no inventar trazabilidad.
            registro = {"archivo": _relativa(destino, origen.parent), "idioma": idioma, "fecha": fecha}
            datos["hii_registros_heredados"] = [*heredados, registro]
            resumen["registros_rescatados"] += 1

    cambios = {}
    for ruta, (anterior, datos) in documentos.items():
        if json.loads(anterior.decode("utf-8-sig")) != datos:
            cambios[ruta] = (anterior, (json.dumps(datos, ensure_ascii=False, indent=2) + "\n").encode("utf-8"))
    return cambios


def retirar_indices_auxiliares(con, base=BASE):
    """Migra bases antiguas una vez; una base actual no genera copias ni escrituras."""
    tablas = {f[0] for f in con.execute("SELECT name FROM main.sqlite_master WHERE type='table'")}
    retirar = [t for t in TABLAS_AUXILIARES if t in tablas]
    resumen = {"tablas_retiradas": [], "vinculos_rescatados": 0, "registros_rescatados": 0,
               "advertencias": [], "respaldo": None}
    if not retirar:
        return resumen
    if con.in_transaction:
        raise RuntimeError("La migración necesita una conexión sin cambios pendientes.")
    base = resolver_ruta(base)
    db = Path(con.execute("PRAGMA database_list").fetchone()[2])
    escritos = []
    try:
        con.execute("BEGIN IMMEDIATE")
        # Otra instancia pudo completar la migración mientras esperábamos el bloqueo.
        tablas = {f[0] for f in con.execute("SELECT name FROM main.sqlite_master WHERE type='table'")}
        retirar = [t for t in TABLAS_AUXILIARES if t in tablas]
        if not retirar:
            con.rollback()
            return resumen
        if con.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise ValueError("La base no superó la comprobación de integridad.")
        copia = base / "RESPALDOS_BASE_DATOS" / ("optimizacion_" + datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid4().hex[:8]) / db.name
        copia.parent.mkdir(parents=True)
        with closing(sqlite3.connect(db.as_uri() + "?mode=ro", uri=True)) as origen, closing(sqlite3.connect(copia)) as respaldo:
            origen.backup(respaldo)
        resumen["respaldo"] = str(copia)
        cambios = _rescatar_hii(con, tablas, base, resumen)
        for vista in ("hii_versiones_idiomas", "recuperacion_actuales"):
            con.execute(f"DROP VIEW IF EXISTS main.{vista}")
        for tabla in retirar:
            con.execute(f"DROP TABLE main.{tabla}")
        if con.execute("PRAGMA foreign_key_check").fetchone():
            raise ValueError("La migración dejaría referencias inválidas.")
        for ruta, (anterior, contenido) in cambios.items():
            if ruta.read_bytes() != anterior:
                raise RuntimeError(f"El JSON cambió durante la migración: {ruta}")
            _escribir_atomico(ruta, contenido)
            escritos.append((ruta, anterior))
        con.commit()
        resumen["tablas_retiradas"] = retirar
        return resumen
    except BaseException:
        con.rollback()
        for ruta, anterior in reversed(escritos):
            _escribir_atomico(ruta, anterior)
        raise


def optimizar(base=BASE, *, destino=None):
    base = resolver_ruta(base)
    db = resolver_ruta(destino, base) if destino is not None else base / "especificaciones.db"
    if not db.is_file():
        raise FileNotFoundError(db)
    antes = db.stat().st_size
    with closing(sqlite3.connect(db, timeout=30)) as con:
        con.execute("PRAGMA foreign_keys=ON")
        resumen = retirar_indices_auxiliares(con, base)
        # VACUUM recupera páginas que DROP deja libres dentro del archivo.
        con.execute("VACUUM")
        if con.execute("PRAGMA integrity_check").fetchone()[0] != "ok" or con.execute("PRAGMA foreign_key_check").fetchone():
            raise ValueError("La base optimizada no superó la validación.")
        con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
    resumen.update(bytes_antes=antes, bytes_despues=db.stat().st_size)
    return resumen


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", type=Path, default=BASE)
    args = parser.parse_args()
    print(json.dumps(optimizar(args.base), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
