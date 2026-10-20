"""Recupera SQLite desde JSON existentes sin modificar especificaciones ni nombres."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sqlite3
import tempfile
import time
import uuid
from collections import defaultdict
from contextlib import closing, contextmanager
from datetime import datetime, timezone
from pathlib import Path
from zipfile import BadZipFile

from rutas_proyecto import BASE, resolver_ruta, ruta_portatil, rutas_configuradas


@contextmanager
def _bloqueo(destino, timeout=30):
    """Coordina recuperaciones entre procesos de este equipo, sin bloquear SQLite en uso."""
    archivo = open(destino.with_name("." + destino.name + ".recuperacion.lock"), "a+b")
    adquirido = False
    try:
        archivo.seek(0, 2)
        if archivo.tell() == 0:
            archivo.write(b"0")
            archivo.flush()
        limite = time.monotonic() + timeout
        while not adquirido:
            archivo.seek(0)
            try:
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(archivo.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(archivo, fcntl.LOCK_EX | fcntl.LOCK_NB)
                adquirido = True
            except OSError:
                if time.monotonic() >= limite:
                    raise TimeoutError("Otra reconstrucción de la base está en curso.")
                time.sleep(.1)
        yield
    finally:
        if adquirido:
            archivo.seek(0)
            if os.name == "nt":
                msvcrt.locking(archivo.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(archivo, fcntl.LOCK_UN)
        archivo.close()


def _texto(valor):
    return "" if valor is None else str(valor).strip()


def _fecha(valor):
    try:
        fecha = datetime.fromisoformat(_texto(valor))
        if fecha.tzinfo is not None:
            fecha = fecha.astimezone(timezone.utc).replace(tzinfo=None)
        return fecha.isoformat(timespec="seconds")
    except ValueError:
        return ""


def _huella(valor):
    texto = _texto(valor).lower()
    return texto if len(texto) == 64 and all(c in "0123456789abcdef" for c in texto) else ""


def _leer_documentos(base, rutas, progreso, resumen):
    from gestor_especificaciones import clave, parsear_fecha_documento
    from generador_especs_hii import leer_idioma_datos
    from especificaciones_wip import es_wip
    archivos = {}
    for tipo, carpeta in (("original", rutas["json_anteriores"]), ("alineada", rutas["json_alineadas"])):
        for ruta in sorted(carpeta.rglob("*")):
            if ruta.is_file() and ruta.suffix.lower() == ".json" and not ruta.name.startswith("."):
                archivos.setdefault(ruta.resolve(), tipo)
    documentos = []
    for indice, (ruta, tipo) in enumerate(archivos.items(), 1):
        if progreso:
            progreso(indice, len(archivos), ruta.name)
        try:
            contenido = ruta.read_bytes()
            datos = json.loads(contenido.decode("utf-8-sig"))
            if not isinstance(datos, dict):
                raise ValueError("El JSON debe contener un objeto.")
            producto, control, documento = (datos.get(c, {}) for c in ("producto", "control_documental", "documento"))
            if not all(isinstance(c, dict) for c in (producto, control, documento)):
                raise ValueError("Producto, control documental o documento inválidos.")
            nombre, codigo = _texto(producto.get("nombre")), _texto(control.get("codigo_unico_producto"))
            if not (nombre or codigo):
                raise ValueError("Faltan el nombre y el código del producto.")
            fecha_doc = _fecha(documento.get("fecha_documento_normalizada"))[:10]
            if not fecha_doc:
                fecha = parsear_fecha_documento(_texto(control.get("fecha")))
                fecha_doc = fecha.isoformat() if fecha else ""
            revision = _texto(control.get("revision"))
            revision_orden = int(revision) if revision.isdigit() and len(revision) < 18 else -1
            # El índice necesita metadatos, no todas las pruebas y auditorías en memoria.
            metadatos = {c: datos[c] for c in ("producto", "documento", "control_documental", "validacion_codigos_200",
                "versiones_idioma", "hii_generadas", "archivo_origen", "hash_sha256_origen", "hash_sha256_catalogo", "formato_origen") if c in datos}
            documentos.append(dict(ruta=ruta, tipo=tipo, datos=metadatos, nombre=nombre or codigo,
                codigo=codigo, cliente=clave(control.get("nombre_cliente")), revision=revision,
                revision_orden=revision_orden, idioma=leer_idioma_datos(datos),
                fecha_doc=fecha_doc, fecha_creacion=_fecha(datos.get("fecha_creacion_json")),
                hash_json=hashlib.sha256(contenido).hexdigest(), wip=es_wip(ruta)))
        except (OSError, UnicodeError, ValueError, TypeError) as exc:
            resumen["omitidos"].append({"archivo": ruta_portatil(ruta, base), "motivo": str(exc)})
    if archivos and not documentos:
        raise ValueError("No se pudo recuperar ningún JSON válido. Se conserva la base anterior.")
    resumen["documentos"] = len(documentos)
    resumen["wip"] = sum(d["wip"] for d in documentos)
    return documentos


def _orden(d):
    return (d["fecha_doc"], d["revision_orden"], d["fecha_creacion"], str(d["ruta"]).casefold())


def _registrar_documentos(con, documentos, base):
    con.executescript("""
        CREATE TABLE recuperacion_documentos (
            ruta TEXT PRIMARY KEY, archivo_json TEXT NOT NULL, tipo TEXT NOT NULL,
            producto_clave TEXT NOT NULL, producto_nombre TEXT NOT NULL, codigo TEXT NOT NULL,
            cliente TEXT NOT NULL, revision TEXT NOT NULL, revision_orden INTEGER NOT NULL,
            idioma TEXT NOT NULL, fecha_documento TEXT, fecha_creacion TEXT,
            hash_json TEXT NOT NULL, en_progreso INTEGER NOT NULL);
        CREATE VIEW recuperacion_actuales AS
            SELECT * FROM (SELECT d.*, ROW_NUMBER() OVER (
                PARTITION BY producto_clave, cliente, idioma, tipo
                ORDER BY fecha_documento DESC, revision_orden DESC, fecha_creacion DESC, ruta DESC
            ) AS orden_actual FROM recuperacion_documentos d WHERE en_progreso=0)
            WHERE orden_actual=1;
    """)
    from gestor_especificaciones import clave
    for d in documentos:
        con.execute("INSERT INTO recuperacion_documentos VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (ruta_portatil(d["ruta"], base), d["ruta"].name, d["tipo"], clave(d["codigo"] or d["nombre"]),
             d["nombre"], d["codigo"], d["cliente"], d["revision"], d["revision_orden"],
             d["idioma"], d["fecha_doc"], d["fecha_creacion"], d["hash_json"], int(d["wip"])))


def _registrar_extraccion(con, documentos, base, rutas):
    from gestor_especificaciones import clave, limpiar, registrar_historial_revisado
    # Original y alineado derivados del mismo JSON cuentan como una extracción.
    versiones = {}
    for d in sorted(documentos, key=lambda d: d["tipo"] != "original"):
        if d["wip"]:
            continue
        identidad = (d["ruta"].name, d["codigo"], d["cliente"], d["idioma"], d["revision"])
        versiones.setdefault(identidad, d)
    grupos = defaultdict(list)
    fuentes_actuales = {}
    for d in versiones.values():
        grupos[clave(d["nombre"])].append(d)
    for producto, grupo in grupos.items():
        for revision, d in enumerate(sorted(grupo, key=_orden)):
            datos = d["datos"]
            control, documento = datos.get("control_documental", {}), datos.get("documento", {})
            validacion = datos.get("validacion_codigos_200", {})
            if not isinstance(validacion, dict):
                validacion = {}
            origen = _texto(datos.get("archivo_origen"))
            hash_fuente = _huella(datos.get("hash_sha256_origen"))
            hash_catalogo = _huella(datos.get("hash_sha256_catalogo")) or None
            # Una extracción sin origen verificable nunca debe provocar un falso "sin cambios".
            ruta_origen = ruta_portatil(resolver_ruta(origen, base), base) if origen else "JSON_RECUPERADO/" + ruta_portatil(d["ruta"], base)
            hash_fuente = hash_fuente or "SIN_HUELLA:" + d["hash_json"]
            fuente = resolver_ruta(origen, base) if origen else None
            stat = fuente.stat() if fuente is not None and fuente.is_file() else None
            try:
                archivo_json = Path(os.path.relpath(d["ruta"], rutas["json_anteriores"])).as_posix()
            except ValueError:
                archivo_json = str(d["ruta"])
            valores = dict(ruta_relativa=ruta_origen, nombre_archivo=Path(origen).name if origen else d["ruta"].name,
                hash_actual=hash_fuente, hash_catalogo=hash_catalogo, tamano_bytes=stat.st_size if stat else 0,
                fecha_modificacion_ns=stat.st_mtime_ns if stat else 0, producto_clave=producto,
                producto_nombre=d["nombre"], nombre_fuente=limpiar(validacion.get("nombre_fuente")),
                resolucion_nombre=limpiar(validacion.get("resolucion_nombre")),
                part_id_fuente=limpiar(control.get("codigo_unico_producto_fuente", d["codigo"])),
                part_id_final=d["codigo"], validacion_catalogo=limpiar(validacion.get("estado")),
                revision_especificacion=d["revision"], codigo_formato=limpiar(documento.get("codigo_formato")),
                revision_formato=limpiar(documento.get("revision_formato")), formato_origen=limpiar(datos.get("formato_origen")),
                idioma=d["idioma"], ultima_revision_json=revision, ultimo_json=archivo_json,
                primera_deteccion=d["fecha_creacion"], ultima_revision_procesada=d["fecha_creacion"],
                estado="OK" if origen and not hash_fuente.startswith("SIN_HUELLA:") and hash_catalogo else "RECUPERADO")
            columnas = list(valores)
            con.execute(f"INSERT INTO archivos ({','.join(columnas)}) VALUES ({','.join('?' for _ in columnas)}) "
                "ON CONFLICT(ruta_relativa) DO UPDATE SET " + ",".join(f"{c}=excluded.{c}" for c in columnas if c not in {"ruta_relativa", "primera_deteccion"}), tuple(valores.values()))
            if ruta_origen not in fuentes_actuales or _orden(d) > fuentes_actuales[ruta_origen][0]:
                fuentes_actuales[ruta_origen] = (_orden(d), valores)
            archivo_id = con.execute("SELECT id FROM archivos WHERE ruta_relativa=?", (ruta_origen,)).fetchone()[0]
            con.execute("""INSERT INTO versiones (archivo_id,producto_clave,revision_json,hash_fuente,hash_catalogo,
                part_id_fuente,part_id_final,validacion_catalogo,nombre_fuente,resolucion_nombre,
                revision_especificacion,codigo_formato,revision_formato,formato_origen,idioma,
                fecha_documento,fecha_creacion,archivo_json) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (archivo_id, producto, revision, hash_fuente, hash_catalogo, valores["part_id_fuente"], d["codigo"],
                 valores["validacion_catalogo"], valores["nombre_fuente"], valores["resolucion_nombre"], d["revision"],
                 valores["codigo_formato"], valores["revision_formato"], valores["formato_origen"], d["idioma"],
                 d["fecha_doc"] or None, d["fecha_creacion"], archivo_json))
    # Un mismo Word puede haber cambiado de nombre de producto entre extracciones.
    # Su huella vigente debe ser la de su versión más reciente, incluso entre grupos.
    for _, valores in fuentes_actuales.values():
        columnas = [c for c in valores if c not in {"ruta_relativa", "primera_deteccion"}]
        con.execute("UPDATE archivos SET " + ",".join(f"{c}=?" for c in columnas) + " WHERE ruta_relativa=?",
                    (*[valores[c] for c in columnas], valores["ruta_relativa"]))
    # Evita que la primera ejecución reescriba JSON ya recuperados solo para reindexarlos.
    registrar_historial_revisado(con, rutas["json_anteriores"], None)


def _registrar_relaciones(con, documentos, base, resumen):
    from generador_especs_hii import clave_texto
    por_ruta = {d["ruta"]: d for d in documentos}
    grupos = defaultdict(list)
    for d in documentos:
        fila = (ruta_portatil(d["ruta"], base), clave_texto(d["codigo"]), clave_texto(d["revision"]), d["cliente"], d["idioma"])
        con.execute("INSERT INTO hii_especificaciones VALUES (?, ?, ?, ?, ?)", fila)
        if d["codigo"] and d["revision"] and not d["wip"]:
            grupos[(d["tipo"], *fila[1:4])].append(fila)
    for d in documentos:
        relaciones = d["datos"].get("versiones_idioma", {})
        if not isinstance(relaciones, dict):
            continue
        for referencias in relaciones.values():
            if isinstance(referencias, dict):
                referencias = [referencias]
            if not isinstance(referencias, list):
                continue
            for referencia in referencias:
                if not isinstance(referencia, dict) or not isinstance(referencia.get("archivo_json"), str):
                    continue
                ruta = resolver_ruta(referencia["archivo_json"], d["ruta"].parent)
                otro = por_ruta.get(ruta)
                if otro and clave_texto(d["codigo"]) == clave_texto(otro["codigo"]) and d["cliente"] == otro["cliente"] and d["idioma"] != otro["idioma"]:
                    pareja = tuple(sorted((ruta_portatil(d["ruta"], base), ruta_portatil(ruta, base))))
                    con.execute("INSERT OR IGNORE INTO hii_traducciones VALUES (?, ?)", pareja)
                else:
                    resumen["advertencias"].append(f"Referencia bilingüe no recuperada: {d['ruta'].name} → {referencia['archivo_json']}")
    for grupo in grupos.values():
        es = [f for f in grupo if f[4] == "ESPANOL"]
        en = [f for f in grupo if f[4] == "INGLES"]
        if len(es) == len(en) == 1:
            pareja = tuple(sorted((es[0][0], en[0][0])))
            if not con.execute("SELECT 1 FROM hii_traducciones WHERE (origen=? AND destino=?) OR (origen=? AND destino=?)", (*pareja, pareja[1], pareja[0])).fetchone():
                con.execute("INSERT INTO hii_traducciones VALUES (?, ?)", pareja)
                con.execute("INSERT INTO hii_enlaces_automaticos VALUES (?, ?)", pareja)
        elif es and en:
            resumen["advertencias"].append("Pareja ESP/ENG ambigua; se conservan todas las versiones sin inventar relaciones: " + ", ".join(f[0] for f in grupo))
    resumen["relaciones"] = con.execute("SELECT COUNT(*) FROM hii_traducciones").fetchone()[0]


def _validar(con):
    resultado = con.execute("PRAGMA integrity_check").fetchall()
    if [f[0] for f in resultado] != ["ok"] or con.execute("PRAGMA foreign_key_check").fetchone():
        raise ValueError("La base reconstruida no superó la validación de integridad.")


def _registrar_hii_generadas(con, documentos, base, resumen):
    """Recupera el índice desde el historial, incluso si el Excel ya no está disponible."""
    registros = []
    for d in documentos:
        historial = d["datos"].get("hii_generadas", [])
        if not isinstance(historial, list):
            resumen["advertencias"].append(f"Historial HII inválido: {d['ruta'].name}")
            continue
        for registro in historial:
            try:
                if not isinstance(registro, dict) or not all(isinstance(registro.get(c), str) and registro[c].strip()
                    for c in ("id", "archivo", "fecha_generacion", "version_generador", "hash_especificacion")):
                    raise ValueError("registro incompleto")
                fecha = datetime.fromisoformat(registro["fecha_generacion"])
                if fecha.tzinfo is None:
                    raise ValueError("fecha sin zona horaria")
                ruta = resolver_ruta(registro["archivo"], d["ruta"].parent)
                registros.append((fecha.timestamp(), str(d["ruta"]), ruta, d, fecha))
            except (ValueError, OSError, OverflowError) as exc:
                resumen["advertencias"].append(f"HII no recuperada de {d['ruta'].name}: {exc}")
    # La tabla existente representa el archivo actual: prevalece la última generación por ruta.
    for _, _, ruta, d, fecha in sorted(registros, key=lambda r: (r[0], r[1])):
        con.execute("""INSERT INTO hii_generadas VALUES (?, ?, ?, ?)
            ON CONFLICT(ruta) DO UPDATE SET especificacion=excluded.especificacion,
            idioma=excluded.idioma, fecha=excluded.fecha""",
            (ruta_portatil(ruta, base), ruta_portatil(d["ruta"], base), d["idioma"], fecha.strftime("%d/%m/%y")))
        if not ruta.is_file():
            resumen["advertencias"].append(f"HII registrada pero no localizada: {ruta}")
    resumen["hii_generadas"] = con.execute("SELECT COUNT(*) FROM hii_generadas").fetchone()[0]


def _copiar_base(origen, destino, timeout=30):
    limite = time.monotonic() + timeout
    def avanzar(estado, pendientes, total):
        if time.monotonic() > limite:
            raise TimeoutError("Se agotó el tiempo de espera al respaldar la base. Cierra otros procesos que la estén utilizando.")
    origen.backup(destino, pages=256, progress=avanzar, sleep=.05)


def _publicar(temporal, destino, base, resumen):
    respaldo = base / "RESPALDOS_BASE_DATOS" / (datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid.uuid4().hex[:8])
    if destino.exists():
        respaldo.mkdir(parents=True)
        try:
            with closing(sqlite3.connect(destino)) as anterior:
                _validar(anterior)
                with closing(sqlite3.connect(respaldo / destino.name)) as copia:
                    _copiar_base(anterior, copia)
        except sqlite3.OperationalError:
            # Un bloqueo o fallo de E/S no demuestra corrupción: no sustituir el archivo.
            raise
        except (sqlite3.DatabaseError, ValueError):
            # Una base dañada solo se sustituye si no tiene diarios que puedan estar activos.
            if any(Path(str(destino) + sufijo).exists() for sufijo in ("-wal", "-shm", "-journal")):
                raise RuntimeError("La base está dañada y tiene archivos auxiliares. Cierra todas las aplicaciones y conserva la base y sus auxiliares antes de retirarlos.")
            shutil.copy2(destino, respaldo / (destino.name + ".corrupta"))
            os.replace(temporal, destino)
        else:
            with closing(sqlite3.connect(temporal)) as origen, closing(sqlite3.connect(destino)) as actual:
                _copiar_base(origen, actual)
                _validar(actual)
        resumen["respaldo"] = str(respaldo)
    else:
        auxiliares = [Path(str(destino) + sufijo) for sufijo in ("-wal", "-shm", "-journal")]
        if any(p.exists() for p in auxiliares):
            raise RuntimeError("Falta la base, pero quedan sus archivos auxiliares. Cierra todos los programas y conserva esos archivos antes de retirarlos; no se reutilizarán automáticamente.")
        # rename no sobrescribe una base creada por otro proceso en Windows.
        temporal.rename(destino)


def motivo_recuperacion(base=BASE, *, destino=None):
    """Consulta registros reales; el tamaño del .db no revela si contiene especificaciones."""
    destino = resolver_ruta(destino, base) if destino is not None else resolver_ruta("especificaciones.db", base)
    if not destino.exists():
        return "ausente"
    try:
        with closing(sqlite3.connect(destino.as_uri() + "?mode=ro", uri=True, timeout=.1)) as con:
            tablas = {f[0] for f in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            for tabla in ("archivos", "versiones", "hii_especificaciones", "recuperacion_documentos", "hii_generadas", "hii_traducciones"):
                if tabla in tablas and con.execute(f"SELECT 1 FROM {tabla} LIMIT 1").fetchone():
                    return None
        return "vacia"
    except sqlite3.Error:
        # Una base bloqueada o ilegible no demuestra que esté vacía.
        # Se conserva y puede recuperarse mediante la opción manual.
        return None


def _hay_json(rutas):
    return any(ruta.is_file() and ruta.suffix.lower() == ".json" and not ruta.name.startswith(".")
        for clave in ("json_anteriores", "json_alineadas") for ruta in rutas[clave].rglob("*"))


def necesita_recuperacion(base=BASE, *, destino=None, rutas=None):
    motivo = motivo_recuperacion(base, destino=destino)
    if motivo == "ausente":
        return True
    if motivo == "vacia":
        rutas = rutas_configuradas(base) if rutas is None else {k: resolver_ruta(v, base) for k, v in rutas.items()}
        return _hay_json(rutas)
    return False


def reconstruir(base=BASE, *, destino=None, rutas=None, reemplazar=False, al_progreso=None, automatica=False):
    """Construye y valida en un archivo aislado; solo después publica la base."""
    base = resolver_ruta(base)
    destino = resolver_ruta(destino, base) if destino is not None else base / "especificaciones.db"
    destino.parent.mkdir(parents=True, exist_ok=True)
    rutas = rutas_configuradas(base) if rutas is None else {k: resolver_ruta(v, base) for k, v in rutas.items()}
    with _bloqueo(destino):
        # Revalidar bajo el bloqueo: otro proceso pudo haberla recuperado primero.
        if automatica:
            if not necesita_recuperacion(base, destino=destino, rutas=rutas):
                return None
        elif destino.exists() and not reemplazar:
            return None
        motivo = motivo_recuperacion(base, destino=destino) if automatica else "manual"
        resumen = {"documentos": 0, "wip": 0, "relaciones": 0, "omitidos": [], "advertencias": [], "base": str(destino), "respaldo": None, "motivo": motivo}
        documentos = _leer_documentos(base, rutas, al_progreso, resumen)
        if destino.exists() and not documentos:
            raise ValueError("No se encontraron JSON para reconstruir. Se conserva la base anterior; revisa las carpetas configuradas.")
        with tempfile.NamedTemporaryFile(prefix=".reconstruccion_", suffix=".db", dir=destino.parent, delete=False) as archivo:
            temporal = Path(archivo.name)
        try:
            from gestor_especificaciones import conectar, leer_catalogo, registrar_catalogo, APP_VERSION
            import generador_hii as hii
            with closing(conectar(temporal, recuperar=False)) as con:
                _registrar_documentos(con, documentos, base)
                _registrar_extraccion(con, documentos, base, rutas)
                catalogo = base / "CODIGOS_200.xlsx"
                if catalogo.is_file():
                    try:
                        registros, digest = leer_catalogo(catalogo)
                        registrar_catalogo(con, registros, digest, datetime.now().isoformat(timespec="seconds"), catalogo)
                        con.execute("INSERT OR REPLACE INTO gestor_estado VALUES ('historial_revisado', ?)", (f"{APP_VERSION}:{digest}",))
                    except (OSError, ValueError, BadZipFile) as exc:
                        resumen["advertencias"].append(f"Catálogo no recuperado: {exc}")
                con.commit()
            with hii.conectar(temporal, recuperar=False) as con:
                _registrar_relaciones(con, documentos, base, resumen)
                _registrar_hii_generadas(con, documentos, base, resumen)
                resumen["versiones"] = con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0]
                _validar(con)
            with closing(sqlite3.connect(temporal)) as con:
                con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                con.execute("PRAGMA journal_mode=DELETE")
            _publicar(temporal, destino, base, resumen)
            resumen["advertencias"].append("Las HII sin historial en los JSON y los eventos que solo estaban en la base anterior no se pueden recuperar con certeza.")
            reporte = base / "REPORTE_RECONSTRUCCION_BASE.json"
            try:
                reporte.write_text(json.dumps(resumen, ensure_ascii=False, indent=2), encoding="utf-8")
            except OSError as exc:
                resumen["advertencias"].append(f"No se pudo guardar el reporte: {exc}")
            return resumen
        finally:
            for ruta in (temporal, Path(str(temporal) + "-wal"), Path(str(temporal) + "-shm"), Path(str(temporal) + "-journal")):
                ruta.unlink(missing_ok=True)


def asegurar_base(base=BASE, *, destino=None, rutas=None, al_progreso=None):
    destino = resolver_ruta(destino, base) if destino is not None else resolver_ruta("especificaciones.db", base)
    if not necesita_recuperacion(base, destino=destino, rutas=rutas):
        return None
    return reconstruir(base, destino=destino, rutas=rutas, al_progreso=al_progreso, automatica=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", type=resolver_ruta, default=BASE)
    parser.add_argument("--reemplazar", action="store_true", help="Reconstruye una base existente y conserva un respaldo.")
    args = parser.parse_args()
    try:
        resultado = reconstruir(args.base, reemplazar=True) if args.reemplazar else asegurar_base(args.base)
        print(json.dumps(resultado, ensure_ascii=False, indent=2) if resultado else "No se requiere reconstrucción automática; la base no se modificó.")
        return 0
    except (OSError, ValueError, sqlite3.Error, RuntimeError) as exc:
        print(f"No se pudo reconstruir la base: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
