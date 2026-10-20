"""Recupera SQLite desde JSON y aparta duplicados confirmados en REVISAR/DUPLICADOS."""
from __future__ import annotations

import argparse
import errno
import hashlib
import json
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import time
import tkinter as tk
import unicodedata
import uuid
from collections import defaultdict
from contextlib import closing, contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from zipfile import BadZipFile

from rutas_proyecto import BASE, resolver_ruta, ruta_portatil, rutas_configuradas, reubicar_ruta_heredada
from especificaciones_wip import es_wip


if __name__ == "__main__":
    sys.modules.setdefault("reconstruir_base_datos", sys.modules[__name__])


EVENTO_CATALOGO = "<<CatalogoEspecificacionesActualizado>>"


def suscribir_catalogo(ventana, recargar):
    """Escucha mientras exista la ventana y libera su suscripción al cerrarla."""
    raiz = ventana._root()

    def actualizar(_evento):
        if ventana.winfo_exists():
            recargar()

    enlace = raiz.bind(EVENTO_CATALOGO, actualizar, add="+")

    def cerrar(evento):
        nonlocal enlace
        if evento.widget is ventana and enlace is not None:
            try:
                raiz.unbind(EVENTO_CATALOGO, enlace)
            except tk.TclError:
                pass  # La raíz puede estar cerrándose junto con las ventanas.
            enlace = None

    ventana.bind("<Destroy>", cerrar, add="+")


def notificar_catalogo(raiz):
    """Encola la recarga tras finalizar los cambios en los archivos y SQLite."""
    raiz.event_generate(EVENTO_CATALOGO, when="tail")


METADATOS = {
    "esquema", "archivo_origen", "archivo_json", "hash_sha256_origen",
    "hash_sha256_catalogo", "formato_origen", "conversion_temporal",
    "revision_json", "fecha_creacion_json", "extractor_version", "version_lector",
    "estructura_detectada", "validacion_codigos_200", "alineacion_metodos_lab",
    "versiones_idioma", "hii_generadas", "hii_registros_heredados",
}


def _normalizar(valor):
    if isinstance(valor, str):
        return re.sub(r"\s+", " ", unicodedata.normalize("NFC", valor)).strip()
    if isinstance(valor, list):
        return [_normalizar(v) for v in valor]
    if isinstance(valor, dict):
        return {k: _normalizar(v) for k, v in valor.items()}
    return valor


def contenido_comparable(datos):
    """Devuelve datos conservadores, o None si la extracción no permite confirmarlos."""
    from generador_especs_hii import leer_idioma_datos
    if not isinstance(datos, dict):
        return None
    producto, control, documento = (datos.get(c) for c in ("producto", "control_documental", "documento"))
    pruebas = datos.get("pruebas")
    if not all(isinstance(v, dict) for v in (producto, control, documento)):
        return None
    if not producto.get("nombre") or not control.get("codigo_unico_producto"):
        return None
    if not isinstance(pruebas, list) or not pruebas:
        return None
    campos = {"prueba", "metodo", "unidades", "especificacion"}
    if any(not isinstance(p, dict) or not campos.issubset(p) for p in pruebas):
        return None
    if not any(p.get("prueba") and p.get("metodo") and p.get("especificacion") for p in pruebas):
        return None
    try:
        idioma = leer_idioma_datos(datos)
    except ValueError:
        return None
    resultado = {k: v for k, v in datos.items() if k not in METADATOS}
    resultado["documento"] = dict(documento)
    resultado["documento"]["idioma"] = idioma
    resultado["control_documental"] = {k: v for k, v in control.items() if not k.endswith("_fuente")}
    resultado["pruebas"] = [
        {k: v for k, v in p.items() if k not in {"numero", "fila_origen"}} for p in pruebas
    ]
    if "alineacion_metodos_lab" in datos:
        alineacion = datos["alineacion_metodos_lab"]
        detalle = alineacion.get("detalle") if isinstance(alineacion, dict) else None
        if (not isinstance(detalle, list) or len(detalle) != len(pruebas)
                or any(not isinstance(d, dict) or not isinstance(d.get("valores_alineados"), dict)
                       or not campos.issubset(d["valores_alineados"]) for d in detalle)):
            return None
        # Esta instantánea es la que muestra el selector, y puede incluir ediciones manuales.
        resultado["valores_alineados"] = [d["valores_alineados"] for d in detalle]
    return _normalizar(resultado)


def _bytes_json(datos):
    return json.dumps(datos, sort_keys=True, ensure_ascii=False, separators=(",", ":"), allow_nan=False).encode("utf-8")


def _sha(ruta):
    digest = hashlib.sha256()
    with ruta.open("rb") as archivo:
        for bloque in iter(lambda: archivo.read(1024 * 1024), b""):
            digest.update(bloque)
    return digest.hexdigest()


def _firma(ruta):
    estado = ruta.stat()
    return estado.st_size, estado.st_mtime_ns


def _dentro(ruta, carpetas):
    return any(ruta.is_relative_to(carpeta) for carpeta in carpetas)


@dataclass
class PlanDuplicados:
    destino: Path
    equivalencias: dict = field(default_factory=dict)
    word_duplicados: dict = field(default_factory=dict)
    firmas: dict = field(default_factory=dict)
    ediciones: dict = field(default_factory=dict)
    revisados: int = 0
    omitidos: int = 0
    advertencias: list = field(default_factory=list)

    def resumen(self):
        return {"json_duplicados": len(self.equivalencias), "word_duplicados": len(self.word_duplicados),
                "revisados": self.revisados, "omitidos": self.omitidos,
                "advertencias": list(self.advertencias), "destino": str(self.destino)}


def _antiguedad_registrada(db, base, rutas):
    json_registrados, word_registrados = {}, {}
    if not db.is_file():
        return json_registrados, word_registrados
    try:
        with closing(sqlite3.connect(db)) as con:
            tablas = {f[0] for f in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if "versiones" in tablas:
                for archivo, fecha in con.execute("SELECT archivo_json,fecha_creacion FROM versiones ORDER BY id"):
                    json_registrados.setdefault(resolver_ruta(archivo, rutas["json_anteriores"]), fecha or "")
            if "archivos" in tablas:
                for archivo, fecha in con.execute("SELECT ruta_relativa,primera_deteccion FROM archivos ORDER BY id"):
                    word_registrados.setdefault(resolver_ruta(archivo, base), fecha or "")
    except sqlite3.DatabaseError:
        # Una reconstrucción debe poder revisar también una base dañada.
        pass
    return json_registrados, word_registrados


def planificar_duplicados(base, rutas, db, al_progreso=None):
    """Solo lee. Separa originales/alineados y protege orígenes compartidos y WIP."""
    base, db = resolver_ruta(base), resolver_ruta(db)
    destino = base / "REVISAR" / "DUPLICADOS"
    plan = PlanDuplicados(destino)
    from gestor_especificaciones import archivo_bloqueado_word
    anteriores, fuentes_registradas = _antiguedad_registrada(db, base, rutas)
    carpetas_word = {rutas["especificaciones_viejas"].resolve(),
                     (rutas["especificaciones_viejas"].parent / "ESPECIFICACIONES_VIEJAS_ENG").resolve()}
    archivos, documentos, referencias = {}, {}, defaultdict(set)
    grupos = defaultdict(list)
    for tipo, carpeta in (("original", rutas["json_anteriores"]), ("alineada", rutas["json_alineadas"])):
        for ruta in sorted(carpeta.rglob("*")):
            ruta = ruta.resolve()
            if ruta.is_file() and ruta.suffix.lower() == ".json" and not ruta.name.startswith(".") and not ruta.is_relative_to(destino):
                archivos.setdefault(ruta, tipo)
    for indice, (ruta, tipo) in enumerate(archivos.items(), 1):
        if al_progreso:
            al_progreso(indice, len(archivos), ruta.name)
        try:
            firma = _firma(ruta)
            bruto = ruta.read_bytes()
            datos = json.loads(bruto.decode("utf-8-sig"))
            if not isinstance(datos, dict):
                raise ValueError("El JSON no es un objeto")
            origen = datos.get("archivo_origen")
            fuente = reubicar_ruta_heredada(origen, base) if isinstance(origen, str) and origen.strip() else None
            if fuente:
                referencias[fuente].add(ruta)
            documentos[ruta] = (datos, fuente, bruto)
            plan.firmas[ruta] = firma
            plan.revisados += 1
            comparable = None if es_wip(ruta) else contenido_comparable(datos)
            if comparable is None:
                plan.omitidos += 1
                continue
            canonico = _bytes_json(comparable)
            grupos[(tipo, hashlib.sha256(canonico).digest())].append((ruta, canonico))
        except (OSError, ValueError, TypeError) as exc:
            plan.omitidos += 1
            plan.advertencias.append(f"No se pudo revisar {ruta.name}: {exc}")
    def orden(ruta):
        fecha = anteriores.get(ruta) or documentos[ruta][0].get("fecha_creacion_json") or "9999"
        return ruta not in anteriores, str(fecha), str(ruta).casefold()
    for tipo in ("original", "alineada"):
        fuentes_conservadas = {fuente for r, (_, fuente, _) in documentos.items()
                               if archivos[r] == "original" and r not in plan.equivalencias and fuente}
        for (tipo_grupo, _), grupo in grupos.items():
            if tipo_grupo != tipo or len(grupo) < 2:
                continue
            por_contenido = {}
            def orden_grupo(item):
                ruta = item[0]
                return (tipo == "alineada" and documentos[ruta][1] not in fuentes_conservadas, *orden(ruta))
            for ruta, contenido in sorted(grupo, key=orden_grupo):
                conservada = por_contenido.setdefault(contenido, ruta)
                if conservada != ruta:
                    plan.equivalencias[ruta] = conservada

    # Un Word con JSON propios solo se mueve si TODOS ellos son copias descartadas.
    # Se verifican los bytes actuales contra la huella de la extracción antes de moverlo.
    candidatos = set()
    huellas_word = {}
    def huella_word(ruta):
        if ruta not in huellas_word:
            firma = _firma(ruta)
            huellas_word[ruta] = _sha(ruta)
            if _firma(ruta) != firma:
                raise RuntimeError(f"El Word cambió durante la revisión: {ruta.name}")
            plan.firmas[ruta] = firma
        return huellas_word[ruta]
    for fuente, refs in referencias.items():
        if refs.issubset(plan.equivalencias) and fuente.is_file() and _dentro(fuente, carpetas_word):
            if archivo_bloqueado_word(fuente) is not None:
                continue
            actual = huella_word(fuente)
            if all(documentos[r][0].get("hash_sha256_origen") == actual for r in refs):
                candidatos.add(fuente)
            else:
                plan.advertencias.append(f"Word modificado desde su extracción; se conserva: {fuente.name}")
    for fuente in candidatos:
        # La copia de referencia sigue disponible aunque no conserve su Word original.
        plan.word_duplicados[fuente] = None

    # Copias Word exactas todavía sin JSON: tamaño primero, SHA-256 solo en tamaños repetidos.
    tamanos = defaultdict(set)
    for carpeta in carpetas_word:
        for ruta in carpeta.rglob("*"):
            ruta = ruta.resolve()
            if ruta.is_file() and ruta.suffix.lower() in {".doc", ".docx", ".docm"} and not ruta.name.startswith("~$") and not ruta.is_relative_to(destino):
                tamanos[_firma(ruta)[0]].add(ruta)
    total_word = sum(len(g) for tamano, g in tamanos.items() if tamano and len(g) > 1)
    indice_word = 0
    for tamano, grupo in tamanos.items():
        if not tamano or len(grupo) < 2:
            continue
        hashes = defaultdict(list)
        for ruta in grupo:
            indice_word += 1
            if al_progreso:
                al_progreso(indice_word, total_word, ruta.name)
            hashes[huella_word(ruta)].append(ruta)
        for iguales in hashes.values():
            iguales.sort(key=lambda r: (r in plan.word_duplicados, not bool(referencias[r] - plan.equivalencias.keys()),
                                       r not in fuentes_registradas, fuentes_registradas.get(r, ""), str(r).casefold()))
            for ruta in iguales[1:]:
                if not referencias[ruta] and archivo_bloqueado_word(ruta) is None:
                    plan.word_duplicados[ruta] = iguales[0]

    # Las referencias de idioma de los JSON conservados deben seguir apuntando a archivos activos.
    for ruta, (datos, _, bruto) in documentos.items():
        if ruta in plan.equivalencias or not isinstance(datos.get("versiones_idioma"), dict):
            continue
        cambio = False
        for referencias_idioma in datos["versiones_idioma"].values():
            for ref in ([referencias_idioma] if isinstance(referencias_idioma, dict) else referencias_idioma if isinstance(referencias_idioma, list) else []):
                if isinstance(ref, dict) and isinstance(ref.get("archivo_json"), str):
                    anterior = resolver_ruta(ref["archivo_json"], ruta.parent)
                    if anterior in plan.equivalencias:
                        nueva = plan.equivalencias[anterior]
                        try:
                            ref["archivo_json"] = Path(os.path.relpath(nueva, ruta.parent)).as_posix()
                        except ValueError:
                            ref["archivo_json"] = str(nueva)
                        cambio = True
        if cambio:
            plan.ediciones[ruta] = (bruto, (json.dumps(datos, ensure_ascii=False, indent=2) + "\n").encode("utf-8"))
    cache = rutas["json_alineadas"] / ".cache_alineacion"
    if plan.equivalencias and cache.is_file():
        try:
            firma = _firma(cache)
            bruto = cache.read_bytes()
            entradas = json.loads(bruto.decode("utf-8"))
            if isinstance(entradas, dict):
                vigentes = {k: v for k, v in entradas.items()
                            if resolver_ruta(k, base) not in plan.equivalencias
                            and not (isinstance(v, dict) and isinstance(v.get("origen"), str)
                                     and resolver_ruta(v["origen"], base) in plan.equivalencias)}
                if len(vigentes) != len(entradas):
                    plan.firmas[cache] = firma
                    plan.ediciones[cache] = (bruto, _bytes_json(vigentes))
        except (OSError, ValueError, TypeError) as exc:
            plan.advertencias.append(f"No se pudo depurar la caché de alineación: {exc}")
    return plan


def _mover_archivo(origen, destino):
    try:
        origen.rename(destino)
    except OSError as exc:
        if exc.errno != errno.EXDEV and getattr(exc, "winerror", None) != 17:
            raise
        # Las carpetas configuradas también pueden estar en otra unidad.
        creado = False
        try:
            with destino.open("xb") as salida:
                creado = True
                with origen.open("rb") as entrada:
                    shutil.copyfileobj(entrada, salida, 1024 * 1024)
            shutil.copystat(origen, destino)
            origen.unlink()
        except BaseException:
            if creado:
                destino.unlink(missing_ok=True)
            raise


def _escribir_atomico(ruta, contenido):
    temporal = None
    try:
        with tempfile.NamedTemporaryFile(prefix=".duplicados_", suffix=".tmp", dir=ruta.parent, delete=False) as archivo:
            temporal = Path(archivo.name)
            archivo.write(contenido)
            archivo.flush()
            os.fsync(archivo.fileno())
        os.replace(temporal, ruta)
    finally:
        if temporal is not None:
            temporal.unlink(missing_ok=True)


@contextmanager
def aplicar_plan(plan):
    """No sobrescribe archivos; deshace movimientos y enlaces si falla SQLite o la publicación."""
    movimientos, editados = [], []
    try:
        for ruta, firma in plan.firmas.items():
            if not ruta.is_file() or _firma(ruta) != firma:
                raise RuntimeError(f"El archivo cambió durante la revisión: {ruta.name}. Vuelve a revisar.")
        for ruta in sorted(set(plan.equivalencias) | set(plan.word_duplicados), key=str):
            plan.destino.mkdir(parents=True, exist_ok=True)
            destino = plan.destino / ruta.name
            indice = 1
            while destino.exists():
                destino = plan.destino / f"{ruta.stem}_{indice:02d}{ruta.suffix}"
                indice += 1
            # Solo se mueven archivos individuales resueltos por el plan, nunca carpetas.
            if ruta.is_relative_to(plan.destino) or not destino.resolve().is_relative_to(plan.destino.resolve()):
                raise ValueError("Destino de duplicados inválido")
            # rename no sobrescribe en Windows y permite revertir sin copiar documentos grandes.
            _mover_archivo(ruta, destino)
            movimientos.append((ruta, destino))
        for ruta, (anterior, nuevo) in plan.ediciones.items():
            editados.append((ruta, anterior))
            _escribir_atomico(ruta, nuevo)
        yield plan.resumen()
    except BaseException as exc:
        errores = []
        for ruta, anterior in reversed(editados):
            try:
                _escribir_atomico(ruta, anterior)
            except OSError as error:
                errores.append(str(error))
        for ruta, destino in reversed(movimientos):
            try:
                if ruta.exists():
                    raise FileExistsError(f"Ya existe otro archivo en {ruta}")
                _mover_archivo(destino, ruta)
            except OSError as error:
                errores.append(str(error))
        if errores:
            raise RuntimeError(f"{exc}; no se pudo revertir por completo: {'; '.join(errores)}") from exc
        raise


def depurar_indices(con, plan, base, rutas):
    """Elimina del catálogo activo las copias trasladadas; no registra sus destinos."""
    tablas = {f[0] for f in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    eliminados = set(plan.equivalencias)
    def rutas_json(valor, carpeta=base):
        return resolver_ruta(valor, carpeta) in eliminados
    if "versiones" in tablas:
        ids = [f[0] for f in con.execute("SELECT id,archivo_json FROM versiones") if rutas_json(f[1], rutas["json_anteriores"])]
        for identidad in ids:
            if "gestor_historial_cache" in tablas:
                con.execute("DELETE FROM gestor_historial_cache WHERE version_id=?", (identidad,))
            con.execute("DELETE FROM versiones WHERE id=?", (identidad,))
    if "archivos" in tablas:
        for identidad, origen, ultimo in con.execute("SELECT id,ruta_relativa,ultimo_json FROM archivos").fetchall():
            movido = resolver_ruta(origen, base) in plan.word_duplicados
            if movido or (ultimo and rutas_json(ultimo, rutas["json_anteriores"])):
                if movido and "versiones" in tablas:
                    if "gestor_historial_cache" in tablas:
                        con.execute("DELETE FROM gestor_historial_cache WHERE version_id IN (SELECT id FROM versiones WHERE archivo_id=?)", (identidad,))
                    con.execute("DELETE FROM versiones WHERE archivo_id=?", (identidad,))
                superviviente = con.execute("SELECT archivo_json,revision_json FROM versiones WHERE archivo_id=? ORDER BY revision_json DESC LIMIT 1", (identidad,)).fetchone() if "versiones" in tablas else None
                if superviviente:
                    con.execute("UPDATE archivos SET ultimo_json=?,ultima_revision_json=? WHERE id=?", (*superviviente, identidad))
                else:
                    con.execute("DELETE FROM archivos WHERE id=?", (identidad,))
    for tabla in ("hii_traducciones", "hii_enlaces_automaticos"):
        if tabla in tablas:
            for origen, destino in con.execute(f"SELECT origen,destino FROM {tabla}").fetchall():
                if rutas_json(origen) or rutas_json(destino):
                    con.execute(f"DELETE FROM {tabla} WHERE origen=? AND destino=?", (origen, destino))
    for tabla, columna in (("hii_generadas", "especificacion"), ("hii_especificaciones", "ruta"), ("recuperacion_documentos", "ruta")):
        if tabla in tablas:
            for (ruta,) in con.execute(f"SELECT {columna} FROM {tabla}").fetchall():
                if rutas_json(ruta):
                    con.execute(f"DELETE FROM {tabla} WHERE {columna}=?", (ruta,))
    # Se recrearán las parejas de idioma al volver a cargar el selector.
    if "gestor_estado" in tablas and eliminados:
        con.execute("DELETE FROM gestor_estado WHERE clave='historial_revisado'")


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


def _leer_documentos(base, rutas, progreso, resumen, *, excluidos=(), ediciones=None):
    from gestor_especificaciones import clave, parsear_fecha_documento
    from generador_especs_hii import leer_idioma_datos
    from especificaciones_wip import es_wip
    archivos = {}
    for tipo, carpeta in (("original", rutas["json_anteriores"]), ("alineada", rutas["json_alineadas"])):
        for ruta in sorted(carpeta.rglob("*")):
            if (ruta.is_file() and ruta.suffix.lower() == ".json" and not ruta.name.startswith(".")
                    and ruta.resolve() not in excluidos and not ruta.resolve().is_relative_to(base / "REVISAR")):
                archivos.setdefault(ruta.resolve(), tipo)
    documentos = []
    for indice, (ruta, tipo) in enumerate(archivos.items(), 1):
        if progreso:
            progreso(indice, len(archivos), ruta.name)
        try:
            contenido = ediciones[ruta][1] if ediciones and ruta in ediciones else ruta.read_bytes()
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
                "versiones_idioma", "hii_generadas", "archivo_origen", "hash_sha256_origen", "hash_sha256_catalogo",
                "formato_origen", "version_lector") if c in datos}
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
            con.execute("UPDATE versiones SET version_lector=? WHERE archivo_json=?",
                        (datos.get("version_lector"), archivo_json))
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
        plan = planificar_duplicados(base, rutas, destino, al_progreso)
        resumen["duplicados"] = plan.resumen()
        documentos = _leer_documentos(base, rutas, al_progreso, resumen,
            excluidos=plan.equivalencias, ediciones=plan.ediciones)
        if destino.exists() and not documentos:
            raise ValueError("No se encontraron JSON para reconstruir. Se conserva la base anterior; revisa las carpetas configuradas.")
        with tempfile.NamedTemporaryFile(prefix=".reconstruccion_", suffix=".db", dir=destino.parent, delete=False) as archivo:
            temporal = Path(archivo.name)
        try:
            from gestor_especificaciones import conectar, leer_catalogo, registrar_catalogo, VERSION_NORMALIZACION
            import generador_hii as hii
            with closing(conectar(temporal, recuperar=False)) as con:
                con.execute("PRAGMA temp_store=MEMORY")
                _registrar_extraccion(con, documentos, base, rutas)
                catalogo = base / "CODIGOS_200.xlsx"
                if catalogo.is_file():
                    try:
                        registros, digest = leer_catalogo(catalogo)
                        registrar_catalogo(con, registros, digest, datetime.now().isoformat(timespec="seconds"), catalogo)
                        con.execute("INSERT OR REPLACE INTO gestor_estado VALUES ('historial_revisado', ?)", (f"{VERSION_NORMALIZACION}:{digest}",))
                    except (OSError, ValueError, BadZipFile) as exc:
                        resumen["advertencias"].append(f"Catálogo no recuperado: {exc}")
                hii.preparar_registro(con)
                _registrar_relaciones(con, documentos, base, resumen)
                _registrar_hii_generadas(con, documentos, base, resumen)
                resumen["versiones"] = con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0]
                _validar(con)
                con.commit()
            with closing(sqlite3.connect(temporal)) as con:
                con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                con.execute("PRAGMA journal_mode=DELETE")
            with aplicar_plan(plan):
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


def revisar_duplicados(base=BASE, *, destino=None, rutas=None, al_progreso=None):
    """Revisa las extracciones y limpia sus índices sin reconstruir el resto del historial."""
    base = resolver_ruta(base)
    db = resolver_ruta(destino, base) if destino is not None else base / "especificaciones.db"
    db.parent.mkdir(parents=True, exist_ok=True)
    rutas = rutas_configuradas(base) if rutas is None else {k: resolver_ruta(v, base) for k, v in rutas.items()}
    with _bloqueo(db):
        plan = planificar_duplicados(base, rutas, db, al_progreso)
        if not plan.equivalencias and not plan.word_duplicados:
            return plan.resumen()
        if not db.exists():
            with aplicar_plan(plan) as resumen:
                return resumen
        with closing(sqlite3.connect(db, timeout=30)) as con:
            con.execute("PRAGMA foreign_keys=ON")
            _validar(con)
            try:
                con.execute("BEGIN IMMEDIATE")
                depurar_indices(con, plan, base, rutas)
                _validar(con)
                with aplicar_plan(plan) as resumen:
                    con.commit()
                return resumen
            except BaseException:
                con.rollback()
                raise


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", type=resolver_ruta, default=BASE)
    parser.add_argument("--reemplazar", action="store_true", help="Reconstruye una base existente y conserva un respaldo.")
    parser.add_argument("--revisar-duplicados", action="store_true", help="Mueve copias confirmadas a REVISAR/DUPLICADOS y las quita del índice activo.")
    args = parser.parse_args()
    try:
        if args.revisar_duplicados:
            resultado = revisar_duplicados(args.base)
        else:
            resultado = reconstruir(args.base, reemplazar=True) if args.reemplazar else asegurar_base(args.base)
        print(json.dumps(resultado, ensure_ascii=False, indent=2) if resultado else "No se requiere reconstrucción automática; la base no se modificó.")
        return 0
    except (OSError, ValueError, sqlite3.Error, RuntimeError) as exc:
        print(f"No se pudo reconstruir la base: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
