"""Relaciones entre especificaciones bilingües y sus HII en SQLite."""
import json
import sqlite3
import os
import tempfile
from contextlib import contextmanager
from pathlib import Path

from generador_especs_hii import clave_texto, leer_idioma_datos


@contextmanager
def conectar(ruta):
    con = sqlite3.connect(ruta, timeout=30)
    try:
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA busy_timeout=30000")
        con.execute("PRAGMA foreign_keys=ON")
        modo_journal = con.execute("PRAGMA journal_mode").fetchone()[0]
        if str(modo_journal).lower() != "wal":
            con.execute("PRAGMA journal_mode=WAL")
        con.executescript("""
            CREATE TABLE IF NOT EXISTS hii_especificaciones (
                ruta TEXT PRIMARY KEY, codigo TEXT NOT NULL, revision TEXT NOT NULL,
                cliente TEXT NOT NULL, idioma TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS hii_traducciones (
                origen TEXT NOT NULL REFERENCES hii_especificaciones(ruta),
                destino TEXT NOT NULL REFERENCES hii_especificaciones(ruta),
                PRIMARY KEY(origen, destino));
            CREATE TABLE IF NOT EXISTS hii_enlaces_automaticos (
                origen TEXT NOT NULL, destino TEXT NOT NULL,
                PRIMARY KEY(origen, destino));
            CREATE TABLE IF NOT EXISTS hii_generadas (
                ruta TEXT PRIMARY KEY, especificacion TEXT NOT NULL REFERENCES hii_especificaciones(ruta),
                idioma TEXT NOT NULL, fecha TEXT NOT NULL);
            CREATE VIEW IF NOT EXISTS hii_versiones_idiomas AS
                SELECT a.ruta AS hii_origen, a.idioma AS idioma_origen,
                       b.ruta AS hii_traduccion, b.idioma AS idioma_traduccion,
                       a.especificacion AS especificacion_origen,
                       b.especificacion AS especificacion_traduccion
                FROM hii_generadas a JOIN hii_generadas b
                JOIN hii_traducciones t ON
                    (a.especificacion=t.origen AND b.especificacion=t.destino) OR
                    (a.especificacion=t.destino AND b.especificacion=t.origen);
        """)
        with con:
            yield con
    finally:
        con.close()


def indexar(con, ruta):
    ruta = Path(ruta).resolve()
    datos = json.loads(ruta.read_text(encoding="utf-8-sig"))
    control = datos.get("control_documental", {})
    fila = (str(ruta), clave_texto(str(control.get("codigo_unico_producto") or "")),
            clave_texto(str(control.get("revision") or "")),
            clave_texto(str(control.get("nombre_cliente") or "")), leer_idioma_datos(datos))
    con.execute("""INSERT INTO hii_especificaciones VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(ruta) DO UPDATE SET codigo=excluded.codigo,
        revision=excluded.revision, cliente=excluded.cliente, idioma=excluded.idioma""", fila)
    return fila


def candidatas(con, origen, idioma, rutas_disponibles):
    fila = indexar(con, origen)
    disponibles = {str(Path(p).resolve()) for p in rutas_disponibles}
    explicitas = con.execute("""SELECT e.* FROM hii_especificaciones e
        JOIN hii_traducciones t ON (e.ruta=t.destino AND t.origen=?)
          OR (e.ruta=t.origen AND t.destino=?) WHERE e.idioma=?""",
        (fila[0], fila[0], idioma)).fetchall()
    validas = [Path(e["ruta"]) for e in explicitas if e["ruta"] in disponibles
               and (e["codigo"], e["cliente"]) == (fila[1], fila[3])]
    if validas:
        return validas
    if not fila[1] or not fila[2]:
        return []
    return [Path(e["ruta"]) for e in con.execute("""SELECT ruta FROM hii_especificaciones
        WHERE codigo=? AND revision=? AND cliente=? AND idioma=? ORDER BY ruta""",
        (*fila[1:4], idioma)) if e["ruta"] in disponibles]


def vincular(con, origen, destino, permitir_revision_distinta=False):
    a, b = indexar(con, origen), indexar(con, destino)
    compatible = (a[1], a[3]) == (b[1], b[3]) and (permitir_revision_distinta or a[2] == b[2])
    if not compatible or a[4] == b[4]:
        raise ValueError("Las traducciones deben tener el mismo código, revisión y cliente, y distinto idioma.")
    con.execute("INSERT OR IGNORE INTO hii_traducciones VALUES (?, ?)", (a[0], b[0]))
    con.execute("DELETE FROM hii_enlaces_automaticos WHERE (origen=? AND destino=?) OR (origen=? AND destino=?)",
                (a[0], b[0], b[0], a[0]))
    exportar_relaciones(con, [origen, destino])


def exportar_relaciones(con, rutas):
    """Referencias relativas al JSON actual; admite más de una traducción confirmada."""
    modificadas = set()
    for ruta in rutas:
        ruta = Path(ruta).resolve()
        datos = json.loads(ruta.read_text(encoding="utf-8-sig"))
        relaciones = {}
        filas = con.execute("""SELECT DISTINCT e.* FROM hii_especificaciones e
            JOIN hii_traducciones t ON (t.origen=? AND t.destino=e.ruta)
                OR (t.destino=? AND t.origen=e.ruta) ORDER BY e.ruta""", (str(ruta), str(ruta)))
        actual = con.execute("SELECT * FROM hii_especificaciones WHERE ruta=?", (str(ruta),)).fetchone()
        for fila in filas:
            destino = Path(fila["ruta"])
            if not destino.is_file() or not actual or fila["idioma"] == actual["idioma"]:
                continue
            if (fila["codigo"], fila["cliente"]) != (actual["codigo"], actual["cliente"]):
                continue
            relaciones.setdefault(fila["idioma"], []).append({
                "archivo_json": Path(os.path.relpath(destino, ruta.parent)).as_posix(),
                "codigo": fila["codigo"], "revision": fila["revision"], "cliente": fila["cliente"],
            })
        if not relaciones and "versiones_idioma" not in datos:
            continue
        if datos.get("versiones_idioma") == relaciones:
            continue
        datos["versiones_idioma"] = relaciones
        temporal = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=ruta.parent,
                prefix=".relaciones_", suffix=".tmp", delete=False) as archivo:
                temporal = Path(archivo.name)
                json.dump(datos, archivo, ensure_ascii=False, indent=2)
                archivo.write("\n")
            os.replace(temporal, ruta)
        finally:
            if temporal is not None:
                temporal.unlink(missing_ok=True)
        modificadas.add(ruta)
    return modificadas


def sincronizar_idiomas(con, carpeta):
    """Autoenlaza solo grupos con una única versión por idioma y revisión."""
    from collections import defaultdict
    grupos = defaultdict(list)
    rutas = []
    for ruta in sorted(Path(carpeta).rglob("*")):
        if not ruta.is_file() or ruta.suffix.lower() != ".json" or ruta.name.startswith("."):
            continue
        try:
            fila = indexar(con, ruta)
        except (OSError, ValueError, AttributeError):
            continue
        rutas.append(ruta.resolve())
        if fila[1] and fila[2] and fila[4] in {"ESPANOL", "INGLES"}:
            grupos[fila[1:4]].append(fila)
    disponibles = {str(r) for r in rutas}
    for a, b in con.execute("SELECT origen, destino FROM hii_enlaces_automaticos").fetchall():
        if a in disponibles or b in disponibles:
            con.execute("DELETE FROM hii_traducciones WHERE origen=? AND destino=?", (a, b))
            con.execute("DELETE FROM hii_enlaces_automaticos WHERE origen=? AND destino=?", (a, b))
    for grupo in grupos.values():
        es = [r for r in grupo if r[4] == "ESPANOL"]
        en = [r for r in grupo if r[4] == "INGLES"]
        if len(es) == len(en) == 1:
            pareja = (es[0][0], en[0][0])
            existe = con.execute("SELECT 1 FROM hii_traducciones WHERE (origen=? AND destino=?) OR (origen=? AND destino=?)",
                                 (*pareja, pareja[1], pareja[0])).fetchone()
            if not existe:
                con.execute("INSERT INTO hii_traducciones VALUES (?, ?)", pareja)
                con.execute("INSERT INTO hii_enlaces_automaticos VALUES (?, ?)", pareja)
    return exportar_relaciones(con, rutas)


def registrar(con, excel, especificacion, idioma, fecha):
    indexar(con, especificacion)
    con.execute("""INSERT INTO hii_generadas VALUES (?, ?, ?, ?)
        ON CONFLICT(ruta) DO UPDATE SET especificacion=excluded.especificacion,
        idioma=excluded.idioma, fecha=excluded.fecha""",
        (str(Path(excel).resolve()), str(Path(especificacion).resolve()), idioma, fecha))
