"""Relaciones entre especificaciones bilingües y sus HII en SQLite."""
import json
import sqlite3
from contextlib import contextmanager
from pathlib import Path

from generador_especs_hii import clave_texto, leer_idioma_datos


@contextmanager
def conectar(ruta):
    con = sqlite3.connect(ruta)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    con.executescript("""
        CREATE TABLE IF NOT EXISTS hii_especificaciones (
            ruta TEXT PRIMARY KEY, codigo TEXT NOT NULL, revision TEXT NOT NULL,
            cliente TEXT NOT NULL, idioma TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS hii_traducciones (
            origen TEXT NOT NULL REFERENCES hii_especificaciones(ruta),
            destino TEXT NOT NULL REFERENCES hii_especificaciones(ruta),
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
    try:
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
               and (e["codigo"], e["revision"], e["cliente"]) == fila[1:4]]
    if validas:
        return validas
    if not fila[1] or not fila[2]:
        return []
    return [Path(e["ruta"]) for e in con.execute("""SELECT ruta FROM hii_especificaciones
        WHERE codigo=? AND revision=? AND cliente=? AND idioma=? ORDER BY ruta""",
        (*fila[1:4], idioma)) if e["ruta"] in disponibles]


def vincular(con, origen, destino):
    a, b = indexar(con, origen), indexar(con, destino)
    if a[1:4] != b[1:4] or a[4] == b[4]:
        raise ValueError("Las traducciones deben tener el mismo código, revisión y cliente, y distinto idioma.")
    con.execute("INSERT OR IGNORE INTO hii_traducciones VALUES (?, ?)", (a[0], b[0]))


def registrar(con, excel, especificacion, idioma, fecha):
    indexar(con, especificacion)
    con.execute("""INSERT INTO hii_generadas VALUES (?, ?, ?, ?)
        ON CONFLICT(ruta) DO UPDATE SET especificacion=excluded.especificacion,
        idioma=excluded.idioma, fecha=excluded.fecha""",
        (str(Path(excel).resolve()), str(Path(especificacion).resolve()), idioma, fecha))
