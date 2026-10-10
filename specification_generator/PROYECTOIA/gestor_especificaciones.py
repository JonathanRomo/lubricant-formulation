#!/usr/bin/env python3
"""Gestor de especificaciones SCHUTZ, Python 3.11."""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unicodedata
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from openpyxl import load_workbook

APP_VERSION = "0.6.0"
PYTHON_REQUERIDO = (3, 11)
ENTRADA = "ESPECIFICACIONES_VIEJAS"
ENTRADA_ENG = "ESPECIFICACIONES_VIEJAS_ENG"
SALIDA = "ESPECIFICACIONES_JSON"
CATALOGO = "CODIGOS_200.xlsx"
DB = "especificaciones.db"
EXTENSIONES = {".doc", ".docx", ".docm"}

COLUMNAS = {
    "PRUEBA": "prueba", "UNIDAD": "unidades", "UNIDADES": "unidades",
    "METODO": "metodo", "ESPECIFICACION": "especificacion",
    "VALOR TIPICO": "valor_tipico", "VALORES TIPICOS": "valor_tipico",
    "TEST": "prueba", "UNITS": "unidades", "UNIT": "unidades",
    "TEST METHOD": "metodo", "METHOD": "metodo",
    "SPECIFICATION": "especificacion", "SPECIFICATIONS": "especificacion",
    "TYPICAL VALUE": "valor_tipico", "TYPICAL VALUES": "valor_tipico",
}
CONTROL = {
    "ELABORO Y APROBO": "elaboro_y_aprobo",
    "ELABORADO Y APROBADO POR": "elaboro_y_aprobo",
    "PREPARED AND APPROVED BY": "elaboro_y_aprobo",
    "ELABORATED AND APPROVED BY": "elaboro_y_aprobo",
    "FOLIO APQP": "folio_apqp", "FOLIO VAP": "folio_vap",
    "APQP FOLIO": "folio_apqp", "VAP FOLIO": "folio_vap",
    "CODIGO UNICO DE PRODUCTO": "codigo_unico_producto",
    "CODIGO DE PRODUCTO": "codigo_unico_producto",
    "PART ID": "codigo_unico_producto",
    "PRODUCT CODE": "codigo_unico_producto",
    "UNIQUE PRODUCT CODE": "codigo_unico_producto",
    "PROTOTYPE BATCH NUMBER": "numero_lote_prototipo",
    "PROTOTYPE BATCH NO": "numero_lote_prototipo",
    "PROTOTYPE PART NUMBER": "codigo_unico_producto",
    "ID": "codigo_unico_producto",
    "FECHA": "fecha", "DATE": "fecha", "REVISION": "revision", "REV": "revision",
    "NOMBRE CLIENTE": "nombre_cliente", "REVISION CLIENTE": "revision_cliente",
    "NUMERO DE REVISION": "revision_cliente",
    "NAME CLIENT": "nombre_cliente", "CLIENT NAME": "nombre_cliente",
    "CUSTOMER NAME": "nombre_cliente", "CUSTOMERS NAME": "nombre_cliente",
    "REVISION CLIENT": "revision_cliente", "CLIENT REVISION": "revision_cliente",
    "CUSTOMER REVISION": "revision_cliente",
    "CUSTOMERS REVISION NUMBER": "revision_cliente",
    "CUSTOMER S REVISION NUMBER": "revision_cliente",
}


def base_app() -> Path:
    return Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent


def limpiar(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return re.sub(r"\s+", " ", str(valor).replace("\xa0", " ")).strip()


def clave(valor) -> str:
    texto = "".join(c for c in unicodedata.normalize("NFD", limpiar(valor).upper())
                    if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", texto)).strip()


def id_clave(valor) -> str:
    return re.sub(r"[^A-Z0-9]", "", limpiar(valor).upper())


def nombre_seguro(valor) -> str:
    texto = re.sub(r'[<>:"/\\|?*]+', " ", limpiar(valor).upper())
    return re.sub(r"\s+", " ", texto).strip(" ._")


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for bloque in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloque)
    return h.hexdigest()


def conectar(path: Path) -> sqlite3.Connection:
    con = sqlite3.connect(path)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    con.execute("PRAGMA journal_mode=WAL")
    con.executescript("""
    CREATE TABLE IF NOT EXISTS archivos(
      id INTEGER PRIMARY KEY, ruta_relativa TEXT UNIQUE NOT NULL,
      nombre_archivo TEXT NOT NULL, hash_actual TEXT NOT NULL,
      hash_catalogo TEXT, tamano_bytes INTEGER NOT NULL,
      fecha_modificacion_ns INTEGER NOT NULL, producto_clave TEXT,
      producto_nombre TEXT, part_id_fuente TEXT, part_id_final TEXT,
      validacion_catalogo TEXT, revision_especificacion TEXT,
      codigo_formato TEXT, revision_formato TEXT, formato_origen TEXT,
      idioma TEXT, ultima_revision_json INTEGER, ultimo_json TEXT,
      primera_deteccion TEXT NOT NULL, ultima_revision_procesada TEXT NOT NULL,
      estado TEXT NOT NULL, ultimo_error TEXT);
    CREATE TABLE IF NOT EXISTS versiones(
      id INTEGER PRIMARY KEY, archivo_id INTEGER NOT NULL,
      producto_clave TEXT NOT NULL, revision_json INTEGER NOT NULL,
      hash_fuente TEXT NOT NULL, hash_catalogo TEXT,
      part_id_fuente TEXT, part_id_final TEXT, validacion_catalogo TEXT,
      revision_especificacion TEXT, codigo_formato TEXT,
      revision_formato TEXT, formato_origen TEXT, idioma TEXT,
      fecha_creacion TEXT NOT NULL, archivo_json TEXT UNIQUE NOT NULL,
      FOREIGN KEY(archivo_id) REFERENCES archivos(id),
      UNIQUE(producto_clave,revision_json));
    CREATE TABLE IF NOT EXISTS catalogo_codigos(
      id INTEGER PRIMARY KEY, hash_catalogo TEXT NOT NULL,
      hoja TEXT NOT NULL, fila_excel INTEGER NOT NULL,
      part_id TEXT NOT NULL, descripcion TEXT NOT NULL,
      part_id_clave TEXT NOT NULL, descripcion_clave TEXT NOT NULL,
      fecha_carga TEXT NOT NULL,
      UNIQUE(hash_catalogo,hoja,fila_excel));
    CREATE TABLE IF NOT EXISTS catalogo_fuentes(
      id INTEGER PRIMARY KEY, nombre_archivo TEXT NOT NULL,
      ruta_archivo TEXT NOT NULL, hash_catalogo TEXT UNIQUE NOT NULL,
      tamano_bytes INTEGER NOT NULL, fecha_modificacion_ns INTEGER NOT NULL,
      fecha_lectura TEXT NOT NULL, total_registros INTEGER NOT NULL);
    """)
    # Migracion no destructiva desde versiones previas.
    requeridas_archivos = {
        "hash_catalogo": "TEXT", "part_id_fuente": "TEXT", "part_id_final": "TEXT",
        "validacion_catalogo": "TEXT", "revision_especificacion": "TEXT",
        "codigo_formato": "TEXT", "revision_formato": "TEXT", "formato_origen": "TEXT",
        "idioma": "TEXT",
    }
    existentes = {r["name"] for r in con.execute("PRAGMA table_info(archivos)")}
    for nombre, tipo in requeridas_archivos.items():
        if nombre not in existentes:
            con.execute(f"ALTER TABLE archivos ADD COLUMN {nombre} {tipo}")
    requeridas_versiones = {k: v for k, v in requeridas_archivos.items() if k != "formato_origen"}
    requeridas_versiones["formato_origen"] = "TEXT"
    existentes = {r["name"] for r in con.execute("PRAGMA table_info(versiones)")}
    for nombre, tipo in requeridas_versiones.items():
        if nombre not in existentes:
            con.execute(f"ALTER TABLE versiones ADD COLUMN {nombre} {tipo}")
    con.commit()
    return con


def leer_catalogo(path: Path) -> tuple[list[dict], str]:
    """Lee CODIGOS_200.xlsx estrictamente en modo solo lectura.

    Esta funcion nunca llama save(), nunca cambia celdas y nunca crea una copia
    del libro. Toda la informacion derivada se devuelve en memoria y se guarda
    exclusivamente en SQLite.
    """
    if not path.exists():
        raise FileNotFoundError(f"No se encontro el catalogo obligatorio: {path.name}")
    digest = sha256(path)
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    registros = []
    for ws in wb.worksheets:
        encabezado = None
        col_id = col_desc = None
        for n_fila, fila in enumerate(ws.iter_rows(values_only=True), 1):
            claves = [clave(x) for x in fila]
            if "ID" in claves and "DESCRIPTION" in claves:
                encabezado = n_fila
                col_id = claves.index("ID")
                col_desc = claves.index("DESCRIPTION")
                break
        if encabezado is None:
            continue
        for n_fila, fila in enumerate(ws.iter_rows(min_row=encabezado + 1, values_only=True), encabezado + 1):
            part_id = limpiar(fila[col_id] if col_id < len(fila) else "")
            descripcion = limpiar(fila[col_desc] if col_desc < len(fila) else "")
            if not part_id and not descripcion:
                continue
            registros.append({"hoja": ws.title, "fila_excel": n_fila,
                              "part_id": part_id, "descripcion": descripcion,
                              "part_id_clave": id_clave(part_id),
                              "descripcion_clave": clave(descripcion)})
    wb.close()
    if not registros:
        raise ValueError("CODIGOS_200.xlsx no contiene columnas ID y Description reconocibles.")
    return registros, digest


def registrar_catalogo(con: sqlite3.Connection, registros: list[dict], digest: str,
                       fecha: str, archivo: Path) -> None:
    stat = archivo.stat()
    con.execute("""INSERT OR IGNORE INTO catalogo_fuentes
      (nombre_archivo,ruta_archivo,hash_catalogo,tamano_bytes,fecha_modificacion_ns,fecha_lectura,total_registros)
      VALUES(?,?,?,?,?,?,?)""",
      (archivo.name, str(archivo.resolve()), digest, stat.st_size, stat.st_mtime_ns, fecha, len(registros)))
    if con.execute("SELECT 1 FROM catalogo_codigos WHERE hash_catalogo=? LIMIT 1", (digest,)).fetchone():
        con.commit()
        return
    con.executemany("""INSERT OR IGNORE INTO catalogo_codigos
      (hash_catalogo,hoja,fila_excel,part_id,descripcion,part_id_clave,descripcion_clave,fecha_carga)
      VALUES(?,?,?,?,?,?,?,?)""",
      [(digest, r["hoja"], r["fila_excel"], r["part_id"], r["descripcion"],
        r["part_id_clave"], r["descripcion_clave"], fecha) for r in registros])
    con.commit()


def convertir_word_legacy(path: Path, tmp: Path) -> Path:
    destino = tmp / f"{path.stem}.docx"
    if sys.platform == "win32":
        word = None
        try:
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
            doc.SaveAs2(str(destino.resolve()), FileFormat=16)
            doc.Close(False)
            word.Quit()
            if destino.exists():
                return destino
        except Exception:
            if word:
                try: word.Quit()
                except Exception: pass
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if exe:
        subprocess.run([exe, "--headless", "--convert-to", "docx", "--outdir", str(tmp), str(path)],
                       capture_output=True, timeout=120, check=False)
        if destino.exists():
            return destino
    raise RuntimeError("No se pudo convertir el archivo .doc o .docm. Instala pywin32 con Microsoft Word o LibreOffice.")


def preparar(path: Path, tmp: Path) -> tuple[Path, bool]:
    extension = path.suffix.lower()
    if extension == ".docx":
        return path, False
    if extension in {".doc", ".docm"}:
        # El original se abre solo para lectura y se convierte a un DOCX temporal.
        # El archivo fuente nunca se guarda, modifica ni sobrescribe.
        return convertir_word_legacy(path, tmp), True
    raise ValueError(f"Formato Word no compatible: {extension}")


def validar_documento_word(path: Path, tmp: Path) -> None:
    if path.stat().st_size == 0:
        raise ValueError("el archivo esta vacio")
    lectura, _ = preparar(path, tmp)
    try:
        doc = Document(lectura)
    except Exception as exc:
        raise ValueError("el archivo no es un documento Word valido") from exc
    if not doc.tables:
        raise ValueError("el documento Word no contiene tablas")
    tiene_contenido = any(limpiar(p.text) for p in doc.paragraphs)
    tiene_contenido = tiene_contenido or any(
        limpiar(celda.text) for tabla in doc.tables for fila in tabla.rows for celda in fila.cells
    )
    if not tiene_contenido:
        raise ValueError("el documento Word no contiene texto")


def texto_ooxml(path: Path) -> str:
    partes = []
    with ZipFile(path) as z:
        for nombre in z.namelist():
            if nombre.startswith("word/") and nombre.endswith(".xml"):
                contenido = z.read(nombre).decode("utf-8", "ignore")
                partes.extend(re.findall(r"<w:t[^>]*>(.*?)</w:t>", contenido, re.S))
                partes.extend(re.findall(r"<a:t[^>]*>(.*?)</a:t>", contenido, re.S))
    return limpiar(html.unescape(" ".join(re.sub(r"<[^>]+>", " ", x) for x in partes)))


def metadatos_formato(path: Path) -> dict:
    texto = texto_ooxml(path)
    datos = {}
    m = re.search(r"\b(R\d{2}-[A-Z]{2,}-[A-Z]{2,}-\d{2})\b", texto, re.I)
    if m:
        datos["codigo_formato"] = m.group(1).upper()
        cerca = texto[m.end():m.end() + 100]
        rev = re.search(r"\b(\d{1,2})\b", cerca)
        if rev: datos["revision_formato"] = rev.group(1).zfill(2)
    exe = shutil.which("tesseract")
    if not exe:
        return datos
    with tempfile.TemporaryDirectory() as td, ZipFile(path) as z:
        for i, nombre in enumerate(n for n in z.namelist() if n.startswith("word/media/")):
            img = Path(td) / f"{i}{Path(nombre).suffix}"
            img.write_bytes(z.read(nombre))
            try:
                r = subprocess.run([exe, str(img), "stdout", "-l", "spa+eng", "--psm", "6"],
                                   capture_output=True, text=True, timeout=20, check=False)
            except Exception:
                continue
            t = limpiar(r.stdout)
            cod = re.search(r"C[OÓ]DIGO\s*[:.]?\s*([A-Z0-9-]+)", t, re.I)
            rev = re.search(r"REV(?:ISI[OÓ]N)?\s*[.:]?\s*([A-Z0-9.-]+)", t, re.I)
            if cod: datos["codigo_formato"] = cod.group(1)
            if rev: datos["revision_formato"] = rev.group(1).strip(".").zfill(2)
    return datos


def contiene_marcadores_ingles(path: Path) -> bool:
    """Detecta idioma por etiquetas del documento, no por el formato visual usado."""
    texto = clave(texto_ooxml(path))
    marcadores_ingles = (
        "TEST", "UNITS", "TEST METHOD", "SPECIFICATION", "TYPICAL VALUES",
        "PART ID", "PRODUCT CODE", "CUSTOMER", "NAME", "DATE", "POSITION",
        "SIGNATURE", "COMPANY",
    )
    marcadores_espanol = (
        "PRUEBA", "UNIDAD", "METODO", "ESPECIFICACION", "VALOR TIPICO",
        "CODIGO UNICO DE PRODUCTO", "FECHA", "NOMBRE CLIENTE",
        "ELABORO Y APROBO",
    )
    ingles = sum(bool(re.search(rf"\b{re.escape(m)}\b", texto)) for m in marcadores_ingles)
    espanol = sum(bool(re.search(rf"\b{re.escape(m)}\b", texto)) for m in marcadores_espanol)
    return ingles > espanol

def destino_movimiento_sin_colision(destino: Path) -> Path:
    """Evita sobrescribir un archivo existente al mover una especificacion ENG."""
    if not destino.exists():
        return destino
    indice = 1
    while True:
        candidato = destino.with_name(f"{destino.stem}_{indice:02d}{destino.suffix}")
        if not candidato.exists():
            return candidato
        indice += 1


def nombre_archivo_json(part_id: str, revision: str, fecha: datetime) -> str:
    codigo = nombre_seguro(part_id) or "SIN_CODIGO_200"
    revision_archivo = nombre_seguro(revision) or "SIN_REVISION"
    if revision_archivo.isdigit():
        revision_archivo = revision_archivo.zfill(2)
    return f"{codigo} {revision_archivo}_{fecha:%d_%m_%y}.JSON"


def extraer(path: Path, idioma: str = "es") -> dict:
    doc = Document(path)
    encabezados = [
        limpiar(p.text) for s in doc.sections for p in s.header.paragraphs
        if limpiar(p.text) and clave(p.text) not in {"ESPECIFICACION", "SPECIFICATION"}
    ]
    producto = encabezados[0] if encabezados else ""
    if not producto:
        candidatos = [
            limpiar(p.text) for p in doc.paragraphs
            if limpiar(p.text) and clave(p.text) not in {"ESPECIFICACION", "SPECIFICATION"}
        ]
        producto = candidatos[0] if candidatos else ""
    if not producto:
        raise ValueError("No se encontro el nombre del producto.")

    tablas = []
    for i, tabla in enumerate(doc.tables):
        if tabla.rows:
            enc = [clave(c.text) for c in tabla.rows[0].cells]
            tablas.append((sum(x in COLUMNAS for x in enc), i, tabla, enc))
    if not tablas:
        raise ValueError("El documento no contiene tablas.")
    puntos, indice, tabla, enc = max(tablas, key=lambda x: x[0])
    if puntos < 3:
        raise ValueError("No se encontro la tabla de pruebas.")
    mapa = [COLUMNAS.get(x) for x in enc]
    pruebas, notas = [], []
    for fila_origen, fila in enumerate(tabla.rows[1:], 2):
        valores = [limpiar(c.text) for c in fila.cells]
        no_vacios = [v for v in valores if v]
        if not no_vacios: continue
        if len({clave(v) for v in no_vacios}) == 1 and re.search(r"\b(?:NOTA|NOTES?)\b", clave(no_vacios[0])):
            notas.append(re.sub(r"^(?:Notas?|Notes?)\s*:\s*", "", no_vacios[0], flags=re.I)); continue
        item = {"numero": len(pruebas) + 1, "fila_origen": fila_origen}
        for j, valor in enumerate(valores):
            if j < len(mapa) and mapa[j]: item[mapa[j]] = valor
        for campo in ("prueba", "unidades", "metodo", "especificacion", "valor_tipico"):
            item.setdefault(campo, "")
        pruebas.append(item)

    control, mejor, puntos_mejor = {}, None, 0
    for i, candidata in enumerate(doc.tables):
        if i == indice or len(candidata.columns) < 2: continue
        puntos = sum(clave(r.cells[0].text) in CONTROL for r in candidata.rows)
        if puntos > puntos_mejor: mejor, puntos_mejor = candidata, puntos
    if mejor:
        for fila in mejor.rows:
            etiqueta = CONTROL.get(clave(fila.cells[0].text))
            if etiqueta: control[etiqueta] = limpiar(fila.cells[1].text)

    pie = " ".join(limpiar(p.text) for s in doc.sections for p in s.footer.paragraphs if limpiar(p.text))
    return {"esquema": "especificacion_producto_es_v3", "producto": {"nombre": producto},
            "documento": {"tipo": "ESPECIFICACION", "idioma": idioma, **metadatos_formato(path)},
            "control_documental": control, "pruebas": pruebas, "notas": notas,
            "datos_institucionales": {"pie_pagina": pie},
            "estructura_detectada": {"numero_tablas": len(doc.tables), "tabla_pruebas_indice": indice,
                                      "filas_pruebas": len(pruebas),
                                      "encabezados": [limpiar(c.text) for c in tabla.rows[0].cells]}}


def validar_catalogo(
    nombre: str, part_id: str, catalogo: list[dict], *, permitir_coincidencia_part_id: bool = False
) -> dict:
    nk, ik = clave(nombre), id_clave(part_id)
    por_nombre = [r for r in catalogo if r["descripcion_clave"] == nk]
    por_id = [r for r in catalogo if ik and r["part_id_clave"] == ik]
    exactos = [r for r in por_nombre if r["part_id_clave"] == ik] if ik else []
    resultado = {"estado": "", "prefijo_archivo": "", "nombre_fuente": nombre,
                 "part_id_fuente": part_id, "part_id_final": part_id,
                 "coincidencias_nombre": len(por_nombre), "coincidencias_part_id": len(por_id)}
    if ik and exactos:
        resultado.update(estado="COINCIDENCIA_EXACTA", registro_catalogo=exactos[0])
    elif permitir_coincidencia_part_id and ik and por_id:
        resultado.update(estado="COINCIDENCIA_PART_ID", part_id_final=por_id[0]["part_id"],
                         registro_catalogo=por_id[0])
    elif not ik and len(por_nombre) == 1:
        resultado.update(estado="PART_ID_COMPLETADO_DESDE_CATALOGO", prefijo_archivo="&&& ",
                         part_id_final=por_nombre[0]["part_id"], registro_catalogo=por_nombre[0])
    elif not ik and len(por_nombre) > 1:
        resultado.update(estado="CONFLICTO_NOMBRE_AMBIGUO", prefijo_archivo="--- ")
    elif (por_nombre and not por_id) or (por_id and not por_nombre):
        resultado.update(estado="CONFLICTO_NOMBRE_PART_ID", prefijo_archivo="--- ")
        if por_nombre: resultado["registro_por_nombre"] = por_nombre[0]
        if por_id: resultado["registro_por_part_id"] = por_id[0]
    elif ik and por_nombre and por_id and not exactos:
        resultado.update(estado="CONFLICTO_NOMBRE_PART_ID", prefijo_archivo="--- ",
                         registro_por_nombre=por_nombre[0], registro_por_part_id=por_id[0])
    else:
        resultado.update(estado="NO_ENCONTRADO_EN_CATALOGO", prefijo_archivo="--- ")
    return resultado


def procesar(base: Path) -> dict:
    entrada, entrada_eng, salida = base / ENTRADA, base / ENTRADA_ENG, base / SALIDA
    entrada.mkdir(exist_ok=True)
    entrada_eng.mkdir(exist_ok=True)
    archivos = sorted(
        ((p, raiz == entrada_eng) for raiz in (entrada, entrada_eng)
         for p in raiz.rglob("*") if p.is_file() and p.suffix.lower() in EXTENSIONES),
        key=lambda item: str(item[0]),
    )
    invalidos = []
    for origen, _ in archivos:
        try:
            with tempfile.TemporaryDirectory(prefix="validar_espec_") as td:
                validar_documento_word(origen, Path(td))
        except Exception as exc:
            invalidos.append(f"{origen.relative_to(base).as_posix()}: {exc}")
    if invalidos:
        detalle = "\n  ".join(invalidos)
        raise ValueError(f"Validacion previa cancelada; no se proceso ningun archivo:\n  {detalle}")

    salida.mkdir(exist_ok=True)
    con = conectar(base / DB)
    ahora_inicio = datetime.now().isoformat(timespec="seconds")
    catalogo, hash_catalogo = leer_catalogo(base / CATALOGO)
    registrar_catalogo(con, catalogo, hash_catalogo, ahora_inicio, base / CATALOGO)
    resumen = {"nuevos": 0, "actualizados": 0, "ignorados": 0, "movidos_eng": 0, "conflictos": 0, "completados": 0, "errores": 0}
    for origen, ya_esta_en_ingles in archivos:
        digest = sha256(origen)
        stat = origen.stat()
        rel = origen.relative_to(base).as_posix()
        ahora = datetime.now(); fecha_iso = ahora.isoformat(timespec="seconds")
        try:
            with tempfile.TemporaryDirectory(prefix="espec_") as td:
                lectura, convertido = preparar(origen, Path(td))
                idioma = "en" if ya_esta_en_ingles or contiene_marcadores_ingles(lectura) else "es"
                datos = extraer(lectura, idioma)
            if idioma == "en" and not ya_esta_en_ingles:
                destino_eng = destino_movimiento_sin_colision(entrada_eng / origen.name)
                shutil.move(str(origen), str(destino_eng))
                origen = destino_eng
                rel = origen.relative_to(base).as_posix()
                con.commit()
                resumen["movidos_eng"] += 1
                print(f"MOVIDO ENG: {rel}")
            anterior = con.execute("SELECT * FROM archivos WHERE ruta_relativa=?", (rel,)).fetchone()
            if anterior is None and idioma == "en":
                anterior_movido = con.execute(
                    """SELECT id FROM archivos
                       WHERE nombre_archivo=? AND hash_actual=? AND estado='MOVIDO_ENG'
                       ORDER BY id DESC LIMIT 1""",
                    (origen.name, digest),
                ).fetchone()
                if anterior_movido:
                    con.execute("UPDATE archivos SET ruta_relativa=? WHERE id=?", (rel, anterior_movido["id"]))
                    con.commit()
                    anterior = con.execute("SELECT * FROM archivos WHERE ruta_relativa=?", (rel,)).fetchone()
            if (anterior and anterior["hash_actual"] == digest
                    and anterior["hash_catalogo"] == hash_catalogo and anterior["estado"] == "OK"):
                resumen["ignorados"] += 1
                continue
            nombre = datos["producto"]["nombre"]
            part_fuente = datos["control_documental"].get("codigo_unico_producto", "")
            validacion = validar_catalogo(
                nombre, part_fuente, catalogo, permitir_coincidencia_part_id=idioma == "en"
            )
            part_final = validacion["part_id_final"]
            if part_final:
                datos["control_documental"]["codigo_unico_producto"] = part_final
            datos["validacion_codigos_200"] = validacion
            pk = clave(nombre)
            ultima = con.execute("SELECT MAX(revision_json) m FROM versiones WHERE producto_clave=?", (pk,)).fetchone()["m"]
            revision = 0 if ultima is None else ultima + 1
            rev_esp = datos["control_documental"].get("revision", "")
            prefijo = validacion["prefijo_archivo"]
            nombre_base = nombre_archivo_json(part_final, rev_esp, ahora)
            archivo_json = nombre_base
            sufijo = 1
            while (salida / archivo_json).exists() or con.execute(
                "SELECT 1 FROM versiones WHERE archivo_json=?", (archivo_json,)
            ).fetchone():
                archivo_json = f"{Path(nombre_base).stem}_{sufijo:02d}.JSON"
                sufijo += 1
            destino = salida / archivo_json
            datos.update(archivo_origen=rel, hash_sha256_origen=digest, hash_sha256_catalogo=hash_catalogo,
                         formato_origen=origen.suffix[1:].lower(), conversion_temporal=convertido,
                         revision_json=f"{revision:02d}", fecha_creacion_json=fecha_iso,
                         archivo_json=archivo_json, extractor_version=APP_VERSION)
            destino.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
            cod_fmt = datos["documento"].get("codigo_formato", "")
            rev_fmt = datos["documento"].get("revision_formato", "")
            con.execute("""INSERT INTO archivos
              (ruta_relativa,nombre_archivo,hash_actual,hash_catalogo,tamano_bytes,fecha_modificacion_ns,
               producto_clave,producto_nombre,part_id_fuente,part_id_final,validacion_catalogo,
               revision_especificacion,codigo_formato,revision_formato,formato_origen,idioma,
               ultima_revision_json,ultimo_json,primera_deteccion,ultima_revision_procesada,estado,ultimo_error)
              VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'OK',NULL)
              ON CONFLICT(ruta_relativa) DO UPDATE SET nombre_archivo=excluded.nombre_archivo,
               hash_actual=excluded.hash_actual,hash_catalogo=excluded.hash_catalogo,tamano_bytes=excluded.tamano_bytes,
               fecha_modificacion_ns=excluded.fecha_modificacion_ns,producto_clave=excluded.producto_clave,
               producto_nombre=excluded.producto_nombre,part_id_fuente=excluded.part_id_fuente,
               part_id_final=excluded.part_id_final,validacion_catalogo=excluded.validacion_catalogo,
               revision_especificacion=excluded.revision_especificacion,codigo_formato=excluded.codigo_formato,
               revision_formato=excluded.revision_formato,formato_origen=excluded.formato_origen,
               idioma=excluded.idioma,ultima_revision_json=excluded.ultima_revision_json,ultimo_json=excluded.ultimo_json,
               ultima_revision_procesada=excluded.ultima_revision_procesada,estado='OK',ultimo_error=NULL""",
              (rel, origen.name, digest, hash_catalogo, stat.st_size, stat.st_mtime_ns, pk, nombre,
               part_fuente, part_final, validacion["estado"], rev_esp, cod_fmt, rev_fmt,
               origen.suffix[1:].lower(), idioma.upper(), revision, archivo_json, fecha_iso, fecha_iso))
            archivo_id = con.execute("SELECT id FROM archivos WHERE ruta_relativa=?", (rel,)).fetchone()["id"]
            con.execute("""INSERT INTO versiones
              (archivo_id,producto_clave,revision_json,hash_fuente,hash_catalogo,part_id_fuente,
               part_id_final,validacion_catalogo,revision_especificacion,codigo_formato,
               revision_formato,formato_origen,idioma,fecha_creacion,archivo_json)
              VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
              (archivo_id, pk, revision, digest, hash_catalogo, part_fuente, part_final,
               validacion["estado"], rev_esp, cod_fmt, rev_fmt, origen.suffix[1:].lower(),
               idioma.upper(), fecha_iso, archivo_json))
            con.commit()
            resumen["actualizados" if anterior else "nuevos"] += 1
            if prefijo == "--- ": resumen["conflictos"] += 1
            if prefijo == "&&& ": resumen["completados"] += 1
            print(f"GENERADO: {archivo_json} [{validacion['estado']}]")
        except Exception as exc:
            con.rollback()
            con.execute("""INSERT INTO archivos
              (ruta_relativa,nombre_archivo,hash_actual,hash_catalogo,tamano_bytes,fecha_modificacion_ns,
               primera_deteccion,ultima_revision_procesada,estado,ultimo_error)
              VALUES(?,?,?,?,?,?,?,?, 'ERROR',?)
              ON CONFLICT(ruta_relativa) DO UPDATE SET hash_actual=excluded.hash_actual,
               hash_catalogo=excluded.hash_catalogo,ultima_revision_procesada=excluded.ultima_revision_procesada,
               estado='ERROR',ultimo_error=excluded.ultimo_error""",
              (rel, origen.name, digest, hash_catalogo, stat.st_size, stat.st_mtime_ns, fecha_iso, fecha_iso, str(exc)))
            con.commit(); resumen["errores"] += 1; print(f"ERROR: {rel}: {exc}")
    con.close(); return resumen


def ver_registro(base: Path) -> None:
    con = conectar(base / DB)
    for r in con.execute("""SELECT estado,ruta_relativa,ultima_revision_json,producto_nombre,
      part_id_fuente,part_id_final,validacion_catalogo,codigo_formato,revision_formato,
      revision_especificacion,idioma FROM archivos ORDER BY ruta_relativa"""):
        rev = "--" if r["ultima_revision_json"] is None else f"{r['ultima_revision_json']:02d}"
        print(f"{r['estado']} | idioma={r['idioma'] or '--'} | JSON {rev} | {r['validacion_catalogo'] or '--'} | "
              f"Part ID fuente={r['part_id_fuente'] or '--'} final={r['part_id_final'] or '--'} | "
              f"formato={r['codigo_formato'] or '--'} rev={r['revision_formato'] or '--'} | {r['ruta_relativa']}")
    con.close()


def main() -> int:
    if sys.version_info[:2] != PYTHON_REQUERIDO:
        raise SystemExit(f"Se requiere Python 3.11. Detectado {sys.version_info.major}.{sys.version_info.minor}.")
    parser = argparse.ArgumentParser()
    parser.add_argument("--ver-registro", action="store_true")
    args = parser.parse_args(); base = base_app()
    if args.ver_registro: ver_registro(base); return 0
    try:
        resumen = procesar(base)
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1
    print("\nRESUMEN")
    for k, v in resumen.items(): print(f"  {k}: {v}")
    return 1 if resumen["errores"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
