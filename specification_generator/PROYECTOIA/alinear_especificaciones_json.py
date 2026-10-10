#!/usr/bin/env python3
"""Alinea nombres de prueba y metodos en JSON usando METODOS_LAB_UNIFICADO.xlsx."""
from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
from copy import deepcopy
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

VERSION = "1.0.0"
PYTHON_REQUERIDO = (3, 11)
ENTRADA = "ESPECIFICACIONES_JSON"
SALIDA = "ESPECIFICACIONES_JSON_ALINEADAS"
MAESTRO = Path("METODOS_LAB") / "METODOS_LAB_UNIFICADO.xlsx"
HOJA = "METODOS"
REPORTE = "REPORTE_ALINEACION.txt"
EXTENSION = ".json"

# Umbrales conservadores para evitar sustituir una prueba por otra incorrecta.
UMBRAL_AUTOMATICO = 0.84
UMBRAL_DUDOSO = 0.72
MARGEN_AMBIGUEDAD = 0.035


def cargar_openpyxl() -> None:
    try:
        import openpyxl  # noqa: F401
    except ModuleNotFoundError as exc:
        if exc.name != "openpyxl":
            raise
        raise RuntimeError(
            "Falta openpyxl. Instala las dependencias en un entorno virtual con "
            "Python 3.11: .venv\\Scripts\\python.exe -m pip install -r requirements.txt"
        ) from exc


def texto(valor: Any) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    salida = str(valor).replace("\xa0", " ").replace("\r", " ").replace("\n", " ").replace("\t", " ")
    salida = "".join(c for c in salida if c.isprintable())
    return re.sub(r"\s+", " ", salida).strip()


def normalizar(valor: Any) -> str:
    salida = texto(valor).upper()
    salida = "".join(c for c in unicodedata.normalize("NFD", salida) if unicodedata.category(c) != "Mn")
    salida = salida.replace("×", "X")
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9%]+", " ", salida)).strip()


def tokens(valor: Any) -> set[str]:
    ignorar = {"DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "EN", "POR", "CON", "PARA", "THE", "OF", "IN", "AND", "BY", "FOR", "WITH", "AT"}
    return {x for x in normalizar(valor).split() if len(x) > 1 and x not in ignorar}


def similitud_texto(a: Any, b: Any) -> float:
    na, nb = normalizar(a), normalizar(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    ta, tb = tokens(a), tokens(b)
    jaccard = len(ta & tb) / len(ta | tb) if ta | tb else 0.0
    secuencia = SequenceMatcher(None, na, nb).ratio()
    contencion = min(len(ta & tb) / len(ta), len(ta & tb) / len(tb)) if ta and tb else 0.0
    return 0.48 * secuencia + 0.37 * jaccard + 0.15 * contencion


def partes_metodo(valor: Any) -> set[str]:
    """Reconoce metodos simples o conjugados: ASTM D123 / ASTM D456, ASTM D123; ISO 10, etc."""
    n = normalizar(valor)
    if not n:
        return set()
    encontrados = set(re.findall(r"\b(?:ASTM|ISO|IEC|NMX|USP|FEUM|EN|DIN|IP)\s*[A-Z]*\s*\d+[A-Z0-9.-]*", n))
    if encontrados:
        return {re.sub(r"\s+", "", x) for x in encontrados}
    return {re.sub(r"\s+", "", n)}


def similitud_metodo(a: Any, b: Any) -> float:
    na, nb = normalizar(a), normalizar(b)
    if not na and not nb:
        return 0.5
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    pa, pb = partes_metodo(a), partes_metodo(b)
    if pa and pb:
        inter = len(pa & pb)
        if inter:
            # Permite que un metodo del JSON sea parte de un metodo conjugado del maestro.
            return 0.90 + 0.10 * inter / max(len(pa), len(pb))
    return similitud_texto(a, b)


def cargar_maestro(path: Path) -> list[dict[str, str]]:
    from openpyxl import load_workbook
    if not path.exists():
        raise FileNotFoundError(f"No se encontro la lista maestra: {path}")
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    if HOJA not in wb.sheetnames:
        raise ValueError(f"El Excel maestro no contiene la hoja {HOJA!r}.")
    ws = wb[HOJA]
    encabezados = {normalizar(c.value): i for i, c in enumerate(ws[1]) if texto(c.value)}
    requeridas = ["METODO", "NOMBRE IQM", "NOMBRE IQM INGLES", "UNIDADES", "UNITS", "ACREDITACION NOMBRE"]
    faltantes = [c for c in requeridas if normalizar(c) not in encabezados]
    if faltantes:
        raise ValueError("Faltan columnas en la hoja METODOS: " + ", ".join(faltantes))
    registros = []
    for numero, fila in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        r = {c: texto(fila[encabezados[normalizar(c)]]) if encabezados[normalizar(c)] < len(fila) else "" for c in requeridas}
        if r["METODO"] or r["NOMBRE IQM"] or r["NOMBRE IQM INGLES"]:
            r["FILA_MAESTRO"] = str(numero)
            registros.append(r)
    wb.close()
    if not registros:
        raise ValueError("La hoja METODOS no contiene registros utilizables.")
    return registros


def puntuar(nombre: str, metodo: str, candidato: dict[str, str]) -> tuple[float, float, float]:
    nombre_es = similitud_texto(nombre, candidato["NOMBRE IQM"])
    nombre_en = similitud_texto(nombre, candidato["NOMBRE IQM INGLES"])
    score_nombre = max(nombre_es, nombre_en)
    score_metodo = similitud_metodo(metodo, candidato["METODO"])
    # El nombre selecciona la familia; el metodo resuelve nombres repetidos.
    if texto(metodo):
        total = 0.68 * score_nombre + 0.32 * score_metodo
        if score_nombre >= 0.96 and score_metodo >= 0.90:
            total = max(total, 0.96)
    else:
        total = 0.88 * score_nombre + 0.12 * score_metodo
    return total, score_nombre, score_metodo


def mejor_coincidencia(nombre: str, metodo: str, maestro: list[dict[str, str]]) -> dict[str, Any]:
    evaluados = []
    for registro in maestro:
        total, sn, sm = puntuar(nombre, metodo, registro)
        evaluados.append((total, sn, sm, registro))
    evaluados.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
    mejor = evaluados[0]
    segundo = evaluados[1] if len(evaluados) > 1 else (0.0, 0.0, 0.0, {})
    margen = mejor[0] - segundo[0]
    ambiguo = margen < MARGEN_AMBIGUEDAD and normalizar(mejor[3]["METODO"]) != normalizar(segundo[3].get("METODO", ""))
    if mejor[0] >= UMBRAL_AUTOMATICO and not ambiguo:
        estado = "ALINEADO"
    elif mejor[0] >= UMBRAL_DUDOSO:
        estado = "REVISION_REQUERIDA"
    else:
        estado = "SIN_COINCIDENCIA"
    return {"estado": estado, "puntaje": round(mejor[0], 4), "puntaje_nombre": round(mejor[1], 4),
            "puntaje_metodo": round(mejor[2], 4), "margen": round(margen, 4), "ambiguo": ambiguo,
            "registro": mejor[3], "segunda_opcion": segundo[3]}


def procesar_json(path: Path, salida: Path, maestro: list[dict[str, str]]) -> dict[str, int]:
    datos = json.loads(path.read_text(encoding="utf-8-sig"))
    resultado = deepcopy(datos)
    pruebas = resultado.get("pruebas")
    if not isinstance(pruebas, list):
        raise ValueError("El JSON no contiene una lista 'pruebas'.")
    conteo = {"alineados": 0, "revision": 0, "sin_coincidencia": 0}
    auditoria = []
    for indice, prueba in enumerate(pruebas, 1):
        if not isinstance(prueba, dict):
            continue
        nombre_original = texto(prueba.get("prueba") or prueba.get("nombre"))
        metodo_original = texto(prueba.get("metodo"))
        match = mejor_coincidencia(nombre_original, metodo_original, maestro)
        registro = match["registro"]
        if match["estado"] == "ALINEADO":
            prueba["prueba"] = registro["NOMBRE IQM"]
            prueba["metodo"] = registro["METODO"]
            conteo["alineados"] += 1
        elif match["estado"] == "REVISION_REQUERIDA":
            conteo["revision"] += 1
        else:
            conteo["sin_coincidencia"] += 1
        auditoria.append({
            "numero_prueba": indice, "estado": match["estado"], "nombre_original": nombre_original,
            "metodo_original": metodo_original, "nombre_maestro_propuesto": registro.get("NOMBRE IQM", ""),
            "metodo_maestro_propuesto": registro.get("METODO", ""), "fila_maestro": registro.get("FILA_MAESTRO", ""),
            "puntaje_total": match["puntaje"], "puntaje_nombre": match["puntaje_nombre"],
            "puntaje_metodo": match["puntaje_metodo"], "ambiguo": match["ambiguo"],
        })
    resultado["alineacion_metodos_lab"] = {
        "version_programa": VERSION, "fecha": datetime.now().isoformat(timespec="seconds"),
        "archivo_maestro": MAESTRO.as_posix(), "resumen": conteo, "detalle": auditoria,
    }
    salida.parent.mkdir(parents=True, exist_ok=True)
    salida.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    return conteo


def ejecutar(base: Path) -> None:
    entrada, salida, maestro_path = base / ENTRADA, base / SALIDA, base / MAESTRO
    if not entrada.exists():
        raise FileNotFoundError(f"No existe la carpeta de entrada: {entrada}")
    salida.mkdir(exist_ok=True)
    maestro = cargar_maestro(maestro_path)
    archivos = sorted(p for p in entrada.rglob("*.json") if p.is_file())
    if not archivos:
        raise FileNotFoundError(f"No se encontraron JSON en {entrada}")
    total = {"archivos": 0, "alineados": 0, "revision": 0, "sin_coincidencia": 0, "errores": 0}
    lineas = [f"REPORTE DE ALINEACION {datetime.now().isoformat(timespec='seconds')}", "=" * 72]
    for origen in archivos:
        rel = origen.relative_to(entrada)
        destino = salida / rel
        try:
            r = procesar_json(origen, destino, maestro)
            total["archivos"] += 1
            for k in ("alineados", "revision", "sin_coincidencia"):
                total[k] += r[k]
            lineas.append(f"OK | {rel.as_posix()} | alineados={r['alineados']} revision={r['revision']} sin_coincidencia={r['sin_coincidencia']}")
            print(f"ALINEADO: {rel}")
        except Exception as exc:
            total["errores"] += 1
            lineas.append(f"ERROR | {rel.as_posix()} | {exc}")
            print(f"ERROR: {rel}: {exc}", file=sys.stderr)
    lineas.extend(["", "RESUMEN"] + [f"{k}: {v}" for k, v in total.items()])
    (salida / REPORTE).write_text("\n".join(lineas) + "\n", encoding="utf-8")
    print("\nRESUMEN")
    for k, v in total.items():
        print(f"  {k}: {v}")


def main() -> int:
    if sys.version_info[:2] != PYTHON_REQUERIDO:
        raise SystemExit(
            f"Se requiere Python 3.11. Detectado "
            f"{sys.version_info.major}.{sys.version_info.minor}."
        )
    parser = argparse.ArgumentParser(description="Alinea pruebas y metodos de JSON con la lista maestra.")
    parser.add_argument("--base", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    try:
        cargar_openpyxl()
        ejecutar(args.base.resolve())
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
