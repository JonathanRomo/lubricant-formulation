#!/usr/bin/env python3
"""Alinea nombres de prueba y metodos en JSON usando METODOS_LAB_UNIFICADO.xlsx."""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import unicodedata
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import pandas as pd
    from rapidfuzz import fuzz
except ModuleNotFoundError as exc:
    if exc.name in {"pandas", "rapidfuzz"}:
        raise SystemExit(
            "Faltan dependencias del alineador. Instala los requisitos con "
            ".venv\\Scripts\\python.exe -m pip install -r requirements.txt"
        ) from None
    raise

VERSION = "1.6.0"
PYTHON_REQUERIDO = (3, 11)
ENTRADA = "ESPECIFICACIONES_JSON"
SALIDA = "ESPECIFICACIONES_JSON_ALINEADAS"
MAESTRO = Path("METODOS_LAB") / "METODOS_LAB_UNIFICADO.xlsx"
HOJA = "METODOS"
REPORTE = "REPORTE_ALINEACION.txt"
EXTENSION = ".json"
ESPACIO_FINO = "\u202f"
ESPACIOS_NUMERICOS = " \u00a0\u2009\u202f"
RANGO_NUMERICO = re.compile(
    r"(?<![\w\d\-\u2013\u2014])(?P<izquierdo>[+-]?\d+(?:[.,]\d+)?)"
    r"\s*[-\u2013\u2014]\s*(?P<derecho>-?\d+(?:[.,]\d+)?)(?!\d|[-\u2013\u2014]\d)"
)
DECIMAL_CON_FRACCION_SEPARADA = re.compile(
    rf"(?<![\w])(?P<entero>[+-]?\d+(?:[{ESPACIOS_NUMERICOS}]\d{{3}})*)"
    rf"(?P<separador>[.,])(?P<fraccion>\d{{3}})[{ESPACIOS_NUMERICOS}]"
    r"(?P<fraccion_final>\d{1,2})(?!\d)"
)
AGRUPACION_MILES_CON_DECIMAL = re.compile(
    r"(?<![\w])(?P<signo>[+-]?)(?P<entero>\d{1,3})"
    r"(?P<grupos>(?:[.,]\d{3})+)(?P<separador_decimal>[.,])"
    r"(?P<fraccion>\d+)(?![\w])"
)
NUMERO_DECIMAL = re.compile(
    rf"(?<![\w])(?P<signo>[+-]?)(?P<entero>\d+(?:[{ESPACIOS_NUMERICOS}]\d{{3}})*|)"
    r"(?P<separador>[.,])(?P<fraccion>\d+)(?![\w])"
)
ENTERO_CON_MILES = re.compile(
    rf"(?<![\w.,])(?P<entero>\d{{1,3}}(?:[{ESPACIOS_NUMERICOS}]\d{{3}})+|\d{{4,}})"
    r"(?![\w.,])"
)
CAMPOS_CONTENIDO_TABLA = (
    "prueba", "nombre", "unidades", "metodo", "especificacion", "valor_tipico"
)
UNIDADES_RECONOCIDAS = (
    "mPa·s", "mPa.s", "mg/kg", "mg/L", "g/L", "mL/min", "mm²/s", "mm2/s",
    "cm-1", "cm³", "cm3", "cm²", "cm2", "mm²", "mm2", "µL", "μL", "uL",
    "mL", "ml", "µm", "μm", "um", "kg", "mg", "µg", "μg", "ug", "mm",
    "cm", "km", "m", "sec", "secs", "min", "mins", "hr", "hrs", "rpm",
    "cSt", "cP", "kPa", "MPa", "Pa", "bar", "mol", "ppm", "ppb", "N",
    "A", "K", "cd", "V", "W", "J", "kJ", "s", "h", "d", "L", "l", "%",
)
UNIDAD_ADYACENTE_A_NUMERO = re.compile(
    r"(?P<numero>(?<!\w)[+-]?(?:\d+(?:[.,]\d+)?|[.,]\d+))\s*"
    r"(?P<unidad>°\s*[CF]|º\s*[CF]|"
    + "|".join(
        sorted(
            (re.escape(unidad) for unidad in UNIDADES_RECONOCIDAS),
            key=len,
            reverse=True,
        )
    )
    + r")(?![A-Za-zµμ])",
    re.IGNORECASE,
)

# Umbrales conservadores para evitar sustituir una prueba por otra incorrecta.
UMBRAL_AUTOMATICO = 0.84
UMBRAL_DUDOSO = 0.72
MARGEN_AMBIGUEDAD = 0.035
COLUMNAS_MAESTRO = (
    "METODO", "NOMBRE IQM", "NOMBRE IQM INGLES", "UNIDADES", "UNITS",
    "ACREDITACION NOMBRE",
)
ALIAS_IDIOMA_INGLES = {"EN", "ENG", "ENGLISH", "INGLES"}
ALIAS_IDIOMA_ESPANOL = {"ES", "ESP", "SPANISH", "ESPANOL"}
ASTM_REFERENCIA_EXPLICITA = re.compile(
    r"(?<![A-Z0-9])ASTM\s*[-]?\s*(?P<letra>[A-Z]?)\s*[-]?\s*"
    r"(?P<numero>\d{2,5})(?P<modificador>[A-Z]?)(?![A-Z0-9])",
    re.IGNORECASE,
)
ASTM_REFERENCIA_D = re.compile(
    r"(?<![A-Z0-9])D\s*[-\u2010-\u2015]?\s*"
    r"(?P<numero>\d{2,5})(?P<modificador>[A-Z]?)(?![A-Z0-9])",
    re.IGNORECASE,
)
ASTM_REFERENCIA_NUMERICA = re.compile(
    r"\s*(?P<numero>\d{3,5})"
    r"(?:\s*[-\u2010-\u2015]\s*\d{1,4}[A-Z]?)?\s*",
    re.IGNORECASE,
)
FARMACOPEIA_USP = re.compile(
    r"\b(?:USP(?:\s*[-/]?\s*NF)?|UNITED\s+STATES\s+PHARMACOPOEIA)\b",
    re.IGNORECASE,
)
FARMACOPEIA_FEUM = re.compile(r"\bFEUM\b", re.IGNORECASE)
FARMACOPEIA_PH_EUR = re.compile(
    r"\b(?:PH\s*\.?\s*EUR\.?|FARMACOPEA\s+EUROPEA|"
    r"EUROPEAN\s+PHARMACOPOEIA)(?=$|[^A-Z])",
    re.IGNORECASE,
)


def texto(valor: Any) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    salida = str(valor).replace("\xa0", " ").replace("\r", " ").replace("\n", " ").replace("\t", " ")
    salida = "".join(c for c in salida if c.isprintable())
    return re.sub(r"\s+", " ", salida).strip()


def _agrupar_digitos(digitos: str) -> str:
    return re.sub(r"\B(?=(?:\d{3})+$)", ESPACIO_FINO, digitos)


def normalizar_formato_numerico(valor: Any) -> str:
    """Aplica formato decimal y de miles sin reinterpretar separadores ambiguos."""
    salida = texto(valor)
    if not salida:
        return ""

    salida = RANGO_NUMERICO.sub(
        lambda coincidencia: (
            f"{coincidencia.group('izquierdo')} - "
            f"{coincidencia.group('derecho')}"
        ),
        salida,
    )

    def normalizar_fraccion_separada(coincidencia: re.Match[str]) -> str:
        entero = re.sub(
            rf"[{ESPACIOS_NUMERICOS}]",
            "",
            coincidencia.group("entero"),
        )
        signo = ""
        if entero[:1] in {"+", "-"}:
            signo, entero = entero[0], entero[1:]
        fraccion = (
            coincidencia.group("fraccion")
            + coincidencia.group("fraccion_final")
        )
        return f"{signo}{_agrupar_digitos(entero)},{fraccion}"

    salida = DECIMAL_CON_FRACCION_SEPARADA.sub(
        normalizar_fraccion_separada, salida
    )

    def normalizar_miles_con_decimal(coincidencia: re.Match[str]) -> str:
        separadores_grupo = re.findall(r"[.,]", coincidencia.group("grupos"))
        separador_decimal = coincidencia.group("separador_decimal")
        if (
            len(set(separadores_grupo)) != 1
            or separadores_grupo[0] == separador_decimal
        ):
            return coincidencia.group(0)
        entero = coincidencia.group("entero") + re.sub(
            r"[.,]", "", coincidencia.group("grupos")
        )
        return (
            f"{coincidencia.group('signo')}{_agrupar_digitos(entero)},"
            f"{coincidencia.group('fraccion')}"
        )

    salida = AGRUPACION_MILES_CON_DECIMAL.sub(
        normalizar_miles_con_decimal, salida
    )

    def normalizar_decimal(coincidencia: re.Match[str]) -> str:
        entero = re.sub(
            rf"[{ESPACIOS_NUMERICOS}]",
            "",
            coincidencia.group("entero"),
        )
        return (
            f"{coincidencia.group('signo')}{_agrupar_digitos(entero or '0')},"
            f"{coincidencia.group('fraccion')}"
        )

    salida = NUMERO_DECIMAL.sub(normalizar_decimal, salida)

    def normalizar_entero(coincidencia: re.Match[str]) -> str:
        digitos = re.sub(
            rf"[{ESPACIOS_NUMERICOS}]",
            "",
            coincidencia.group("entero"),
        )
        return _agrupar_digitos(digitos)

    return ENTERO_CON_MILES.sub(normalizar_entero, salida)


def normalizar_espacio_unidades(valor: Any) -> str:
    """Separa valores de símbolos de unidad reconocidos en celdas de tablas."""
    salida = "" if valor is None else str(valor).strip()

    def separar_unidad(coincidencia: re.Match[str]) -> str:
        unidad = re.sub(r"\s+", "", coincidencia.group("unidad")).replace("º", "°")
        if unidad.casefold() == "ml":
            unidad = "mL"
        return f"{coincidencia.group('numero')} {unidad}"

    return UNIDAD_ADYACENTE_A_NUMERO.sub(separar_unidad, salida)


def normalizar(valor: Any) -> str:
    salida = texto(valor).upper()
    salida = "".join(c for c in unicodedata.normalize("NFD", salida) if unicodedata.category(c) != "Mn")
    salida = salida.replace("×", "X")
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9%]+", " ", salida)).strip()


def similitud_texto(a: Any, b: Any) -> float:
    na, nb = normalizar(a), normalizar(b)
    if not na or not nb:
        return 0.0
    referencias_a, referencias_b = set(referencias_astm(a)), set(referencias_astm(b))
    if referencias_a or referencias_b:
        if not referencias_a or not referencias_b:
            return 0.0
        interseccion = referencias_a & referencias_b
        if referencias_a == referencias_b:
            return 1.0
        if interseccion and (referencias_a <= referencias_b or referencias_b <= referencias_a):
            return 0.90 + 0.10 * len(interseccion) / max(
                len(referencias_a), len(referencias_b)
            )
        if interseccion:
            return 0.50 * len(interseccion) / len(referencias_a | referencias_b)
        return 0.0
    if na == nb:
        return 1.0
    secuencia = fuzz.ratio(na, nb) / 100
    orden_tokens = fuzz.token_sort_ratio(na, nb) / 100
    conjunto_tokens = fuzz.token_set_ratio(na, nb) / 100
    return 0.75 * secuencia + 0.15 * orden_tokens + 0.10 * conjunto_tokens


def referencias_astm(valor: Any) -> tuple[str, ...]:
    """Extrae designaciones ASTM, ignorando guiones, espacios y ediciones/años."""
    original = texto(valor).replace("\u2212", "-")
    encontradas: list[str] = []
    for coincidencia in ASTM_REFERENCIA_EXPLICITA.finditer(original):
        letra = (coincidencia.group("letra") or "D").upper()
        numero = coincidencia.group("numero")
        modificador = (coincidencia.group("modificador") or "").upper()
        encontradas.append(f"{letra}{numero}{modificador}")

    for coincidencia in ASTM_REFERENCIA_D.finditer(original):
        encontradas.append(
            f"D{coincidencia.group('numero')}{(coincidencia.group('modificador') or '').upper()}"
        )

    # Algunas filas omiten D/ASTM en el segundo método de una conjugación:
    # "ASTM D445 / 7042".
    segmentos = re.split(r"[/+;]", original)
    if any(referencias for referencias in encontradas):
        for segmento in segmentos[1:]:
            coincidencia = ASTM_REFERENCIA_NUMERICA.fullmatch(segmento)
            if coincidencia:
                encontradas.append(f"D{coincidencia.group('numero')}")

    return tuple(dict.fromkeys(encontradas))


def referencias_astm_fuera_maestro(
    metodo: Any, referencias_maestras: set[str]
) -> tuple[str, ...]:
    return tuple(
        referencia
        for referencia in referencias_astm(metodo)
        if referencia not in referencias_maestras
    )


def referencias_farmacopea(valor: Any) -> tuple[str, ...]:
    """Extrae farmacopeas, ignorando ediciones, capítulos y texto accesorio."""
    original = texto(valor)
    normalizado = normalizar(original)
    encontradas: list[tuple[int, str]] = []

    # Casos de farmacopeas: USP/NF y ediciones, FEUM/MGA, y PH. EUR. con
    # ediciones en español o inglés y capítulos como 2.2.9.
    for patron, referencia in (
        (FARMACOPEIA_USP, "USP"),
        (FARMACOPEIA_FEUM, "FEUM"),
        (FARMACOPEIA_PH_EUR, "Ph Eur"),
    ):
        encontradas.extend(
            (coincidencia.start(), referencia)
            for coincidencia in patron.finditer(normalizado)
        )

    encontradas.sort(key=lambda coincidencia: coincidencia[0])
    return tuple(dict.fromkeys(referencia for _, referencia in encontradas))


def metodo_referencias_canonico(valor: Any) -> str:
    referencias = [
        *(f"ASTM {referencia}" for referencia in referencias_astm(valor)),
        *referencias_farmacopea(valor),
    ]
    if referencias:
        return " / ".join(dict.fromkeys(referencias))
    return texto(valor)


def metodo_maestro_canonico(valor: Any) -> str:
    """Normaliza citas ASTM y farmacopeas del maestro sin ediciones ni capítulos."""
    original = texto(valor).replace("\u2212", "-")
    referencias = [
        *(f"ASTM {referencia}" for referencia in referencias_astm(original)),
        *referencias_farmacopea(original),
    ]
    if not referencias:
        return original
    return " / ".join(dict.fromkeys(referencias))


def partes_metodo(valor: Any) -> set[str]:
    """Reconoce metodos simples o conjugados: ASTM D123 / ASTM D456, ASTM D123; ISO 10, etc."""
    referencias = referencias_astm(valor)
    farmacopeas = referencias_farmacopea(valor)
    referencias_normalizadas = {
        *(f"ASTM {referencia}" for referencia in referencias),
        *farmacopeas,
    }
    if referencias_normalizadas:
        return referencias_normalizadas
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
        return 0.0
    return similitud_texto(a, b)


def cargar_maestro(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        raise FileNotFoundError(f"No se encontro la lista maestra: {path}")
    with pd.ExcelFile(path, engine="openpyxl") as libro:
        if HOJA not in libro.sheet_names:
            raise ValueError(f"El Excel maestro no contiene la hoja {HOJA!r}.")
        tabla = pd.read_excel(
            libro, sheet_name=HOJA, dtype=str, keep_default_na=False
        )

    encabezados: dict[str, int] = {}
    for indice, encabezado in enumerate(tabla.columns):
        clave_encabezado = normalizar(encabezado)
        if clave_encabezado:
            encabezados.setdefault(clave_encabezado, indice)
    faltantes = [columna for columna in COLUMNAS_MAESTRO if normalizar(columna) not in encabezados]
    if faltantes:
        raise ValueError("Faltan columnas en la hoja METODOS: " + ", ".join(faltantes))

    filas = []
    for numero, (_, fila) in enumerate(tabla.iterrows(), 2):
        registro = {
            columna: texto(fila.iloc[encabezados[normalizar(columna)]])
            for columna in COLUMNAS_MAESTRO
        }
        if any(registro[columna] for columna in ("METODO", "NOMBRE IQM", "NOMBRE IQM INGLES")):
            registro["FILA_MAESTRO"] = str(numero)
            filas.append(registro)

    if filas:
        # Pandas elimina filas totalmente repetidas del maestro, conservando la primera referencia.
        marco = pd.DataFrame(filas)
        marco = marco.drop_duplicates(subset=list(COLUMNAS_MAESTRO), keep="first")
        registros = marco.to_dict(orient="records")
    else:
        registros = []
    if not registros:
        raise ValueError("La hoja METODOS no contiene registros utilizables.")
    return registros


def puntuar(nombre: str, metodo: str, candidato: dict[str, str]) -> tuple[float, float, float]:
    nombre_es = similitud_texto(
        nombre, limpiar_marca_acreditacion(candidato["NOMBRE IQM"])
    )
    nombre_en = similitud_texto(
        nombre, limpiar_marca_acreditacion(candidato["NOMBRE IQM INGLES"])
    )
    score_nombre = max(nombre_es, nombre_en)
    score_metodo = similitud_metodo(metodo, candidato["METODO"])
    # El nombre selecciona la familia; el metodo resuelve nombres repetidos.
    if texto(metodo):
        total = 0.68 * score_nombre + 0.32 * score_metodo
        if score_nombre >= 0.96 and score_metodo >= 0.90:
            total = max(total, 0.96)
        elif (
            referencias_astm(metodo)
            and referencias_astm(candidato["METODO"])
            and score_metodo >= 0.90
            and score_nombre >= 0.45
        ):
            total = max(total, UMBRAL_AUTOMATICO)
    else:
        total = 0.88 * score_nombre + 0.12 * score_metodo
    return total, score_nombre, score_metodo


def idioma_especificacion(datos: dict[str, Any]) -> str:
    documento = datos.get("documento")
    idioma = documento.get("idioma", "") if isinstance(documento, dict) else ""
    idioma_normalizado = normalizar(idioma).replace(" ", "")
    if idioma_normalizado in ALIAS_IDIOMA_INGLES:
        return "INGLES"
    if idioma_normalizado in ALIAS_IDIOMA_ESPANOL:
        return "ESPANOL"
    return "DESCONOCIDO"


def limpiar_marca_acreditacion(nombre: Any) -> str:
    """Quita asteriscos de acreditacion solo al inicio del nombre de prueba."""
    return re.sub(r"^\s*\*+\s*", "", texto(nombre))


def nombres_canonicos(registro: dict[str, str], idioma: str) -> tuple[str, str]:
    if idioma == "INGLES":
        return (
            limpiar_marca_acreditacion(registro.get("NOMBRE IQM INGLES", "")),
            registro.get("UNITS", "") or registro.get("UNIDADES", ""),
        )
    if idioma == "ESPANOL":
        return (
            limpiar_marca_acreditacion(registro.get("NOMBRE IQM", "")),
            registro.get("UNIDADES", "") or registro.get("UNITS", ""),
        )
    return (
        limpiar_marca_acreditacion(
            registro.get("NOMBRE IQM", "") or registro.get("NOMBRE IQM INGLES", "")
        ),
        registro.get("UNIDADES", "") or registro.get("UNITS", ""),
    )


def registros_ambiguos(a: dict[str, str], b: dict[str, str], idioma: str) -> bool:
    nombre_a, unidades_a = nombres_canonicos(a, idioma)
    nombre_b, unidades_b = nombres_canonicos(b, idioma)
    metodo_a = metodo_maestro_canonico(a.get("METODO", ""))
    metodo_b = metodo_maestro_canonico(b.get("METODO", ""))
    if normalizar(metodo_a) != normalizar(metodo_b):
        return True
    return (
        bool(normalizar(nombre_a))
        and normalizar(nombre_a) == normalizar(nombre_b)
        and normalizar(unidades_a) != normalizar(unidades_b)
    )


def mejor_coincidencia(
    nombre: str,
    metodo: str,
    maestro: list[dict[str, str]],
    idioma: str = "DESCONOCIDO",
) -> dict[str, Any]:
    if not maestro:
        raise ValueError("La lista maestra no contiene registros.")
    if not texto(nombre):
        return {
            "estado": "SIN_COINCIDENCIA", "puntaje": 0.0, "puntaje_nombre": 0.0,
            "puntaje_metodo": 0.0, "margen": 0.0, "ambiguo": False,
            "registro": {}, "segunda_opcion": {},
        }

    evaluados = []
    for registro in maestro:
        total, sn, sm = puntuar(nombre, metodo, registro)
        evaluados.append((total, sn, sm, registro))
    evaluados.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
    mejor = evaluados[0]
    segundo = evaluados[1] if len(evaluados) > 1 else (0.0, 0.0, 0.0, {})
    margen = mejor[0] - segundo[0]
    ambiguo = (
        margen < MARGEN_AMBIGUEDAD
        and registros_ambiguos(mejor[3], segundo[3], idioma)
    )
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
    if not isinstance(datos, dict):
        raise ValueError("La raiz del JSON debe ser un objeto.")
    resultado = deepcopy(datos)
    pruebas = resultado.get("pruebas")
    if not isinstance(pruebas, list):
        raise ValueError("El JSON no contiene una lista 'pruebas'.")
    idioma = idioma_especificacion(resultado)
    notas = resultado.get("notas")
    if isinstance(notas, list):
        resultado["notas"] = [
            normalizar_espacio_unidades(nota) if isinstance(nota, str) else nota
            for nota in notas
        ]
    conteo = {"alineados": 0, "revision": 0, "sin_coincidencia": 0}
    auditoria = []
    mensajes_astm_no_encontradas = []
    referencias_maestras = {
        referencia
        for registro in maestro
        for referencia in referencias_astm(registro.get("METODO", ""))
    }
    for indice, prueba in enumerate(pruebas, 1):
        if not isinstance(prueba, dict):
            raise ValueError(f"El elemento {indice} de 'pruebas' debe ser un objeto.")
        nombre_original = texto(prueba.get("prueba") or prueba.get("nombre"))
        metodo_original = texto(prueba.get("metodo"))
        unidades_originales = texto(prueba.get("unidades"))
        valor_tipico_original = texto(prueba.get("valor_tipico"))
        astm_no_encontradas = referencias_astm_fuera_maestro(
            metodo_original, referencias_maestras
        )
        mensaje_astm = ""
        if astm_no_encontradas:
            referencias_texto = ", ".join(
                f"ASTM {referencia}" for referencia in astm_no_encontradas
            )
            nombre_prueba = f" ({nombre_original})" if nombre_original else ""
            mensaje_astm = (
                f"Prueba {indice}{nombre_prueba}: Referencias no encontradas "
                f"en la lista maestra: {referencias_texto}."
            )
            mensajes_astm_no_encontradas.append(mensaje_astm)
        match = mejor_coincidencia(nombre_original, metodo_original, maestro, idioma)
        registro = match["registro"]
        nombre_maestro, unidad_maestra = nombres_canonicos(registro, idioma)
        metodo_maestro_fuente = registro.get("METODO", "")
        metodo_maestro = metodo_maestro_canonico(metodo_maestro_fuente)
        if match["estado"] == "ALINEADO" and not nombre_maestro:
            match["estado"] = "REVISION_REQUERIDA"
        if match["estado"] == "ALINEADO":
            prueba["prueba"] = nombre_maestro
            prueba["metodo"] = metodo_maestro
            if unidad_maestra:
                prueba["unidades"] = normalizar_formato_numerico(unidad_maestra)
            if "nombre" in prueba:
                prueba["nombre"] = nombre_maestro
            conteo["alineados"] += 1
        elif match["estado"] == "REVISION_REQUERIDA":
            conteo["revision"] += 1
        else:
            conteo["sin_coincidencia"] += 1
        for campo in CAMPOS_CONTENIDO_TABLA:
            valor_celda = prueba.get(campo)
            if isinstance(valor_celda, str):
                if campo in {"unidades", "valor_tipico"}:
                    valor_celda = normalizar_formato_numerico(valor_celda)
                prueba[campo] = normalizar_espacio_unidades(valor_celda)
        valores_alineados = {
            "prueba": texto(registro.get("NOMBRE IQM")) or nombre_original,
            "unidades": (
                normalizar_espacio_unidades(normalizar_formato_numerico(unidad_maestra))
                if unidad_maestra
                else "" if prueba.get("unidades") is None else str(prueba["unidades"])
            ),
            "metodo": (
                metodo_maestro
                or metodo_referencias_canonico(metodo_original)
                or metodo_original
            ),
            "especificacion": (
                "" if prueba.get("especificacion") is None else str(prueba["especificacion"])
            ),
            "valor_tipico": (
                "" if prueba.get("valor_tipico") is None else str(prueba["valor_tipico"])
            ),
        }
        auditoria.append({
            "numero_prueba": indice, "estado": match["estado"], "nombre_original": nombre_original,
            "metodo_original": metodo_original, "nombre_maestro_propuesto": registro.get("NOMBRE IQM", ""),
            "nombre_maestro_ingles_propuesto": registro.get("NOMBRE IQM INGLES", ""),
            "metodo_original_normalizado": metodo_referencias_canonico(
                metodo_original
            ),
            "metodo_maestro_fuente": metodo_maestro_fuente,
            "metodo_maestro_propuesto": metodo_maestro,
            "referencias_astm_originales": list(referencias_astm(metodo_original)),
            "referencias_astm_no_encontradas": list(astm_no_encontradas),
            "mensaje_astm_no_encontradas": mensaje_astm,
            "referencias_astm_maestras": list(referencias_astm(metodo_maestro_fuente)),
            "referencias_farmacopea_originales": list(
                referencias_farmacopea(metodo_original)
            ),
            "referencias_farmacopea_maestras": list(
                referencias_farmacopea(metodo_maestro_fuente)
            ),
            "fila_maestro": registro.get("FILA_MAESTRO", ""),
            "nombre_segunda_opcion": match["segunda_opcion"].get("NOMBRE IQM", ""),
            "metodo_segunda_opcion": match["segunda_opcion"].get("METODO", ""),
            "fila_segunda_opcion": match["segunda_opcion"].get("FILA_MAESTRO", ""),
            "margen_puntaje": match["margen"],
            "unidades_originales": unidades_originales,
            "valor_tipico_original": valor_tipico_original,
            "unidades_maestras_propuestas": unidad_maestra, "idioma": idioma,
            "puntaje_total": match["puntaje"], "puntaje_nombre": match["puntaje_nombre"],
            "puntaje_metodo": match["puntaje_metodo"], "ambiguo": match["ambiguo"],
            "valores_alineados": valores_alineados,
        })
    resultado["alineacion_metodos_lab"] = {
        "version_programa": VERSION, "fecha": datetime.now().isoformat(timespec="seconds"),
        "archivo_maestro": MAESTRO.as_posix(), "resumen": conteo, "detalle": auditoria,
        "mensajes_astm_no_encontradas": mensajes_astm_no_encontradas,
    }
    salida.parent.mkdir(parents=True, exist_ok=True)
    archivo_temporal = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=salida.parent,
            prefix=f".{salida.name}.",
            suffix=".tmp",
            delete=False,
        ) as temporal:
            archivo_temporal = Path(temporal.name)
            json.dump(resultado, temporal, ensure_ascii=False, indent=2)
            temporal.write("\n")
        os.replace(archivo_temporal, salida)
    finally:
        if archivo_temporal is not None:
            archivo_temporal.unlink(missing_ok=True)
    return conteo


def listar_json(carpeta: Path) -> list[Path]:
    return sorted(
        (path for path in carpeta.rglob("*") if path.is_file() and path.suffix.lower() == EXTENSION),
        key=lambda path: str(path.relative_to(carpeta)).casefold(),
    )


def ejecutar(base: Path) -> dict[str, int]:
    entrada, salida, maestro_path = base / ENTRADA, base / SALIDA, base / MAESTRO
    if not entrada.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de entrada: {entrada}")
    maestro = cargar_maestro(maestro_path)
    archivos = listar_json(entrada)
    if not archivos:
        raise FileNotFoundError(f"No se encontraron JSON en {entrada}")
    salida.mkdir(parents=True, exist_ok=True)
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
    return total


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
        resumen = ejecutar(args.base.resolve())
        return 1 if resumen["errores"] else 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
