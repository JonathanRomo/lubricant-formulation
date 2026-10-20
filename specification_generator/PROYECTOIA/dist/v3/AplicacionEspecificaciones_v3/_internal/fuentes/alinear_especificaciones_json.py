#!/usr/bin/env python3
"""Alinea nombres de prueba y metodos en JSON usando METODOS_LAB_UNIFICADO.xlsx."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import tempfile
import unicodedata
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any, Callable
from rutas_proyecto import BASE, resolver_ruta, ruta_portatil, rutas_configuradas, reubicar_ruta_heredada, carpeta_listas

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

VERSION = "1.13.1"
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
NOTAS_ESTANDARIZADAS = (
    (
        re.compile(
            r"(?<!\w)(?:Notas?\s*:\s*)?La\s+versi[oó]n\s+vigente\s+del\s+"
            r"m[eé]todo\s+aplicable\s+corresponder[aá]\s+a\s+la\s+"
            r"versi[oó]n\s+emitida\s+proporcionada\s+por\s+el\s+"
            r"laboratorio\.?(?!\w)",
            re.IGNORECASE,
        ),
        "La versión vigente del método aplicable corresponderá a la "
        "versión emitida proporcionada por el laboratorio.",
    ),
    (
        re.compile(
            r"(?<!\w)(?:Notes?\s*:\s*)?The\s+active\s+version\s+of\s+the\s+"
            r"applicable\s+method\s+shall\s+be\s+the\s+version\s+"
            r"issued\s+and\s+provided\s+by\s+the\s+laboratory\.?(?!\w)",
            re.IGNORECASE,
        ),
        "The active version of the applicable method shall be the "
        "version issued and provided by the laboratory.",
    ),
)
UNIDADES_RECONOCIDAS = (
    "mPa·s", "mPa.s", "mg/kg", "mg/L", "g/L", "mL/min", "mm²/s", "mm2/s",
    "kg/m³", "kg/m3", "kg/m^3", "m/s", "N·m", "N.m", "N*m", "Pa·s",
    "Pa.s", "Pa*s", "s⁻¹", "s-1", "s^-1", "cm-1", "cm³", "cm3", "cm^3",
    "cm²", "cm2", "cm^2", "mm³", "mm3", "mm^3", "mm²", "mm2", "mm^2",
    "µL", "μL", "uL", "mL", "ml", "µm", "μm", "um", "kg", "kgs", "mg",
    "mgs", "µg", "μg", "ug", "mm", "cm", "km", "m", "sec", "secs", "min",
    "mins", "hr", "hrs", "rpm", "cSt", "cP", "kPa", "MPa", "Pa", "bar",
    "mol", "ppm", "ppb", "N", "A", "K", "cd", "V", "W", "J", "kJ", "s",
    "h", "d", "L", "l", "%",
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
MAXIMO_ESPECIFICACION = re.compile(
    r"(?<!\w)M[ÁA]X(?:IMO)?\.?(?!\w)",
    re.IGNORECASE,
)
MINIMO_ESPECIFICACION_COMPLETO = re.compile(
    r"(?<!\w)M[IÍ]NIMO(?!\w)",
    re.IGNORECASE,
)
MINIMO_ESPECIFICACION_ABREVIADO = re.compile(
    r"(?<!\w)M[IÍ]N\.?(?!\w)"
)
TERMINOS_TIPICO_ES = re.compile(
    r"(?<!\w)T[IÍ]PIC[OA]S?(?!\w)",
    re.IGNORECASE,
)
TERMINOS_TIPICO_EN = re.compile(
    r"(?<!\w)TYPICALS?(?!\w)",
    re.IGNORECASE,
)
TERMINOS_PASA_ES = re.compile(
    r"(?<!\w)PAS(?:A|AN|[OÓ]|ARON|ANDO|AD[OA]S?)(?!\w)",
    re.IGNORECASE,
)
TERMINOS_PASS_EN = re.compile(
    r"(?<!\w)(?:PASSES|PASSED|PASSING|PASS)(?!\w)",
    re.IGNORECASE,
)
TERMINOS_REPORTAR_ES = re.compile(
    r"(?<!\w)REPORT(?:AR|A|AN|[OÓ]|ARON|ANDO|AD[OA]S?)(?!\w)",
    re.IGNORECASE,
)
TERMINOS_REPORT_EN = re.compile(
    r"(?<!\w)REPORT(?:S|ED|ING)?(?!\w)",
    re.IGNORECASE,
)
TERMINOS_NEGATIVO_ES = re.compile(
    r"(?<!\w)NEGATIV[OA]S?(?!\w)",
    re.IGNORECASE,
)
TERMINOS_NEGATIVE_EN = re.compile(
    r"(?<!\w)NEGATIVES?(?!\w)",
    re.IGNORECASE,
)
FRAGMENTOS_PROTEGIDOS_MEDICION = re.compile(
    r"(?<!\w)(?:\d{1,4}[./-]\d{1,2}[./-]\d{1,4})(?![\w.,])"
    r"|(?<!\w)(?:\d{1,3}\.){3}\d{1,3}(?!\w)"
    r"|(?<!\w)(?:v(?:er(?:sion)?)?\s*\d+(?:\.\d+)+|"
    r"\d{1,4}\.\d{1,2}\.\d{1,2})(?!\w)"
    r"|(?<!\w)(?:ASTM|ISO|IEC|NMX|USP|FEUM|EN|DIN|IP|SAE|API|JASO|MIL)"
    r"\s*(?:[A-Z]\s*)?\d+(?:[./-][A-Z0-9]+)*(?!\w)"
    r"|(?<!\w)(?!(?:cm|mm|kg|mg)-\d)(?=[A-Z0-9_-]*\d)"
    r"[A-Z]{2,}[A-Z0-9]*(?:[-_][A-Z0-9]+)+(?!\w)"
    r"|(?<!\w)(?:SN|S/N|ID|SERIAL(?:\s+NO\.?)?)\s*[:#-]?\s*"
    r"[A-Z0-9][A-Z0-9._/-]*(?!\w)",
    re.IGNORECASE,
)
NUMERO_AMBIGUO_MEDICION = re.compile(
    r"(?<![\w.,])(?P<valor>[+-]?\d+(?:[ \u00a0\u2009\u202f]\d{3})*)"
    r"[.,]\d{3}(?![\w.,])"
)
NUMERO_MEDICION = re.compile(
    r"(?<!\w)(?P<signo>[+−-]?)"
    r"(?P<valor>(?:\d[\d., \u00a0\u2009\u202f]*\d|\d|[.,]\d+))"
    r"(?!\w)"
)
NOTACION_CIENTIFICA_E = re.compile(
    r"(?<!\w)(?P<coef>[+-]?(?:\d+(?:[.,]\d+)?|[.,]\d+))"
    r"\s*[eE]\s*(?P<exp>[+-]?\d+)(?!\w)"
)
NOTACION_CIENTIFICA_X = re.compile(
    r"(?<!\w)(?P<coef>[+-]?(?:\d+(?:[.,]\d+)?|[.,]\d+))"
    r"\s*[x×]\s*10(?:\s*\^\s*|\s*)(?P<exp>[+-]?\d+)(?!\w)",
    re.IGNORECASE,
)

# Umbrales conservadores para evitar sustituir una prueba por otra incorrecta.
UMBRAL_AUTOMATICO = 0.84
UMBRAL_DUDOSO = 0.72
MARGEN_AMBIGUEDAD = 0.035
COLUMNAS_MAESTRO = (
    "METODO", "MÉTODO INGLÉS", "NOMBRE IQM", "NOMBRE IQM INGLES",
    "UNIDADES", "UNITS", "ACREDITACION NOMBRE",
)
COLUMNAS_MAESTRO_REQUERIDAS = (
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


def normalizar_especificacion(valor: Any) -> str:
    """Normaliza máximo/mínimo sin confundir «min» minúscula con minutos."""
    salida = "" if valor is None else str(valor)
    salida = MAXIMO_ESPECIFICACION.sub("MÁXIMO", salida)
    salida = MINIMO_ESPECIFICACION_COMPLETO.sub("MÍNIMO", salida)
    return MINIMO_ESPECIFICACION_ABREVIADO.sub("MÍNIMO", salida)


def normalizar_terminos_especificacion(valor: Any) -> str:
    """Corrige términos frecuentes en especificaciones y valores típicos."""
    salida = texto(valor)

    def normalizar_tipico_es(coincidencia: re.Match[str]) -> str:
        palabra = coincidencia.group(0).upper()
        return "TÍ" + palabra[2:]

    salida = TERMINOS_NEGATIVO_ES.sub("NEGATIVA", salida)
    salida = TERMINOS_NEGATIVE_EN.sub("NEGATIVE", salida)
    for patron in (
        TERMINOS_TIPICO_ES,
        TERMINOS_TIPICO_EN,
        TERMINOS_PASA_ES,
        TERMINOS_PASS_EN,
        TERMINOS_REPORTAR_ES,
        TERMINOS_REPORT_EN,
    ):
        if patron is TERMINOS_TIPICO_ES:
            salida = patron.sub(normalizar_tipico_es, salida)
        else:
            salida = patron.sub(
                lambda coincidencia: coincidencia.group(0).upper(), salida
            )
    return salida


def _normalizar_simbolos_unidades(valor: str) -> str:
    compuestas = (
        (r"mPa\s*[.*·]\s*s", "mPa·s", 0),
        (r"mPa", "mPa", 0),
        (r"kg\s*/\s*m(?:\s*\^\s*|\s*)3|kg\s*/\s*m³", "kg/m³", re.IGNORECASE),
        (r"mg\s*/\s*l(?:iter|itre)?s?", "mg/L", re.IGNORECASE),
        (r"m\s*/\s*s(?:ec(?:ond)?s?)?", "m/s", re.IGNORECASE),
        (r"N\s*[.*·]\s*m", "N·m", re.IGNORECASE),
        (r"Pa\s*[.*·]\s*s", "Pa·s", re.IGNORECASE),
        (r"s\s*(?:\^\s*-?1|-1|⁻¹)", "s⁻¹", re.IGNORECASE),
        (r"cm\s*(?:\^\s*-1|-1|⁻¹)", "cm⁻¹", re.IGNORECASE),
        (r"(?:mm)\s*(?:\^\s*2|2|²)", "mm²", re.IGNORECASE),
        (r"(?:mm)\s*(?:\^\s*3|3|³)", "mm³", re.IGNORECASE),
        (r"(?:cm)\s*(?:\^\s*2|2|²)", "cm²", re.IGNORECASE),
        (r"(?:cm)\s*(?:\^\s*3|3|³)", "cm³", re.IGNORECASE),
        (r"m\s*(?:\^\s*2|2|²)", "m²", re.IGNORECASE),
        (r"m\s*(?:\^\s*3|3|³)", "m³", re.IGNORECASE),
        (r"º\s*C|°\s*C|deg(?:ree)?s?\s+C(?:elsius)?", "°C", re.IGNORECASE),
        (r"megapascals?", "MPa", re.IGNORECASE),
        (r"MPa", "MPa", 0),
        (r"MPA", "MPa", 0),
        (r"kilopascals?", "kPa", re.IGNORECASE),
        (r"kPa", "kPa", 0),
        (r"pascals?|pa", "Pa", re.IGNORECASE),
        (r"kilograms?|kgs?", "kg", re.IGNORECASE),
        (r"milligrams?|mgs?", "mg", re.IGNORECASE),
        (r"micrograms?|mcgs?|[µμ]g|ugs?", "µg", re.IGNORECASE),
        (r"millilit(?:er|re)s?|mls?", "mL", re.IGNORECASE),
        (r"microliters?|microlitres?|[µμ]ls?|uls?", "µL", re.IGNORECASE),
        (r"liters?|litres?|lts?", "L", re.IGNORECASE),
        (r"l", "L", 0),
        (r"kilometers?|kilometres?|kms?", "km", re.IGNORECASE),
        (r"millimeters?|millimetres?", "mm", re.IGNORECASE),
        (r"centimeters?|centimetres?", "cm", re.IGNORECASE),
        (r"meters?|metres?", "m", re.IGNORECASE),
        (r"m", "m", re.IGNORECASE),
        (r"grams?|gs?", "g", re.IGNORECASE),
        (r"seconds?|secs?", "s", re.IGNORECASE),
        (r"minutes?|mins?", "min", re.IGNORECASE),
        (r"hours?|hrs?", "h", re.IGNORECASE),
    )
    salida = valor
    for patron, simbolo, flags in compuestas:
        salida = re.sub(
            rf"(?<![^\W\d_])(?:{patron})\.?(?![^\W\d_])",
            simbolo,
            salida,
            flags=flags,
        )
    return salida


def _agrupar_entero(digitos: str) -> str:
    digitos = re.sub(r"\D", "", digitos)
    if not digitos:
        return "0"
    primer_grupo = len(digitos) % 3 or 3
    grupos = [digitos[:primer_grupo]]
    grupos.extend(
        digitos[indice : indice + 3]
        for indice in range(primer_grupo, len(digitos), 3)
    )
    return " ".join(grupos)


def _formatear_coeficiente_cientifico(coeficiente: str) -> str:
    signo = ""
    if coeficiente[:1] in "+-−":
        signo, coeficiente = coeficiente[0], coeficiente[1:]
    if coeficiente.startswith((".", ",")):
        coeficiente = "0" + coeficiente
    if "." in coeficiente or "," in coeficiente:
        entero, fraccion = re.split(r"[.,]", coeficiente, maxsplit=1)
        numero = f"{_agrupar_entero(entero)},{fraccion}"
    else:
        numero = _agrupar_entero(coeficiente)
    if signo == "-":
        signo = "−"
    return f"{signo} {numero}" if signo else numero


def _exponente_superindice(exponente: str) -> str:
    traduccion = str.maketrans("0123456789+-", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻")
    return exponente.translate(traduccion)


def _formatear_numero_medicion(numero: str) -> tuple[str, bool]:
    compacto = re.sub(r"[ \u00a0\u2009\u202f]", "", numero)
    separadores = [
        (indice, caracter)
        for indice, caracter in enumerate(compacto)
        if caracter in ".,"
    ]
    if not separadores:
        return _agrupar_entero(compacto), False

    tipos = {caracter for _, caracter in separadores}
    if len(tipos) > 1:
        indice_decimal = separadores[-1][0]
        parte_entera = compacto[:indice_decimal]
        fraccion = compacto[indice_decimal + 1 :]
        if not fraccion.isdigit():
            return numero, True
        digitos_entero = re.sub(r"[.,]", "", parte_entera)
        if not digitos_entero.isdigit():
            return numero, True
        return f"{_agrupar_entero(digitos_entero)},{fraccion}", False

    separador = separadores[0][1]
    partes = compacto.split(separador)
    if len(partes) > 2:
        if len(partes[0]) not in range(1, 4) or any(
            len(parte) != 3 or not parte.isdigit() for parte in partes[1:]
        ):
            return numero, True
        return _agrupar_entero("".join(partes)), False

    parte_entera, fraccion = partes
    if not fraccion.isdigit():
        return numero, True
    return f"{_agrupar_entero(parte_entera)},{fraccion}", False


def normalizar_valor_medicion(valor: Any) -> tuple[str, list[str]]:
    """Normaliza mediciones y conserva valores ambiguos para revisión."""
    original = texto(valor)
    if not original:
        return original, []

    protegidos: list[str] = []

    def guardar_fragmento(coincidencia: re.Match[str]) -> str:
        indice = len(protegidos)
        protegidos.append(coincidencia.group(0))
        return f"QZPROTEGIDO{indice}QZ"

    salida = FRAGMENTOS_PROTEGIDOS_MEDICION.sub(guardar_fragmento, original)
    salida = _normalizar_simbolos_unidades(salida)
    salida = normalizar_espacio_unidades(salida)
    salida = RANGO_NUMERICO.sub(
        lambda coincidencia: (
            f"{coincidencia.group('izquierdo')} - "
            f"{coincidencia.group('derecho')}"
        ),
        salida,
    )
    salida = re.sub(
        r"(?<!\w)-(?=(?:\d|[.,]\d))",
        "− ",
        salida,
    )
    salida = re.sub(
        r"(?<=\d)\s*([−+±<>≤≥])\s*(?=[+\-−]?(?:\d|[.,]\d))",
        r" \1 ",
        salida,
    )
    salida = re.sub(
        r"([−+±<>≤≥])\s*(?=(?:\d|[.,]\d))",
        r"\1 ",
        salida,
    )

    ambiguos: list[str] = []

    def guardar_ambiguo(coincidencia: re.Match[str]) -> str:
        indice = len(ambiguos)
        ambiguos.append(coincidencia.group(0))
        return f"QZAMBIGUO{indice}QZ"

    salida = NUMERO_AMBIGUO_MEDICION.sub(guardar_ambiguo, salida)

    def normalizar_cientifico(coincidencia: re.Match[str]) -> str:
        coeficiente = _formatear_coeficiente_cientifico(
            coincidencia.group("coef")
        )
        exponente = _exponente_superindice(coincidencia.group("exp"))
        return f"{coeficiente} × 10{exponente}"

    salida = NOTACION_CIENTIFICA_E.sub(normalizar_cientifico, salida)
    salida = NOTACION_CIENTIFICA_X.sub(normalizar_cientifico, salida)

    def normalizar_numero(coincidencia: re.Match[str]) -> str:
        numero, ambiguo = _formatear_numero_medicion(
            coincidencia.group("valor")
        )
        if ambiguo:
            ambiguos.append(coincidencia.group(0))
        signo = coincidencia.group("signo")
        if signo == "-":
            signo = "− "
        elif signo:
            signo += " "
        return f"{signo}{numero}"

    salida = NUMERO_MEDICION.sub(normalizar_numero, salida)
    for indice, numero in enumerate(ambiguos):
        salida = salida.replace(f"QZAMBIGUO{indice}QZ", numero)
    for indice, fragmento in enumerate(protegidos):
        salida = salida.replace(f"QZPROTEGIDO{indice}QZ", fragmento)
    return salida, ambiguos


def normalizar(valor: Any) -> str:
    salida = texto(valor).upper()
    salida = "".join(c for c in unicodedata.normalize("NFD", salida) if unicodedata.category(c) != "Mn")
    salida = salida.replace("×", "X")
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9%]+", " ", salida)).strip()


def limpiar_notas_estandarizadas(
    pruebas: list[Any], notas: list[Any]
) -> tuple[list[Any], list[Any]]:
    """Separa las notas estándar de las filas y campos de pruebas."""
    claves_notas = {
        normalizar(canonica)
        for nota in notas
        if isinstance(nota, str)
        for patron, canonica in NOTAS_ESTANDARIZADAS
        if patron.search(nota)
    }
    pruebas_limpias = []
    for prueba in pruebas:
        if not isinstance(prueba, dict):
            pruebas_limpias.append(prueba)
            continue
        for campo in CAMPOS_CONTENIDO_TABLA:
            valor = prueba.get(campo)
            if not isinstance(valor, str):
                continue
            for patron, canonica in NOTAS_ESTANDARIZADAS:
                if not patron.search(valor):
                    continue
                clave_nota = normalizar(canonica)
                if clave_nota not in claves_notas:
                    notas.append(canonica)
                    claves_notas.add(clave_nota)
                valor = patron.sub("", valor)
            valor = re.sub(r"\s+([.,;:])", r"\1", valor)
            prueba[campo] = valor.strip(" \t\r\n,;:")
        if any(texto(prueba.get(campo)) for campo in CAMPOS_CONTENIDO_TABLA):
            pruebas_limpias.append(prueba)
    return pruebas_limpias, notas


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


def metodo_maestro_para_idioma(registro: dict[str, str], idioma: str) -> str:
    """Devuelve el método del maestro en el idioma solicitado, con fallback."""
    idioma_normalizado = normalizar(idioma).replace(" ", "")
    if idioma_normalizado in ALIAS_IDIOMA_INGLES:
        metodo_ingles = registro.get("MÉTODO INGLÉS", "") or registro.get(
            "METODO INGLES", ""
        )
        if texto(metodo_ingles):
            return texto(metodo_ingles)
    return texto(registro.get("METODO", ""))


def metodos_maestros(registro: dict[str, str]) -> tuple[str, ...]:
    """Devuelve métodos distintos de ambos idiomas para validar referencias."""
    return tuple(
        dict.fromkeys(
            metodo
            for metodo in (
                texto(registro.get("METODO", "")),
                texto(
                    registro.get("MÉTODO INGLÉS", "")
                    or registro.get("METODO INGLES", "")
                ),
            )
            if metodo
        )
    )


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
    faltantes = [
        columna
        for columna in COLUMNAS_MAESTRO_REQUERIDAS
        if normalizar(columna) not in encabezados
    ]
    if faltantes:
        raise ValueError("Faltan columnas en la hoja METODOS: " + ", ".join(faltantes))

    filas = []
    for numero, (_, fila) in enumerate(tabla.iterrows(), 2):
        registro = {
            columna: (
                texto(fila.iloc[encabezados[normalizar(columna)]])
                if normalizar(columna) in encabezados
                else ""
            )
            for columna in COLUMNAS_MAESTRO
        }
        if not registro["MÉTODO INGLÉS"]:
            registro["MÉTODO INGLÉS"] = registro["METODO"]
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


def puntuar(
    nombre: str,
    metodo: str,
    candidato: dict[str, str],
    idioma: str = "DESCONOCIDO",
) -> tuple[float, float, float]:
    nombre_es = similitud_texto(
        nombre, limpiar_marca_acreditacion(candidato["NOMBRE IQM"])
    )
    nombre_en = similitud_texto(
        nombre, limpiar_marca_acreditacion(candidato["NOMBRE IQM INGLES"])
    )
    score_nombre = max(nombre_es, nombre_en)
    metodo_candidato = metodo_maestro_para_idioma(candidato, idioma)
    score_metodo = similitud_metodo(metodo, metodo_candidato)
    # El nombre selecciona la familia; el metodo resuelve nombres repetidos.
    if texto(metodo):
        total = 0.68 * score_nombre + 0.32 * score_metodo
        if score_nombre >= 0.96 and score_metodo >= 0.90:
            total = max(total, 0.96)
        elif (
            referencias_astm(metodo)
            and referencias_astm(metodo_candidato)
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


def unidad_original_esta_entre_alternativas(
    unidad_original: Any, unidad_maestra: Any
) -> bool:
    original = normalizar(normalizar_espacio_unidades(unidad_original))
    if not original:
        return False
    return any(
        original == normalizar(normalizar_espacio_unidades(alternativa))
        for alternativa in alternativas_unidades(unidad_maestra)
    )


def alternativas_unidades(unidades: Any) -> list[str]:
    return [
        alternativa.strip()
        for alternativa in re.split(
            r"\s*,\s*|\s+/\s+|\s+(?:o|or)\s+",
            texto(unidades),
            flags=re.IGNORECASE,
        )
        if alternativa.strip()
    ]


def registros_ambiguos(a: dict[str, str], b: dict[str, str], idioma: str) -> bool:
    nombre_a, unidades_a = nombres_canonicos(a, idioma)
    nombre_b, unidades_b = nombres_canonicos(b, idioma)
    metodo_a = metodo_maestro_canonico(metodo_maestro_para_idioma(a, idioma))
    metodo_b = metodo_maestro_canonico(metodo_maestro_para_idioma(b, idioma))
    if normalizar(metodo_a) != normalizar(metodo_b):
        return True
    return (
        bool(normalizar(nombre_a))
        and normalizar(nombre_a) == normalizar(nombre_b)
        and {
            normalizar(alternativa)
            for alternativa in alternativas_unidades(unidades_a)
        }
        != {
            normalizar(alternativa)
            for alternativa in alternativas_unidades(unidades_b)
        }
    )


class MaestroPreparado(list):
    """Datos derivados y coincidencias reutilizables, aislados por lote."""
    def __init__(self, registros):
        super().__init__(deepcopy(registros))
        self.rasgos = {}
        self.coincidencias = {}
        self.candidatos = []
        self.referencias = frozenset(
            referencia for registro in self for metodo in metodos_maestros(registro)
            for referencia in referencias_astm(metodo)
        )
        for registro in self:
            es = limpiar_marca_acreditacion(registro["NOMBRE IQM"])
            en = limpiar_marca_acreditacion(registro["NOMBRE IQM INGLES"])
            metodos = {idioma: metodo_maestro_para_idioma(registro, idioma)
                       for idioma in ("ESPANOL", "INGLES", "DESCONOCIDO")}
            for valor in (es, en, *metodos.values()):
                self.preparar_texto(valor)
            self.candidatos.append((registro, es, en, metodos))

    def preparar_texto(self, valor):
        clave = texto(valor)
        if clave not in self.rasgos:
            self.rasgos[clave] = (normalizar(clave), frozenset(referencias_astm(clave)),
                                 frozenset(partes_metodo(clave)))
        return self.rasgos[clave]

    def similitud(self, a, b):
        na, ra, _ = self.preparar_texto(a)
        nb, rb, _ = self.preparar_texto(b)
        if not na or not nb:
            return 0.0
        if ra or rb:
            if not ra or not rb:
                return 0.0
            inter = ra & rb
            if ra == rb:
                return 1.0
            if inter and (ra <= rb or rb <= ra):
                return 0.90 + 0.10 * len(inter) / max(len(ra), len(rb))
            return 0.50 * len(inter) / len(ra | rb) if inter else 0.0
        if na == nb:
            return 1.0
        return (0.75 * (fuzz.ratio(na, nb) / 100)
                + 0.15 * (fuzz.token_sort_ratio(na, nb) / 100)
                + 0.10 * (fuzz.token_set_ratio(na, nb) / 100))

    def puntuar(self, nombre, metodo, candidato, idioma):
        _, es, en, metodos = candidato
        sn = max(self.similitud(nombre, es), self.similitud(nombre, en))
        mc = metodos.get(idioma)
        if mc is None:
            mc = metodo_maestro_para_idioma(candidato[0], idioma)
        na, ra, pa = self.preparar_texto(metodo)
        nb, rb, pb = self.preparar_texto(mc)
        if not na and not nb:
            sm = 0.5
        elif not na or not nb:
            sm = 0.0
        elif na == nb:
            sm = 1.0
        elif pa and pb:
            inter = len(pa & pb)
            sm = 0.90 + 0.10 * inter / max(len(pa), len(pb)) if inter else 0.0
        else:
            sm = self.similitud(metodo, mc)
        if texto(metodo):
            total = 0.68 * sn + 0.32 * sm
            if sn >= 0.96 and sm >= 0.90:
                total = max(total, 0.96)
            elif ra and rb and sm >= 0.90 and sn >= 0.45:
                total = max(total, UMBRAL_AUTOMATICO)
        else:
            total = 0.88 * sn + 0.12 * sm
        return total, sn, sm


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

    clave_cache = (nombre, metodo, idioma)
    preparado = isinstance(maestro, MaestroPreparado)
    if preparado and clave_cache in maestro.coincidencias:
        return deepcopy(maestro.coincidencias[clave_cache])
    mejor = segundo = None
    for indice, registro in enumerate(maestro):
        total, sn, sm = (maestro.puntuar(nombre, metodo, maestro.candidatos[indice], idioma)
                         if preparado else puntuar(nombre, metodo, registro, idioma))
        candidato = (total, sn, sm, registro)
        if mejor is None or candidato[:3] > mejor[:3]:
            segundo, mejor = mejor, candidato
        elif segundo is None or candidato[:3] > segundo[:3]:
            segundo = candidato
    segundo = segundo if segundo is not None else (0.0, 0.0, 0.0, {})
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
    resultado_match = {"estado": estado, "puntaje": round(mejor[0], 4), "puntaje_nombre": round(mejor[1], 4),
            "puntaje_metodo": round(mejor[2], 4), "margen": round(margen, 4), "ambiguo": ambiguo,
            "registro": mejor[3], "segunda_opcion": segundo[3]}
    if preparado:
        maestro.coincidencias[clave_cache] = deepcopy(resultado_match)
    return resultado_match


def procesar_json(path: Path, salida: Path, maestro: list[dict[str, str]]) -> dict[str, int]:
    from generador_hii import conservar_historial
    if not isinstance(maestro, MaestroPreparado):
        maestro = MaestroPreparado(maestro)
    datos = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(datos, dict):
        raise ValueError("La raiz del JSON debe ser un objeto.")
    resultado = deepcopy(datos)
    resultado["FECHA DE REANALISIS"] = ""
    resultado["Desarrollador"] = ""
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
    elif notas is None:
        resultado["notas"] = []
    else:
        resultado["notas"] = [notas]
    resultado["pruebas"], resultado["notas"] = limpiar_notas_estandarizadas(
        pruebas, resultado["notas"]
    )
    pruebas = resultado["pruebas"]
    conteo = {"alineados": 0, "revision": 0, "sin_coincidencia": 0}
    auditoria = []
    mensajes_astm_no_encontradas = []
    referencias_maestras = maestro.referencias
    for indice, prueba in enumerate(pruebas, 1):
        if not isinstance(prueba, dict):
            raise ValueError(f"El elemento {indice} de 'pruebas' debe ser un objeto.")
        prueba["PERIODIDICAD DE ANALISIS"] = next(
            (prueba[campo] for campo in ("PERIODIDICAD DE ANALISIS", "periodicidad", "periodicidad_analisis")
             if texto(prueba.get(campo))), "▶")
        nombre_original = texto(prueba.get("prueba") or prueba.get("nombre"))
        metodo_original = texto(prueba.get("metodo"))
        unidades_originales = texto(prueba.get("unidades"))
        valor_tipico_original = texto(prueba.get("valor_tipico"))
        limpieza_especificacion, ambiguos_especificacion = normalizar_valor_medicion(
            normalizar_terminos_especificacion(
                normalizar_especificacion(prueba.get("especificacion"))
            )
        )
        limpieza_valor_tipico, ambiguos_valor_tipico = normalizar_valor_medicion(
            normalizar_terminos_especificacion(prueba.get("valor_tipico"))
        )
        observaciones_limpieza = []
        if ambiguos_especificacion:
            observaciones_limpieza.append(
                "Separador numérico ambiguo en especificacion: "
                + ", ".join(ambiguos_especificacion)
            )
        if ambiguos_valor_tipico:
            observaciones_limpieza.append(
                "Separador numérico ambiguo en valor_tipico: "
                + ", ".join(ambiguos_valor_tipico)
            )
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
        metodo_maestro_fuente = metodo_maestro_para_idioma(registro, idioma)
        metodo_maestro = metodo_maestro_canonico(metodo_maestro_fuente)
        if len(referencias_astm(metodo_original)) > len(
            referencias_astm(metodo_maestro_fuente)
        ):
            metodo_maestro = metodo_referencias_canonico(metodo_original)
        if match["estado"] == "ALINEADO" and not nombre_maestro:
            match["estado"] = "REVISION_REQUERIDA"
        if observaciones_limpieza:
            match["estado"] = "REVISION_REQUERIDA"
        alternativas_maestras = alternativas_unidades(unidad_maestra)
        unidades_limpias = (
            alternativas_maestras[0] if alternativas_maestras else unidad_maestra
        )
        if match["estado"] == "ALINEADO":
            prueba["prueba"] = nombre_maestro
            prueba["metodo"] = metodo_maestro
            if unidades_limpias:
                if unidad_original_esta_entre_alternativas(
                    unidades_originales, unidad_maestra
                ):
                    unidades_limpias = unidades_originales
                prueba["unidades"] = normalizar_formato_numerico(unidades_limpias)
            if "nombre" in prueba:
                prueba["nombre"] = nombre_maestro
            conteo["alineados"] += 1
        elif match["estado"] == "REVISION_REQUERIDA":
            conteo["revision"] += 1
        else:
            conteo["sin_coincidencia"] += 1
        if prueba.get("especificacion") is not None:
            prueba["especificacion"] = limpieza_especificacion
        if prueba.get("valor_tipico") is not None:
            prueba["valor_tipico"] = limpieza_valor_tipico
        for campo in CAMPOS_CONTENIDO_TABLA:
            valor_celda = prueba.get(campo)
            if isinstance(valor_celda, str):
                if campo == "unidades":
                    valor_celda = normalizar_formato_numerico(valor_celda)
                    valor_celda = normalizar_espacio_unidades(valor_celda)
                    prueba[campo] = valor_celda
                elif campo not in {"especificacion", "valor_tipico"}:
                    prueba[campo] = normalizar_espacio_unidades(valor_celda)
        valores_alineados = {
            "prueba": texto(registro.get("NOMBRE IQM")) or nombre_original,
            "unidades": (
                normalizar_espacio_unidades(
                    normalizar_formato_numerico(unidades_limpias)
                )
                if unidades_limpias
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
            "metodo_segunda_opcion": metodo_maestro_para_idioma(
                match["segunda_opcion"], idioma
            ),
            "fila_segunda_opcion": match["segunda_opcion"].get("FILA_MAESTRO", ""),
            "margen_puntaje": match["margen"],
            "unidades_originales": unidades_originales,
            "valor_tipico_original": valor_tipico_original,
            "observaciones_limpieza": observaciones_limpieza,
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
    conservar_historial(resultado, path, salida)
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


def ejecutar(
    base: Path,
    archivos_seleccionados: list[Path] | None = None,
    al_progreso: Callable[[int, int, str], None] | None = None,
    carpeta_entrada: Path | None = None,
    carpeta_salida: Path | None = None,
    destinos_seleccionados: list[Path] | None = None,
    forzar: bool = False,
) -> dict[str, int]:
    base = resolver_ruta(base)
    rutas = rutas_configuradas(base)
    entrada = resolver_ruta(carpeta_entrada if carpeta_entrada is not None else rutas["json_anteriores"], base)
    salida = resolver_ruta(carpeta_salida if carpeta_salida is not None else rutas["json_alineadas"], base)
    if archivos_seleccionados is not None:
        archivos_seleccionados = [resolver_ruta(ruta, base) for ruta in archivos_seleccionados]
    if destinos_seleccionados is not None:
        destinos_seleccionados = [resolver_ruta(ruta, base) for ruta in destinos_seleccionados]
    carpeta_maestro = carpeta_listas(base)
    maestro_path = base / MAESTRO
    if MAESTRO.parent == Path("METODOS_LAB") and carpeta_maestro != base.resolve() / "METODOS_LAB":
        maestro_path = carpeta_maestro / MAESTRO.name
    if not entrada.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de entrada: {entrada}")
    maestro = MaestroPreparado(cargar_maestro(maestro_path))
    hash_maestro = hashlib.sha256(maestro_path.read_bytes()).hexdigest()
    if destinos_seleccionados is not None and archivos_seleccionados is None:
        raise ValueError("Los destinos seleccionados requieren archivos de entrada.")
    if (
        destinos_seleccionados is not None
        and len(destinos_seleccionados) != len(archivos_seleccionados or ())
    ):
        raise ValueError("Cada archivo de entrada debe tener un destino alineado.")
    if archivos_seleccionados is None:
        archivos = [(path, None) for path in listar_json(entrada)]
    elif destinos_seleccionados is None:
        archivos = [
            (path, None)
            for path in archivos_seleccionados
            if path.is_file() and path.suffix.lower() == EXTENSION
        ]
        archivos.sort(key=lambda par: str(par[0]))
    else:
        archivos = [
            (origen, destino)
            for origen, destino in zip(
                archivos_seleccionados, destinos_seleccionados
            )
            if origen.is_file() and origen.suffix.lower() == EXTENSION
        ]
        archivos.sort(key=lambda par: str(par[0]))
    if not archivos:
        raise FileNotFoundError(f"No se encontraron JSON en {entrada}")
    archivos_relativos = []
    for origen, destino_seleccionado in archivos:
        try:
            origen.relative_to(entrada)
        except ValueError as exc:
            raise ValueError(
                f"El archivo seleccionado no pertenece a la carpeta de entrada: {origen}"
            ) from exc
        if destino_seleccionado is None:
            relativo_salida = origen.relative_to(entrada)
        else:
            try:
                relativo_salida = destino_seleccionado.relative_to(salida)
            except ValueError as exc:
                raise ValueError(
                    "El destino seleccionado no pertenece a la carpeta de salida: "
                    f"{destino_seleccionado}"
                ) from exc
            if ".." in relativo_salida.parts:
                raise ValueError(
                    f"El destino seleccionado no es válido: {destino_seleccionado}"
                )
        archivos_relativos.append((origen, relativo_salida))
    salida.mkdir(parents=True, exist_ok=True)
    ruta_cache = salida / ".cache_alineacion"
    try:
        cache = json.loads(ruta_cache.read_text(encoding="utf-8"))
        if not isinstance(cache, dict):
            cache = {}
    except (OSError, ValueError):
        cache = {}
    # La caché también debe sobrevivir al traslado de la carpeta del proyecto.
    cache = {
        ruta_portatil(reubicar_ruta_heredada(clave, base), base):
        dict(entrada_cache, origen=ruta_portatil(reubicar_ruta_heredada(entrada_cache["origen"], base), base))
        for clave, entrada_cache in cache.items()
        if isinstance(entrada_cache, dict) and isinstance(entrada_cache.get("origen"), str)
    }
    total = {"archivos": 0, "alineados": 0, "revision": 0, "sin_coincidencia": 0, "errores": 0, "omitidos": 0}
    lineas = [f"REPORTE DE ALINEACION {datetime.now().isoformat(timespec='seconds')}", "=" * 72]
    for indice, (origen, rel) in enumerate(archivos_relativos, 1):
        if al_progreso is not None:
            al_progreso(indice - 1, len(archivos), str(rel))
        destino = salida / rel
        try:
            clave = ruta_portatil(destino, base)
            hash_origen = hashlib.sha256(origen.read_bytes()).hexdigest()
            entrada_cache = cache.get(clave, {})
            vigente = (not forzar and isinstance(entrada_cache, dict)
                       and entrada_cache.get("origen") == ruta_portatil(origen, base)
                       and entrada_cache.get("hash_origen") == hash_origen
                       and entrada_cache.get("hash_maestro") == hash_maestro
                       and entrada_cache.get("version") == VERSION
                       and destino.is_file()
                       and entrada_cache.get("hash_destino") == hashlib.sha256(destino.read_bytes()).hexdigest())
            if vigente:
                total["omitidos"] += 1
                lineas.append(f"SIN CAMBIOS | {rel.as_posix()}")
                print(f"SIN CAMBIOS: {rel}")
                continue
            r = procesar_json(origen, destino, maestro)
            hash_destino = hashlib.sha256(destino.read_bytes()).hexdigest()
            cache[clave] = {"origen": ruta_portatil(origen, base), "version": VERSION,
                            "hash_origen": hash_destino if origen.resolve() == destino.resolve() else hash_origen,
                            "hash_maestro": hash_maestro, "hash_destino": hash_destino}
            total["archivos"] += 1
            for k in ("alineados", "revision", "sin_coincidencia"):
                total[k] += r[k]
            lineas.append(f"OK | {rel.as_posix()} | alineados={r['alineados']} revision={r['revision']} sin_coincidencia={r['sin_coincidencia']}")
            print(f"ALINEADO: {rel}")
        except Exception as exc:
            total["errores"] += 1
            lineas.append(f"ERROR | {rel.as_posix()} | {exc}")
            print(f"ERROR: {rel}: {exc}", file=sys.stderr)
        finally:
            if al_progreso is not None:
                al_progreso(indice, len(archivos), str(rel))
    import generador_hii as hii
    with hii.conectar(base / "especificaciones.db") as relaciones:
        cambios_relaciones = hii.sincronizar_idiomas(relaciones, salida)
    for origen, rel in archivos_relativos:
        destino = salida / rel
        clave = ruta_portatil(destino, base)
        if clave in cache and destino.resolve() in cambios_relaciones:
            cache[clave]["hash_destino"] = hashlib.sha256(destino.read_bytes()).hexdigest()
            if origen.resolve() == destino.resolve():
                cache[clave]["hash_origen"] = cache[clave]["hash_destino"]
    temporal_cache = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=salida,
                                         prefix=".cache_alineacion.", suffix=".tmp", delete=False) as archivo:
            temporal_cache = Path(archivo.name)
            json.dump(cache, archivo, ensure_ascii=False)
        os.replace(temporal_cache, ruta_cache)
    finally:
        if temporal_cache is not None:
            temporal_cache.unlink(missing_ok=True)
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
    parser.add_argument("--base", type=resolver_ruta, default=BASE)
    parser.add_argument("--forzar", action="store_true", help="Realinea incluso archivos sin cambios.")
    args = parser.parse_args()
    try:
        resumen = ejecutar(args.base, forzar=args.forzar)
        return 1 if resumen["errores"] else 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
