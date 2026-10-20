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
from decimal import Decimal
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

VERSION = "2.1.35"
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
    rf"(?P<fraccion_final>\d{{1,3}}(?:[{ESPACIOS_NUMERICOS}]\d{{1,3}})*)(?!\d)"
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
    r",\d{3}(?![\w.,])"
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


_NUMERO_RANGO_PRECISION = rf"[+\-−]?\s*\d+(?:[{ESPACIOS_NUMERICOS}]\d{{3}})*(?:[.,]\d+(?:[{ESPACIOS_NUMERICOS}]\d{{1,3}})*)?"
_RANGO_PRECISION = re.compile(
    rf"(?<![\w.,+\-−])(?P<izquierdo>{_NUMERO_RANGO_PRECISION})"
    rf"\s*[-–—]\s*(?P<derecho>{_NUMERO_RANGO_PRECISION})(?![\w.,]|[-–—]\d)"
)


def _igualar_precision_rango(coincidencia):
    extremos = [coincidencia.group(campo).strip() for campo in ("izquierdo", "derecho")]
    precision = max(len(re.sub(rf"[{ESPACIOS_NUMERICOS}]", "", v.split(",", 1)[1]))
                    if "," in v else 0 for v in extremos)
    if not precision:
        return " - ".join(extremos)
    salida = []
    for valor in extremos:
        entero, _, fraccion = valor.partition(",")
        fraccion = re.sub(rf"[{ESPACIOS_NUMERICOS}]", "", fraccion)
        salida.append(entero + "," + fraccion.ljust(precision, "0"))
    return " - ".join(salida)


def _formatear_precision_decimal(salida):
    salida = _RANGO_PRECISION.sub(_igualar_precision_rango, salida)
    def agrupar(coincidencia):
        fraccion = re.sub(rf"[{ESPACIOS_NUMERICOS}]", "", coincidencia.group("fraccion"))
        return "," + ESPACIO_FINO.join(fraccion[i:i + 3] for i in range(0, len(fraccion), 3))
    return re.sub(rf",(?P<fraccion>\d+(?:[{ESPACIOS_NUMERICOS}]\d{{1,3}})*)(?!\w)", agrupar, salida)


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
            + re.sub(rf"[{ESPACIOS_NUMERICOS}]", "", coincidencia.group("fraccion_final"))
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

    return _formatear_precision_decimal(ENTERO_CON_MILES.sub(normalizar_entero, salida))


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


def _formatear_numero_medicion(numero: str, *, coma_de_miles: bool = False) -> tuple[str, bool]:
    compacto = re.sub(r"[ \u00a0\u2009\u202f]", "", numero)
    if coma_de_miles and "," in compacto:
        grupos = re.fullmatch(r"(?P<entero>\d{1,3}(?:,\d{3})+)(?:\.(?P<fraccion>\d+))?", compacto)
        if grupos is None:
            return numero, True
        entero = grupos.group("entero").replace(",", "").lstrip("0") or "0"
        fraccion = grupos.group("fraccion")
        return _agrupar_entero(entero) + (f",{fraccion}" if fraccion is not None else ""), False
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


def normalizar_valor_medicion(valor: Any, *, coma_de_miles: bool = False) -> tuple[str, list[str]]:
    """Usa coma decimal y conserva comas de tres cifras ambiguas para revisión."""
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

    def reservar_ambiguo(fragmento: str) -> str:
        indice = len(ambiguos)
        ambiguos.append(fragmento)
        return f"QZAMBIGUO{indice}QZ"

    def guardar_ambiguo(coincidencia: re.Match[str]) -> str:
        # Un punto único es decimal. Con coma, una parte entera cero tampoco
        # puede representar una agrupación de miles significativa.
        parte_entera = re.sub(r"\D", "", coincidencia.group("valor"))
        if not parte_entera.strip("0"):
            return coincidencia.group(0)
        return reservar_ambiguo(coincidencia.group(0))

    def normalizar_rango_explicito(coincidencia):
        extremos = [coincidencia.group(c).strip() for c in ("izquierdo", "derecho")]
        # Una precisión distinta de tres cifras, un punto decimal o una parte
        # entera cero aportan contexto decimal. Dos comas ambiguas siguen a revisión.
        def decimal_claro(valor):
            compacto = re.sub(rf"[{ESPACIOS_NUMERICOS}]", "", valor).lstrip("+−-")
            if "." in compacto:
                return True
            entero, separador, fraccion = compacto.partition(",")
            return bool(separador) and (len(fraccion) != 3 or not entero.strip("0"))
        if not any(decimal_claro(v) for v in extremos):
            return coincidencia.group(0)
        valores = []
        for extremo in extremos:
            compacto = extremo.strip()
            signo = compacto[0] if compacto[0] in "+−-" else ""
            numero, ambiguo = _formatear_numero_medicion(compacto[len(signo):])
            if ambiguo:
                return coincidencia.group(0)
            valores.append(("− " if signo in {"-", "−"} else "+ " if signo else "") + numero)
        indice = len(protegidos)
        protegidos.append(_formatear_precision_decimal(" - ".join(valores)))
        return f"QZPROTEGIDO{indice}QZ"

    if not coma_de_miles:
        salida = _RANGO_PRECISION.sub(normalizar_rango_explicito, salida)

    if not coma_de_miles:
        salida = NUMERO_AMBIGUO_MEDICION.sub(guardar_ambiguo, salida)

    def normalizar_cientifico(coincidencia: re.Match[str]) -> str:
        entrada = coincidencia.group("coef")
        if coma_de_miles and "," in entrada:
            signo = entrada[0] if entrada[0] in "+-" else ""
            numero, ambiguo = _formatear_numero_medicion(entrada[len(signo):], coma_de_miles=True)
            if ambiguo:
                return reservar_ambiguo(coincidencia.group(0))
            entrada = signo + numero.replace(" ", "")
        coeficiente = _formatear_precision_decimal(_formatear_coeficiente_cientifico(entrada))
        exponente = _exponente_superindice(coincidencia.group("exp"))
        # La coma de salida ya es decimal: no debe volver a leerse como miles.
        indice = len(protegidos)
        protegidos.append(f"{coeficiente} × 10{exponente}")
        return f"QZPROTEGIDO{indice}QZ"

    salida = NOTACION_CIENTIFICA_E.sub(normalizar_cientifico, salida)
    salida = NOTACION_CIENTIFICA_X.sub(normalizar_cientifico, salida)

    def normalizar_numero(coincidencia: re.Match[str]) -> str:
        numero, ambiguo = _formatear_numero_medicion(
            coincidencia.group("valor"), coma_de_miles=coma_de_miles,
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
    salida = _formatear_precision_decimal(salida)
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


def metodo_alineado_canonico(original: Any, maestro: Any) -> str:
    """Normaliza el método sin añadir farmacopeas ausentes en el original."""
    resultado = metodo_maestro_canonico(maestro)
    if len(referencias_astm(original)) > len(referencias_astm(maestro)):
        resultado = metodo_referencias_canonico(original)
    if referencias_astm(original) and re.search(r"\bNMX\b", texto(original), re.IGNORECASE):
        # El apartado NMX pertenece a la cita original y no a otra fila del maestro.
        apartados = re.findall(r"\bNMX\b[^/]+", texto(original), re.IGNORECASE)
        resultado = " / ".join([*(f"ASTM {r}" for r in referencias_astm(original)),
                                *referencias_farmacopea(original),
                                *(re.sub(r"\s+", " ", p).strip() for p in apartados)])
    originales = referencias_farmacopea(original)
    maestras = referencias_farmacopea(maestro)
    if originales and maestras:
        # Mantiene el orden del maestro, pero solo las farmacopeas de la fuente.
        partes = [parte for parte in resultado.split(" / ")
                  if parte not in maestras or parte in originales]
        partes.extend(ref for ref in originales if ref not in partes)
        resultado = " / ".join(partes)
    return resultado


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
    referencias_normalizadas = set(referencias_normativas(valor))
    if referencias_normalizadas:
        return referencias_normalizadas
    n = normalizar(valor)
    if not n:
        return set()
    if n == 'ORGANOLEPTIC':
        n = 'ORGANOLEPTICO'
    interno = re.fullmatch(r'(?:METODO INTERNO|INTERNAL METHOD)(?:\s+(\d+))?', n)
    if interno:
        return {'INTERNO:' + (interno[1] or '')}
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


def referencias_normativas(valor: Any) -> frozenset[str]:
    """Identidad de las normas, sin confundir número de norma y edición."""
    n = normalizar(valor).replace("–", "-").replace("—", "-")
    referencias = {*(f"ASTM {r}" for r in referencias_astm(valor)), *referencias_farmacopea(valor)}
    for numero, rama in re.findall(r"\bNOM\s*[- ]?\s*(\d{2,3})\s*[- ]?\s*([A-Z]+\s*\d*)", n):
        rama = re.sub(r"\s+", "", rama)
        referencias.add(f"NOM-{numero.zfill(3)}-{rama}")
    for letra, numero, organismo in re.findall(
            r"\bNMX\s*[- ]\s*([A-Z]+)\s*[- ]\s*(\d+)(?:\s*[- ]\s*([A-Z]+))?", n):
        referencias.add(f"NMX-{letra}-{numero}" + (f"-{organismo}" if organismo else ""))
    for familia, numero in re.findall(r"\b(ISO|IEC|DIN|EN|IP|SAE|JIS|API)\s*[- ]?\s*([A-Z]*\s*\d+(?:-\d{1,2})?)", n):
        numero = re.sub(r"\s+", "", numero)
        referencias.add(f"{familia} {numero}")
    return frozenset(referencias)


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
    # Recupera un encabezado sobrescrito en el diseño fijo de siete columnas.
    # Solo la columna D puede faltar y las otras seis deben conservar su posición.
    esquema = ('METODO', 'MÉTODO INGLÉS', 'NOMBRE IQM', 'NOMBRE IQM INGLES', 'UNIDADES', 'UNITS', 'ACREDITACION NOMBRE')
    if (faltantes == ['NOMBRE IQM INGLES'] and len(tabla.columns) == 7
            and all(normalizar(tabla.columns[i]) == normalizar(c) for i, c in enumerate(esquema) if i != 3)):
        encabezados[normalizar('NOMBRE IQM INGLES')] = 3
        faltantes = []
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
        nombre_ingles = limpiar_marca_acreditacion(registro.get("NOMBRE IQM INGLES", ""))
        if normalizar(nombre_ingles) == 'VISUAL APPEAREANCE' and normalizar(registro.get('NOMBRE IQM', '')) == 'APARIENCIA VISUAL':
            nombre_ingles = 'VISUAL APPEARANCE'
        return (
            nombre_ingles,
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
    original = normalizar(normalizar_espacio_unidades(normalizar_unidad_declarada(unidad_original, 'ESPANOL')))
    if not original:
        return False
    return any(
        original == normalizar(normalizar_espacio_unidades(normalizar_unidad_declarada(alternativa, 'ESPANOL')))
        for alternativa in alternativas_unidades(unidad_maestra)
    )


def normalizar_unidad_declarada(valor, idioma):
    """Alias inequívocos de unidad; también se aplican a métodos ausentes."""
    clave = re.sub(r'\s+', '', texto(valor)).upper()
    if clave in {'NA', 'N/A'}:
        return 'NA'
    if clave in {'AD', 'A/D', 'ADIMENSIONAL', 'DIMENSIONLESS'}:
        return 'DIMENSIONLESS' if normalizar(idioma) in ALIAS_IDIOMA_INGLES else 'ADIMENSIONAL'
    return texto(valor)


def resolver_unidad_declarada(original, propuesta, idioma, usar_maestro=False):
    # La unidad de una coincidencia aceptada tiene prioridad sobre los alias fuente.
    if usar_maestro:
        return normalizar_unidad_declarada(propuesta, idioma)
    if normalizar_unidad_declarada(original, 'ESPANOL') in {'NA', 'ADIMENSIONAL'}:
        return normalizar_unidad_declarada(original, idioma)
    return normalizar_unidad_declarada(propuesta, idioma)


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


def _grupo_prueba_normas(registro):
    """Relaciona métodos alternativos solo para una misma prueba y unidades."""
    return (
        normalizar(limpiar_marca_acreditacion(registro.get("NOMBRE IQM", ""))),
        normalizar(limpiar_marca_acreditacion(registro.get("NOMBRE IQM INGLES", ""))),
        tuple(sorted(normalizar(normalizar_unidad_declarada(u, 'ESPANOL')) for u in alternativas_unidades(registro.get("UNIDADES", "")))),
        tuple(sorted(normalizar(normalizar_unidad_declarada(u, 'ESPANOL')) for u in alternativas_unidades(registro.get("UNITS", "") or registro.get("UNIDADES", "")))),
    )


TEMPERATURA_PRUEBA = re.compile(
    r"(?<![\w.,])(?P<valor>[+\-\u2212\u2013\u2014]?\s*\d+(?:\s*[.,]\s*\d+)?|X)"
    r"\s*[°º]?\s*(?P<escala>[CF])\b", re.IGNORECASE,
)


def _texto_temperaturas(nombre):
    numero = r'[+\-\u2212\u2013\u2014]?\s*\d+(?:\s*[.,]\s*\d+)?|X'
    return re.sub(r'(?<![\w.,])(' + numero + r')\s*/\s*(' + numero + r')\s*[°º]?\s*([CF])\b',
                  lambda m: m[1] + ' °' + m[3] + '/' + m[2] + ' °' + m[3], texto(nombre), flags=re.IGNORECASE)


def temperaturas_prueba(nombre):
    """Conserva signo, escala y temperaturas de referencia, sin conversiones."""
    return [(None if m['valor'].upper() == 'X' else
             Decimal(re.sub(r"\s+", "", m['valor']).translate(str.maketrans('−–—,', '---.'))),
             m['escala'].upper()) for m in TEMPERATURA_PRUEBA.finditer(_texto_temperaturas(nombre))]


def magnitud_viscosidad(valor):
    n = unicodedata.normalize('NFKC', texto(valor)).upper()
    cin = bool(re.search(r"\bCST\b|MM\s*\^?2\s*/\s*(?:S\b|SEG\b|SEC\b)|CINEM[AÁ]T|KINEMATIC", n))
    din = bool(re.search(r"\bCP\b|\bCPS\b|MPA\s*[.*·]?\s*S\b|DIN[AÁ]MIC|DYNAMIC|BROOKFIELD|APARENTE|APPARENT", n))
    return 'conflicto' if cin and din else 'cinematica' if cin else 'dinamica' if din else ''


def _tipo_carbono(nombre):
    """Reconoce familias y simbolos de carbono sin enumerar nombres completos."""
    n = normalizar(nombre)
    # Los simbolos deben ser tokens completos: CA no debe coincidir en CADENAS.
    tipos = set(re.findall(r'\b(CA|CN|CP)\b', n))
    patrones = ((r'\bAROMAT', 'CA'),
                (r'\b(?:NAFT\s*EN|NAPHTH?EN)', 'CN'),
                (r'\b(?:PARAFIN|PARAFFIN)', 'CP'))
    tipos.update(tipo for patron, tipo in patrones if re.search(patron, n))
    return next(iter(tipos)) if len(tipos) == 1 else ''


def _saybolt_fahrenheit(nombre, metodo):
    if referencias_normativas(metodo) != {'ASTM D2161'} or 'SAYBOLT' not in normalizar(nombre):
        return None
    temperaturas = temperaturas_prueba(nombre)
    if not temperaturas:
        return None
    for celsius, fahrenheit in ((Decimal('37.8'), 100), (Decimal('98.9'), 210)):
        nominales = {(celsius, 'C'), (Decimal(fahrenheit), 'F')}
        if set(temperaturas) <= nominales:
            return fahrenheit
    return None


def equivalencias_confirmadas(nombre, metodo):
    """Alias aprobados por el laboratorio; no cambian la identidad de otras normas."""
    n, m = normalizar(nombre), normalizar(metodo)
    refs = referencias_normativas(metodo)
    nombre_busqueda, metodo_busqueda, reglas = nombre, metodo, []
    if re.fullmatch(r'FIXED OILS FATS (?:AND )?(?:ROSIN|RESIN|RESINS)(?: NO OILY OR SOLID MATTER (?:SAPARATES|SEPARATES))?(?: B)?', n):
        nombre_busqueda = 'ACEITES FIJOS, GRASAS Y RESINA'
        reglas.append('FIXED OILS, FATS AND ROSIN/RESIN = ACEITES FIJOS, GRASAS Y RESINA')
    if n in {'COLOR VISUAL', 'VISUAL COLOR', 'DETERMINACION DE COLOR'} and m == 'VISUAL':
        nombre_busqueda = 'COLOR'
        reglas.append('COLOR VISUAL / VISUAL COLOR / DETERMINACION DE COLOR = COLOR, VISUAL')
    if (n == 'CONTENIDO DE AGUA' and 'ASTM D1533' in refs
            and all(ref == 'ASTM D1533' or ref.startswith('NMX ') for ref in refs)):
        nombre_busqueda = 'CONTENIDO DE AGUA POR KARL FISCHER (LIQUIDOS AISLANTES)'
        reglas.append('CONTENIDO DE AGUA, ASTM D1533 = KARL FISCHER (LIQUIDOS AISLANTES)')
    if n == 'CONTENIDO DE AGUA PROCEDIMIENTO A' and m in {'', 'NA', 'N A'}:
        nombre_busqueda = 'CONTENIDO DE AGUA POR KARL FISCHER (COULOM\u00c9TRICO)'
        metodo_busqueda = 'ASTM D6304'
        reglas.append('CONTENIDO DE AGUA, PROCEDIMIENTO A sin metodo = KARL FISCHER COULOM\u00c9TRICO, ASTM D6304')
    procedimiento = re.search(r'\b(?:PROCEDIMIENTO|PROCEDURE)\s+([ABC])\b', m)
    if (refs == {'ASTM D6304'} and procedimiento and procedimiento[1] == 'A'
            and re.search(r'\b(?:AGUA|HUMEDAD|WATER)\b', n)
            and not re.search(r'\b(?:PROCEDIMIENTO|PROCEDURE)\s+[BC]\b', n)):
        nombre_busqueda = 'CONTENIDO DE AGUA POR KARL FISCHER (COULOM\u00c9TRICO)'
        reglas.append('ASTM D6304 PROCEDIMIENTO A: AGUA / HUMEDAD / WATER = KARL FISCHER COULOM\u00c9TRICO')
    acidez = normalizar(limpiar_marca_acreditacion(nombre)).rstrip(' %').strip()
    if (re.fullmatch(r'(?:ACIDEZ|ACIDITY)(?: AND ALKALINITY| NO RED OR PINK COLOR IS PRODUCED)?', acidez)
            and (refs and refs <= {'USP', 'FEUM', 'Ph Eur'} or m in {'', 'NA', 'N A'})):
        nombre_busqueda, metodo_busqueda = 'ACIDEZ', 'USP / FEUM'
        reglas.append('ACIDEZ CONFIRMADA = ACIDEZ, USP / FEUM')
    if (refs == {'ISO 4406'} and re.search(
            r'\b(?:CODIGO DE LIMPIEZA|NIVEL DE LIMPIEZA|PUREZA DE PARTICULAS?|'
            r'AUTOMATIC CLEANING CODE|CODING THE LEVEL OF CONTAMINATION|CLEANLINESS CODE)\b', n)):
        nombre_busqueda = 'C\u00d3DIGO DE LIMPIEZA'
        reglas.append('ISO 4406: variantes de limpieza y contaminacion = C\u00d3DIGO DE LIMPIEZA')
    if (re.fullmatch(r'ALKALINITY THE SOLUTION DOES NOT A(?:C)?QUIRE A PINK COLOR(?: B)?', n)
            and refs and refs <= {'USP', 'FEUM'}):
        nombre_busqueda = 'ALCALINIDAD'
        reglas.append('ALKALINITY con nota de color rosa = ALCALINIDAD')
    azufre = normalizar(limpiar_marca_acreditacion(nombre))
    if (azufre in {'AZUFRE', 'AZUFRE Y SULFUROS', 'LIMITE DE COMPUESTOS DE AZUFRE'}
            and refs and refs <= {'USP', 'FEUM'}):
        nombre_busqueda, metodo_busqueda = 'L\u00cdMITE DE COMPUESTOS DE AZUFRE', 'USP'
        reglas.append('AZUFRE CONFIRMADO: USP/FEUM = LIMITE DE COMPUESTOS DE AZUFRE, USP')
    nombre_inflamacion = re.sub(r',?\s*[\u00b0\u00ba]\s*C\s*$', '', limpiar_marca_acreditacion(nombre), flags=re.IGNORECASE)
    if normalizar(nombre_inflamacion) in {'PUNTO DE INFLAMACION', 'FLASH POINT'} and refs == {'ASTM D92'}:
        nombre_busqueda = 'PUNTO DE INFLAMACI\u00d3N A 101,3 kPa, COPA ABIERTA CLEVELAND'
        reglas.append('PUNTO DE INFLAMACION, ASTM D92 = CLEVELAND A 101,3 kPa')
    if (re.fullmatch(r'PUNTO\s*DE ANILINA(?: METODO E)?', n)
            and refs == {'ASTM D611'}
            and re.search(r'\b(?:METODO|METHOD) E\b', m)
            and not re.search(r'\b(?:METODO|METHOD) [A-DF-Z]\b', n + ' ' + m)):
        nombre_busqueda = 'PUNTO DE ANILINA, M\u00c9TODO E'
        reglas.append('PUNTODE ANILINA, ASTM D611 METODO E = PUNTO DE ANILINA, METODO E')
    if normalizar(limpiar_marca_acreditacion(nombre)) == 'PUNTO DE ESCURRIMIENTO' and refs == {'ASTM D97'}:
        nombre_busqueda = 'PUNTO DE FLUIDEZ'
        reglas.append('PUNTO DE ESCURRIMIENTO, ASTM D97 = PUNTO DE FLUIDEZ')
    viscosidad_confirmada = re.fullmatch(
        r'VISCOSIDAD\s+A\s+(40|100)\s*[\u00b0\u00ba]?\s*C', texto(nombre), re.IGNORECASE)
    if viscosidad_confirmada and refs == {'ASTM D445'}:
        temperatura = viscosidad_confirmada[1]
        nombre_busqueda = f'VISCOSIDAD CINEM\u00c1TICA A {temperatura} \u00b0C'
        reglas.append(f'VISCOSIDAD A {temperatura} C, ASTM D445 = VISCOSIDAD CINEMATICA A {temperatura} C')
    separacion_agua_aceite = re.fullmatch(
        r'SEPARACI[O\u00d3]N\s+AGUA[-\s]+ACEITE\s+A\s+(54|82)\s*[\u00b0\u00ba]?\s*C', texto(nombre), re.IGNORECASE)
    if separacion_agua_aceite and refs == {'ASTM D1401'}:
        temperatura = separacion_agua_aceite[1]
        nombre_busqueda = f'SEPARABILIDAD AGUA-ACEITE A {temperatura} \u00b0C'
        reglas.append(f'SEPARACION AGUA ACEITE A {temperatura} C, ASTM D1401 = SEPARABILIDAD AGUA-ACEITE A {temperatura} C')
    if n == 'DEMULSIBILIDAD' and refs == {'ASTM D1401'}:
        nombre_busqueda = 'SEPARABILIDAD AGUA-ACEITE A 54 \u00b0C'
        reglas.append('DEMULSIBILIDAD, ASTM D1401 = SEPARABILIDAD AGUA-ACEITE A 54 C')
    separabilidad_confirmada = re.fullmatch(
        r'SEPARABIL(?:I)?DAD\s+DEL\s+AGUA\s+Y\s+DEL\s+ACEITE\s+Y\s+FLUIDOS\s+SINT[E\u00c9]TICOS\s+A\s+54\s*[\u00b0\u00ba]?\s*C',
        texto(nombre), re.IGNORECASE)
    if separabilidad_confirmada and (refs == {'ASTM D1401'} or m in {'', 'NA', 'N A'}):
        nombre_busqueda, metodo_busqueda = 'SEPARABILIDAD AGUA-ACEITE A 54 \u00b0C', 'ASTM D1401'
        reglas.append('SEPARABILDAD DEL AGUA Y DEL ACEITE Y FLUIDOS SINTETICOS A 54 C = SEPARABILIDAD AGUA-ACEITE A 54 C, ASTM D1401')
    densidad_confirmada = re.fullmatch(
        r'DENSIDAD\s+RELATIVA\s+A\s+15[.,]6\s*/\s*15[.,]6\s*[\u00b0\u00ba]\s*C', texto(nombre), re.IGNORECASE)
    if densidad_confirmada and refs == {'ASTM D1298', 'ASTM D7042'}:
        nombre_busqueda = 'DENSIDAD RELATIVA 15,6 \u00b0C/15,6 \u00b0C'
        metodo_busqueda = 'ASTM D1298'
        reglas.append('DENSIDAD RELATIVA 15,6/15,6 C: conjunto D1298/D7042 confirmado para fila D1298')
    espuma_secuencia = re.fullmatch(
        r'(?:TENDENCIA A LA ESPUMA )?SECUENCIA (III|II|I)',
        normalizar(limpiar_marca_acreditacion(nombre)))
    if espuma_secuencia and (refs == {'ASTM D892'} or m in {'', 'NA', 'N A', '---'}):
        nombre_busqueda = f'FORMACI\u00d3N DE ESPUMA (SECUENCIA {espuma_secuencia[1]})'
        metodo_busqueda = 'ASTM D892'
        reglas.append(f'SECUENCIA {espuma_secuencia[1]} confirmada = FORMACION DE ESPUMA, ASTM D892')
    if (refs == {'ASTM D972'} and re.search(
            r'\b(?:VOLATILIDAD|VOLATILITY|PERDIDA(?:D)? POR EVAPORACION)\b',
            normalizar(limpiar_marca_acreditacion(nombre)))):
        nombre_busqueda = 'P\u00c9RDIDA POR EVAPORACI\u00d3N DE GRASAS LUBRICANTES Y ACEITES'
        reglas.append('ASTM D972: volatilidad / perdida por evaporacion = PERDIDA POR EVAPORACION DE GRASAS LUBRICANTES Y ACEITES')
    if n == 'IV' and refs == {'ASTM D2270'}:
        nombre_busqueda = 'ÍNDICE DE VISCOSIDAD'
        reglas.append('IV = ÍNDICE DE VISCOSIDAD (ASTM D2270)')
    if n == 'COMPUESTOS CARBONIZABLES' and referencias_farmacopea(metodo):
        nombre_busqueda = 'SUSTANCIAS FÁCILMENTE CARBONIZABLES'
        reglas.append('COMPUESTOS CARBONIZABLES = SUSTANCIAS FÁCILMENTE CARBONIZABLES')
    polinucleares = normalizar(limpiar_marca_acreditacion(nombre))
    if re.fullmatch(r'(?:LIMITE DE )?COMPUESTOS POLINUCLEARES', polinucleares) and referencias_farmacopea(metodo):
        nombre_busqueda = 'LÍMITE DE HIDROCARBUROS POLICÍCLICOS AROMÁTICOS'
        reglas.append('COMPUESTOS POLINUCLEARES = HIDROCARBUROS POLICÍCLICOS AROMÁTICOS')
        if refs and refs <= {'USP', 'FEUM'}:
            metodo_busqueda = 'USP'
            reglas.append('POLINUCLEARES CONFIRMADOS: USP/FEUM = USP del maestro')
    if n in {'APARIENCIA', 'ASPECTO', 'ASPCTO', 'APARIENCIA VISUAL', 'APPEARANCE', 'VISUAL APPEARANCE', 'VISUAL APPEAREANCE'} and m in {'', 'NA', 'N A', 'VISUAL', 'ORGANOLEPTICO', 'ORGANOLEPTIC'}:
        nombre_busqueda, metodo_busqueda = 'APARIENCIA VISUAL', 'ORGANOLÉPTICO'
        reglas.append('APARIENCIA / ASPECTO / ASPCTO, VISUAL = APARIENCIA VISUAL, ORGANOLÉPTICO')
    if re.search(r'\bLOVIBOND\b', n) and (m in {'', 'NA', 'N A', 'LOVIBOND'} or refs == {'USP'}):
        nombre_busqueda, metodo_busqueda = 'COLOR LOVIBOND RYBN', 'IP17 A'
        reglas.append('COLOR LOVIBOND, LOVIBOND = COLOR LOVIBOND RYBN, IP17 A')
    nombre_humedad = limpiar_marca_acreditacion(nombre)
    humedad_confirmada = normalizar(nombre_humedad) in {
        'HUMEDAD', 'HUMIDITY', 'HUMIDITY SPLATTER TEST',
        'HUMIDITY CREPITATION', 'CREPITATION HUMIDITY', 'HUMEDAD POR CREPITACION'}
    if humedad_confirmada and not refs and m in {
            '', 'NA', 'N A', 'METODO DEL CLIENTE', 'CLIENT METHOD',
            'INS CCA 07', 'INS CCA 07 00', '91300092'}:
        nombre_busqueda, metodo_busqueda = 'HUMEDAD POR CREPITACI\u00d3N', 'M\u00c9TODO INTERNO (91300029)'
        reglas.append('HUMEDAD / HUMIDITY = HUMEDAD POR CREPITACI\u00d3N, 91300029')
    if re.search(r'\b(?:HUMEDAD|AGUA|MOISTURE)\b', n) and m in {'INS CCA 07 00', '91300092'}:
        nombre_busqueda, metodo_busqueda = 'HUMEDAD POR CREPITACIÓN', 'MÉTODO INTERNO (91300029)'
        reglas.append('INS-CCA-07-00 / 91300092 = HUMEDAD POR CREPITACIÓN, 91300029')
    if not refs and re.search(r'(?<!\d)91300092(?!\d)', texto(metodo)):
        nombre_busqueda, metodo_busqueda = 'HUMEDAD POR CREPITACIÓN', 'MÉTODO INTERNO (91300029)'
        reglas.append('IDENTIFICADOR 91300092 = HUMEDAD POR CREPITACIÓN, sin depender del nombre')
    if not refs and (m in {'INS CCA 07', 'INS CCA 07 00'} or re.search(r'(?<!\d)91300029(?!\d)', texto(metodo))):
        nombre_busqueda, metodo_busqueda = 'HUMEDAD POR CREPITACIÓN', 'MÉTODO INTERNO (91300029)'
        reglas.append('IDENTIFICADOR CREPITACIÓN = INS-CCA-07 / INS-CCA-07-00 / 91300029')
    saybolt = _saybolt_fahrenheit(nombre, metodo)
    if saybolt is not None:
        nombre_busqueda = f'VISCOSIDAD SAYBOLT A {saybolt} °F'
        reglas.append(f'SAYBOLT ASTM D2161: equivalencia nominal de temperatura a {saybolt} °F')
    if refs == {'ASTM E1252'}:
        nombre_busqueda = 'ESPECTRO INFRARROJO'
        reglas.append('ASTM E1252 = ESPECTRO INFRARROJO, sin depender del nombre')
    if refs == {'ASTM D2272'}:
        variante = 'A' if re.search(r'\b(?:METODO|METHOD|PROCEDIMIENTO|PROCEDURE) A\b', normalizar(texto(nombre) + ' ' + texto(metodo))) else 'B'
        nombre_busqueda = 'ESTABILIDAD A LA OXIDACIÓN, MÉTODO ' + variante
        reglas.append('ASTM D2272 = ESTABILIDAD A LA OXIDACIÓN, MÉTODO ' + variante)
    tipo = _tipo_carbono(nombre)
    if tipo and refs == {'ASTM D2140'}:
        nombre_busqueda = {'CA': 'PORCENTAJE DE CARBONOS EN ANILLOS AROMÁTICOS (%CA)',
                           'CN': 'PORCENTAJE DE CARBONOS EN ANILLOS NAFTÉNICOS (%CN)',
                           'CP': 'PORCENTAJE DE CARBONOS EN CADENAS PARAFÍNICAS (%CP)'}[tipo]
        reglas.append(f'ASTM D2140: componente {tipo} vinculado al maestro')
    secuencia = re.fullmatch(r'(?:SEC|SECUENCIA|SEQUENCE)\s*(III|II|I|1|2|3)(?:\s+(?:TENDENCIA\s+ESTABILIDAD|TENDENCY\s+STABILITY))?', n)
    sin_metodo = not m or m in {'NA', 'N A'}
    if secuencia and (refs == {'ASTM D892'} or (sin_metodo and ('TENDENCIA ESTABILIDAD' in n or 'TENDENCY STABILITY' in n))):
        sec = {'1': 'I', '2': 'II', '3': 'III'}.get(secuencia[1], secuencia[1])
        nombre_busqueda = f'FORMACIÓN DE ESPUMA (SECUENCIA {sec})'
        if sin_metodo:
            metodo_busqueda = 'ASTM D892'
        reglas.append(f'SEC {sec}, TENDENCIA/ESTABILIDAD = FORMACIÓN DE ESPUMA, ASTM D892')
    if n in {'TENDENCIA A LA ESPUMA', 'ESPUMA', 'ESTABILIDAD A LA ESPUMACION', 'PRUEBA A LA ESPUMA'} and (sin_metodo or refs == {'ASTM D892'}):
        nombre_busqueda, metodo_busqueda = 'FORMACIÓN DE ESPUMA', 'ASTM D892'
        reglas.append('ESPUMA: ASTM D892 reconocido; secuencia no especificada')
    return nombre_busqueda, metodo_busqueda, reglas


def _apartados_nmx(metodo):
    return {normalizar(p).replace(' ', '') for p in
            re.findall(r'\bNMX\b[^/]+', texto(metodo), re.IGNORECASE)}


def _candidato_con_condiciones(nombre, unidades, registro, identidad=None):
    """Filtra magnitudes y subtipos; no adapta ni bloquea temperaturas."""
    identidad = nombre if identidad is None else identidad
    magnitudes = {magnitud_viscosidad(v) for v in (identidad, unidades)} - {''}
    if 'conflicto' in magnitudes or len(magnitudes) > 1:
        return None
    magnitud = next(iter(magnitudes), '')
    relativa = r'PESO ESPECIFICO|DENSIDAD (?:RELATIVA|ESPECIFICA)|GRAVEDAD ESPECIFICA|SPECIFIC GRAVITY|RELATIVE DENSITY'
    nombres_destino = ' '.join(texto(registro.get(k)) for k in ('NOMBRE IQM', 'NOMBRE IQM INGLES'))
    tipo_carbono = _tipo_carbono(identidad)
    if tipo_carbono and _tipo_carbono(nombres_destino) != tipo_carbono:
        return None
    if re.search(relativa, normalizar(identidad)) and not re.search(relativa, normalizar(nombres_destino)):
        return None
    patron_secuencia = r'\b(?:SEC|SECUENCIA|SEQUENCE)\s*(III|II|I|[123])\b'
    sec_fuente = re.search(patron_secuencia, normalizar(identidad))
    sec_destino = re.search(patron_secuencia, normalizar(nombres_destino))
    if sec_fuente and sec_destino:
        equivalentes = {'1': 'I', '2': 'II', '3': 'III'}
        if equivalentes.get(sec_fuente[1], sec_fuente[1]) != equivalentes.get(sec_destino[1], sec_destino[1]):
            return None
    destino = magnitud_viscosidad(' '.join(texto(registro.get(k)) for k in
                                          ('NOMBRE IQM', 'NOMBRE IQM INGLES', 'UNIDADES', 'UNITS')))
    if magnitud and destino and magnitud != destino:
        return None
    return dict(registro)


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
        self.normas_candidatas = [{i: referencias_normativas(metodo_maestro_para_idioma(r, i))
                                  for i in ("ESPANOL", "INGLES", "DESCONOCIDO")} for r in self]
        self.grupos_normas = [_grupo_prueba_normas(r) for r in self]
        self.normas_por_grupo = {}
        for grupo, normas in zip(self.grupos_normas, self.normas_candidatas):
            for lengua, referencias in normas.items():
                self.normas_por_grupo.setdefault((grupo, lengua), set()).update(referencias)
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


def es_separador_pruebas(nombre, metodo=''):
    n = normalizar(limpiar_marca_acreditacion(nombre))
    return n == 'PRUEBA CADA LOTE' or (n == 'TENDENCIA A LA ESPUMA' and texto(metodo) == '---')


def mejor_coincidencia(
    nombre: str,
    metodo: str,
    maestro: list[dict[str, str]],
    idioma: str = "DESCONOCIDO",
    unidades: str = "",
) -> dict[str, Any]:
    if not maestro:
        raise ValueError("La lista maestra no contiene registros.")
    if es_separador_pruebas(nombre, metodo):
        return {'estado': 'REVISION_REQUERIDA', 'puntaje': 0.0, 'puntaje_nombre': 0.0,
                'puntaje_metodo': 0.0, 'margen': 0.0, 'ambiguo': False,
                'registro': {}, 'segunda_opcion': {}, 'fila_separadora': True,
                'tipo_pendiente': 'SEPARADOR_NO_ES_PRUEBA',
                'mensaje_separador': 'Separador: ' + texto(nombre) + ' no es una prueba.'}
    nombre_busqueda, metodo_busqueda, reglas = equivalencias_confirmadas(nombre, metodo)
    if not texto(nombre_busqueda):
        return {
            "estado": "SIN_COINCIDENCIA", "puntaje": 0.0, "puntaje_nombre": 0.0,
            "puntaje_metodo": 0.0, "margen": 0.0, "ambiguo": False,
            "registro": {}, "segunda_opcion": {},
        }

    clave_cache = (nombre, metodo, idioma, unidades)
    preparado = isinstance(maestro, MaestroPreparado)
    if preparado and clave_cache in maestro.coincidencias:
        return deepcopy(maestro.coincidencias[clave_cache])
    mejor = segundo = None
    grupos_contextuales = {}
    originales = referencias_normativas(metodo_busqueda)
    metodo_vigencia = (metodo_referencias_canonico(metodo)
                       if originales and originales == set(referencias_farmacopea(metodo))
                       and re.search(r"\b(?:VIGENTES?|CURRENT|CURRENCY)\b", normalizar(metodo)) else '')
    compatibles = set()
    compuestos_verificados = set()
    parciales = False
    apartados = _apartados_nmx(metodo_busqueda)
    normas_por_grupo = {}
    grupos = maestro.grupos_normas if preparado else [_grupo_prueba_normas(r) for r in maestro]
    if originales:
        astm_compuesto = (len(originales) > 1 and any(r.startswith("ASTM ") for r in originales)
                          and all(r.startswith(("ASTM ", "NMX-")) for r in originales))
        normas_por_grupo = maestro.normas_por_grupo if preparado else {}
        if astm_compuesto and not preparado:
            for grupo, registro in zip(grupos, maestro):
                normas_por_grupo.setdefault((grupo, idioma), set()).update(
                    referencias_normativas(metodo_maestro_para_idioma(registro, idioma)))
        apartados_por_grupo = {}
        if apartados:
            for grupo, registro in zip(grupos, maestro):
                apartados_por_grupo.setdefault(grupo, set()).update(
                    _apartados_nmx(metodo_maestro_para_idioma(registro, idioma)))
        for indice, registro in enumerate(maestro):
            normas = (maestro.normas_candidatas[indice].get(idioma) if preparado else None)
            if normas is None:
                normas = referencias_normativas(metodo_maestro_para_idioma(registro, idioma))
            if apartados and not apartados <= apartados_por_grupo.get(grupos[indice], set()):
                parciales = parciales or bool(originales & normas)
                continue
            if originales <= normas:
                compatibles.add(indice)
            elif astm_compuesto and originales & normas and originales <= normas_por_grupo.get((grupos[indice], idioma), set()):
                compatibles.add(indice)
                compuestos_verificados.add(indice)
            elif originales & normas:
                parciales = True
        if not compatibles:
            globales = {ref for r in maestro for m in metodos_maestros(r) for ref in referencias_normativas(m)}
            ausentes = sorted(originales - globales)
            resultado = {
                "estado": "REVISION_REQUERIDA" if parciales else "SIN_COINCIDENCIA",
                "puntaje": 0.0, "puntaje_nombre": 0.0, "puntaje_metodo": 0.0,
                "margen": 0.0, "ambiguo": False, "registro": {}, "segunda_opcion": {},
                "metodo_no_encontrado": True,
                "mensaje_metodo": "Método no encontrado en lista maestra: " + texto(metodo),
                "referencias_originales": sorted(originales),
                "referencias_no_encontradas": ausentes,
                "tipo_pendiente": "REFERENCIA_AUSENTE" if ausentes else "CONJUNTO_NO_APROBADO",
            }
            if preparado:
                maestro.coincidencias[clave_cache] = deepcopy(resultado)
            return resultado
    elif texto(metodo_busqueda) and normalizar(metodo_busqueda) not in {'', 'NA', 'N A'}:
        compatibles = {i for i, r in enumerate(maestro)
                       if partes_metodo(metodo_busqueda) == partes_metodo(metodo_maestro_para_idioma(r, idioma))}
        if not compatibles:
            resultado = {'estado': 'SIN_COINCIDENCIA', 'puntaje': 0.0, 'puntaje_nombre': 0.0,
                         'puntaje_metodo': 0.0, 'margen': 0.0, 'ambiguo': False, 'registro': {},
                         'segunda_opcion': {}, 'metodo_no_encontrado': True,
                         'tipo_pendiente': 'METODO_SIN_EQUIVALENCIA',
                         'mensaje_metodo': 'Método no encontrado en lista maestra: ' + texto(metodo)}
            if preparado:
                maestro.coincidencias[clave_cache] = deepcopy(resultado)
            return resultado
    for indice, registro in enumerate(maestro):
        if originales and indice not in compatibles:
            continue
        if not originales and compatibles and indice not in compatibles:
            continue
        if originales == {'ASTM D2272'}:
            variante = re.search(r'\bMETODO ([AB])\b', normalizar(nombre_busqueda))
            if variante and not re.search(r'\b(?:METODO|METHOD) ' + variante[1] + r'\b',
                                          normalizar(registro.get('NOMBRE IQM', '') + ' ' + registro.get('NOMBRE IQM INGLES', ''))):
                continue
        identidad_metodo = any(regla.startswith(('ASTM E1252 =', 'ASTM D2272 =', 'IDENTIFICADOR 91300092 =', 'IDENTIFICADOR CREPITACIÓN =')) for regla in reglas)
        nombre_condiciones = nombre_busqueda if any(regla.startswith('SAYBOLT ASTM D2161:') for regla in reglas) else nombre
        contextual = _candidato_con_condiciones(nombre_condiciones, unidades, registro, nombre_busqueda if identidad_metodo else None)
        if contextual is None:
            continue
        grupos_contextuales[id(contextual)] = grupos[indice]
        total, sn, sm = (maestro.puntuar(nombre_busqueda, metodo_busqueda,
                                        (contextual, limpiar_marca_acreditacion(contextual['NOMBRE IQM']),
                                         limpiar_marca_acreditacion(contextual['NOMBRE IQM INGLES']),
                                         maestro.candidatos[indice][3]), idioma)
                         if preparado else puntuar(nombre_busqueda, metodo_busqueda, contextual, idioma))
        candidato = (total, sn, sm, contextual)
        def prioridad(c):
            return (c[0], unidad_original_esta_entre_alternativas(unidades, nombres_canonicos(c[3], idioma)[1]), c[1], c[2])
        if mejor is None or prioridad(candidato) > prioridad(mejor):
            segundo, mejor = mejor, candidato
        elif segundo is None or prioridad(candidato) > prioridad(segundo):
            segundo = candidato
    if mejor is None:
        resultado = {'estado': 'REVISION_REQUERIDA', 'puntaje': 0.0, 'puntaje_nombre': 0.0,
                     'puntaje_metodo': 0.0, 'margen': 0.0, 'ambiguo': False, 'registro': {},
                     'segunda_opcion': {}, 'condiciones_incompatibles': True,
                     'tipo_pendiente': 'CONDICIONES_INCOMPATIBLES',
                     'mensaje_condiciones': 'No hay candidato compatible con las unidades o el subtipo de prueba originales.'}
        if metodo_vigencia:
            resultado['metodo_normalizado'] = metodo_vigencia
        if preparado:
            maestro.coincidencias[clave_cache] = deepcopy(resultado)
        return resultado
    segundo = segundo if segundo is not None else (0.0, 0.0, 0.0, {})
    margen = mejor[0] - segundo[0]
    ambiguo = (
        margen < MARGEN_AMBIGUEDAD
        and registros_ambiguos(mejor[3], segundo[3], idioma)
    )
    if ambiguo and compuestos_verificados:
        # Dos ASTM alternativos de la misma prueba no son dos pruebas ambiguas.
        # Las copias pueden tener X vinculadas; las filas maestras mantienen el grupo.
        grupo_mejor = grupos_contextuales[id(mejor[3])]
        grupo_segundo = grupos_contextuales[id(segundo[3])]
        if grupo_mejor == grupo_segundo and originales <= normas_por_grupo.get((grupo_mejor, idioma), set()):
            ambiguo = False
    if (ambiguo and normalizar(nombres_canonicos(mejor[3], idioma)[0]) == normalizar(nombres_canonicos(segundo[3], idioma)[0])
            and unidad_original_esta_entre_alternativas(unidades, nombres_canonicos(mejor[3], idioma)[1])
            and not unidad_original_esta_entre_alternativas(unidades, nombres_canonicos(segundo[3], idioma)[1])):
        ambiguo = False
    if mejor[0] >= UMBRAL_AUTOMATICO and not ambiguo:
        estado = "ALINEADO"
    elif mejor[0] >= UMBRAL_DUDOSO:
        estado = "REVISION_REQUERIDA"
    else:
        estado = "SIN_COINCIDENCIA"
    resultado_match = {"estado": estado, "puntaje": round(mejor[0], 4), "puntaje_nombre": round(mejor[1], 4),
            "puntaje_metodo": round(mejor[2], 4), "margen": round(margen, 4), "ambiguo": ambiguo,
            "registro": mejor[3], "segunda_opcion": segundo[3], 'reglas_equivalencia': reglas,
            'nombre_busqueda': nombre_busqueda, 'metodo_busqueda': metodo_busqueda}
    if estado != 'ALINEADO':
        resultado_match['tipo_pendiente'] = 'COINCIDENCIA_AMBIGUA' if ambiguo else 'NOMBRE_SIN_EQUIVALENCIA'
        if (reglas and mejor[1] >= 0.96 and ambiguo
                and normalizar(nombres_canonicos(mejor[3], idioma)[0]) == normalizar(nombres_canonicos(segundo[3], idioma)[0])
                and metodo_maestro_canonico(metodo_maestro_para_idioma(mejor[3], idioma)) == metodo_maestro_canonico(metodo_maestro_para_idioma(segundo[3], idioma))):
            # El alias de nombre está confirmado aunque A/NA no se puedan resolver.
            resultado_match['nombre_normalizado'] = nombres_canonicos(mejor[3], idioma)[0]
            resultado_match['metodo_normalizado'] = metodo_alineado_canonico(metodo, metodo_maestro_para_idioma(mejor[3], idioma))
            resultado_match['tipo_pendiente'] = 'UNIDADES_POR_REVISAR'
    if metodo_vigencia:
        # La farmacopea existe entre los candidatos compatibles, aunque el nombre
        # de la prueba no alcance el umbral. Limpia solo la cita de la fuente.
        resultado_match["metodo_normalizado"] = metodo_vigencia
    if 'ESPUMA: ASTM D892 reconocido; secuencia no especificada' in reglas:
        resultado_match.update(estado='REVISION_REQUERIDA', tipo_pendiente='SECUENCIA_POR_DEFINIR',
                               nombre_normalizado='FOAMING CHARACTERISTICS' if idioma == 'INGLES' else 'FORMACIÓN DE ESPUMA',
                               metodo_normalizado='ASTM D892')
    if originales == {'ASTM D2140'} and not _tipo_carbono(nombre) and _tipo_carbono(mejor[3].get('NOMBRE IQM', '')):
        resultado_match.update(estado='REVISION_REQUERIDA', tipo_pendiente='COMPONENTE_CARBONO_POR_DEFINIR',
                               metodo_normalizado='ASTM D2140')
    if resultado_match['estado'] == 'ALINEADO' and 'ACIDEZ CONFIRMADA = ACIDEZ, USP / FEUM' in reglas:
        resultado_match['metodo_confirmado'] = 'USP / FEUM'
    if resultado_match['estado'] == 'ALINEADO' and 'POLINUCLEARES CONFIRMADOS: USP/FEUM = USP del maestro' in reglas:
        resultado_match['metodo_confirmado'] = 'USP'
    if resultado_match['estado'] == 'ALINEADO' and 'AZUFRE CONFIRMADO: USP/FEUM = LIMITE DE COMPUESTOS DE AZUFRE, USP' in reglas:
        resultado_match['metodo_confirmado'] = 'USP'
    if preparado:
        maestro.coincidencias[clave_cache] = deepcopy(resultado_match)
    return resultado_match


def es_prueba_brookfield(nombre: str, metodo: str) -> bool:
    """Reconoce el nombre aproximado o D2983, sin corregir otros números ASTM."""
    for valor in (nombre, metodo):
        if "D2983" in referencias_astm(valor):
            return True
        compacto = re.sub(r"[^A-Z0-9]", "", normalizar(valor))
        if re.search(r"(?:ASTM|ATSM|ASMT)D?2983(?![A-Z0-9])", compacto):
            return True
    palabras = normalizar(nombre).split()
    for inicio in range(len(palabras)):
        for longitud in (1, 2):
            candidato = "".join(palabras[inicio:inicio + longitud])
            if 8 <= len(candidato) <= 12 and fuzz.ratio(candidato, "BROOKFIELD") >= 85:
                return True
    return False


def _entrada_medicion_original(prueba: dict, detalle_anterior: Any, campo: str) -> Any:
    """Reutiliza el origen de una cifra ya normalizada si la fila sigue intacta."""
    actual = prueba.get(campo)
    if not isinstance(detalle_anterior, dict):
        return actual
    valores = detalle_anterior.get("valores_alineados")
    if not isinstance(valores, dict) or valores.get(campo) != actual:
        return actual
    for clave, original in (("prueba", "nombre_original"), ("metodo", "metodo_original")):
        vigente = texto(prueba.get(clave) or (prueba.get("nombre") if clave == "prueba" else ""))
        permitidos = [detalle_anterior.get(original), valores.get(clave)]
        if clave == "prueba":
            permitidos.extend((detalle_anterior.get("nombre_maestro_propuesto"),
                               detalle_anterior.get("nombre_maestro_ingles_propuesto")))
        if normalizar(vigente) not in {normalizar(limpiar_marca_acreditacion(valor)) for valor in permitidos}:
            return actual
    clave = {"prueba": "nombre_original", "unidades": "unidades_originales"}.get(campo, f"{campo}_original")
    return detalle_anterior.get(clave, actual)


def separar_encabezados_pruebas(pruebas: list[Any]) -> tuple[list[Any], list[dict]]:
    """Excluye encabezados reconocidos y conserva las filas originales para auditoría."""
    from generador_especs_hii import _es_fila_encabezado_pruebas

    conservadas, omitidas = [], []
    for indice, prueba in enumerate(pruebas, 1):
        es_encabezado = isinstance(prueba, dict) and _es_fila_encabezado_pruebas(
            {**prueba, "prueba": prueba.get("prueba") or prueba.get("nombre", "")}
        )
        if es_encabezado:
            omitidas.append({
                "indice_original": indice,
                "motivo": "Encabezado de tabla, no es una prueba.",
                "fila_original": deepcopy(prueba),
            })
        else:
            conservadas.append(prueba)
    return conservadas, omitidas


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
    alineacion_anterior = datos.get("alineacion_metodos_lab")
    detalle_anterior = alineacion_anterior.get("detalle", []) if isinstance(alineacion_anterior, dict) else []
    if not isinstance(detalle_anterior, list):
        detalle_anterior = []
    # Asocia la auditoría antes de omitir encabezados o filas de notas.
    auditoria_por_fila = {id(prueba): detalle for prueba, detalle in zip(pruebas, detalle_anterior)}
    pruebas, encabezados_nuevos = separar_encabezados_pruebas(pruebas)
    encabezados_omitidos = (deepcopy(alineacion_anterior.get("encabezados_omitidos", []))
                           if isinstance(alineacion_anterior, dict) else [])
    if not isinstance(encabezados_omitidos, list):
        encabezados_omitidos = []
    for encabezado in encabezados_nuevos:
        if encabezado not in encabezados_omitidos:
            encabezados_omitidos.append(encabezado)
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
        fila_original = deepcopy(prueba)
        if encabezados_nuevos and "numero" in prueba:
            prueba["numero"] = indice
        prueba["PERIODIDICAD DE ANALISIS"] = next(
            (prueba[campo] for campo in ("PERIODIDICAD DE ANALISIS", "periodicidad", "periodicidad_analisis")
             if texto(prueba.get(campo))), "▶")
        anterior = auditoria_por_fila.get(id(prueba))
        nombre_original = texto(_entrada_medicion_original(prueba, anterior, "prueba") or prueba.get("nombre"))
        metodo_original = texto(_entrada_medicion_original(prueba, anterior, "metodo"))
        unidades_originales = texto(_entrada_medicion_original(prueba, anterior, "unidades"))
        coma_de_miles = es_prueba_brookfield(nombre_original, metodo_original)
        anterior = auditoria_por_fila.get(id(prueba))
        especificacion_original = _entrada_medicion_original(prueba, anterior, "especificacion")
        valor_tipico_original = texto(_entrada_medicion_original(prueba, anterior, "valor_tipico"))
        limpieza_especificacion, ambiguos_especificacion = normalizar_valor_medicion(
            normalizar_terminos_especificacion(
                normalizar_especificacion(especificacion_original)
            ), coma_de_miles=coma_de_miles,
        )
        limpieza_valor_tipico, ambiguos_valor_tipico = normalizar_valor_medicion(
            normalizar_terminos_especificacion(valor_tipico_original), coma_de_miles=coma_de_miles,
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
        match = mejor_coincidencia(nombre_original, metodo_original, maestro, idioma, unidades_originales)
        registro = match["registro"]
        nombre_maestro, unidad_maestra = nombres_canonicos(registro, idioma)
        metodo_maestro_fuente = metodo_maestro_para_idioma(registro, idioma)
        metodo_maestro = match.get('metodo_confirmado') or metodo_alineado_canonico(metodo_original, metodo_maestro_fuente)
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
            "prueba": nombre_maestro or nombre_original,
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
        if match["estado"] != "ALINEADO":
            metodo_limpio = match.get("metodo_normalizado", metodo_original)
            nombre_limpio = match.get("nombre_normalizado", nombre_original)
            prueba.update(prueba=nombre_limpio, metodo=metodo_limpio, unidades=unidades_originales)
            if 'nombre' in prueba:
                prueba['nombre'] = nombre_limpio
            valores_alineados.update(prueba=nombre_limpio, metodo=metodo_limpio, unidades=unidades_originales)
        if match.get("metodo_no_encontrado"):
            # No se exportan propuestas ajenas al método. Conserva la fila fuente.
            fuente = {"prueba": nombre_original, "metodo": metodo_original, "unidades": unidades_originales,
                      "especificacion": texto(especificacion_original), "valor_tipico": valor_tipico_original}
            prueba.update(fuente)
            if "nombre" in fila_original:
                prueba["nombre"] = fila_original["nombre"]
            valores_alineados = dict(fuente)
        unidad_salida = resolver_unidad_declarada(unidades_originales, prueba.get('unidades', ''), idioma, usar_maestro=match['estado'] == 'ALINEADO')
        prueba['unidades'] = unidad_salida
        valores_alineados['unidades'] = unidad_salida
        if (normalizar_unidad_declarada(unidades_originales, 'ESPANOL') in {'NA', 'ADIMENSIONAL'}
                and unidad_salida != unidades_originales):
            match.setdefault('reglas_equivalencia', []).append(f'UNIDAD DECLARADA: {unidades_originales} = {unidad_salida}')
        auditoria.append({
            "fila_separadora": bool(match.get("fila_separadora")),
            "mensaje_separador": match.get("mensaje_separador", ""),
            "metodo_no_encontrado": bool(match.get("metodo_no_encontrado")),
            "mensaje_metodo": match.get("mensaje_metodo", ""),
            "numero_prueba": indice, "estado": match["estado"], "nombre_original": nombre_original,
            "metodo_original": metodo_original, "nombre_maestro_propuesto": registro.get("NOMBRE IQM", ""),
            "reglas_equivalencia": match.get("reglas_equivalencia", []),
            "condiciones_incompatibles": bool(match.get("condiciones_incompatibles")),
            "mensaje_condiciones": match.get("mensaje_condiciones", ""),
            "tipo_pendiente": match.get("tipo_pendiente", ""),
            "referencias_no_encontradas": match.get("referencias_no_encontradas", []),
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
            "especificacion_original": texto(especificacion_original),
            "regla_formato_numerico": "COMA_DE_MILES_BROOKFIELD_D2983" if coma_de_miles else "GENERAL",
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
        "encabezados_omitidos": encabezados_omitidos,
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
