"""Resumen redactado de JSON originales o alineados, calculado solo en memoria."""
from collections import Counter, defaultdict
import hashlib
import json
import re
import unicodedata


# Edita aquí las frases del comparador. No se cargan plantillas desde archivos.
TEXTOS_COMPARACION = [
    {"id": "campo", "texto": "{campo}: cambió de «{antes}» a «{despues}»."},
    {"id": "campo_agregado", "texto": "{campo}: se agregó «{despues}»."},
    {"id": "campo_retirado", "texto": "{campo}: se retiró «{antes}»."},
    {"id": "prueba", "texto": "En {prueba}, {campo} cambió de «{antes}» a «{despues}»."},
    {"id": "prueba_agregado", "texto": "En {prueba}, se agregó {campo}: «{despues}»."},
    {"id": "prueba_retirado", "texto": "En {prueba}, se retiró {campo}: «{antes}»."},
    {"id": "agregada", "texto": "Se agregó la prueba {prueba}. {detalle}"},
    {"id": "retirada", "texto": "Se retiró la prueba {prueba}. {detalle}"},
    {"id": "nota_agregada", "texto": "Se agregó la nota: «{nota}»."},
    {"id": "nota_retirada", "texto": "Se retiró la nota: «{nota}»."},
    {"id": "ambiguas", "texto": "Algunas pruebas repetidas no tienen una correspondencia única; se muestran como retiradas y agregadas."},
    {"id": "sin_cambios", "texto": "No se detectaron diferencias en el contenido comparado."},
]
_TEXTOS = {item["id"]: item["texto"] for item in TEXTOS_COMPARACION}
CAMPOS_PRUEBA = (
    ("prueba", "el nombre de la prueba"), ("unidades", "el campo de unidades"),
    ("metodo", "el método"), ("especificacion", "la especificación"),
    ("valor_tipico", "el valor típico"), ("periodicidad", "la periodicidad de análisis"),
)
ETIQUETAS_CONTROL = {
    "elaboro_y_aprobo": "Elaboró y aprobó", "folio_apqp": "Folio APQP",
    "folio_vap": "Folio VAP", "codigo_unico_producto": "Código / Part ID",
    "fecha": "Fecha", "revision": "Revisión", "nombre_cliente": "Nombre del cliente",
    "revision_cliente": "Revisión del cliente",
}


def _texto(valor):
    return "" if valor is None else str(valor).strip()


def _clave(valor):
    valor = unicodedata.normalize("NFKD", _texto(valor).upper())
    return " ".join("".join(c for c in valor if not unicodedata.combining(c)).split())


def _contenido_alineado(datos):
    from generador_especs_hii import _separar_notas_de_pruebas

    pruebas = []
    originales = datos.get("pruebas", [])
    if not isinstance(originales, list):
        originales = []
    detalle_pruebas = datos.get("alineacion_metodos_lab", {}).get("detalle")
    if detalle_pruebas is None:
        detalle_pruebas = [{"valores_alineados": fila} for fila in originales if isinstance(fila, dict)]
    for indice, detalle in enumerate(detalle_pruebas):
        fila = dict(detalle["valores_alineados"])
        if "prueba" not in fila:
            fila["prueba"] = fila.get("nombre", "")
        original = originales[indice] if indice < len(originales) and isinstance(originales[indice], dict) else {}
        fila["periodicidad"] = next(
            (_texto(objeto.get(campo)) for objeto in (original, fila)
             for campo in ("PERIODIDICAD DE ANALISIS", "periodicidad", "periodicidad_analisis")
             if _texto(objeto.get(campo))), "",
        )
        pruebas.append(fila)
    pruebas, notas = _separar_notas_de_pruebas(pruebas)
    notas = list(dict.fromkeys(_texto(n) for n in [*datos.get("notas", []), *notas] if _texto(n)))
    return pruebas, notas


def _emparejar(antes, despues, maestro):
    """Empareja sin depender de la posición y sin adivinar entre duplicados."""
    from alinear_especificaciones_json import metodo_referencias_canonico

    nombres = defaultdict(set)
    for indice, registro in enumerate(maestro):
        for campo in ("NOMBRE IQM", "NOMBRE IQM INGLES"):
            if registro.get(campo):
                nombres[_clave(registro[campo])].add(indice)

    def identidad(fila):
        nombre = _clave(fila.get("prueba"))
        candidatos = nombres.get(nombre, set())
        return ("maestro", next(iter(candidatos))) if len(candidatos) == 1 else nombre

    def metodo(fila):
        return _clave(metodo_referencias_canonico(_texto(fila.get("metodo"))))

    def temperatura(fila):
        return tuple(re.sub(r"\s+", "", t) for t in re.findall(
            r"[+-]?\d+(?:[.,]\d+)?\s*°?\s*[CF]\b", _clave(fila.get("prueba"))))

    pendientes_a, pendientes_b = set(range(len(antes))), set(range(len(despues)))
    parejas = []
    # Las filas idénticas se consumen una a una, también cuando están repetidas.
    for a in sorted(pendientes_a):
        firma = tuple(_texto(antes[a].get(c)) for c, _ in CAMPOS_PRUEBA)
        b = next((b for b in sorted(pendientes_b)
                  if firma == tuple(_texto(despues[b].get(c)) for c, _ in CAMPOS_PRUEBA)), None)
        if b is not None:
            parejas.append((a, b))
            pendientes_a.remove(a)
            pendientes_b.remove(b)

    def unir(clave, validar=lambda a, b: True):
        grupos_a, grupos_b = defaultdict(list), defaultdict(list)
        for indices, filas, grupos in ((pendientes_a, antes, grupos_a), (pendientes_b, despues, grupos_b)):
            for indice in sorted(indices):
                valor = clave(filas[indice])
                if valor:
                    grupos[valor].append(indice)
        for valor, grupo in grupos_a.items():
            otro = grupos_b.get(valor, [])
            if len(grupo) == len(otro) == 1 and validar(antes[grupo[0]], despues[otro[0]]):
                a, b = grupo[0], otro[0]
                parejas.append((a, b))
                pendientes_a.remove(a)
                pendientes_b.remove(b)

    unir(lambda f: (identidad(f), metodo(f), _clave(f.get("unidades"))) if identidad(f) else None)
    unir(lambda f: (identidad(f), metodo(f)) if identidad(f) else None)
    unir(identidad)

    metodos_a = Counter(metodo(f) for f in antes)
    metodos_b = Counter(metodo(f) for f in despues)

    def admite_metodo(a, b):
        ia, ib = identidad(a), identidad(b)
        return (metodos_a[metodo(a)] == metodos_b[metodo(b)] == 1
                and temperatura(a) == temperatura(b)
                and not (isinstance(ia, tuple) and isinstance(ib, tuple) and ia != ib))

    unir(lambda f: metodo(f) if metodo(f) not in {"", "NA", "N/A", "-"} else None, admite_metodo)
    ambiguas = any(identidad(antes[a]) == identidad(despues[b])
                   for a in pendientes_a for b in pendientes_b)
    return sorted(parejas), sorted(pendientes_a), sorted(pendientes_b), ambiguas


def resumir_diferencias_alineadas(antes, despues, maestro=()):
    """Devuelve una lista de frases: Especificación de origen → destino.

    Compara el texto guardado, incluidos cambios de formato; omite la auditoría
    del alineador. Los números se muestran sin interpretar mejoras o conversiones.
    """
    cambios = []

    def cambio(campo, a, b, prueba=None):
        a, b = _texto(a), _texto(b)
        if a == b:
            return
        tipo = "prueba" if prueba else "campo"
        if not a:
            tipo += "_agregado"
        elif not b:
            tipo += "_retirado"
        cambios.append(_TEXTOS[tipo].format(campo=campo, antes=a, despues=b, prueba=prueba))

    cambio("Nombre del producto", antes.get("producto", {}).get("nombre"), despues.get("producto", {}).get("nombre"))
    cambio("Idioma", antes.get("documento", {}).get("idioma"), despues.get("documento", {}).get("idioma"))
    control_a, control_b = antes.get("control_documental", {}), despues.get("control_documental", {})
    for campo in dict.fromkeys([*control_a, *control_b]):
        cambio(ETIQUETAS_CONTROL.get(campo, campo.replace("_", " ").capitalize()), control_a.get(campo), control_b.get(campo))
    for campo in ("FECHA DE REANALISIS", "Desarrollador"):
        cambio("Fecha de reanálisis" if campo == "FECHA DE REANALISIS" else campo, antes.get(campo), despues.get(campo))
    cambio("Pie de página", antes.get("datos_institucionales", {}).get("pie_pagina"), despues.get("datos_institucionales", {}).get("pie_pagina"))

    filas_a, notas_a = _contenido_alineado(antes)
    filas_b, notas_b = _contenido_alineado(despues)
    parejas, retiradas, agregadas, ambiguas = _emparejar(filas_a, filas_b, maestro)

    def etiqueta(fila, indice):
        return f"«{_texto(fila.get('prueba')) or 'Sin nombre'}» (fila {indice + 1})"

    for a, b in parejas:
        for campo, nombre in CAMPOS_PRUEBA:
            cambio(nombre, filas_a[a].get(campo), filas_b[b].get(campo), etiqueta(filas_a[a], a))
    if ambiguas:
        cambios.append(_TEXTOS["ambiguas"])
    for tipo, indices, filas in (("retirada", retiradas, filas_a), ("agregada", agregadas, filas_b)):
        for indice in indices:
            fila = filas[indice]
            detalle = " · ".join(f"{nombre.capitalize()}: «{_texto(fila.get(campo))}»"
                                 for campo, nombre in CAMPOS_PRUEBA[1:] if _texto(fila.get(campo)))
            cambios.append(_TEXTOS[tipo].format(prueba=etiqueta(fila, indice), detalle=detalle).rstrip())
    for tipo, notas, otras in (("nota_retirada", notas_a, notas_b), ("nota_agregada", notas_b, notas_a)):
        cambios.extend(_TEXTOS[tipo].format(nota=nota) for nota in notas if nota not in otras)
    return cambios or [_TEXTOS["sin_cambios"]]


def formatear_resumen_diferencias(antes, despues, origen, destino, frases):
    """Incluye las huellas del contenido completo de los dos JSON comparados."""
    def huella(datos):
        canonico = json.dumps(datos, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return hashlib.sha256(canonico.encode("utf-8")).hexdigest()

    return (
        "CAMBIOS ENTRE ESPECIFICACIONES\n"
        f"Hash JSON origen (SHA-256, contenido): {huella(antes)}\n"
        f"Hash JSON destino (SHA-256, contenido): {huella(despues)}\n\n"
        f"Origen: {origen}\nDestino: {destino}\n\n"
        + "\n".join(f"• {frase}" for frase in frases)
    )
