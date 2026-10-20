"""Lectura estructural de especificaciones Word, sin escrituras ni automatización."""
from __future__ import annotations

import re
import unicodedata
from xml.etree import ElementTree
from zipfile import ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def clave(texto):
    texto = "".join(c for c in unicodedata.normalize("NFD", texto.upper())
                    if unicodedata.category(c) != "Mn")
    return re.sub(r"[^A-Z0-9]+", " ", texto).strip()


def texto(elemento):
    partes = []

    def visitar(nodo):
        if nodo.tag == W + "t":
            partes.append(nodo.text or "")
        elif nodo.tag in {W + "br", W + "tab", W + "cr"}:
            partes.append(" ")
        elif nodo.tag not in {W + "tbl", W + "del", W + "moveFrom"}:
            for hijo in nodo:
                visitar(hijo)
            if nodo.tag == W + "p":
                partes.append(" ")

    visitar(elemento)
    return re.sub(r"\s+", " ", "".join(partes).replace("\xa0", " ")).strip()


def leer_partes(path):
    with ZipFile(path) as archivo:
        return {nombre: ElementTree.fromstring(archivo.read(nombre))
                for nombre in archivo.namelist()
                if re.fullmatch(r"word/(?:document|header\d*|footer\d*)\.xml", nombre)}


def filas(tabla):
    resultado = []
    anteriores = []
    for fila in tabla.findall(W + "tr"):
        antes = fila.find(f"{W}trPr/{W}gridBefore")
        valores = [""] * (int(antes.get(W + "val", "0")) if antes is not None else 0)
        for celda in fila.findall(W + "tc"):
            span = celda.find(f"{W}tcPr/{W}gridSpan")
            ancho = int(span.get(W + "val", "1")) if span is not None else 1
            valor = texto(celda)
            fusion = celda.find(f"{W}tcPr/{W}vMerge")
            if fusion is not None and fusion.get(W + "val") != "restart" and len(anteriores) > len(valores):
                valor = anteriores[len(valores)]
            valores.extend([valor] * ancho)
        resultado.append(valores)
        anteriores = valores
    return resultado


def campo_columna(valor, columnas):
    etiqueta = clave(valor)
    bilingues = {"PRUEBA TEST": "prueba", "UNIDADES UNITS": "unidades",
                 "METODO TEST METHOD": "metodo", "ESPECIFICACION SPECIFICATION": "especificacion"}
    return columnas.get(etiqueta) or bilingues.get(etiqueta)


def seleccionar_tabla(raiz, columnas):
    tablas = [(elemento, filas(elemento)) for elemento in raiz.iter(W + "tbl")]
    candidatas = []
    for indice, (elemento, renglones) in enumerate(tablas):
        for encabezado, valores in enumerate(renglones):
            mapa = [campo_columna(v, columnas) for v in valores]
            campos = set(mapa) - {None}
            if "prueba" in campos and {"especificacion", "valor_tipico"}.intersection(campos) and len(campos) >= 3:
                candidatas.append((len(campos), len(renglones) - encabezado, indice, encabezado, mapa))
    if not candidatas:
        raise ValueError("No se encontro la tabla de pruebas con nombre y especificacion reconocibles.")
    _, _, indice, encabezado, mapa = max(candidatas, key=lambda c: (c[0], c[1]))
    return tablas, indice, encabezado, mapa


def etiqueta_control(valor, etiquetas):
    # CUSTOMER’S y CUSTOMER'S son la misma etiqueta que CUSTOMERS.
    valor = re.sub(r"(?<=[A-Za-z])['’´](?=[A-Za-z])", "", valor)
    return etiquetas.get(clave(valor))


def es_registro_calidad(renglones):
    encabezados = {clave(v) for fila in renglones[:3] for v in fila}
    return ("REGISTRO DE CALIDAD" in encabezados or "QUALITY RECORD" in encabezados
            or {"CODIGO", "RETENER", "RESPONSABLE"}.issubset(encabezados))


def metadatos_formato(partes):
    datos = {}
    for nombre, raiz in partes.items():
        for tabla in raiz.iter(W + "tbl"):
            renglones = filas(tabla)
            if not es_registro_calidad(renglones) and "footer" not in nombre:
                continue
            for i, valores in enumerate(renglones[:-1]):
                etiquetas = [clave(v) for v in valores]
                if "CODIGO" in etiquetas and "REVISION" in etiquetas:
                    siguiente = renglones[i + 1]
                    for campo, etiqueta in (("codigo_formato", "CODIGO"), ("revision_formato", "REVISION")):
                        posicion = etiquetas.index(etiqueta)
                        if posicion < len(siguiente) and siguiente[posicion]:
                            datos[campo] = siguiente[posicion]
        if "footer" in nombre:
            contenido = " ".join(texto(p) for p in raiz.iter(W + "p"))
            codigo = re.search(r"\bC[OÓ]DIGO\s*[:.]?\s*(R\d{2}-[A-Z]+-[A-Z]+-\d{2}|\d{6,})\b", contenido, re.I)
            revision = re.search(r"\bREV(?:ISI[OÓ]N)?\s*[.:]?\s*(\d{1,2})\b", contenido, re.I)
            if codigo:
                datos.setdefault("codigo_formato", codigo.group(1))
            if revision:
                datos.setdefault("revision_formato", revision.group(1).zfill(2))
    if "codigo_formato" not in datos:
        contenido = " ".join(texto(p) for raiz in partes.values() for p in raiz.iter(W + "p"))
        codigo = re.search(r"\bR\d{2}-[A-Z]{2,}-[A-Z]{2,}-\d{2}\b", contenido, re.I)
        if codigo:
            datos["codigo_formato"] = codigo.group().upper()
            revision = re.search(r"\bREV(?:ISI[OÓ]N)?\s*[.:]?\s*(\d{1,2})\b", contenido[codigo.end():codigo.end() + 100], re.I)
            if revision:
                datos["revision_formato"] = revision.group(1).zfill(2)
    return datos


def nombre_producto(partes, limite, posiciones, columnas, catalogo):
    candidatos = []
    for nombre, raiz in partes.items():
        if "footer" in nombre:
            continue
        tras_titulo = False
        for parrafo in raiz.iter(W + "p"):
            if nombre == "word/document.xml" and posiciones[id(parrafo)] >= limite:
                break
            valor = texto(parrafo)
            normal = clave(valor)
            if normal in {"ESPECIFICACION", "SPECIFICATION"}:
                tras_titulo = True
                continue
            if not normal:
                continue
            if (normal in columnas or re.search(r"\b(?:CHEMICAL|LUBRICANTS|LUBRICACION|REGISTRO DE CALIDAD)\b", normal)
                    or normal.startswith(("CODIGO ", "REVISION ", "ACEPTACION "))):
                continue
            valor = re.sub(r"^(?:ESPECIFICACI[OÓ]N|SPECIFICATION)\s*[:\-]?\s+", "", valor, flags=re.I)
            candidatos.append((valor, tras_titulo, "header" in nombre))
            tras_titulo = False
    if not candidatos:
        raise ValueError("No se encontro el nombre del producto.")
    descripciones = {r["descripcion_clave"] for r in catalogo or []}
    encontrados = [(v, t, h) for v, t, h in candidatos if clave(v) in descripciones]
    if encontrados:
        preferidos = [v for v, t, _ in encontrados if t] or [v for v, _, _ in encontrados]
        if len({clave(v) for v in preferidos}) > 1:
            raise ValueError("El nombre del producto es ambiguo entre varios titulos del catalogo.")
        return preferidos[0]
    return next((v for v, t, _ in candidatos if t),
                next((v for v, _, h in candidatos if h), candidatos[0][0]))


def extraer(partes, columnas, etiquetas, catalogo=None):
    raiz = partes["word/document.xml"]
    tablas, indice, encabezado, mapa = seleccionar_tabla(raiz, columnas)
    tabla, renglones = tablas[indice]
    posiciones = {id(nodo): i for i, nodo in enumerate(raiz.iter())}
    filas_xml = tabla.findall(W + "tr")
    limite = posiciones[id(filas_xml[encabezado])]
    producto = nombre_producto(partes, limite, posiciones, columnas, catalogo)
    pruebas, notas, control_incrustado = [], [], {}
    for numero, valores in enumerate(renglones[encabezado + 1:], encabezado + 2):
        presentes = list(dict.fromkeys(v for v in valores if v))
        if not presentes:
            continue
        if len(presentes) == 1 and re.search(r"\b(?:NOTAS?|NOTES?)\b", clave(presentes[0])):
            notas.append(re.sub(r"^(?:Notas?|Notes?)\s*:\s*", "", presentes[0], flags=re.I))
            continue
        campos_fila = {campo_columna(v, columnas) for v in valores} - {None}
        if {"prueba", "especificacion"}.issubset(campos_fila):
            continue  # Encabezado repetido al cambiar de página.
        item = {"numero": len(pruebas) + 1, "fila_origen": numero}
        campo_control = etiqueta_control(valores[0], etiquetas) if valores else None
        if campo_control and len([v for v in valores if v]) == 2 and len(valores) > 1:
            if not re.fullmatch(r"[-–—.\s]*", valores[1]):
                control_incrustado[campo_control] = valores[1]
            continue
        if len(presentes) == 1 and (
            clave(presentes[0]) in {"CARACTERISTICAS DE ESPUMADO", "FOAMING CHARACTERISTICS"}
            or re.match(r"^(?:CARBON TYPE ANALYSIS|CLAY GEL ANALYSIS|ANALISIS DE TIPO DE CARBONO)\b", clave(presentes[0]))
        ):
            notas.append(presentes[0])
            continue
        campos = ("prueba", "unidades", "metodo", "especificacion", "valor_tipico")
        if "tipo_prueba" in mapa:
            campos += ("tipo_prueba",)
        for campo in campos:
            celdas = list(dict.fromkeys(v for j, v in enumerate(valores) if j < len(mapa) and mapa[j] == campo))
            if campo == "prueba" and len([v for v in celdas if v]) > 1:
                raise ValueError("La tabla contiene columnas PRUEBA duplicadas con contenidos distintos; requiere revision.")
            item[campo] = " ".join(celdas)
        if not item["prueba"] or not any(item[c] for c in ("metodo", "especificacion", "valor_tipico")):
            raise ValueError(f"Fila {numero} de la tabla de pruebas incompleta o sin nombre de prueba.")
        pruebas.append(item)
    if not pruebas:
        raise ValueError("La tabla de pruebas no contiene pruebas validas.")

    candidatas = []
    for i, (elemento, renglones_control) in enumerate(tablas):
        if i == indice or es_registro_calidad(renglones_control):
            continue
        control = {}
        for valores in renglones_control:
            unicos = []
            for valor in valores:
                if not unicos or valor != unicos[-1]:
                    unicos.append(valor)
            if len(unicos) >= 2:
                campo = etiqueta_control(unicos[0], etiquetas)
                if campo:
                    control[campo] = unicos[1]
        anclas = {"elaboro_y_aprobo", "codigo_unico_producto", "folio_apqp", "folio_vap", "nombre_cliente"}
        if len(control) >= 2 and anclas.intersection(control):
            candidatas.append((len(control), elemento, control))
    control = max(candidatas, key=lambda c: c[0])[2] if candidatas else {}
    for campo, valor in control_incrustado.items():
        if control.get(campo) and control[campo] != valor:
            raise ValueError(f"Valores contradictorios de control documental: {campo}.")
        control[campo] = valor

    dentro_tablas = {id(p) for elemento, _ in tablas for p in elemento.iter(W + "p")}
    fin_pruebas = max(posiciones[id(n)] for n in tabla.iter())
    for p in raiz.iter(W + "p"):
        if id(p) in dentro_tablas or posiciones[id(p)] <= fin_pruebas:
            continue
        valor = texto(p)
        if re.search(r"^(?:NOTAS?|NOTES?)\b|^[◊◇♦◈]|\b(?:CARACTERISTICA CRITICA|CRITICAL CHARACTERISTIC)\b", clave(valor) if not valor.startswith(("◊", "◇", "♦", "◈")) else valor, re.I):
            notas.append(re.sub(r"^(?:Notas?|Notes?)\s*:\s*", "", valor, flags=re.I))
    pie = " ".join(texto(p) for nombre, root in partes.items() if "footer" in nombre
                   for p in root.iter(W + "p") if texto(p))
    return {"producto": {"nombre": producto}, "control_documental": control,
            "pruebas": pruebas, "notas": list(dict.fromkeys(notas)),
            "datos_institucionales": {"pie_pagina": pie},
            "estructura_detectada": {"numero_tablas": len(tablas), "tabla_pruebas_indice": indice,
                                     "filas_pruebas": len(pruebas), "encabezados": renglones[encabezado]}}
