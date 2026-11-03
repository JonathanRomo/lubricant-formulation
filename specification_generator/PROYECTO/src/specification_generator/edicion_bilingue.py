"""Borradores de un producto y traducciones controladas, sin servicios externos."""
from __future__ import annotations

import copy
import re
import uuid

# Ampliar aquí las equivalencias aprobadas por el laboratorio.
PATRONES_VALORES = (
    ("MÁXIMO", "MAXIMUM"), ("MAXIMO", "MAXIMUM"),
    ("MÍNIMO", "MINIMUM"), ("MINIMO", "MINIMUM"),
    ("TÍPICO", "TYPICAL"), ("TIPICO", "TYPICAL"),
    ("GRIS", "GRAY"), ("PASA", "PASS"),
)
CAMPOS = ("prueba", "unidades", "metodo", "especificacion", "valor_tipico")


def traducir_patrones(texto, origen, destino):
    """Devuelve texto y certeza; no altera cifras, signos ni códigos técnicos."""
    if not texto.strip():
        return texto, True
    salida = texto
    resto = texto
    usados = set()
    for es, en in PATRONES_VALORES:
        fuente, objetivo = (es, en) if origen == "es" else (en, es)
        if fuente in usados:
            continue
        usados.add(fuente)
        patron = re.compile(r"(?<!\w)" + re.escape(fuente) + r"(?!\w)", re.I)
        salida = patron.sub(objetivo, salida)
        resto = patron.sub("", resto)
    resto = re.sub(r"\b(?:ASTM\s+[A-Z]\s*\d+(?:-\d+)?|INS-[A-Z0-9-]+|N/A|NA)\b", "", resto, flags=re.I)
    # Solo datos numéricos y signos: no aceptar palabras desconocidas como traducidas.
    seguro = re.fullmatch(r"[\d\s.,+−\-–—/<>=≤≥±%():;°]*", resto) is not None
    return salida, seguro


class BorradoresBilingues:
    def __init__(self, contenido, traducir):
        self.traducir = traducir
        self.borradores = {}
        self.pendientes = {"es": set(), "en": set()}
        self.revisiones = {"es": set(), "en": set()}
        self.automaticas = {"es": set(), "en": set()}
        self.manuales = {"es": set(), "en": set()}
        self.actualizar(contenido["idioma"], contenido)

    def actualizar(self, idioma, contenido):
        """Propaga cambios; conserva traducciones manuales y pide revisarlas."""
        nuevo = copy.deepcopy(contenido)
        for fila in nuevo["pruebas"]:
            fila.setdefault("_id_bilingue", uuid.uuid4().hex)
        anterior = self.borradores.get(idioma)
        otro = "en" if idioma == "es" else "es"
        contraparte = self.borradores.get(otro)
        if contraparte is None:
            contraparte = copy.deepcopy(nuevo)
            contraparte["idioma"] = otro
            for fila in contraparte["pruebas"]:
                fuente = copy.deepcopy(fila)
                for campo in CAMPOS:
                    clave = (fila["_id_bilingue"], campo)
                    fila[campo], seguro = self.traducir(campo, fuente.get(campo, ""), idioma, otro, fuente)
                    self.automaticas[otro].add(clave)
                    if not seguro:
                        self.pendientes[otro].add(clave)
            self.borradores[otro] = contraparte
            for campo in ("notas", "elaboro_y_aprobo"):
                valor = self._texto(nuevo, campo)
                traducido, seguro = self.traducir(campo, valor, idioma, otro, {})
                self._poner(contraparte, campo, traducido)
                self.automaticas[otro].add(("cabecera", campo))
                if not seguro:
                    self.pendientes[otro].add(("cabecera", campo))
        else:
            previas = {f["_id_bilingue"]: f for f in anterior["pruebas"]} if anterior else {}
            destinos = {f["_id_bilingue"]: f for f in contraparte["pruebas"]}
            filas = []
            for fila in nuevo["pruebas"]:
                identidad = fila["_id_bilingue"]
                previa = previas.get(identidad, {})
                destino = destinos.get(identidad, {"_id_bilingue": identidad})
                for campo in CAMPOS:
                    clave = (identidad, campo)
                    valor = fila.get(campo, "")
                    if campo in destino and valor == previa.get(campo, ""):
                        continue
                    if not valor.strip():
                        # Una celda vacía no borra su texto en el otro idioma.
                        # El asterisco de la interfaz nunca entra en estos datos.
                        if destino.get(campo, "").strip():
                            self.pendientes[idioma].add(clave)
                            self.automaticas[idioma].add(clave)
                        else:
                            destino[campo] = ""
                            for lengua in (idioma, otro):
                                self.pendientes[lengua].discard(clave)
                                self.revisiones[lengua].discard(clave)
                        continue
                    completando = clave in self.pendientes[idioma]
                    revisando = clave in self.revisiones[idioma]
                    if valor.strip():
                        self.manuales[idioma].add(clave)
                        self.pendientes[idioma].discard(clave)
                        self.revisiones[idioma].discard(clave)
                        self.automaticas[idioma].discard(clave)
                    traducido, seguro = self.traducir(campo, valor, idioma, otro, fila)
                    compartido = seguro and not re.search(r"[A-Za-zÁÉÍÓÚáéíóú]", valor)
                    if completando and not revisando and not compartido and destino.get(campo, "").strip() and clave not in self.automaticas[otro]:
                        continue
                    if compartido or not valor.strip() or campo not in destino or clave in self.automaticas[otro] or not destino[campo].strip():
                        destino[campo] = traducido
                        self.automaticas[otro].add(clave)
                        if seguro:
                            self.pendientes[otro].discard(clave)
                        else:
                            self.pendientes[otro].add(clave)
                    elif seguro and destino[campo] == traducido:
                        self.pendientes[otro].discard(clave)
                    else:
                        self.pendientes[otro].add(clave)
                        self.revisiones[otro].add(clave)
                destino["periodicidad"] = fila.get("periodicidad", "")
                filas.append(destino)
            contraparte["pruebas"] = filas
            for campo in ("notas", "elaboro_y_aprobo"):
                valor = self._texto(nuevo, campo)
                if anterior is None or valor != self._texto(anterior, campo):
                    clave = ("cabecera", campo)
                    completando = clave in self.pendientes[idioma]
                    revisando = clave in self.revisiones[idioma]
                    if valor.strip():
                        self.manuales[idioma].add(clave)
                        self.pendientes[idioma].discard(clave)
                        self.revisiones[idioma].discard(clave)
                        self.automaticas[idioma].discard(clave)
                    if completando and not revisando and self._texto(contraparte, campo).strip():
                        continue
                    traducido, seguro = self.traducir(campo, valor, idioma, otro, {})
                    if clave in self.automaticas[otro] or not self._texto(contraparte, campo).strip():
                        self._poner(contraparte, campo, traducido)
                        if seguro:
                            self.pendientes[otro].discard(clave)
                        else:
                            self.pendientes[otro].add(clave)
                    else:
                        self.pendientes[otro].add(clave)
                        self.revisiones[otro].add(clave)
        contraparte["nombre_producto"] = nuevo["nombre_producto"]
        contraparte["desarrollador"] = nuevo.get("desarrollador", "")
        contraparte["fecha_reanalisis"] = nuevo.get("fecha_reanalisis", "")
        for campo, valor in nuevo["control"].items():
            if campo != "elaboro_y_aprobo":
                contraparte["control"][campo] = valor
        self.borradores[idioma] = nuevo
        for lengua in (idioma, otro):
            ids = {f["_id_bilingue"] for f in self.borradores[lengua]["pruebas"]} | {"cabecera"}
            self.pendientes[lengua] = {c for c in self.pendientes[lengua] if c[0] in ids}
        return copy.deepcopy(nuevo)

    @staticmethod
    def _texto(contenido, campo):
        return "\n".join(contenido["notas"]) if campo == "notas" else contenido["control"].get(campo, "")

    @staticmethod
    def _poner(contenido, campo, valor):
        if campo == "notas":
            contenido["notas"] = valor.splitlines()
        else:
            contenido["control"][campo] = valor

    def confirmar(self, idioma, clave):
        self.manuales[idioma].add(clave)
        self.pendientes[idioma].discard(clave)
        self.revisiones[idioma].discard(clave)
        self.automaticas[idioma].discard(clave)

    def actualizar_traducciones_automaticas(self, idioma=None, completar_sin_traducir=False):
        """Refresca equivalencias sin reemplazar traducciones manuales."""
        # Leer ambos originales antes de actualizar para no traducir de vuelta
        # el resultado que acabamos de producir en el otro idioma.
        originales = copy.deepcopy(self.borradores)
        for destino in ((idioma,) if idioma else ("es", "en")):
            origen = "en" if destino == "es" else "es"
            fuente = originales[origen]
            objetivo = self.borradores[destino]
            filas = {f["_id_bilingue"]: f for f in fuente["pruebas"]}
            celdas = [(f["_id_bilingue"], c, f, filas[f["_id_bilingue"]])
                      for f in objetivo["pruebas"] for c in CAMPOS]
            celdas += [("cabecera", c, None, {}) for c in ("notas", "elaboro_y_aprobo")]
            for identidad, campo, fila, referencia in celdas:
                clave = (identidad, campo)
                if clave in self.manuales[destino]:
                    continue
                actual = self._texto(objetivo, campo) if fila is None else fila.get(campo, "")
                texto = self._texto(fuente, campo) if fila is None else referencia.get(campo, "")
                automatica = clave in self.automaticas[destino]
                if not automatica and not completar_sin_traducir:
                    continue
                if not automatica and clave in self.revisiones[destino]:
                    continue
                if not actual.strip() and texto.strip():
                    continue
                if automatica or actual.strip() == texto.strip():
                    traducido, seguro = self.traducir(campo, texto, origen, destino, referencia)
                else:
                    # Normalizar el texto propio de una pareja guardada, sin
                    # sustituir cifras ni traducciones diferentes por las del otro.
                    traducido, seguro = self.traducir(campo, actual, destino, destino, fila or {})
                # Una equivalencia nueva incierta no reemplaza un texto existente.
                if not seguro:
                    if automatica:
                        self.pendientes[destino].add(clave)
                    continue
                if fila is None:
                    self._poner(objetivo, campo, traducido)
                else:
                    fila[campo] = traducido
                if automatica or actual.strip() == texto.strip():
                    self.automaticas[destino].add(clave)
                self.pendientes[destino].discard(clave)
                self.revisiones[destino].discard(clave)

    def cargar_pareja(self, contenido):
        """Relaciona filas por nombre traducido y método, nunca por posición."""
        idioma = contenido["idioma"]
        generado = self.borradores[idioma]
        cargado = copy.deepcopy(contenido)
        disponibles = list(cargado["pruebas"])
        def identidad(fila):
            valores = []
            for campo in ("prueba", "metodo"):
                original = fila.get(campo, "")
                valor, seguro = self.traducir(campo, original, idioma, idioma, fila)
                valores.append((valor if seguro else original).strip().casefold())
            return tuple(valores)
        identidades = {id(f): identidad(f) for f in disponibles}
        for fila in generado["pruebas"]:
            candidatas = [f for f in disponibles if f.get("prueba", "").strip().casefold() == fila.get("prueba", "").strip().casefold()
                          and f.get("metodo", "").strip().casefold() == fila.get("metodo", "").strip().casefold()]
            if len(candidatas) != 1:
                clave = identidad(fila)
                candidatas = [f for f in disponibles if identidades[id(f)] == clave]
                if len(candidatas) != 1:
                    raise ValueError("No se pueden relacionar de forma inequívoca las pruebas de ESP y ENG. Revisa la pareja antes de editarla.")
            pareja = candidatas[0]
            pareja["_id_bilingue"] = fila["_id_bilingue"]
            disponibles.remove(pareja)
        if disponibles:
            raise ValueError("La pareja ESP / ENG tiene pruebas distintas. Revisa ambos JSON antes de editarlos como pareja.")
        por_id = {f["_id_bilingue"]: f for f in cargado["pruebas"]}
        cargado["pruebas"] = [por_id[f["_id_bilingue"]] for f in generado["pruebas"]]
        self.borradores[idioma] = cargado
        self.pendientes[idioma].clear()
        self.automaticas[idioma].clear()
        for indice, fila in enumerate(cargado["pruebas"]):
            for campo in CAMPOS:
                if generado["pruebas"][indice].get(campo, "").strip() and not fila.get(campo, "").strip():
                    self.pendientes[idioma].add((fila["_id_bilingue"], campo))
                if campo in {"especificacion", "valor_tipico"} and re.findall(r"[+-]?\d+(?:[.,]\d+)?", generado["pruebas"][indice].get(campo, "")) != re.findall(r"[+-]?\d+(?:[.,]\d+)?", fila.get(campo, "")):
                    self.pendientes[idioma].add((fila["_id_bilingue"], campo))
                    self.revisiones[idioma].add((fila["_id_bilingue"], campo))
        if self._texto(generado, "notas").strip() and not self._texto(cargado, "notas").strip():
            self.pendientes[idioma].add(("cabecera", "notas"))


def guardar_pareja_json(documentos):
    """Prepara ambos archivos antes de reemplazarlos; restaura ante errores de E/S."""
    import json
    import os
    import tempfile
    from pathlib import Path
    temporales = {}
    respaldos = {}
    escritos = []
    try:
        for ruta, datos in documentos.items():
            ruta = Path(ruta)
            respaldos[ruta] = ruta.read_bytes() if ruta.exists() else None
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=ruta.parent,
                                             prefix=".bilingue_", suffix=".tmp", delete=False) as archivo:
                temporales[ruta] = Path(archivo.name)
                json.dump(datos, archivo, ensure_ascii=False, indent=2)
                archivo.write("\n")
        for ruta, temporal in temporales.items():
            os.replace(temporal, ruta)
            escritos.append(ruta)
    except Exception:
        for ruta in reversed(escritos):
            if respaldos[ruta] is None:
                ruta.unlink(missing_ok=True)
            else:
                ruta.write_bytes(respaldos[ruta])
        raise
    finally:
        for temporal in temporales.values():
            temporal.unlink(missing_ok=True)
