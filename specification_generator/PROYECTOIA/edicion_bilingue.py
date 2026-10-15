"""Estado y sincronización de borradores; las traducciones las provee el editor."""
from __future__ import annotations

import copy
import re
import uuid

# Nombres de campos de la tabla, sin equivalencias de traducción.
CAMPOS = ('prueba', 'unidades', 'metodo', 'especificacion', 'valor_tipico')


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
                            self.manuales[idioma].discard(clave)
                        else:
                            destino[campo] = ""
                            for lengua in (idioma, otro):
                                self.pendientes[lengua].discard(clave)
                                self.revisiones[lengua].discard(clave)
                        continue
                    completando = clave in self.pendientes[idioma]
                    revisando = clave in self.revisiones[idioma]
                    if valor.strip():
                        self.pendientes[idioma].discard(clave)
                        self.revisiones[idioma].discard(clave)
                        self.automaticas[idioma].discard(clave)
                        self.manuales[idioma].add(clave)
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
                        self.pendientes[idioma].discard(clave)
                        self.revisiones[idioma].discard(clave)
                        self.automaticas[idioma].discard(clave)
                        self.manuales[idioma].add(clave)
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
            self.manuales[lengua] = {c for c in self.manuales[lengua] if c[0] in ids}
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
        self.pendientes[idioma].discard(clave)
        self.revisiones[idioma].discard(clave)
        self.automaticas[idioma].discard(clave)
        self.manuales[idioma].add(clave)

    def actualizar_traducciones_automaticas(self, idioma_destino=None, completar_sin_traducir=False):
        """Recalcula sugerencias del diccionario sin reemplazar correcciones manuales."""
        fuentes = copy.deepcopy(self.borradores)
        for idioma in ("es", "en") if idioma_destino is None else (idioma_destino,):
            otro = "en" if idioma == "es" else "es"
            origenes = {f["_id_bilingue"]: f for f in fuentes[otro]["pruebas"]}
            for fila in self.borradores[idioma]["pruebas"]:
                fuente = origenes.get(fila["_id_bilingue"])
                if fuente is None:
                    continue
                for campo in CAMPOS:
                    marca = (fila["_id_bilingue"], campo)
                    sin_traducir = (not fila.get(campo, "").strip()
                                    or fila.get(campo, "").strip().casefold() == fuente.get(campo, "").strip().casefold()
                                    or (marca in self.pendientes[idioma] and marca not in self.revisiones[idioma]))
                    if marca in self.manuales[idioma] or (marca not in self.automaticas[idioma] and not (completar_sin_traducir and sin_traducir)):
                        continue
                    self.automaticas[idioma].add(marca)
                    fila[campo], seguro = self.traducir(campo, fuente.get(campo, ""), otro, idioma, fuente)
                    if seguro:
                        self.pendientes[idioma].discard(marca)
                    else:
                        self.pendientes[idioma].add(marca)
            for campo in ("notas", "elaboro_y_aprobo"):
                marca = ("cabecera", campo)
                valor_actual = self._texto(self.borradores[idioma], campo).strip()
                valor_origen = self._texto(fuentes[otro], campo).strip()
                sin_traducir = not valor_actual or valor_actual.casefold() == valor_origen.casefold()
                if marca not in self.manuales[idioma] and (marca in self.automaticas[idioma] or (completar_sin_traducir and sin_traducir)):
                    self.automaticas[idioma].add(marca)
                    valor, seguro = self.traducir(campo, self._texto(fuentes[otro], campo), otro, idioma, {})
                    self._poner(self.borradores[idioma], campo, valor)
                    if seguro:
                        self.pendientes[idioma].discard(marca)
                    else:
                        self.pendientes[idioma].add(marca)

    def cargar_pareja(self, contenido):
        """Relaciona filas por nombre traducido y método, nunca por posición."""
        idioma = contenido["idioma"]
        generado = self.borradores[idioma]
        cargado = copy.deepcopy(contenido)
        disponibles = list(cargado["pruebas"])
        for fila in generado["pruebas"]:
            candidatas = [f for f in disponibles if f.get("prueba", "").strip().casefold() == fila.get("prueba", "").strip().casefold()
                          and f.get("metodo", "").strip().casefold() == fila.get("metodo", "").strip().casefold()]
            if len(candidatas) != 1:
                candidatas = [f for f in disponibles if fila.get("metodo", "").strip()
                              and f.get("metodo", "").strip().casefold() == fila["metodo"].strip().casefold()]
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
        if self._texto(generado, "notas").strip() and not self._texto(cargado, "notas").strip():
            self.pendientes[idioma].add(("cabecera", "notas"))


def guardar_pareja_json(documentos):
    """Prepara ambos archivos antes de reemplazarlos; restaura ante errores de E/S."""
    import json
    import os
    import tempfile
    from pathlib import Path
    from rutas_proyecto import resolver_ruta
    temporales = {}
    respaldos = {}
    escritos = []
    try:
        for ruta, datos in documentos.items():
            ruta = Path(ruta)
            if not ruta.is_absolute():
                ruta = resolver_ruta(ruta)
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
