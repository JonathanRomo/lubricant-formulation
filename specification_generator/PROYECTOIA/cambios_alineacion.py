"""Diferencias de alineación calculadas en memoria; no modifica documentos."""
import json
from pathlib import Path
import tkinter as tk
import re
import unicodedata
from collections import defaultdict
from decimal import Decimal
from estilo_interfaz import ttk
import estilo_interfaz as ui

CAMPOS = ("prueba", "unidades", "metodo", "especificacion", "valor_tipico")
ORIGINALES = {"prueba": "nombre_original", "metodo": "metodo_original",
              "unidades": "unidades_originales", "valor_tipico": "valor_tipico_original"}


def _clave_comparacion(valor):
    valor = unicodedata.normalize("NFKD", texto(valor).upper())
    return " ".join("".join(c for c in valor if not unicodedata.combining(c)).split())


def _valor_comparable(valor, idioma, traducir):
    valor = texto(valor)
    if idioma == "en" and traducir is not None:
        valor, _ = traducir(valor, "en", "es")
    valor = _clave_comparacion(valor).replace("−", "-").replace("–", "-").replace("—", "-")
    def numero(coincidencia):
        original = coincidencia.group()
        partes = re.split(r"[.,]", original)
        # BETA: no adivinar si tres decimales representan separadores de miles.
        if len(partes) == 2 and len(partes[1]) == 3 and len(partes[0].lstrip("+-")) <= 3:
            return original
        decimal = Decimal(original.replace(",", "."))
        return format(decimal.normalize(), "f") if decimal else "0"
    return re.sub(r"(?<![\w.,])[+-]?\d+(?:[.,]\d+)?(?![\w.,])", numero, valor)


def comparar_valores_pruebas(especificaciones, maestro=(), traducir=None):
    # FUNCIÓN EN ESTADO BETA: comparación conservadora, sin conversiones de
    # unidades ni coincidencias ambiguas. Solo resalta especificación y valor típico.
    from alinear_especificaciones_json import metodo_referencias_canonico
    nombres = defaultdict(set)
    for indice, registro in enumerate(maestro):
        for campo in ("NOMBRE IQM", "NOMBRE IQM INGLES"):
            if registro.get(campo):
                nombres[_clave_comparacion(registro[campo])].add(indice)
    metodos = defaultdict(lambda: defaultdict(list))
    cambios = [{} for _ in especificaciones]
    for panel, especificacion in enumerate(especificaciones):
        idioma = especificacion.get("idioma", "es")
        for fila, prueba in enumerate(especificacion["pruebas"]):
            metodo = texto(prueba.get("metodo"))
            if idioma == "en" and traducir:
                metodo, _ = traducir(metodo, "en", "es")
            metodo = _clave_comparacion(metodo_referencias_canonico(metodo))
            if not metodo or metodo in {"NA", "N/A", "-"}:
                continue
            nombre = _clave_comparacion(prueba.get("prueba", prueba.get("nombre")))
            candidatos = nombres.get(nombre, set())
            if len(candidatos) == 1:
                nombre = ("maestro", next(iter(candidatos)))
            elif idioma == "en" and traducir:
                nombre = _clave_comparacion(traducir(texto(prueba.get("prueba")), "en", "es")[0])
            unidades = re.sub(r"\s+", "", _clave_comparacion(prueba.get("unidades")))
            if unidades == "N/A":
                unidades = "NA"
            metodos[metodo][panel].append((fila, nombre, unidades, prueba, idioma))
    comunes, distintas_unidades = 0, 0
    def comparar(grupo):
        nonlocal comunes, distintas_unidades
        por_unidad = defaultdict(list)
        for panel, entrada in grupo:
            por_unidad[entrada[2]].append((panel, entrada))
        if len(por_unidad) > 1:
            distintas_unidades += 1
        for filas in por_unidad.values():
            if len(filas) < 2:
                continue
            comunes += 1
            for campo in ("especificacion", "valor_tipico"):
                valores = {_valor_comparable(f[3].get(campo), f[4], traducir) for _, f in filas}
                if len(valores) > 1:
                    for panel, f in filas:
                        cambios[panel].setdefault(f[0], set()).add(campo)
    for paneles in metodos.values():
        unicas = [(p, filas[0]) for p, filas in paneles.items() if len(filas) == 1]
        identidades = {f[1] for _, f in unicas if isinstance(f[1], tuple)}
        temperaturas = {tuple(re.sub(r"\s+", "", t) for t in re.findall(r"\d+(?:[.,]\d+)?\s*°?\s*[CF]\b", _clave_comparacion(f[3].get("prueba"))))
                        for _, f in unicas}
        temperaturas.discard(())
        admite_metodo_solo = len(identidades) <= 1 and len(temperaturas) <= 1
        if len(unicas) == len(paneles) and len(unicas) >= 2 and admite_metodo_solo:
            comparar(unicas)
            continue
        por_nombre = defaultdict(lambda: defaultdict(list))
        for panel, filas in paneles.items():
            for fila in filas:
                if fila[1]:
                    por_nombre[fila[1]][panel].append(fila)
        usados = set()
        for paneles_nombre in por_nombre.values():
            grupo = [(p, filas[0]) for p, filas in paneles_nombre.items() if len(filas) == 1]
            if len(grupo) >= 2:
                comparar(grupo)
                usados.update((p, f[0]) for p, f in grupo)
        # El método solo basta cuando aparece una vez por especificación.
        if len(unicas) >= 2 and admite_metodo_solo and not any((p, f[0]) in usados for p, f in unicas):
            comparar(unicas)
    return cambios, {"comunes": comunes, "celdas": sum(len(c) for p in cambios for c in p.values()),
                     "unidades_distintas": distintas_unidades}


def texto(valor):
    return "" if valor is None else str(valor).strip()


def diferencias_filas(originales, alineadas):
    return [{campo for campo in CAMPOS if campo in original and texto(original.get(campo)) != texto(alineada.get(campo))}
            for original, alineada in zip(originales, alineadas)]


def original_disponible(datos, ruta, carpeta):
    if ruta is None or carpeta is None:
        return None
    from rutas_proyecto import resolver_ruta
    candidata = resolver_ruta(carpeta) / Path(ruta).name
    try:
        original = json.loads(candidata.read_text(encoding="utf-8-sig"))
        a, b = datos.get("control_documental", {}), original.get("control_documental", {})
        if any(texto(a.get(c)) != texto(b.get(c)) for c in ("codigo_unico_producto", "revision", "nombre_cliente")):
            return None
        if datos.get("documento", {}).get("idioma") != original.get("documento", {}).get("idioma"):
            return None
        return original
    except (OSError, ValueError, AttributeError):
        return None


def cambios_documento(datos, visibles, original=None):
    detalle = datos.get("alineacion_metodos_lab", {}).get("detalle", [])
    pruebas = (original or datos).get("pruebas", [])
    disponibles = []
    for indice, auditoria in enumerate(detalle):
        alineada = auditoria.get("valores_alineados", {})
        fuente = dict(pruebas[indice]) if indice < len(pruebas) and isinstance(pruebas[indice], dict) else {}
        if original is None:
            for campo, clave in ORIGINALES.items():
                if clave in auditoria:
                    fuente[campo] = auditoria[clave]
            if "prueba" not in fuente and "nombre" in fuente:
                fuente["prueba"] = fuente["nombre"]
        cambios = diferencias_filas([fuente], [alineada])[0]
        disponibles.append((tuple(texto(alineada.get(c)) for c in CAMPOS), cambios))
    resultado = []
    for fila in visibles:
        firma = tuple(texto(fila.get(c)) for c in CAMPOS)
        candidatas = [cambios for guardada, cambios in disponibles if guardada == firma]
        resultado.append(set(candidatas[0]) if candidatas and all(c == candidatas[0] for c in candidatas) else set())
    return resultado


class ResaltadorCambios:
    """Texto azul por celda, con el fondo y la fuente propios del Treeview."""
    def __init__(self, tabla, cambios, vertical, horizontal):
        self.tabla, self.cambios = tabla, cambios
        self.marcas, self.pendiente = [], None
        tabla.configure(yscrollcommand=lambda *a: (vertical(*a), self.programar()),
                        xscrollcommand=lambda *a: (horizontal(*a), self.programar()))
        for evento in ("<Configure>", "<<TreeviewSelect>>", "<ButtonRelease-1>"):
            tabla.bind(evento, lambda _e: self.programar(), add="+")
        tabla.bind("<Destroy>", self.cerrar, add="+")
        self.programar()

    def programar(self):
        if self.pendiente is None and self.tabla.winfo_exists():
            self.pendiente = self.tabla.after_idle(self.dibujar)

    def establecer_cambios(self, cambios):
        self.cambios = cambios
        self.programar()

    def cerrar(self, evento):
        if evento.widget is self.tabla and self.pendiente is not None:
            self.tabla.after_cancel(self.pendiente)
            self.pendiente = None

    def dibujar(self):
        self.pendiente = None
        if not self.tabla.winfo_exists():
            return
        for marca in self.marcas:
            marca.destroy()
        self.marcas.clear()
        estilo = ttk.Style(self.tabla)
        nombre = self.tabla.cget("style") or "Treeview"
        fondo = estilo.lookup(nombre, "fieldbackground") or ui.paleta()["panel"]
        fuente = estilo.lookup(nombre, "font") or (ui.PREFERENCIAS["fuente"], 10)
        seleccion = set(self.tabla.selection())
        for iid in self.tabla.get_children():
            valores = self.tabla.item(iid, "values")
            for columna, campo in enumerate(CAMPOS, 1):
                if campo not in self.cambios.get(iid, set()):
                    continue
                caja = self.tabla.bbox(iid, f"#{columna}")
                if not caja:
                    continue
                x, y, ancho, alto = caja
                izquierda = max(0, x)
                ancho = min(x + ancho, self.tabla.winfo_width()) - izquierda
                if ancho <= 0:
                    continue
                seleccionado = iid in seleccion
                color_fondo = estilo.lookup(nombre, "background", ("selected",)) if seleccionado else fondo
                marca = tk.Label(self.tabla, text=valores[columna - 1] or "*", anchor="w", padx=2,
                                 background=color_fondo, foreground="#bfdbfe" if seleccionado else "#3b82f6",
                                 font=fuente, borderwidth=0)
                marca.place(x=izquierda, y=y, width=ancho, height=alto)
                marca.bind("<Button-1>", lambda _e, i=iid: self.seleccionar(i))
                marca.bind("<MouseWheel>", lambda e: self.tabla.event_generate("<MouseWheel>", delta=e.delta))
                self.marcas.append(marca)

    def seleccionar(self, iid):
        self.tabla.selection_set(iid)
        self.tabla.focus(iid)
        self.tabla.focus_set()
