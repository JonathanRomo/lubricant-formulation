"""Cobertura dirigida de requisitos entre especificaciones JSON alineadas."""
from __future__ import annotations

import json
import math
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass
from itertools import combinations
from pathlib import Path
from tkinter import TclError, messagebox

import estilo_interfaz as ui
from estilo_interfaz import ctk, ttk


ESTADOS_ACEPTADOS = {"ALINEADO", "EDICION_MANUAL"}
SIN_DATO = {"", "-", "--", "---", "NA", "N/A", "N D", "N/D", "NO APLICA", "NO APLICABLE"}
SIN_REQUISITO = SIN_DATO | {"REPORTAR", "REPORT", "INFORMAR", "INFORMATIVE"}
PATRON_NUMERO = r"[+-]?(?:\d{1,3}(?: \d{3})+|\d+)(?:[.,]\d+(?: \d)*)?"
TEXTO_AYUDA_COBERTURA = (
    "Cada candidato se compara contra TODOS los requisitos del producto consultado. "
    "Un rango candidato cuenta como cubierto solo si queda íntegramente dentro del límite solicitado; "
    "un valor típico dentro del límite es indicativo, no una garantía."
)
TEXTO_AYUDA_MATRIZ = (
    "Selecciona uno o varios productos; cada fila será un producto elegido y cada columna una variable "
    "técnica alineada: prueba, unidad y método. "
    "La vista Valores típicos conserva los datos faltantes como nulos y los muestra como N/D. "
    "La vista Disponibilidad usa 1 cuando existe un valor típico y 0 cuando está ausente."
)
TEXTO_AYUDA_MULTIDIMENSIONAL = (
    "Selecciona al menos dos productos de una misma categoría o familia técnica. "
    "Se compararán todas las parejas. La distancia técnica usa valores típicos y requisitos comparables; "
    "una ventaja solo se declara cuando ambos documentos indican la misma dirección de desempeño "
    "(mínimo: mayor es favorable; máximo: menor es favorable). La confianza combina superposición, "
    "cantidad de pruebas compartidas y proporción de datos realmente evaluables."
)


def normalizar(valor: object) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or "").translate(
        str.maketrans({"−": "-", "–": "-", "—": "-", "≤": "<=", "≥": ">="})))
    texto = "".join(letra for letra in texto if not unicodedata.combining(letra))
    return " ".join(texto.upper().split())


def numero(valor: str) -> float:
    return float(valor.replace(" ", "").replace(",", "."))


def dato_numerico(valor: object) -> float | None:
    texto = normalizar(valor)
    texto = re.sub(r"(?<!\d)([+-])\s+(?=\d)", r"\1", texto)
    if re.fullmatch(PATRON_NUMERO, texto):
        return numero(texto)
    return None


def requisito(valor: object) -> tuple[str, object] | None:
    texto = normalizar(valor)
    texto = re.sub(r"(?<!\d)([+-])\s+(?=\d)", r"\1", texto)
    if texto in SIN_REQUISITO:
        return None
    if re.fullmatch(rf"{PATRON_NUMERO}\s*(TIPICO|TYPICAL)", texto):
        return None
    rango = re.fullmatch(rf"({PATRON_NUMERO})\s*(?:-|A|TO)\s*({PATRON_NUMERO})", texto)
    if rango:
        a, b = numero(rango[1]), numero(rango[2])
        return ("intervalo", (min(a, b), max(a, b)))
    limite = re.fullmatch(rf"({PATRON_NUMERO})\s*(MINIMO|MINIMUM|MIN|MAXIMO|MAXIMUM|MAX)", texto)
    if limite:
        return ("minimo" if limite[2].startswith("MIN") else "maximo", numero(limite[1]))
    if dato_numerico(texto) is not None:
        return ("puntual", numero(texto))
    equivalentes = {"NEGATIVE": "NEGATIVA", "PASS": "PASA", "POSITIVE": "POSITIVA"}
    return ("texto", equivalentes.get(texto, texto))


def peso(prueba: dict) -> float:
    periodicidad = str(prueba.get("PERIODIDICAD DE ANALISIS", ""))
    return 3.0 if "◊" in periodicidad else 1.0


def clave_prueba(prueba: dict, detalle: dict) -> tuple[str, str, str] | None:
    if detalle.get("estado") not in ESTADOS_ACEPTADOS:
        return None
    nombre = normalizar(detalle.get("nombre_maestro_propuesto") or prueba.get("prueba"))
    unidades = normalizar(prueba.get("unidades"))
    metodo = normalizar(detalle.get("metodo_maestro_propuesto") or prueba.get("metodo"))
    # NA es una unidad declarada válida para pruebas cualitativas o adimensionales.
    if not nombre or not metodo or not unidades:
        return None
    return (nombre, unidades, metodo)


@dataclass
class Producto:
    ruta: Path
    nombre: str
    codigo: str
    idioma: str
    fecha: str
    categoria: str
    uso: str
    pruebas: dict[tuple[str, str, str], dict]

    @property
    def numero_pruebas(self) -> int:
        return sum(len(filas_grupo(fila)) for fila in self.pruebas.values())

    @property
    def etiqueta(self) -> str:
        return f"{self.nombre} · {self.codigo or 'sin código'} · {self.numero_pruebas} pruebas"


def filas_grupo(fila: dict) -> list[dict]:
    """Una clave alineada no garantiza que las condiciones de sus filas sean iguales."""
    return fila.get("_filas_repetidas", [fila])


def leer_producto(ruta: Path) -> Producto | None:
    datos = json.loads(ruta.read_text(encoding="utf-8-sig"))
    if not isinstance(datos, dict) or not isinstance(datos.get("pruebas"), list):
        return None
    if not isinstance(datos.get("alineacion_metodos_lab"), dict):
        return None
    producto = datos.get("producto") or {}
    documento = datos.get("documento") or {}
    control = datos.get("control_documental") or {}
    if not all(isinstance(campo, dict) for campo in (producto, documento, control)):
        return None
    nombre = str(producto.get("nombre") or "").strip()
    if not nombre:
        return None
    detalles = {d.get("numero_prueba"): d for d in datos["alineacion_metodos_lab"].get("detalle", [])
                if isinstance(d, dict)}
    pruebas = {}
    for fila in datos["pruebas"]:
        if not isinstance(fila, dict):
            continue
        detalle = detalles.get(fila.get("numero"), {})
        clave = clave_prueba(fila, detalle)
        if clave:
            if clave in pruebas:
                pruebas[clave] = {"_filas_repetidas": [*filas_grupo(pruebas[clave]), fila]}
            else:
                pruebas[clave] = fila
    if not pruebas:
        return None
    return Producto(
        ruta=ruta, nombre=nombre,
        codigo=str(control.get("codigo_unico_producto") or "").strip(),
        idioma=str(documento.get("idioma") or "").lower(),
        fecha=str(documento.get("fecha_documento_normalizada") or ""),
        categoria=str(producto.get("categoria") or producto.get("categoría") or "").strip(),
        uso=str(producto.get("uso") or producto.get("aplicacion") or producto.get("aplicación") or "").strip(),
        pruebas=pruebas,
    )


def cargar_catalogo(carpeta: Path) -> tuple[list[Producto], int]:
    if not carpeta.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de JSON alineados: {carpeta}")
    encontrados: dict[str, Producto] = {}
    descartados = 0
    for ruta in sorted(carpeta.iterdir()):
        if ruta.suffix.lower() != ".json" or ruta.stem.lower().endswith("_trabajo"):
            continue
        try:
            producto = leer_producto(ruta)
        except (OSError, UnicodeError, ValueError, TypeError, KeyError):
            producto = None
        if producto is None:
            descartados += 1
            continue
        identidad = producto.codigo or normalizar(producto.nombre)
        anterior = encontrados.get(identidad)
        prioridad = (producto.idioma == "es", producto.fecha, producto.numero_pruebas, ruta.name)
        if anterior is None or prioridad > (anterior.idioma == "es", anterior.fecha,
                                             anterior.numero_pruebas, anterior.ruta.name):
            encontrados[identidad] = producto
    return sorted(encontrados.values(), key=lambda p: (p.nombre.casefold(), p.codigo)), descartados


def limites(valor: tuple[str, object]) -> tuple[float, float] | None:
    tipo, dato = valor
    if tipo == "intervalo":
        return dato
    if tipo == "minimo":
        return (dato, math.inf)
    if tipo == "maximo":
        return (-math.inf, dato)
    if tipo == "puntual":
        return (dato, dato)
    return None


def metodo_compatible(a: str, b: str) -> bool:
    if a == b:
        return True
    astm_a = set(re.findall(r"ASTM\s*([A-Z]\d+[A-Z]?)", a))
    astm_b = set(re.findall(r"ASTM\s*([A-Z]\d+[A-Z]?)", b))
    return bool(astm_a and astm_b and astm_a & astm_b)


def satisface(valor: tuple[str, object], esperado: tuple[str, object]) -> bool:
    l_valor, l_esperado = limites(valor), limites(esperado)
    if l_valor is not None and l_esperado is not None:
        return l_valor[0] >= l_esperado[0] and l_valor[1] <= l_esperado[1]
    return valor == esperado if valor[0] == esperado[0] == "texto" else False


def evaluar_requisito(origen: dict, candidato: dict) -> tuple[str, str]:
    """La cobertura garantizada exige que TODO el rango candidato satisfaga el origen."""
    esperado = requisito(origen.get("especificacion"))
    if esperado is None:
        return "SIN_REQUISITO", "El producto consultado no declara un límite evaluable."
    declarado = requisito(candidato.get("especificacion"))
    tipico_numerico = dato_numerico(candidato.get("valor_tipico"))
    tipico_texto = requisito(candidato.get("valor_tipico"))
    tipico = (("puntual", tipico_numerico) if tipico_numerico is not None
              else tipico_texto if tipico_texto and tipico_texto[0] == "texto" else None)
    if declarado is not None and satisface(declarado, esperado):
        if tipico is not None and not satisface(tipico, esperado):
            return "INCONSISTENTE", "La especificación cubre, pero el valor típico declarado queda fuera."
        return "CUBRE", "La especificación completa del candidato queda dentro del requisito consultado."
    if tipico is not None and satisface(tipico, esperado):
        return "INDICATIVO", "El valor típico cae dentro, pero la especificación no garantiza el requisito."
    if tipico is not None and not satisface(tipico, esperado):
        return "FUERA", "El valor típico declarado queda fuera del requisito consultado."
    if declarado is None:
        return "SIN_DATO", "El candidato no declara un límite evaluable ni un típico aplicable."
    la, lb = limites(esperado), limites(declarado)
    if la is not None and lb is not None:
        if lb[1] < la[0] or lb[0] > la[1]:
            return "FUERA", "Los intervalos declarados no se intersectan."
        return "PARCIAL", "Los rangos se cruzan o el candidato permite valores fuera del requisito."
    if esperado[0] == declarado[0] == "texto":
        opuestos = {"POSITIVA", "NEGATIVA"}
        if {esperado[1], declarado[1]} == opuestos:
            return "FUERA", "Los resultados cualitativos declarados son opuestos."
        return "SIN_DATO", "Los requisitos textuales difieren y no hay una jerarquía verificable."
    return "SIN_DATO", "Los formatos de especificación no permiten verificar la cobertura."


@dataclass
class FilaAnalisis:
    nombre: str
    origen: dict
    candidato: dict | None
    estado: str
    motivo: str


@dataclass
class Comparacion:
    candidato: Producto
    filas: list[FilaAnalisis]
    requisitos: int
    cubiertos: int
    indicativos: int
    fuera: int
    parciales: int
    faltantes: int
    sin_dato: int
    compartidas: int
    cobertura: float

    @property
    def dictamen(self) -> str:
        if self.requisitos == 0:
            return "Sin requisitos evaluables"
        if self.cubiertos == self.requisitos:
            return "Cubre lo declarado"
        if self.cubiertos == 0:
            return "Sin cobertura demostrada"
        return "Cobertura parcial"


def comparar(origen: Producto, candidato: Producto) -> Comparacion:
    """Comparación dirigida: cambiar origen y candidato puede cambiar el resultado."""
    por_nombre: dict[str, list[tuple[tuple[str, str, str], dict]]] = {}
    for clave, fila in candidato.pruebas.items():
        por_nombre.setdefault(clave[0], []).append((clave, fila))
    filas: list[FilaAnalisis] = []
    requisitos = cubiertos = indicativos = fuera = parciales = faltantes = sin_dato = compartidas = 0
    peso_total = peso_cubierto = 0.0
    for clave, fila in ((clave, fila) for clave, grupo in origen.pruebas.items() for fila in filas_grupo(grupo)):
        nombre, unidad, metodo = clave
        esperado = requisito(fila.get("especificacion"))
        if esperado is not None:
            requisitos += 1
            peso_total += peso(fila)
        opciones = por_nombre.get(nombre, [])
        compatibles = [(k, f) for k, f in opciones if k[1] == unidad and metodo_compatible(metodo, k[2])]
        if len(compatibles) > 1 or (compatibles and len(filas_grupo(compatibles[0][1])) > 1):
            compartidas += 1
            rival = None
            estado, motivo = "NO_COMPARABLE", "Hay varias filas candidatas para esta prueba; revisa sus condiciones y límites. No se elige la primera."
        elif compatibles:
            compartidas += 1
            _, rival = compatibles[0]
            estado, motivo = evaluar_requisito(fila, rival)
        else:
            rival = None
            if opciones:
                estado, motivo = "NO_COMPARABLE", "La unidad o el método difiere; no se presume equivalencia."
            else:
                estado, motivo = "AUSENTE", "El candidato no incluye esta prueba alineada."
        filas.append(FilaAnalisis(nombre, fila, rival, estado, motivo))
        if esperado is None:
            continue
        if estado == "CUBRE":
            cubiertos += 1
            peso_cubierto += peso(fila)
        elif estado == "INDICATIVO":
            indicativos += 1
        elif estado == "FUERA":
            fuera += 1
        elif estado == "PARCIAL":
            parciales += 1
        elif estado in {"AUSENTE", "NO_COMPARABLE"}:
            faltantes += 1
        else:
            sin_dato += 1
    return Comparacion(candidato, filas, requisitos, cubiertos, indicativos, fuera,
                       parciales, faltantes, sin_dato, compartidas,
                       peso_cubierto / peso_total if peso_total else 0.0)


def distancia_relativa(a: float, b: float) -> float:
    return min(abs(a - b) / max(abs(a), abs(b), 1e-9), 1.0)


def distancia_fila(a: dict, b: dict) -> tuple[float | None, str]:
    valor_a, valor_b = dato_numerico(a.get("valor_tipico")), dato_numerico(b.get("valor_tipico"))
    if valor_a is not None and valor_b is not None:
        return distancia_relativa(valor_a, valor_b), "valores típicos"
    tipico_a, tipico_b = requisito(a.get("valor_tipico")), requisito(b.get("valor_tipico"))
    if tipico_a and tipico_b and tipico_a[0] == tipico_b[0] == "texto":
        return (0.0 if tipico_a[1] == tipico_b[1] else 1.0), "valores típicos cualitativos"
    req_a, req_b = requisito(a.get("especificacion")), requisito(b.get("especificacion"))
    if req_a is None or req_b is None:
        return None, "sin datos comparables"
    if req_a[0] == req_b[0] == "intervalo":
        centro_a, centro_b = sum(req_a[1]) / 2, sum(req_b[1]) / 2
        ancho_a, ancho_b = req_a[1][1] - req_a[1][0], req_b[1][1] - req_b[1][0]
        return (0.75 * distancia_relativa(centro_a, centro_b)
                + 0.25 * distancia_relativa(ancho_a, ancho_b)), "rangos especificados"
    if req_a[0] == req_b[0] and req_a[0] in {"minimo", "maximo", "puntual"}:
        return distancia_relativa(req_a[1], req_b[1]), "límites especificados"
    if req_a[0] == req_b[0] == "texto":
        return (0.0 if req_a[1] == req_b[1] else 1.0), "requisitos cualitativos"
    return None, "formatos no comparables"


def ventaja_dirigida(nombre: str, a: dict, b: dict) -> tuple[str, str] | None:
    """Devuelve el lado favorecido solo cuando mínimo/máximo define la dirección."""
    req_a, req_b = requisito(a.get("especificacion")), requisito(b.get("especificacion"))
    if req_a is None or req_b is None or req_a[0] != req_b[0] or req_a[0] not in {"minimo", "maximo"}:
        return None
    valor_a, valor_b = dato_numerico(a.get("valor_tipico")), dato_numerico(b.get("valor_tipico"))
    fuente = "valor típico"
    if valor_a is None or valor_b is None:
        valor_a, valor_b = float(req_a[1]), float(req_b[1])
        fuente = "límite declarado"
    if valor_a == valor_b:
        return None
    lado = "A" if ((req_a[0] == "minimo" and valor_a > valor_b)
                   or (req_a[0] == "maximo" and valor_a < valor_b)) else "B"
    mejor, referencia = (valor_a, valor_b) if lado == "A" else (valor_b, valor_a)
    diferencia = abs(mejor - referencia) / max(abs(referencia), 1e-9)
    if diferencia < 0.02:
        return None
    direccion = "mayor" if req_a[0] == "minimo" else "menor"
    unidad = str(a.get("unidades") or "").strip()
    descripcion = (f"{nombre}: {diferencia:.1%} favorable ({fuente} {direccion}; "
                   f"A={valor_a:g}, B={valor_b:g}{' ' + unidad if unidad else ''})")
    return lado, descripcion


@dataclass
class ComparacionMultidimensional:
    producto_a: Producto
    producto_b: Producto
    compartidas: int
    total_union: int
    superposicion: float
    distancia: float | None
    evaluables: int
    ventajas_a: list[str]
    ventajas_b: list[str]
    diferencias: list[str]
    exclusivas_a: list[str]
    exclusivas_b: list[str]
    confianza: str
    confianza_puntaje: float


def comparar_multidimensional(a: Producto, b: Producto) -> ComparacionMultidimensional:
    claves_a, claves_b = set(a.pruebas), set(b.pruebas)
    compartidas = claves_a & claves_b
    union = claves_a | claves_b
    pesos = distancia_total = 0.0
    evaluables = 0
    ventajas_a: list[str] = []
    ventajas_b: list[str] = []
    diferencias: list[str] = []
    for clave in sorted(compartidas):
        fila_a, fila_b = a.pruebas[clave], b.pruebas[clave]
        if len(filas_grupo(fila_a)) > 1 or len(filas_grupo(fila_b)) > 1:
            diferencias.append(f"{clave[0]}: varias filas con la misma clave; requiere revisión de condiciones y límites.")
            continue
        distancia, fuente = distancia_fila(fila_a, fila_b)
        if distancia is not None:
            importancia = max(peso(fila_a), peso(fila_b))
            distancia_total += distancia * importancia
            pesos += importancia
            evaluables += 1
            if distancia >= 0.10:
                diferencias.append(
                    f"{clave[0]}: distancia {distancia:.1%} por {fuente}; "
                    f"A={fila_a.get('valor_tipico') or fila_a.get('especificacion') or 'N/D'}, "
                    f"B={fila_b.get('valor_tipico') or fila_b.get('especificacion') or 'N/D'}"
                )
        ventaja = ventaja_dirigida(clave[0], fila_a, fila_b)
        if ventaja:
            (ventajas_a if ventaja[0] == "A" else ventajas_b).append(ventaja[1])
    superposicion = len(compartidas) / len(union) if union else 0.0
    distancia_tecnica = distancia_total / pesos if pesos else None
    proporcion_evaluable = evaluables / len(compartidas) if compartidas else 0.0
    confianza_puntaje = (
        0.45 * superposicion
        + 0.30 * min(len(compartidas) / 12, 1.0)
        + 0.25 * proporcion_evaluable
    )
    confianza = "Alta" if confianza_puntaje >= 0.72 else "Media" if confianza_puntaje >= 0.45 else "Baja"
    exclusivas_a = sorted({clave[0] for clave in claves_a - claves_b})
    exclusivas_b = sorted({clave[0] for clave in claves_b - claves_a})
    return ComparacionMultidimensional(
        a, b, len(compartidas), len(union), superposicion, distancia_tecnica, evaluables,
        ventajas_a, ventajas_b, diferencias, exclusivas_a, exclusivas_b,
        confianza, confianza_puntaje,
    )


class _ListaSelectorProductos(ctk.CTkFrame):
    """Una tabla reutilizable: filtrar no crea ni destruye controles por producto."""

    def __init__(self, padre, valores, consulta, seleccionados, elegir, *, multiple=False):
        colores = ui.paleta()
        super().__init__(padre, fg_color=colores["panel"], corner_radius=0)
        self.consulta = consulta
        self.seleccionados = seleccionados
        self.elegir = elegir
        self.multiple = multiple
        self._pendiente = None
        self._filas = {str(i): valor for i, valor in enumerate(valores)}
        self._busqueda = [(fila, normalizar(valor)) for fila, valor in self._filas.items()]
        self._visibles = list(self._filas)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.tabla = ttk.Treeview(
            self, columns=("marca", "producto"), show="", selectmode="browse",
            displaycolumns=("marca", "producto") if multiple else ("producto",),
        )
        self.tabla.column("marca", width=36, minwidth=36, stretch=False, anchor="center")
        self.tabla.column("producto", width=580, minwidth=200, stretch=True)
        self.tabla.grid(row=0, column=0, sticky="nsew")
        scroll = ctk.CTkScrollbar(
            self, command=self.tabla.yview, fg_color=colores["panel"],
            button_color=colores["acento"], button_hover_color="#10446f",
        )
        scroll.grid(row=0, column=1, sticky="ns")
        self.tabla.configure(yscrollcommand=scroll.set)
        self.estado = ctk.CTkLabel(self, text="", anchor="w", text_color=colores["secundario"])
        self.estado.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4)
        marcados = seleccionados()
        for fila, valor in self._filas.items():
            self.tabla.insert("", "end", iid=fila, values=("☑" if valor in marcados else "☐", valor))
            if not multiple and valor in marcados:
                self.tabla.selection_set(fila)
                self.tabla.focus(fila)
                self.tabla.see(fila)
        self.tabla.bind("<Button-1>", self._clic)
        self.tabla.bind("<Return>", self.activar)
        self.tabla.bind("<space>", self.activar)
        self._traza = consulta.trace_add("write", self._programar_filtro)
        self.aplicar_filtro()

    def _programar_filtro(self, *_):
        if self._pendiente is not None:
            self.after_cancel(self._pendiente)
        self._pendiente = self.after(120, self.aplicar_filtro)

    def aplicar_filtro(self):
        if self._pendiente is not None:
            self.after_cancel(self._pendiente)
            self._pendiente = None
        filtro = normalizar(self.consulta.get())
        self._visibles = [fila for fila, texto in self._busqueda if filtro in texto]
        # set_children separa las filas ocultas y reutiliza las existentes al volver.
        self.tabla.set_children("", *self._visibles)
        self.estado.configure(text=(f"{len(self._visibles)} de {len(self._filas)} productos"
                                    if self._visibles else "No se encontraron productos."))
        if self._visibles and self.tabla.focus() not in self._visibles:
            self.tabla.focus(self._visibles[0])
            self.tabla.selection_set(self._visibles[0])
        if self._visibles:
            self.tabla.see(self.tabla.focus())

    def valores_visibles(self):
        # Marcar visibles y Enter usan la consulta vigente, incluso antes de 120 ms.
        self.aplicar_filtro()
        return [self._filas[fila] for fila in self._visibles]

    def actualizar_marcas(self, filas=None):
        marcados = self.seleccionados()
        for fila in self._filas if filas is None else filas:
            self.tabla.set(fila, "marca", "☑" if self._filas[fila] in marcados else "☐")

    def enfocar(self, _evento=None):
        self.aplicar_filtro()
        self.tabla.focus_set()
        return "break"

    def _clic(self, evento):
        fila = self.tabla.identify_row(evento.y)
        if fila:
            self.tabla.focus_set()
            self.tabla.focus(fila)
            self.tabla.selection_set(fila)
            self._elegir_fila(fila)
        return "break"

    def activar(self, _evento=None):
        self.aplicar_filtro()
        fila = self.tabla.focus()
        if fila in self._visibles:
            self._elegir_fila(fila)
        return "break"

    def _elegir_fila(self, fila):
        self.elegir(self._filas[fila])
        if self.multiple:
            self.actualizar_marcas([fila])

    def destroy(self):
        if self._pendiente is not None:
            self.after_cancel(self._pendiente)
            self._pendiente = None
        self.consulta.trace_remove("write", self._traza)
        super().destroy()


class SelectorProductos(ctk.CTkFrame):
    """Selector buscable con desplazamiento CTk, sin controles de flecha nativos."""

    def __init__(self, padre, command=None):
        super().__init__(padre, fg_color="transparent")
        self.command = command
        self.valores: list[str] = []
        self.valor = ""
        self.popup = None
        self._foco_pendiente = None
        colores = ui.paleta()
        self.boton = ctk.CTkButton(
            self, text="Cargando...", anchor="w", height=30, corner_radius=6,
            fg_color=colores["panel"], hover_color=colores["seleccion"],
            text_color=colores["texto"], border_width=1, border_color=colores["borde"],
            font=(ui.PREFERENCIAS["fuente"], 12), command=self.alternar,
        )
        self.boton.pack(fill="both", expand=True)

    def establecer_valores(self, valores):
        self.valores = list(valores)
        if self.valor not in self.valores and self.valores:
            self.set(self.valores[0])

    def get(self):
        return self.valor

    def set(self, valor):
        self.valor = str(valor)
        self.boton.configure(text=self.valor)

    def alternar(self):
        if self.popup is not None and self.popup.winfo_exists():
            self.cerrar()
        else:
            self.abrir()

    def abrir(self):
        if not self.valores:
            return
        colores = ui.paleta()
        popup = self.popup = ctk.CTkToplevel(self)
        popup.withdraw()
        popup.overrideredirect(True)
        popup.transient(self.winfo_toplevel())
        popup.configure(fg_color=colores["borde"])
        popup.bind("<Escape>", lambda _evento: self.cerrar())
        popup.bind("<FocusOut>", self._revisar_foco, add="+")

        panel = ctk.CTkFrame(popup, fg_color=colores["panel"], corner_radius=7)
        panel.pack(fill="both", expand=True, padx=1, pady=1)
        cabecera = ctk.CTkFrame(panel, fg_color="transparent")
        cabecera.pack(fill="x", padx=10, pady=(10, 7))
        consulta = ctk.StringVar(master=popup)
        entrada = ctk.CTkEntry(
            cabecera, textvariable=consulta, placeholder_text="Buscar por nombre, código o número de pruebas…",
            height=34, fg_color=colores["fondo"], text_color=colores["texto"],
            border_color=colores["borde"], font=(ui.PREFERENCIAS["fuente"], 12),
        )
        entrada.pack(side="left", fill="x", expand=True)
        cerrar = ctk.CTkButton(
            cabecera, text="×", width=34, height=34, corner_radius=6,
            fg_color=colores["seleccion"], hover_color=colores["borde"],
            text_color=colores["texto"], font=(ui.PREFERENCIAS["fuente"], 17),
            command=self.cerrar,
        )
        cerrar.pack(side="left", padx=(8, 0))

        lista = self.lista = _ListaSelectorProductos(
            panel, self.valores, consulta, lambda: {self.valor}, self.elegir,
        )
        lista.pack(fill="both", expand=True, padx=(8, 5), pady=(0, 8))
        entrada.bind("<Down>", lista.enfocar)
        entrada.bind("<Return>", lista.activar)
        popup.update_idletasks()
        ancho = max(self.winfo_width(), 660)
        alto = min(470, max(300, self.winfo_screenheight() - 120))
        x = min(self.winfo_rootx(), self.winfo_screenwidth() - ancho - 12)
        debajo = self.winfo_rooty() + self.winfo_height() + 5
        y = debajo if debajo + alto <= self.winfo_screenheight() - 12 else max(12, self.winfo_rooty() - alto - 5)
        popup.geometry(f"{ancho}x{alto}+{max(0, x)}+{y}")
        popup.deiconify()
        popup.lift()
        entrada.focus_set()

    def _revisar_foco(self, _evento=None):
        if self.popup is None:
            return
        if self._foco_pendiente is not None:
            self.after_cancel(self._foco_pendiente)
        self._foco_pendiente = self.after(80, self._cerrar_si_perdio_foco)

    def _cerrar_si_perdio_foco(self):
        self._foco_pendiente = None
        if self.popup is None or not self.popup.winfo_exists():
            return
        enfocado = self.popup.focus_displayof()
        if enfocado is None or enfocado.winfo_toplevel() is not self.popup:
            self.cerrar()

    def elegir(self, valor):
        cambio = valor != self.valor
        self.set(valor)
        self.cerrar()
        if cambio and self.command:
            self.command(valor)

    def cerrar(self):
        if self._foco_pendiente is not None:
            self.after_cancel(self._foco_pendiente)
            self._foco_pendiente = None
        if self.popup is not None:
            try:
                self.popup.destroy()
            except TclError:
                pass
            self.popup = None


class SelectorMultipleProductos(ctk.CTkFrame):
    """Selección buscable de una o varias filas para la matriz."""

    def __init__(self, padre, command=None):
        super().__init__(padre, fg_color="transparent")
        self.command = command
        self.valores: list[str] = []
        self.seleccionados: set[str] = set()
        self._seleccion_al_abrir: set[str] = set()
        self.popup = None
        self._foco_pendiente = None
        colores = ui.paleta()
        self.boton = ctk.CTkButton(
            self, text="Seleccionar productos · 0", anchor="w", height=30, corner_radius=6,
            fg_color=colores["panel"], hover_color=colores["seleccion"],
            text_color=colores["texto"], border_width=1, border_color=colores["borde"],
            font=(ui.PREFERENCIAS["fuente"], 12), command=self.alternar,
        )
        self.boton.pack(fill="both", expand=True)

    def establecer_valores(self, valores):
        self.valores = list(valores)
        self.seleccionados.intersection_update(self.valores)
        self._actualizar_texto()

    def get(self):
        return set(self.seleccionados)

    def _actualizar_texto(self):
        cantidad = len(self.seleccionados)
        texto = f"Productos seleccionados · {cantidad}" if cantidad else "Seleccionar productos · 0"
        self.boton.configure(text=texto)

    def alternar(self):
        if self.popup is not None and self.popup.winfo_exists():
            self.cerrar()
        else:
            self.abrir()

    def abrir(self):
        if not self.valores:
            return
        self._seleccion_al_abrir = set(self.seleccionados)
        colores = ui.paleta()
        popup = self.popup = ctk.CTkToplevel(self)
        popup.withdraw()
        popup.overrideredirect(True)
        popup.transient(self.winfo_toplevel())
        popup.configure(fg_color=colores["borde"])
        popup.bind("<Escape>", lambda _evento: self.cerrar())
        popup.bind("<FocusOut>", self._revisar_foco, add="+")

        panel = ctk.CTkFrame(popup, fg_color=colores["panel"], corner_radius=7)
        panel.pack(fill="both", expand=True, padx=1, pady=1)
        cabecera = ctk.CTkFrame(panel, fg_color="transparent")
        cabecera.pack(fill="x", padx=10, pady=(10, 7))
        consulta = ctk.StringVar(master=popup)
        entrada = ctk.CTkEntry(
            cabecera, textvariable=consulta, placeholder_text="Buscar productos por nombre o código…",
            height=34, fg_color=colores["fondo"], text_color=colores["texto"],
            border_color=colores["borde"], font=(ui.PREFERENCIAS["fuente"], 12),
        )
        entrada.pack(fill="x", expand=True)
        def alternar_valor(valor):
            if valor in self.seleccionados:
                self.seleccionados.discard(valor)
            else:
                self.seleccionados.add(valor)
            self._actualizar_texto()

        lista = self.lista = _ListaSelectorProductos(
            panel, self.valores, consulta, self.get, alternar_valor, multiple=True,
        )
        lista.pack(fill="both", expand=True, padx=(8, 5), pady=(0, 6))
        entrada.bind("<Down>", lista.enfocar)
        entrada.bind("<Return>", lista.activar)
        acciones = ctk.CTkFrame(panel, fg_color="transparent")
        acciones.pack(fill="x", padx=10, pady=(0, 10))
        ui.Boton(acciones, text="Marcar visibles", variante="secundario", fondo=colores["panel"],
                 command=lambda: self._marcar_visibles(True)).pack(side="left")
        ui.Boton(acciones, text="Quitar visibles", variante="secundario", fondo=colores["panel"],
                 command=lambda: self._marcar_visibles(False)).pack(side="left", padx=8)
        ui.Boton(acciones, text="Listo", fondo=colores["panel"], command=self.cerrar).pack(side="right")
        ctk.CTkLabel(panel, text="El análisis se actualiza al cerrar este selector.",
                     text_color=colores["secundario"]).pack(fill="x", padx=10, pady=(0, 6))

        popup.update_idletasks()
        ancho = max(self.winfo_width(), 700)
        alto = min(500, max(320, self.winfo_screenheight() - 120))
        x = min(self.winfo_rootx(), self.winfo_screenwidth() - ancho - 12)
        debajo = self.winfo_rooty() + self.winfo_height() + 5
        y = debajo if debajo + alto <= self.winfo_screenheight() - 12 else max(12, self.winfo_rooty() - alto - 5)
        popup.geometry(f"{ancho}x{alto}+{max(0, x)}+{y}")
        popup.deiconify()
        popup.lift()
        entrada.focus_set()

    def _marcar_visibles(self, marcar):
        visibles = self.lista.valores_visibles()
        if marcar:
            self.seleccionados.update(visibles)
        else:
            self.seleccionados.difference_update(visibles)
        self.lista.actualizar_marcas()
        self._actualizar_texto()

    def _notificar(self):
        self._actualizar_texto()
        if self.command:
            self.command(set(self.seleccionados))

    def _revisar_foco(self, _evento=None):
        if self.popup is not None:
            if self._foco_pendiente is not None:
                self.after_cancel(self._foco_pendiente)
            self._foco_pendiente = self.after(80, self._cerrar_si_perdio_foco)

    def _cerrar_si_perdio_foco(self):
        self._foco_pendiente = None
        if self.popup is None or not self.popup.winfo_exists():
            return
        enfocado = self.popup.focus_displayof()
        if enfocado is None or enfocado.winfo_toplevel() is not self.popup:
            self.cerrar()

    def cerrar(self):
        if self._foco_pendiente is not None:
            self.after_cancel(self._foco_pendiente)
            self._foco_pendiente = None
        if self.popup is not None:
            try:
                self.popup.destroy()
            except TclError:
                pass
            self.popup = None
            if self.seleccionados != self._seleccion_al_abrir:
                self._notificar()


class VentanaAnalisis:
    def __init__(self, padre, carpeta_json: Path):
        self.carpeta_json = Path(carpeta_json)
        self.productos: list[Producto] = []
        self.comparaciones: list[Comparacion] = []
        self.comparaciones_multidimensionales: list[ComparacionMultidimensional] = []
        self.variables_matriz: list[tuple[str, str, str]] = []
        self.matriz_valores: dict[str, dict[tuple[str, str, str], object | None]] = {}
        self.ventana = ui.Ventana(padre)
        self.ventana.title("ANÁLISIS · Cobertura de especificaciones")
        self.ventana.geometry("1260x800")
        self.ventana.minsize(980, 620)
        self.ventana.configure(fg_color=ui.paleta()["fondo"])

        from menu_principal import BarraPestanas
        self.pestanas = BarraPestanas(
            self.ventana, fg_color=("#ffffff", "#202c3d"),
            segmented_button_selected_color=ui.paleta()["acento"],
            segmented_button_unselected_color=("#425a77", "#29384d"),
        )
        self.pestanas.pack(fill="both", expand=True, padx=18, pady=18)
        primera = self.pestanas.add("Productos equivalentes")
        segunda = self.pestanas.add("Matriz producto × prueba")
        tercera = self.pestanas.add("Comparación multidimensional")
        self.pestanas.set("Productos equivalentes")
        self._construir_primera(primera)
        self._construir_segunda(segunda)
        self._construir_tercera(tercera)
        self.actualizar_catalogo()

    def _construir_primera(self, pagina):
        controles = ctk.CTkFrame(pagina, fg_color="transparent")
        controles.pack(fill="x", padx=18, pady=(14, 8))
        ctk.CTkLabel(controles, text="Producto consultado:").pack(side="left", padx=(0, 8))
        self.selector = SelectorProductos(controles, command=lambda _: self.actualizar_resultados())
        self.selector.pack(side="left", fill="x", expand=True, padx=(0, 12))
        fondo_controles = ui.paleta()["panel"]
        ui.Boton(
            controles, text="Actualizar catálogo", command=self.actualizar_catalogo,
            width=22, fondo=fondo_controles,
        ).pack(side="left")
        informacion = ui.Boton(
            controles, text="ⓘ", cuadrado=True, height=30, corner_radius=6,
            fondo=fondo_controles, font=(ui.PREFERENCIAS["fuente"], 15, "bold"),
            command=lambda: messagebox.showinfo(
                "Cómo se calcula la cobertura", TEXTO_AYUDA_COBERTURA, parent=self.ventana
            ),
        )
        informacion.pack(side="left", padx=(8, 0))
        ui.AyudaHover(informacion, TEXTO_AYUDA_COBERTURA)
        self.estado = ctk.CTkLabel(pagina, text="", anchor="w", justify="left", wraplength=1110,
                                   text_color=ui.paleta()["secundario"])
        self.estado.pack(fill="x", padx=18, pady=(0, 10))

        contenedor = ctk.CTkFrame(pagina, fg_color="transparent")
        contenedor.pack(fill="both", expand=True, padx=18, pady=(0, 9))
        columnas = ("producto", "dictamen", "cobertura", "cumple", "tipico", "parcial", "fuera", "faltan")
        self.tabla = ttk.Treeview(contenedor, columns=columnas, show="headings", height=11)
        for clave, titulo, ancho in (
            ("producto", "Producto candidato", 300), ("dictamen", "Resultado", 180),
            ("cobertura", "Cobertura", 95), ("cumple", "Cubre", 75),
            ("tipico", "Típico", 75), ("parcial", "Parcial", 75),
            ("fuera", "Fuera", 70), ("faltan", "Faltan", 80),
        ):
            self.tabla.heading(clave, text=titulo)
            self.tabla.column(clave, width=ancho, minwidth=65, stretch=(clave == "producto"))
        scroll = ctk.CTkScrollbar(
            contenedor, orientation="vertical", command=self.tabla.yview,
            fg_color=ui.paleta()["panel"], button_color=ui.paleta()["acento"],
            button_hover_color="#10446f", width=14,
        )
        self.tabla.configure(yscrollcommand=scroll.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tabla.bind("<<TreeviewSelect>>", self.mostrar_detalle)
        self.detalle = ctk.CTkTextbox(pagina, height=230, wrap="word",
                                      fg_color=ui.paleta()["fondo"], text_color=ui.paleta()["texto"],
                                      border_color=ui.paleta()["borde"], border_width=1)
        self.detalle.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        self._escribir_detalle("Selecciona un candidato para revisar cada límite y valor típico.")
        ctk.CTkLabel(
            pagina, anchor="w", justify="left", wraplength=1110,
            text="Cobertura = requisitos del consultado garantizados por los rangos declarados del candidato, "
                 "ponderando ×3 las pruebas marcadas como críticas. Comparar especificaciones no sustituye una validación de uso.",
            text_color=ui.paleta()["secundario"],
        ).pack(fill="x", padx=18, pady=(0, 12))

    def _construir_segunda(self, pagina):
        colores = ui.paleta()
        controles = ctk.CTkFrame(pagina, fg_color="transparent")
        controles.pack(fill="x", padx=18, pady=(14, 8))
        self.modo_matriz = ctk.StringVar(master=pagina, value="Valores típicos")
        modos = ctk.CTkSegmentedButton(
            controles, values=("Valores típicos", "Disponibilidad 1/0"),
            variable=self.modo_matriz, command=lambda _valor: self._actualizar_vista_matriz(),
            selected_color=colores["acento"], selected_hover_color="#10446f",
            unselected_color=colores["seleccion"], unselected_hover_color=colores["borde"],
            text_color=colores["texto"], height=30, border_width=0,
        )
        modos.pack(side="left")
        self.selector_productos_matriz = SelectorMultipleProductos(
            controles, command=lambda _seleccion: self._actualizar_vista_matriz()
        )
        self.selector_productos_matriz.pack(side="left", fill="x", expand=True, padx=12)
        self.boton_quitar_matriz = ui.Boton(
            controles, text="Quitar de matriz", width=18, variante="secundario",
            fondo=colores["panel"], state="disabled", command=self._quitar_productos_matriz,
        )
        self.boton_quitar_matriz.pack(side="left", padx=(0, 8))
        informacion = ui.Boton(
            controles, text="ⓘ", cuadrado=True, height=30, corner_radius=6,
            fondo=colores["panel"], font=(ui.PREFERENCIAS["fuente"], 15, "bold"),
            command=lambda: messagebox.showinfo(
                "Cómo se construye la matriz", TEXTO_AYUDA_MATRIZ, parent=self.ventana
            ),
        )
        informacion.pack(side="left")
        ui.AyudaHover(informacion, TEXTO_AYUDA_MATRIZ)

        filtros = ctk.CTkFrame(pagina, fg_color="transparent")
        filtros.pack(fill="x", padx=18, pady=(0, 8))
        self.busqueda_prueba = ctk.StringVar(master=pagina)
        buscador = ctk.CTkEntry(
            filtros, textvariable=self.busqueda_prueba,
            placeholder_text="Filtrar columnas por prueba, unidad o método…",
            height=30, fg_color=colores["panel"], text_color=colores["texto"],
            border_color=colores["borde"], font=(ui.PREFERENCIAS["fuente"], 12),
        )
        buscador.pack(fill="x", expand=True)
        self.busqueda_prueba.trace_add("write", lambda *_: self._actualizar_vista_matriz())

        self.estado_matriz = ctk.CTkLabel(
            pagina, text="Preparando matriz…", anchor="w", justify="left",
            text_color=colores["secundario"],
        )
        self.estado_matriz.pack(fill="x", padx=18, pady=(0, 10))

        contenedor = ctk.CTkFrame(pagina, fg_color="transparent")
        contenedor.pack(fill="both", expand=True, padx=18, pady=(0, 14))
        contenedor.grid_rowconfigure(0, weight=1)
        contenedor.grid_columnconfigure(0, weight=1)
        self.tabla_matriz = ttk.Treeview(contenedor, show="headings", selectmode="extended")
        scroll_vertical = ctk.CTkScrollbar(
            contenedor, orientation="vertical", command=self.tabla_matriz.yview,
            fg_color=colores["panel"], button_color=colores["acento"],
            button_hover_color="#10446f", width=14,
        )
        scroll_horizontal = ctk.CTkScrollbar(
            contenedor, orientation="horizontal", command=self.tabla_matriz.xview,
            fg_color=colores["panel"], button_color=colores["acento"],
            button_hover_color="#10446f", height=14,
        )
        self.tabla_matriz.configure(
            yscrollcommand=scroll_vertical.set, xscrollcommand=scroll_horizontal.set,
        )
        self.tabla_matriz.grid(row=0, column=0, sticky="nsew")
        scroll_vertical.grid(row=0, column=1, sticky="ns", padx=(4, 0))
        scroll_horizontal.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.tabla_matriz.bind("<Shift-MouseWheel>", self._desplazar_matriz_horizontal, add="+")
        self.tabla_matriz.bind("<<TreeviewSelect>>", self._actualizar_boton_quitar_matriz, add="+")
        self.tabla_matriz.bind("<Delete>", lambda _evento: self._quitar_productos_matriz(), add="+")

    def _construir_tercera(self, pagina):
        colores = ui.paleta()
        controles = ctk.CTkFrame(pagina, fg_color="transparent")
        controles.pack(fill="x", padx=18, pady=(14, 8))
        ctk.CTkLabel(controles, text="Categoría o grupo:").pack(side="left", padx=(0, 8))
        self.selector_grupo_multidimensional = SelectorMultipleProductos(
            controles, command=lambda _seleccion: self._actualizar_multidimensional()
        )
        self.selector_grupo_multidimensional.pack(side="left", fill="x", expand=True, padx=(0, 8))
        informacion = ui.Boton(
            controles, text="ⓘ", cuadrado=True, height=30, corner_radius=6,
            fondo=colores["panel"], font=(ui.PREFERENCIAS["fuente"], 15, "bold"),
            command=lambda: messagebox.showinfo(
                "Cómo funciona la comparación multidimensional",
                TEXTO_AYUDA_MULTIDIMENSIONAL, parent=self.ventana,
            ),
        )
        informacion.pack(side="left")
        ui.AyudaHover(informacion, TEXTO_AYUDA_MULTIDIMENSIONAL)
        self.estado_multidimensional = ctk.CTkLabel(
            pagina, text="Selecciona al menos dos productos de la misma categoría.",
            anchor="w", justify="left", text_color=colores["secundario"],
        )
        self.estado_multidimensional.pack(fill="x", padx=18, pady=(0, 10))

        contenedor = ctk.CTkFrame(pagina, fg_color="transparent")
        contenedor.pack(fill="both", expand=True, padx=18, pady=(0, 8))
        contenedor.grid_rowconfigure(0, weight=1)
        contenedor.grid_columnconfigure(0, weight=1)
        columnas = ("a", "b", "compartidas", "superposicion", "distancia",
                    "ventaja_a", "ventaja_b", "diferencias", "exclusivas", "confianza")
        self.tabla_multidimensional = ttk.Treeview(
            contenedor, columns=columnas, show="headings", height=9, selectmode="browse"
        )
        for clave, titulo, ancho in (
            ("a", "Producto A", 230), ("b", "Producto B", 230),
            ("compartidas", "Compartidas", 90), ("superposicion", "Superposición", 100),
            ("distancia", "Distancia", 85), ("ventaja_a", "Ventajas A", 85),
            ("ventaja_b", "Ventajas B", 85), ("diferencias", "Diferencias", 90),
            ("exclusivas", "Exclusivas A/B", 105), ("confianza", "Confianza", 90),
        ):
            self.tabla_multidimensional.heading(clave, text=titulo)
            self.tabla_multidimensional.column(
                clave, width=ancho, minwidth=70, stretch=clave in {"a", "b"}, anchor="w" if clave in {"a", "b"} else "center"
            )
        scroll_vertical = ctk.CTkScrollbar(
            contenedor, orientation="vertical", command=self.tabla_multidimensional.yview,
            fg_color=colores["panel"], button_color=colores["acento"],
            button_hover_color="#10446f", width=14,
        )
        scroll_horizontal = ctk.CTkScrollbar(
            contenedor, orientation="horizontal", command=self.tabla_multidimensional.xview,
            fg_color=colores["panel"], button_color=colores["acento"],
            button_hover_color="#10446f", height=14,
        )
        self.tabla_multidimensional.configure(
            yscrollcommand=scroll_vertical.set, xscrollcommand=scroll_horizontal.set,
        )
        self.tabla_multidimensional.grid(row=0, column=0, sticky="nsew")
        scroll_vertical.grid(row=0, column=1, sticky="ns", padx=(4, 0))
        scroll_horizontal.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.tabla_multidimensional.bind("<<TreeviewSelect>>", self._mostrar_detalle_multidimensional)
        self.tabla_multidimensional.bind(
            "<Shift-MouseWheel>",
            lambda evento: self._desplazar_tabla_horizontal(self.tabla_multidimensional, evento), add="+",
        )
        self.detalle_multidimensional = ctk.CTkTextbox(
            pagina, height=250, wrap="word", fg_color=colores["fondo"],
            text_color=colores["texto"], border_color=colores["borde"], border_width=1,
        )
        self.detalle_multidimensional.pack(fill="both", expand=True, padx=18, pady=(0, 14))
        self._escribir_detalle_multidimensional(
            "Selecciona una categoría o grupo de productos para comparar todas sus parejas."
        )

    @staticmethod
    def _desplazar_tabla_horizontal(tabla, evento):
        tabla.xview_scroll((-1 if evento.delta > 0 else 1) * 3, "units")
        return "break"

    def _actualizar_multidimensional(self):
        seleccion = self.selector_grupo_multidimensional.get()
        productos = [producto for producto in self.productos if producto.etiqueta in seleccion]
        self.tabla_multidimensional.delete(*self.tabla_multidimensional.get_children())
        self.comparaciones_multidimensionales = []
        if len(productos) < 2:
            self.estado_multidimensional.configure(
                text=f"{len(productos)} producto seleccionado. Selecciona al menos dos de la misma categoría."
            )
            self._escribir_detalle_multidimensional(
                "La comparación se genera cuando el grupo contiene dos o más productos."
            )
            return
        resultados = [comparar_multidimensional(a, b) for a, b in combinations(productos, 2)]
        self.comparaciones_multidimensionales = sorted(
            resultados,
            key=lambda r: (r.confianza_puntaje, r.superposicion,
                           -(r.distancia if r.distancia is not None else 2.0)),
            reverse=True,
        )
        for indice, resultado in enumerate(self.comparaciones_multidimensionales):
            distancia = f"{resultado.distancia:.1%}" if resultado.distancia is not None else "N/D"
            self.tabla_multidimensional.insert("", "end", iid=str(indice), values=(
                f"{resultado.producto_a.nombre} · {resultado.producto_a.codigo}",
                f"{resultado.producto_b.nombre} · {resultado.producto_b.codigo}",
                f"{resultado.compartidas}/{resultado.total_union}",
                f"{resultado.superposicion:.1%}", distancia,
                len(resultado.ventajas_a), len(resultado.ventajas_b),
                len(resultado.diferencias),
                f"{len(resultado.exclusivas_a)}/{len(resultado.exclusivas_b)}",
                resultado.confianza,
            ))
        parejas = len(self.comparaciones_multidimensionales)
        self.estado_multidimensional.configure(
            text=f"{len(productos)} productos en el grupo · {parejas} comparaciones por parejas. "
                 "Una distancia menor indica mayor cercanía técnica."
        )
        primera = self.tabla_multidimensional.get_children()[0]
        self.tabla_multidimensional.selection_set(primera)
        self.tabla_multidimensional.focus(primera)
        self._mostrar_detalle_multidimensional()

    def _escribir_detalle_multidimensional(self, texto):
        self.detalle_multidimensional.configure(state="normal")
        self.detalle_multidimensional.delete("1.0", "end")
        self.detalle_multidimensional.insert("1.0", texto)
        self.detalle_multidimensional.configure(state="disabled")

    def _mostrar_detalle_multidimensional(self, _evento=None):
        seleccion = self.tabla_multidimensional.selection()
        if not seleccion:
            return
        resultado = self.comparaciones_multidimensionales[int(seleccion[0])]
        a, b = resultado.producto_a, resultado.producto_b
        distancia = f"{resultado.distancia:.1%}" if resultado.distancia is not None else "no calculable"
        lineas = [
            f"PRODUCTO A: {a.nombre} ({a.codigo})",
            f"PRODUCTO B: {b.nombre} ({b.codigo})",
            "",
            f"Comparten {resultado.compartidas} de {resultado.total_union} variables técnicas "
            f"({resultado.superposicion:.1%} de superposición).",
            f"Distancia técnica: {distancia}, basada en {resultado.evaluables} pruebas evaluables.",
            f"Confianza: {resultado.confianza} ({resultado.confianza_puntaje:.0%}).",
        ]

        def agregar_seccion(titulo, elementos, vacio):
            lineas.extend(("", titulo))
            lineas.extend(f"• {texto}" for texto in elementos)
            if not elementos:
                lineas.append(f"• {vacio}")

        agregar_seccion(
            "VENTAJAS DE A SOBRE B", resultado.ventajas_a,
            "No hay ventajas direccionales demostrables con los datos disponibles.",
        )
        agregar_seccion(
            "VENTAJAS DE B SOBRE A", resultado.ventajas_b,
            "No hay ventajas direccionales demostrables con los datos disponibles.",
        )
        agregar_seccion(
            "DIFERENCIAS RELEVANTES", resultado.diferencias,
            "No se detectaron diferencias evaluables mayores al 10 %.",
        )
        agregar_seccion("PRUEBAS EXCLUSIVAS DE A", resultado.exclusivas_a, "Ninguna.")
        agregar_seccion("PRUEBAS EXCLUSIVAS DE B", resultado.exclusivas_b, "Ninguna.")
        lineas.extend(("", "ESCENARIOS DE USO"))
        if resultado.ventajas_a:
            prioridades = ", ".join(texto.split(":", 1)[0] for texto in resultado.ventajas_a[:4])
            lineas.append(f"• A puede convenir cuando se priorizan: {prioridades}.")
        if resultado.ventajas_b:
            prioridades = ", ".join(texto.split(":", 1)[0] for texto in resultado.ventajas_b[:4])
            lineas.append(f"• B puede convenir cuando se priorizan: {prioridades}.")
        if not resultado.ventajas_a and not resultado.ventajas_b:
            lineas.append("• Los datos no justifican una preferencia direccional entre A y B.")
        lineas.append(
            "• Las propiedades con rango objetivo dependen de la aplicación; deben validarse contra el escenario real."
        )
        self._escribir_detalle_multidimensional("\n".join(lineas))

    def _desplazar_matriz_horizontal(self, evento):
        pasos = -1 if evento.delta > 0 else 1
        self.tabla_matriz.xview_scroll(pasos * 3, "units")
        return "break"

    def _actualizar_boton_quitar_matriz(self, _evento=None):
        estado = "normal" if self.tabla_matriz.selection() else "disabled"
        self.boton_quitar_matriz.configure(state=estado)

    def _quitar_productos_matriz(self):
        filas = self.tabla_matriz.selection()
        etiquetas = {
            self._filas_productos_matriz[fila]
            for fila in filas if fila in self._filas_productos_matriz
        }
        if not etiquetas:
            return
        self.selector_productos_matriz.seleccionados.difference_update(etiquetas)
        self.selector_productos_matriz._notificar()

    @staticmethod
    def _valor_matriz(fila):
        if len(filas_grupo(fila)) > 1:
            valores = [f.get("valor_tipico") for f in filas_grupo(fila)]
            if all(normalizar(valor) in SIN_DATO for valor in valores):
                return None
            return "Revisar varias filas: " + " | ".join(
                "N/D" if normalizar(valor) in SIN_DATO else str(valor) for valor in valores)
        valor = fila.get("valor_tipico")
        if normalizar(valor) in SIN_DATO:
            return None
        return valor

    def _reconstruir_matriz(self):
        frecuencia = Counter(clave for producto in self.productos for clave in producto.pruebas)
        self.variables_matriz = sorted(
            frecuencia, key=lambda clave: (-frecuencia[clave], clave[0], clave[1], clave[2])
        )
        self.matriz_valores = {
            producto.etiqueta: {
                clave: self._valor_matriz(fila)
                for clave, fila in producto.pruebas.items()
            }
            for producto in self.productos
        }
        self._actualizar_vista_matriz()

    def _actualizar_vista_matriz(self):
        if not hasattr(self, "tabla_matriz"):
            return
        seleccion = self.selector_productos_matriz.get()
        productos_visibles = [producto for producto in self.productos if producto.etiqueta in seleccion]
        disponibles = {
            clave for producto in productos_visibles for clave in producto.pruebas
        }
        filtro = normalizar(self.busqueda_prueba.get())
        variables = [
            clave for clave in self.variables_matriz
            if clave in disponibles and (not filtro or filtro in normalizar(" ".join(clave)))
        ]
        columnas = ("producto", *(f"v{indice}" for indice in range(len(variables))))
        self.tabla_matriz.delete(*self.tabla_matriz.get_children())
        self._filas_productos_matriz = {}
        self.boton_quitar_matriz.configure(state="disabled")
        self.tabla_matriz.configure(columns=columnas)
        self.tabla_matriz.heading("producto", text="Producto")
        self.tabla_matriz.column("producto", width=285, minwidth=210, stretch=False, anchor="w")
        for indice, (nombre, unidad, metodo) in enumerate(variables):
            columna = f"v{indice}"
            self.tabla_matriz.heading(columna, text=f"{nombre}\n{unidad} · {metodo}")
            self.tabla_matriz.column(columna, width=210, minwidth=120, stretch=False, anchor="center")
        binaria = self.modo_matriz.get() == "Disponibilidad 1/0"
        for fila, producto in enumerate(productos_visibles):
            valores_producto = self.matriz_valores.get(producto.etiqueta, {})
            valores = [
                (1 if valor is not None else 0) if binaria else (valor if valor is not None else "N/D")
                for valor in (valores_producto.get(clave) for clave in variables)
            ]
            identificador = f"producto_{fila}"
            self._filas_productos_matriz[identificador] = producto.etiqueta
            self.tabla_matriz.insert(
                "", "end", iid=identificador,
                values=(f"{producto.nombre} · {producto.codigo}", *valores),
                tags=("fila_alterna",) if fila % 2 else (),
            )
        modo = "disponibilidad" if binaria else "valores típicos"
        if productos_visibles:
            texto_estado = (
                f"{len(productos_visibles)} productos seleccionados × {len(disponibles)} variables técnicas · "
                f"{len(variables)} columnas visibles · vista de {modo}. "
                "N/D se conserva internamente como valor nulo."
            )
        else:
            texto_estado = "Selecciona uno o varios productos para construir la matriz."
        self.estado_matriz.configure(text=texto_estado)

    def _escribir_detalle(self, texto):
        self.detalle.configure(state="normal")
        self.detalle.delete("1.0", "end")
        self.detalle.insert("1.0", texto)
        self.detalle.configure(state="disabled")

    def actualizar_catalogo(self):
        anterior = self.selector.get()
        try:
            productos, descartados = cargar_catalogo(self.carpeta_json)
        except OSError as exc:
            self.estado.configure(text=str(exc))
            return
        self.productos = productos
        etiquetas = [p.etiqueta for p in productos]
        self.selector.establecer_valores(etiquetas or ["Sin productos disponibles"])
        self.selector_productos_matriz.establecer_valores(etiquetas)
        self.selector_grupo_multidimensional.establecer_valores(etiquetas)
        inicial = (anterior if anterior in etiquetas else
                   max(productos, key=lambda p: p.numero_pruebas).etiqueta if productos else "Sin productos disponibles")
        self.selector.set(inicial)
        self.estado.configure(text=f"Catálogo completo: {len(productos)} productos únicos de JSON alineados. "
                                   f"{descartados} archivos omitidos. Selecciona un candidato para ver qué cubre.")
        self._reconstruir_matriz()
        self._actualizar_multidimensional()
        self.actualizar_resultados()

    def actualizar_resultados(self):
        self.tabla.delete(*self.tabla.get_children())
        self.comparaciones = []
        origen = next((p for p in self.productos if p.etiqueta == self.selector.get()), None)
        if origen is None:
            self._escribir_detalle("No hay productos alineados disponibles para comparar.")
            return
        resultados = (comparar(origen, p) for p in self.productos if p is not origen)
        self.comparaciones = sorted((r for r in resultados if r.cobertura > 0),
                                    key=lambda r: (r.cobertura, r.cubiertos, r.indicativos,
                                                   -r.fuera, -r.faltantes), reverse=True)
        for indice, resultado in enumerate(self.comparaciones):
            self.tabla.insert("", "end", iid=str(indice), values=(
                f"{resultado.candidato.nombre} · {resultado.candidato.codigo}",
                resultado.dictamen, "<1%" if resultado.cobertura < 0.005 else f"{resultado.cobertura:.0%}",
                f"{resultado.cubiertos}/{resultado.requisitos}", resultado.indicativos,
                resultado.parciales, resultado.fuera, resultado.faltantes + resultado.sin_dato,
            ))
        if self.comparaciones:
            primero = self.tabla.get_children()[0]
            self.tabla.selection_set(primero)
            self.tabla.focus(primero)
            self.mostrar_detalle()
        else:
            self._escribir_detalle("No hay candidatos con cobertura demostrada para este producto.")

    def mostrar_detalle(self, _evento=None):
        seleccion = self.tabla.selection()
        if not seleccion:
            return
        origen = next((p for p in self.productos if p.etiqueta == self.selector.get()), None)
        if origen is None:
            return
        resultado = self.comparaciones[int(seleccion[0])]
        candidato = resultado.candidato
        lineas = [
            f"CONSULTADO: {origen.nombre} ({origen.codigo})",
            f"CANDIDATO: {candidato.nombre} ({candidato.codigo})",
            f"RESULTADO: {resultado.dictamen} · cobertura ponderada {resultado.cobertura:.1%}",
            f"Requisitos cubiertos {resultado.cubiertos}/{resultado.requisitos}; "
            f"típico indicativo {resultado.indicativos}; parcial {resultado.parciales}; "
            f"fuera {resultado.fuera}; faltan/no verificables "
            f"{resultado.faltantes + resultado.sin_dato}.",
            f"Pruebas alineadas compartidas: {resultado.compartidas}/{origen.numero_pruebas}. "
            "El valor típico no garantiza el cumplimiento de todos los lotes.",
            f"Categoría: {origen.categoria or 'no documentada'} / {candidato.categoria or 'no documentada'}. "
            f"Uso: {origen.uso or 'no documentado'} / {candidato.uso or 'no documentado'}.",
            "",
        ]
        secciones = (
            ("CUBRE", "REQUISITOS CUBIERTOS"),
            ("INDICATIVO", "TÍPICO DENTRO, SIN GARANTÍA DE ESPECIFICACIÓN"),
            ("PARCIAL", "RANGOS QUE SE CRUZAN, SIN GARANTÍA"),
            ("FUERA", "VALORES FUERA DEL REQUISITO"),
            ("INCONSISTENTE", "DATOS INTERNAMENTE CONTRADICTORIOS"),
            ("AUSENTE", "PRUEBAS FALTANTES"),
            ("NO_COMPARABLE", "UNIDAD O MÉTODO DISTINTO"),
            ("SIN_DATO", "SIN DATOS EVALUABLES"),
            ("SIN_REQUISITO", "PRUEBAS SIN LÍMITE EN EL CONSULTADO"),
        )
        for estado, titulo in secciones:
            grupo = [fila for fila in resultado.filas if fila.estado == estado]
            if not grupo:
                continue
            lineas.extend((titulo, ""))
            for fila in grupo:
                fuente = fila.origen
                rival = fila.candidato or {}
                lineas.append(f"• {fila.nombre} [{fuente.get('unidades', '')}]"
                              f"{' · CRÍTICA' if peso(fuente) > 1 else ''}")
                lineas.append(f"  Requerido: {fuente.get('especificacion') or 'sin dato'}  |  "
                              f"Candidato: {rival.get('especificacion') or 'sin dato'}")
                lineas.append(f"  Típico candidato: {rival.get('valor_tipico') or 'sin dato'}  |  {fila.motivo}")
            lineas.append("")
        self._escribir_detalle("\n".join(lineas))
