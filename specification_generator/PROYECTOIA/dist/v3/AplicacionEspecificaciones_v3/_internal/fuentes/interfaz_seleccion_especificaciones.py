#!/usr/bin/env python3
"""Selector con busqueda y vista previa de especificaciones en espanol."""
from __future__ import annotations
import estilo_interfaz as ui
from rutas_proyecto import BASE, resolver_ruta, rutas_configuradas, ruta_maestro
from especificaciones_wip import es_wip
from cambios_alineacion import cambios_documento, ResaltadorCambios


import json
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import unicodedata
import webbrowser
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from tkinter import messagebox
from estilo_interfaz import ttk

from typing import Any, Callable

from alinear_especificaciones_json import (
    MAESTRO,
    VERSION as VERSION_ALINEADOR,
    cargar_maestro,
    ejecutar as ejecutar_alineacion,
    metodos_maestros,
    referencias_astm,
    referencias_astm_fuera_maestro,
)
from generador_especs_hii import (
    _separar_notas_de_pruebas,
    clave_texto,
    detectar_json,
    leer_datos_json,
    leer_idioma_datos,
    preparar_control_generacion,
)


JSON_DIR_ANTERIORES = BASE / "ESPECIFICACIONES_JSON"
JSON_DIR_NUEVAS = BASE / "ESPECIFICACIONES_JSON_ALINEADAS"

CAMPOS_CONTROL = (
    ("Elaboró y aprobó", "elaboro_y_aprobo"),
    ("Folio APQP", "folio_apqp"),
    ("Folio VAP", "folio_vap"),
    ("Código único de producto", "codigo_unico_producto"),
    ("Fecha", "fecha"),
    ("Revisión", "revision"),
    ("Nombre del cliente", "nombre_cliente"),
    ("Revisión del cliente", "revision_cliente"),
)
CAMPOS_PRUEBA = (
    ("Prueba", "prueba"),
    ("Unidades", "unidades"),
    ("Método", "metodo"),
    ("Especificación", "especificacion"),
    ("Valor típico", "valor_tipico"),
)
CAMPOS_PRUEBA_INGLES = (
    ("Test", "prueba"),
    ("Units", "unidades"),
    ("Method", "metodo"),
    ("Specification", "especificacion"),
    ("Typical value", "valor_tipico"),
)
CAMPOS_CONTROL_INGLES = {
    "elaboro_y_aprobo": "Prepared and approved by",
    "folio_apqp": "APQP number",
    "folio_vap": "VAP number",
    "codigo_unico_producto": "Unique product code",
    "fecha": "Date",
    "revision": "Revision",
    "nombre_cliente": "Customer name",
    "revision_cliente": "Customer revision",
}


@dataclass(frozen=True)
class Especificacion:
    ruta: Path
    nombre_producto: str
    codigo: str
    cliente: str
    revision: str
    busqueda: str
    idioma: str = "ESPANOL"
    version_alineador: str = ""
    versiones_idioma: tuple[Path, ...] = ()
    edicion_bilingue: bool = False
    en_progreso: bool = False

    @property
    def etiqueta(self) -> str:
        identidad = " — ".join(valor for valor in (self.codigo, self.nombre_producto) if valor)
        detalles = []
        if self.revision:
            detalles.append(f"Rev. {self.revision}")
        if self.cliente:
            detalles.append(self.cliente)
        if detalles:
            identidad = f"{identidad}  |  {' · '.join(detalles)}"
        return identidad or self.ruta.name


@dataclass(frozen=True)
class Catalogo:
    especificaciones: tuple[Especificacion, ...]
    errores: tuple[str, ...]


@dataclass(frozen=True)
class VistaPrevia:
    nombre_producto: str
    control: tuple[tuple[str, str], ...]
    pruebas: tuple[dict[str, Any], ...]
    notas: tuple[str, ...]
    mensajes_astm_no_encontradas: tuple[str, ...]
    mensajes_control: tuple[str, ...]
    pie_pagina: str
    idioma: str = "ESPANOL"


def encabezados_tabla(idioma: str) -> tuple[tuple[str, str], ...]:
    return CAMPOS_PRUEBA_INGLES if idioma == "INGLES" else CAMPOS_PRUEBA


class OrdenadorColumnasTreeview:
    def __init__(self, tabla: ttk.Treeview, encabezados: dict[str, str]):
        self.tabla = tabla
        self.encabezados = dict(encabezados)
        self.orden_original: tuple[str, ...] = ()
        self.columna_ordenada: str | None = None
        self._firma_encabezados = None
        self.registrar_orden_original()
        self._actualizar_encabezados()
        self.tabla.bind("<Configure>", self._al_cambiar_tamano, add="+")
        self.tabla.bind("<ButtonRelease-1>", self._al_cambiar_tamano, add="+")

    def establecer_encabezados(self, encabezados: dict[str, str]) -> None:
        self.encabezados = dict(encabezados)
        self._actualizar_encabezados()

    def registrar_orden_original(
        self, elementos: tuple[str, ...] | list[str] | None = None
    ) -> None:
        self.orden_original = tuple(
            self.tabla.get_children("") if elementos is None else elementos
        )
        if self.columna_ordenada is not None:
            self._aplicar_orden(self.columna_ordenada)

    def _actualizar_encabezados(self) -> None:
        estilo = ttk.Style(self.tabla)
        estilo_tabla = str(self.tabla.cget("style")) or "Treeview"
        fuente_estilo = estilo.lookup(f"{estilo_tabla}.Heading", "font") or "TkDefaultFont"
        firma = (tuple(self.encabezados.items()), self.columna_ordenada, str(fuente_estilo),
                 tuple(int(self.tabla.column(c, "width")) for c in self.encabezados),
                 str(self.tabla.tk.call("tk", "scaling")) if hasattr(self.tabla, "tk") else None)
        if firma == self._firma_encabezados:
            return
        self._firma_encabezados = firma
        fuente = tkfont.Font(
            self.tabla,
            font=fuente_estilo,
        )
        ancho_espacio = max(1, fuente.measure(" "))
        for columna, titulo in self.encabezados.items():
            indicador = "▲" if self.columna_ordenada == columna else "▼"
            ancho = int(self.tabla.column(columna, "width"))
            hueco = max(
                1,
                round(
                    (
                        ancho / 2
                        - fuente.measure(titulo) / 2
                        - fuente.measure(indicador)
                        - 6
                    )
                    / ancho_espacio
                ),
            )
            self.tabla.heading(
                columna,
                text=f"{titulo}{' ' * hueco}{indicador}",
                anchor="e",
                command=lambda c=columna: self.alternar(c),
            )

    def _al_cambiar_tamano(self, _evento: tk.Event | None = None) -> None:
        if self.encabezados and self.tabla.winfo_exists():
            self._actualizar_encabezados()

    def _aplicar_orden(self, columna: str) -> None:
        elementos = self.tabla.get_children("")

        def clave_alfabetica(elemento: str) -> str:
            normalizado = unicodedata.normalize(
                "NFKD", str(self.tabla.set(elemento, columna))
            )
            sin_acentos = "".join(
                caracter
                for caracter in normalizado
                if unicodedata.category(caracter) != "Mn"
            )
            return sin_acentos.casefold()

        nuevos = sorted(
            (elemento for elemento in elementos if elemento != "nueva"),
            key=clave_alfabetica,
        )
        if "nueva" in elementos:
            nuevos.append("nueva")
        for indice, elemento in enumerate(nuevos):
            self.tabla.move(elemento, "", indice)

    def alternar(self, columna: str) -> None:
        if columna not in self.encabezados:
            return
        if self.columna_ordenada == columna:
            for indice, elemento in enumerate(self.orden_original):
                if self.tabla.exists(elemento):
                    self.tabla.move(elemento, "", indice)
            self.columna_ordenada = None
        else:
            self.columna_ordenada = columna
            self._aplicar_orden(columna)
        self._actualizar_encabezados()


def etiqueta_control(campo: str, etiqueta: str, idioma: str) -> str:
    if idioma == "INGLES":
        return CAMPOS_CONTROL_INGLES.get(campo, etiqueta)
    return etiqueta


class _DetallesSelector(ttk.Frame):
    """Conserva etiquetas y secciones al cambiar el documento mostrado."""

    def __init__(self, padre):
        super().__init__(padre, padding=20, cursor="arrow")
        self._widgets = {}
        self._configuraciones = {}
        self._escala = 1.0
        self._ui_preservar_fuentes = True

    def _etiqueta(self, clave, padre, texto, *, tamano=9, negrita=False, color=None, ancho=780):
        if clave not in self._widgets:
            self._widgets[clave] = ttk.Label(padre, justify="left")
        etiqueta = self._widgets[clave]
        configuracion = dict(text=texto, wraplength=ancho,
            foreground=color or ui.paleta()["texto"],
            font=(ui.PREFERENCIAS["fuente"], max(7, round(tamano * self._escala)),
                  "bold" if negrita else "normal"))
        if self._configuraciones.get(clave) != configuracion:
            etiqueta.configure(**configuracion)
            self._configuraciones[clave] = configuracion
        return etiqueta

    def _grupo(self, clave, titulo):
        if clave not in self._widgets:
            self._widgets[clave] = ttk.LabelFrame(self, padding=10)
        grupo = self._widgets[clave]
        if str(grupo.cget("text")) != titulo:
            grupo.configure(text=titulo)
        for hijo in grupo.winfo_children():
            if hijo.winfo_manager() == "grid":
                hijo.grid_remove()
            else:
                hijo.pack_forget()
        grupo.pack(fill="x", pady=(0, 14))
        return grupo

    def actualizar(self, previa, origen, *, alineada, escala):
        self._escala = escala
        for hijo in self.winfo_children():
            hijo.pack_forget()
        secundario = ui.paleta()["secundario"]
        if previa.idioma == "INGLES":
            self._etiqueta("idioma", self, "English version", color=secundario).pack(anchor="w")
        self._etiqueta("origen", self, f"Origen: {origen}", color=secundario).pack(anchor="w", pady=(0, 12))
        for clave, titulo, mensajes in (
                ("avisos", "Avisos del control documental", previa.mensajes_control),
                ("astm", "Pruebas ASTM no encontradas en la lista maestra", previa.mensajes_astm_no_encontradas)):
            if mensajes:
                grupo = self._grupo(clave, titulo)
                for i, mensaje in enumerate(mensajes):
                    self._etiqueta((clave, i), grupo, f"\u2022 {mensaje}", color="#9a3412").pack(anchor="w", pady=2)
        grupo = self._grupo("control", "Document control" if previa.idioma == "INGLES" else "Control documental")
        for i, (etiqueta, valor) in enumerate(previa.control):
            fila, columna = divmod(i, 2)
            self._etiqueta(("campo", i), grupo, f"{etiqueta}:", negrita=True).grid(
                row=fila, column=columna * 2, sticky="w", padx=(4, 5), pady=4)
            self._etiqueta(("valor", i), grupo, valor, ancho=280).grid(
                row=fila, column=columna * 2 + 1, sticky="w", padx=(0, 18), pady=4)
        grupo.columnconfigure(1, weight=1)
        grupo.columnconfigure(3, weight=1)
        if alineada:
            texto = ("This table uses valores_alineados from alineacion_metodos_lab.detalle."
                     if previa.idioma == "INGLES" else
                     "Esta tabla se construye con los valores_alineados guardados en alineacion_metodos_lab.detalle del JSON alineado.")
            self._etiqueta("alineacion", self, texto, color=secundario, ancho=800).pack(anchor="w", pady=(0, 7))
        if previa.notas:
            self._etiqueta("notas", self, "Notas", tamano=11, negrita=True,
                           color=ui.paleta()["alineado"]).pack(anchor="w", pady=(14, 4))
            for i, nota in enumerate(previa.notas):
                self._etiqueta(("nota", i), self, f"\u2022 {nota}", ancho=800).pack(anchor="w", pady=2)
        if previa.pie_pagina:
            if "separador" not in self._widgets:
                self._widgets["separador"] = ttk.Separator(self)
            self._widgets["separador"].pack(fill="x", pady=(16, 8))
            self._etiqueta("pie", self, previa.pie_pagina, color=secundario, ancho=800).pack(anchor="w")


class _TiradorDivisor(tk.Canvas):
    """Agarre plano que delega el arrastre al separador nativo de Tk."""

    def __init__(self, divisor):
        super().__init__(divisor, width=8, height=26, borderwidth=0,
                         highlightthickness=0, cursor="sb_h_double_arrow",
                         background=ui.paleta()["fondo"])
        self.divisor = divisor
        self._tarea = None
        for y in (8, 13, 18):
            self.create_line(2, y, 6, y, fill="#8090a0")
        for secuencia in ("<ButtonPress-1>", "<B1-Motion>", "<ButtonRelease-1>"):
            self.bind(secuencia, lambda evento, s=secuencia: self._reenviar(evento, s))
        divisor.bind("<Configure>", self._programar_posicion, add="+")
        divisor.bind("<ButtonRelease-1>", self._programar_posicion, add="+")
        self.bind("<Destroy>", self._al_destruir, add="+")
        self._programar_posicion()

    def _reenviar(self, evento, secuencia):
        self.divisor.event_generate(secuencia, x=self.winfo_x() + evento.x,
                                    y=self.winfo_y() + evento.y, state=evento.state)
        return "break"

    def _programar_posicion(self, _evento=None):
        if self._tarea is not None:
            self.after_cancel(self._tarea)
        self._tarea = self.after_idle(self._posicionar)

    def _posicionar(self):
        self._tarea = None
        if len(self.divisor.panes()) < 2:
            self.place_forget()
            return
        x, _ = self.divisor.sash_coord(0)
        self.place(x=x, y=self.divisor.winfo_height() // 2, anchor="w",
                   width=int(self.divisor.cget("sashwidth")), height=26)
        tk.Misc.tkraise(self)

    def _al_destruir(self, evento):
        if evento.widget is self and self._tarea is not None:
            self.after_cancel(self._tarea)
            self._tarea = None


def crear_divisor_redimensionable(
    contenedor: tk.Misc, orientacion: str, *, mostrar_tirador: bool = False,
) -> tk.PanedWindow:
    cursor = "sb_h_double_arrow" if orientacion == tk.HORIZONTAL else "sb_v_double_arrow"
    divisor = tk.PanedWindow(
        contenedor,
        orient=orientacion,
        sashwidth=8 if mostrar_tirador else 5,
        handlesize=0,
        showhandle=False,
        sashrelief=tk.FLAT,
        background="#f0f0f0",
        borderwidth=0,
        relief=tk.FLAT,
        cursor="arrow",
        sashcursor=cursor,
        opaqueresize=False,
    )
    if mostrar_tirador and orientacion == tk.HORIZONTAL:
        divisor._tirador_agarre = _TiradorDivisor(divisor)
    return divisor


@dataclass
class PanelEspecificaciones:
    catalogo: Catalogo
    sugerencias: list[Especificacion]
    documentos_marcados: set[Path]
    entrada_busqueda: ttk.Entry
    lista: tk.Listbox
    mensaje_busqueda: ttk.Label
    estado_catalogo: ttk.Label
    marco_previa: ttk.Frame
    canvas_previa: tk.Canvas | None = None
    tabla_pruebas: ttk.Treeview | None = None
    boton_contraparte: ttk.Button | None = None
    boton_idioma: ttk.Button | None = None
    boton_editar: ttk.Button | None = None
    boton_alinear: ttk.Button | None = None
    boton_realinear: ttk.Button | None = None
    boton_comparar: ttk.Button | None = None
    boton_compactar: ttk.Button | None = None
    previa_compacta: bool = False
    separador_previa: tk.PanedWindow | None = None
    area_detalles: ttk.Frame | None = None
    posicion_detalles: tuple[int, int] | None = None
    widgets_previa: dict[str, Any] = field(default_factory=dict)
    cache_previa: dict[tuple, tuple] = field(default_factory=dict)
    firma_vista: tuple | None = None
    busqueda_pendiente: str | None = None
    zoom_previa: float = 1.44
    especificacion_actual: Especificacion | None = None


def _texto(valor: Any) -> str:
    return "" if valor is None else str(valor).strip()


def _aplicar_zoom_vista_previa(contenedor: tk.Misc, factor: float) -> None:
    estilo = ttk.Style(contenedor)
    etiqueta_marco = f"Preview{id(contenedor)}.TLabelframe"
    fuente_marco = ("Segoe UI", max(8, round(9 * factor)), "bold")
    estilo.configure(etiqueta_marco, font=("Segoe UI", max(8, round(10 * factor))))
    estilo.configure(f"{etiqueta_marco}.Label", font=fuente_marco)
    for widget in contenedor.winfo_children():
        if isinstance(widget, ttk.Label):
            nombre_estilo = widget.cget("style") or "TLabel"
            especificacion_fuente = (
                widget.cget("font") or estilo.lookup(nombre_estilo, "font") or "TkDefaultFont"
            )
            fuente = tkfont.Font(root=widget, font=especificacion_fuente)
            tamanio = max(7, round(abs(int(fuente.actual("size"))) * factor))
            widget.configure(
                font=(
                    fuente.actual("family"),
                    tamanio,
                    fuente.actual("weight"),
                )
            )
        elif isinstance(widget, ttk.LabelFrame):
            widget.configure(style=etiqueta_marco)
        _aplicar_zoom_vista_previa(widget, factor)


def _referencias_maestras_astm() -> frozenset[str]:
    ruta = ruta_maestro(BASE)
    info = ruta.stat()
    return _referencias_maestras_desde_archivo(ruta, info.st_mtime_ns, info.st_size)


@lru_cache(maxsize=4)
def _referencias_maestras_desde_archivo(ruta, fecha, tamano) -> frozenset[str]:
    registros = cargar_maestro(ruta)
    return frozenset(
        referencia
        for registro in registros
        for metodo in metodos_maestros(registro)
        for referencia in referencias_astm(metodo)
    )


def _mensajes_astm_no_encontradas(
    datos: dict[str, Any], pruebas: list[dict[str, Any]]
) -> tuple[str, ...]:
    alineacion = datos.get("alineacion_metodos_lab")
    alineacion = alineacion if isinstance(alineacion, dict) else {}
    mensajes = alineacion.get("mensajes_astm_no_encontradas")
    if isinstance(mensajes, list):
        mensajes_validos = tuple(
            mensaje for mensaje in mensajes if isinstance(mensaje, str) and mensaje
        )
        if mensajes_validos:
            return mensajes_validos

    detalle = alineacion.get("detalle")
    if isinstance(detalle, list):
        mensajes_guardados = []
        for elemento in detalle:
            if not isinstance(elemento, dict):
                continue
            mensaje = elemento.get("mensaje_astm_no_encontradas")
            if isinstance(mensaje, str) and mensaje:
                mensajes_guardados.append(mensaje)
                continue
            referencias = elemento.get("referencias_astm_no_encontradas")
            if not isinstance(referencias, list) or not referencias:
                continue
            numero = elemento.get("numero_prueba", indice)
            nombre = _texto(elemento.get("nombre_original"))
            sufijo = f" ({nombre})" if nombre else ""
            metodos = ", ".join(f"ASTM {referencia}" for referencia in referencias)
            mensajes_guardados.append(
                f"Prueba {numero}{sufijo}: Referencias no encontradas "
                f"en la lista maestra: {metodos}."
            )
        if mensajes_guardados:
            return tuple(mensajes_guardados)

    referencias_maestras = _referencias_maestras_astm()
    resultado = []
    for indice, prueba in enumerate(pruebas, 1):
        metodo = prueba.get("metodo", "")
        faltantes = referencias_astm_fuera_maestro(metodo, referencias_maestras)
        if not faltantes:
            continue
        nombre = _texto(prueba.get("prueba") or prueba.get("nombre"))
        sufijo = f" ({nombre})" if nombre else ""
        referencias = ", ".join(f"ASTM {referencia}" for referencia in faltantes)
        resultado.append(
            f"Prueba {indice}{sufijo}: Referencias no encontradas "
            f"en la lista maestra: {referencias}."
        )
    return tuple(resultado)


def cargar_catalogo(carpeta: Path, *, archivos: list[Path] | None = None) -> Catalogo:
    """Indexa JSON en español e inglés; reporta archivos inválidos."""
    if not carpeta.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de especificaciones: {carpeta}")

    especificaciones = []
    errores = []
    for ruta in detectar_json(carpeta) if archivos is None else archivos:
        try:
            datos = leer_datos_json(ruta)
            idioma = leer_idioma_datos(datos)
            producto = datos.get("producto")
            producto = producto if isinstance(producto, dict) else {}
            control = datos.get("control_documental")
            control = control if isinstance(control, dict) else {}
            nombre = _texto(producto.get("nombre")) or ruta.stem
            codigo = _texto(control.get("codigo_unico_producto"))
            cliente = _texto(control.get("nombre_cliente"))
            revision = _texto(control.get("revision"))
            alineacion = datos.get("alineacion_metodos_lab")
            alineacion = alineacion if isinstance(alineacion, dict) else {}
            version_alineador = _texto(alineacion.get("version_programa"))
            pruebas = datos.get("pruebas")
            nombres_pruebas = (
                " ".join(
                    _texto(prueba.get("prueba") or prueba.get("nombre"))
                    for prueba in pruebas
                    if isinstance(prueba, dict)
                )
                if isinstance(pruebas, list)
                else ""
            )
            campos = (
                codigo,
                nombre,
                cliente,
                revision,
                ruta.stem,
                ruta.relative_to(carpeta).as_posix(),
                nombres_pruebas,
            )
            especificaciones.append(
                Especificacion(
                    ruta=ruta,
                    nombre_producto=nombre,
                    codigo=codigo,
                    cliente=cliente,
                    revision=revision,
                    busqueda=clave_texto(" ".join(campos)),
                    idioma=idioma,
                    version_alineador=version_alineador,
                    en_progreso=es_wip(ruta),
                    edicion_bilingue=(datos.get("documento") or {}).get("edicion_bilingue") == 1
                    if isinstance(datos.get("documento"), dict) else False,
                    versiones_idioma=tuple(
                        (ruta.parent / referencia["archivo_json"]).resolve()
                        for referencias in (datos.get("versiones_idioma") or {}).values()
                        if isinstance(referencias, list)
                        for referencia in referencias
                        if isinstance(referencia, dict)
                        and isinstance(referencia.get("archivo_json"), str)
                    ) if isinstance(datos.get("versiones_idioma"), dict) else (),
                )
            )
        except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
            errores.append(f"{ruta.name}: {exc}")

    return Catalogo(
        especificaciones=tuple(especificaciones),
        errores=tuple(errores),
    )


def parejas_idioma(especificaciones: tuple[Especificacion, ...]) -> dict[Path, Especificacion]:
    """Agrupa parejas vigentes o del editor bilingüe, explícitas y de igual revisión."""
    por_ruta = {e.ruta.resolve(): e for e in especificaciones}
    candidatas: dict[Path, set[Path]] = {}
    for origen in especificaciones:
        if origen.version_alineador != VERSION_ALINEADOR and not origen.edicion_bilingue and not origen.en_progreso:
            continue
        for ruta in origen.versiones_idioma:
            destino = por_ruta.get(ruta)
            if destino is None or (destino.version_alineador != VERSION_ALINEADOR and not destino.edicion_bilingue and not destino.en_progreso):
                continue
            if {origen.idioma, destino.idioma} != {"ESPANOL", "INGLES"}:
                continue
            if tuple(map(clave_texto, (origen.codigo, origen.revision, origen.cliente))) != tuple(
                map(clave_texto, (destino.codigo, destino.revision, destino.cliente))
            ):
                continue
            candidatas.setdefault(origen.ruta, set()).add(destino.ruta)
            candidatas.setdefault(destino.ruta, set()).add(origen.ruta)
    return {
        ruta: por_ruta[next(iter(destinos)).resolve()]
        for ruta, destinos in candidatas.items()
        if len(destinos) == 1 and len(candidatas[next(iter(destinos))]) == 1
    }


def _parejas_panel(panel):
    if getattr(panel, "_catalogo_parejas", None) is not panel.catalogo:
        panel._catalogo_parejas = panel.catalogo
        panel._parejas_idioma = parejas_idioma(panel.catalogo.especificaciones)
    return panel._parejas_idioma


def buscar_especificaciones(
    especificaciones: tuple[Especificacion, ...] | list[Especificacion],
    consulta: str,
    limite: int | None = None,
) -> list[Especificacion]:
    """Coincidencia sin acentos, rankeada por prefijo, producto y metadatos."""
    consulta_normalizada = clave_texto(consulta)
    if not consulta_normalizada:
        resultados = list(especificaciones)
        return resultados if limite is None else resultados[:limite]

    tokens = consulta_normalizada.split()
    coincidencias: list[tuple[tuple[int, int, str], Especificacion]] = []
    for especificacion in especificaciones:
        campos = especificacion.busqueda
        if not all(token in campos for token in tokens):
            continue
        nombre = clave_texto(especificacion.nombre_producto)
        codigo = clave_texto(especificacion.codigo)
        if codigo.startswith(consulta_normalizada):
            rango = 0
        elif nombre.startswith(consulta_normalizada):
            rango = 1
        elif campos.startswith(consulta_normalizada):
            rango = 2
        else:
            rango = 3
        coincidencias.append(
            ((rango, campos.find(consulta_normalizada), nombre), especificacion)
        )

    coincidencias.sort(key=lambda item: item[0])
    resultados = [especificacion for _, especificacion in coincidencias]
    return resultados if limite is None else resultados[:limite]


def buscar_contraparte(
    especificacion: Especificacion,
    candidatas: tuple[Especificacion, ...] | list[Especificacion],
) -> Especificacion | None:
    """Busca la otra versión sin confundir revisiones del mismo producto."""
    mismo_nombre = [
        candidata
        for candidata in candidatas
        if candidata.ruta.name.casefold() == especificacion.ruta.name.casefold()
    ]

    def compatibles(candidata: Especificacion) -> bool:
        if especificacion.idioma != candidata.idioma:
            return False
        if especificacion.codigo and candidata.codigo:
            if clave_texto(especificacion.codigo) != clave_texto(candidata.codigo):
                return False
            if clave_texto(especificacion.revision) != clave_texto(candidata.revision):
                return False
            return not (
                especificacion.cliente
                and candidata.cliente
                and clave_texto(especificacion.cliente)
                != clave_texto(candidata.cliente)
            )
        return (
            clave_texto(especificacion.nombre_producto)
            == clave_texto(candidata.nombre_producto)
            and clave_texto(especificacion.revision)
            == clave_texto(candidata.revision)
            and (
                not especificacion.cliente
                or not candidata.cliente
                or clave_texto(especificacion.cliente)
                == clave_texto(candidata.cliente)
            )
        )

    coincidencias_archivo = [
        candidata for candidata in mismo_nombre if compatibles(candidata)
    ]
    if len(coincidencias_archivo) == 1:
        return coincidencias_archivo[0]
    if coincidencias_archivo:
        return None

    coincidencias_metadatos = [
        candidata
        for candidata in candidatas
        if candidata not in mismo_nombre and compatibles(candidata)
    ]
    return coincidencias_metadatos[0] if len(coincidencias_metadatos) == 1 else None


def requiere_realineacion(version_guardada: str, version_actual: str = VERSION_ALINEADOR) -> bool:
    def componentes(version: str) -> tuple[int, ...] | None:
        coincidencia = re.fullmatch(r"\s*[vV]?(\d+(?:\.\d+)*)\s*", version)
        if coincidencia is None:
            return None
        return tuple(int(parte) for parte in coincidencia.group(1).split("."))

    guardada = componentes(version_guardada)
    actual = componentes(version_actual)
    if actual is None:
        raise ValueError(f"Versión actual del alineador inválida: {version_actual!r}")
    if guardada is None:
        return True
    longitud = max(len(guardada), len(actual))
    guardada += (0,) * (longitud - len(guardada))
    actual += (0,) * (longitud - len(actual))
    return guardada < actual


def especificaciones_pendientes_realineacion(
    seleccionadas: list[Especificacion],
) -> list[Especificacion]:
    return [
        especificacion
        for especificacion in seleccionadas
        if requiere_realineacion(especificacion.version_alineador)
    ]


def preparar_vista_previa(ruta: Path, *, datos: dict | None = None) -> VistaPrevia:
    """Interpreta un JSON con las mismas transformaciones del generador Word."""
    if datos is None:
        datos = leer_datos_json(ruta)
    idioma = leer_idioma_datos(datos)

    producto = datos.get("producto")
    nombre = producto.get("nombre") if isinstance(producto, dict) else None
    if not isinstance(nombre, str) or not nombre.strip():
        raise ValueError("Falta el campo de texto 'producto.nombre'.")

    pruebas = datos.get("pruebas")
    if not isinstance(pruebas, list) or any(not isinstance(item, dict) for item in pruebas):
        raise ValueError("El campo 'pruebas' debe ser una lista de objetos.")
    pruebas_validas, notas_en_pruebas = _separar_notas_de_pruebas(pruebas)

    control = datos.get("control_documental", {})
    if not isinstance(control, dict):
        raise ValueError("El campo 'control_documental' debe ser un objeto.")
    mensajes_control = ()
    try:
        control_generado = preparar_control_generacion(control)
    except ValueError as exc:
        if "control_documental.revision" not in str(exc):
            raise
        control_generado = dict(control)
        mensajes_control = (str(exc),)
    notas = datos.get("notas", [])
    if not isinstance(notas, list):
        raise ValueError("El campo 'notas' debe ser una lista.")

    institucionales = datos.get("datos_institucionales", {})
    if not isinstance(institucionales, dict):
        raise ValueError("El campo 'datos_institucionales' debe ser un objeto.")
    pie = institucionales.get("pie_pagina")
    if pie is not None and not isinstance(pie, str):
        raise ValueError("El campo 'datos_institucionales.pie_pagina' debe ser texto.")

    campos_control = tuple(
        (etiqueta_control(campo, etiqueta, idioma), _texto(control_generado.get(campo)))
        for etiqueta, campo in CAMPOS_CONTROL
        if _texto(control_generado.get(campo))
    )
    notas_finales = tuple(
        dict.fromkeys(_texto(nota) for nota in [*notas, *notas_en_pruebas] if _texto(nota))
    )
    return VistaPrevia(
        nombre_producto=nombre.strip(),
        control=campos_control,
        pruebas=tuple(pruebas_validas),
        notas=notas_finales,
        mensajes_astm_no_encontradas=_mensajes_astm_no_encontradas(datos, pruebas),
        mensajes_control=mensajes_control,
        pie_pagina=pie or "",
        idioma=idioma,
    )


def preparar_vista_previa_alineada(ruta: Path, *, datos: dict | None = None) -> VistaPrevia:
    """Muestra exclusivamente la instantanea de valores guardada en la auditoria alineada."""
    if datos is None:
        datos = leer_datos_json(ruta)
    if not isinstance(datos, dict):
        raise ValueError("El JSON debe contener un objeto.")
    idioma = leer_idioma_datos(datos)

    producto = datos.get("producto")
    nombre = producto.get("nombre") if isinstance(producto, dict) else None
    if not isinstance(nombre, str) or not nombre.strip():
        raise ValueError("Falta el campo de texto 'producto.nombre'.")

    alineacion = datos.get("alineacion_metodos_lab")
    detalle = alineacion.get("detalle") if isinstance(alineacion, dict) else None
    if not isinstance(detalle, list) or any(not isinstance(item, dict) for item in detalle):
        raise ValueError(
            "El JSON alineado no contiene 'alineacion_metodos_lab.detalle'. "
            "Vuelve a ejecutar la alineacion para regenerarlo."
        )
    pruebas_alineadas = []
    for indice, elemento in enumerate(detalle, 1):
        valores = elemento.get("valores_alineados")
        if not isinstance(valores, dict):
            raise ValueError(
                "El detalle de alineacion no contiene 'valores_alineados'. "
                "Vuelve a ejecutar la alineacion para regenerarlo."
            )
        pruebas_alineadas.append(
            {
                campo: _texto(valores.get(campo))
                for campo in ("prueba", "unidades", "metodo", "especificacion", "valor_tipico")
            }
        )
    pruebas_validas, notas_en_pruebas = _separar_notas_de_pruebas(pruebas_alineadas)

    control = datos.get("control_documental", {})
    if not isinstance(control, dict):
        raise ValueError("El campo 'control_documental' debe ser un objeto.")
    mensajes_control = ()
    try:
        preparar_control_generacion(control)
    except ValueError as exc:
        if "control_documental.revision" not in str(exc):
            raise
        mensajes_control = (str(exc),)

    notas = datos.get("notas", [])
    if not isinstance(notas, list):
        raise ValueError("El campo 'notas' debe ser una lista.")
    institucionales = datos.get("datos_institucionales", {})
    if not isinstance(institucionales, dict):
        raise ValueError("El campo 'datos_institucionales' debe ser un objeto.")
    pie = institucionales.get("pie_pagina")
    if pie is not None and not isinstance(pie, str):
        raise ValueError("El campo 'datos_institucionales.pie_pagina' debe ser texto.")

    campos_control = tuple(
        (etiqueta_control(campo, etiqueta, idioma), _texto(control.get(campo)))
        for etiqueta, campo in CAMPOS_CONTROL
        if _texto(control.get(campo))
    )
    notas_finales = tuple(
        dict.fromkeys(_texto(nota) for nota in [*notas, *notas_en_pruebas] if _texto(nota))
    )
    return VistaPrevia(
        nombre_producto=nombre.strip(),
        control=campos_control,
        pruebas=tuple(pruebas_validas),
        notas=notas_finales,
        mensajes_astm_no_encontradas=_mensajes_astm_no_encontradas(datos, pruebas_alineadas),
        mensajes_control=mensajes_control,
        pie_pagina=pie or "",
        idioma=idioma,
    )


class SelectorEspecificaciones:
    def __init__(
        self,
        raiz: tk.Tk | tk.Toplevel,
        carpeta_json: Path | None = None,
        carpeta_json_anterior: Path | None = None,
        al_actualizar: Callable[[Path], None] | None = None,
        al_editar: Callable[[Path], None] | None = None,
        al_comparar: Callable[[list[Path]], None] | None = None,
    ):
        self.raiz = raiz
        self.al_actualizar = al_actualizar
        self.al_editar = al_editar
        self.al_comparar = al_comparar
        rutas = rutas_configuradas(BASE)
        self.carpetas_json = {
            "anteriores": resolver_ruta(carpeta_json_anterior if carpeta_json_anterior is not None else rutas["json_anteriores"], BASE),
            "nuevas": resolver_ruta(carpeta_json if carpeta_json is not None else rutas["json_alineadas"], BASE),
        }
        self.paneles: dict[str, PanelEspecificaciones] = {}
        self.zoom_aplicacion = ui.PREFERENCIAS["escala"]
        self._escala_tk_base = float(self.raiz.tk.call("tk", "scaling"))
        self._alineacion_en_proceso = False
        self._cola_alineacion: queue.Queue[tuple[str, Any]] | None = None
        self._ventana_progreso_alineacion: tk.Toplevel | None = None
        self._etiqueta_progreso_alineacion: ttk.Label | None = None
        self._barra_progreso_alineacion: ttk.Progressbar | None = None
        self._lotes_pendientes_alineacion: list[
            tuple[list[Path], list[Path] | None]
        ] = []
        self._archivos_actuales_alineacion: set[Path] = set()
        self._resumen_alineacion: dict[str, int] = {}
        self._configurar_ventana()
        self._crear_interfaz()
        self.recargar_catalogos()

    def _configurar_ventana(self) -> None:
        self.raiz.title("Selector de especificaciones")
        self.raiz.geometry("1250x800")
        self.raiz.minsize(950, 620)
        self.raiz.configure(background="#edf2f7")
        self.raiz.option_add("*Font", ("Segoe UI", 10))

    def _crear_interfaz(self) -> None:
        estilo = ttk.Style(self.raiz)
        estilo.configure("Titulo.TLabel", font=("Segoe UI", 17, "bold"), foreground="#ffffff")
        estilo.configure("Subtitulo.TLabel", font=("Segoe UI", 10), foreground="#dbeafe")
        estilo.configure("Seccion.TLabel", font=("Segoe UI", 12, "bold"), foreground=ui.paleta()["alineado"])
        estilo.configure("Treeview", rowheight=25, font=("Segoe UI", 9))
        estilo.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        encabezado = tk.Frame(self.raiz, background="#0752b9", padx=24, pady=5)
        encabezado.pack(fill="x")
        controles_zoom = ui.ctk.CTkFrame(encabezado, fg_color="#0752b9")
        controles_zoom.pack(side="right", padx=(12, 0))
        ttk.Button(
            controles_zoom, text="A−", width=4, height=26, fondo="#0752b9", command=lambda: self._cambiar_zoom_aplicacion(-0.1)
        ).pack(side="left", padx=2)
        ttk.Button(
            controles_zoom, text="A+", width=4, height=26, fondo="#0752b9", command=lambda: self._cambiar_zoom_aplicacion(0.1)
        ).pack(side="left", padx=2)
        ttk.Label(
            encabezado,
            text="SELECCIONAR ESPECIFICACIÓN",
            style="Titulo.TLabel",
            background="#0752b9",
        ).pack(anchor="w")
        self.pestanas = ttk.Notebook(self.raiz)
        self.pestanas.pack(fill="both", expand=True, padx=18, pady=(0, 5))
        for clave, titulo in (
            ("anteriores", "Anteriores · JSON extraídos"),
            ("nuevas", "Nuevas · JSON alineados"),
        ):
            self._crear_pestana(clave, titulo)
        self.pestanas.select(1)

    def _crear_pestana(self, clave: str, titulo: str) -> None:
        pestaña = ttk.Frame(self.pestanas)
        self.pestanas.add(pestaña, text=titulo)
        contenido = crear_divisor_redimensionable(pestaña, tk.HORIZONTAL, mostrar_tirador=True)
        contenido.pack(fill="both", expand=True, padx=4, pady=8)
        panel_busqueda = ttk.Frame(contenido, padding=12, cursor="arrow")
        panel_previa = ttk.Frame(contenido, padding=12, cursor="arrow")
        contenido.add(panel_busqueda, minsize=220, stretch="always")
        contenido.add(panel_previa, minsize=280, stretch="always")

        ttk.Label(panel_busqueda, text="Buscar especificación", style="Seccion.TLabel").pack(
            anchor="w", pady=(0, 8)
        )
        entrada = ttk.Entry(panel_busqueda)
        entrada.pack(fill="x", pady=(0, 8))

        marco_lista = ttk.Frame(panel_busqueda)
        marco_lista.pack(fill="both", expand=True)
        lista = tk.Listbox(
            marco_lista,
            exportselection=False,
            activestyle="dotbox",
            background="#ffffff",
            relief="solid",
            borderwidth=1,
            selectbackground="#dbeafe",
            selectforeground="#123b70",
            font=("Segoe UI", 9),
        )
        barra_lista = ttk.Scrollbar(marco_lista, orient="vertical", command=lista.yview)
        lista.configure(yscrollcommand=barra_lista.set)
        lista.bind("<ButtonRelease-1>", lambda evento, c=clave: self._clic_en_lista(c, evento))
        lista.bind("<<ListboxSelect>>", lambda _evento, c=clave: self._al_seleccionar(c))
        lista.pack(side="left", fill="both", expand=True)
        barra_lista.pack(side="right", fill="y")
        mensaje = ttk.Label(panel_busqueda, text="", wraplength=280)
        mensaje.pack(anchor="w", pady=(5, 0))
        ttk.Button(
            panel_busqueda,
            text="Actualizar lista",
            height=28, variante="secundario",
            command=self.recargar_catalogos,
        ).pack(fill="x", pady=(10, 0))
        estado = ttk.Label(panel_busqueda, text="", wraplength=280)
        estado.pack(anchor="w", pady=(8, 0))
        boton_alinear = None
        boton_realinear = None
        boton_comparar = None
        if clave == "anteriores":
            acciones_alinear = ttk.Frame(panel_busqueda)
            acciones_alinear.pack(fill="x", pady=(12, 3))
            boton_alinear = ttk.Button(
                acciones_alinear,
                text="Alinear especificaciones",
                height=28,
                command=lambda c=clave: self._alinear_seleccionadas(c),
                state="disabled",
            )
            boton_alinear.pack(side="left", fill="x", expand=True)
            ttk.Button(
                acciones_alinear,
                text="ⓘ",
                width=3, height=28, variante="secundario",
                command=self.abrir_resumen_alineador,
            ).pack(side="left", padx=(5, 0))
        else:
            boton_realinear = ttk.Button(
                panel_busqueda,
                text="Realinear a versión vigente",
                height=28,
                command=lambda c=clave: self._realinear_seleccionadas(c),
                state="disabled",
            )
            boton_realinear.pack(fill="x", pady=(12, 0))
            boton_comparar = ttk.Button(panel_busqueda, text="Comparar especificaciones", height=28, command=lambda c=clave: self._comparar_seleccionadas(c), state="disabled")
            boton_comparar.pack(fill="x", pady=(6, 0))

        barra_vista = ttk.Frame(panel_previa)
        barra_vista.pack(fill="x", pady=(0, 9))
        ttk.Label(
            barra_vista,
            text=(
                "Vista previa · Datos extraídos"
                if clave == "anteriores"
                else "Vista previa · JSON alineado"
            ),
            style="Seccion.TLabel",
        ).pack(anchor="w")
        acciones_vista = ttk.Frame(barra_vista)
        acciones_vista.pack(fill="x", pady=(6, 0))
        boton_compactar = ttk.Button(
            acciones_vista, text="Ø", width=3, height=24,
            variante="secundario", state="disabled",
            command=lambda c=clave: self._alternar_previa_compacta(c),
        )
        boton_editar = ttk.Button(
            acciones_vista,
            text="Editar",
            width=10, height=28,
            command=lambda c=clave: self._editar_especificacion(c),
            state="disabled",
        )
        boton_contraparte = ttk.Button(
            acciones_vista,
            text="Ver alineada →" if clave == "anteriores" else "← Ver original",
            width=16, height=28, variante="secundario",
            command=lambda c=clave: self._ir_a_contraparte(c),
            state="disabled",
        )
        ttk.Button(
            acciones_vista,
            text="−",
            width=3, height=28, variante="secundario",
            command=lambda c=clave: self._cambiar_zoom_previa(c, -0.1),
        ).pack(side="right", padx=(6, 0))
        ttk.Button(
            acciones_vista,
            text="+",
            width=3, height=28, variante="secundario",
            command=lambda c=clave: self._cambiar_zoom_previa(c, 0.1),
        ).pack(side="right")
        ttk.Button(
            acciones_vista,
            text="×",
            width=3, height=28, variante="secundario",
            command=lambda c=clave: self._cerrar_vista_previa(c),
        ).pack(side="right", padx=(0, 6))
        boton_compactar.pack(side="left", padx=(0, 6))
        boton_editar.pack(side="left")
        boton_contraparte.pack(side="left", padx=(8, 0))
        boton_idioma = None
        if clave == "nuevas":
            boton_idioma = ttk.Button(
                acciones_vista, text="ENG", width=6, height=28,
                variante="secundario", command=self._cambiar_idioma,
                state="disabled",
            )
        marco_previa = ttk.Frame(panel_previa)
        marco_previa.pack(fill="both", expand=True)
        ttk.Label(
            marco_previa,
            text="Selecciona una especificación para ver cómo se interpreta en el formato.",
            justify="center",
            anchor="center",
            wraplength=650,
        ).pack(fill="both", expand=True)

        self.paneles[clave] = PanelEspecificaciones(
            catalogo=Catalogo((), ()),
            sugerencias=[],
            documentos_marcados=set(),
            entrada_busqueda=entrada,
            lista=lista,
            mensaje_busqueda=mensaje,
            estado_catalogo=estado,
            marco_previa=marco_previa,
            boton_contraparte=boton_contraparte,
            boton_idioma=boton_idioma,
            boton_editar=boton_editar,
            boton_compactar=boton_compactar,
            boton_alinear=boton_alinear,
            boton_realinear=boton_realinear,
            boton_comparar=boton_comparar,
        )
        self._actualizar_estado_accion(clave)
        entrada.bind("<Destroy>", lambda _e, c=clave: self._cancelar_sugerencias(c), add="+")
        entrada.bind("<KeyRelease>", lambda _evento, c=clave: self._programar_sugerencias(c))
        entrada.bind("<Return>", lambda _evento, c=clave: self._seleccionar_primera_sugerencia(c))
        entrada.bind("<Down>", lambda _evento, c=clave: self._enfocar_sugerencias(c))
        lista.bind("<Return>", lambda _evento, c=clave: self._seleccionar_sugerencia_actual(c))

    def recargar_catalogos(self) -> None:
        for clave, carpeta in self.carpetas_json.items():
            panel = self.paneles[clave]
            try:
                panel.catalogo = cargar_catalogo(carpeta)
            except (OSError, ValueError) as exc:
                panel.catalogo = Catalogo((), ())
                panel.estado_catalogo.configure(text=f"No se pudo cargar:\n{exc}")
                self._actualizar_estado_accion(clave)
                continue
            self._actualizar_sugerencias(clave)
            disponibles = len(panel.catalogo.especificaciones)
            inglesas = sum(
                especificacion.idioma == "INGLES"
                for especificacion in panel.catalogo.especificaciones
            )
            texto_estado = f"{disponibles} especificaciones ({inglesas} en inglés)"
            if panel.catalogo.errores:
                texto_estado += f"\n{len(panel.catalogo.errores)} archivo(s) con error"
            panel.estado_catalogo.configure(text=texto_estado)
        for clave in self.paneles:
            self._actualizar_boton_idioma(clave)
            self._actualizar_boton_contraparte(clave)
            self._actualizar_estado_accion(clave)
            panel = self.paneles[clave]
            if panel.especificacion_actual is not None:
                actual = next((e for e in panel.catalogo.especificaciones
                               if e.ruta == panel.especificacion_actual.ruta), None)
                if actual is None:
                    self._cerrar_vista_previa(clave)
                else:
                    self._mostrar_previa(clave, actual)

    def _cancelar_sugerencias(self, clave):
        panel = self.paneles[clave]
        if getattr(panel, "busqueda_pendiente", None) is not None:
            self.raiz.after_cancel(panel.busqueda_pendiente)
            panel.busqueda_pendiente = None

    def _programar_sugerencias(self, clave: str) -> None:
        panel = self.paneles[clave]
        if getattr(panel, "busqueda_pendiente", None) is not None:
            self.raiz.after_cancel(panel.busqueda_pendiente)
        panel.busqueda_pendiente = self.raiz.after(150, lambda: self._actualizar_sugerencias(clave))

    def _actualizar_sugerencias(self, clave: str) -> None:
        panel = self.paneles[clave]
        if getattr(panel, "busqueda_pendiente", None) is not None:
            self.raiz.after_cancel(panel.busqueda_pendiente)
            panel.busqueda_pendiente = None
        parejas = _parejas_panel(panel) if clave == "nuevas" else {}
        resultados = buscar_especificaciones(
            panel.catalogo.especificaciones,
            panel.entrada_busqueda.get(),
        )
        panel.sugerencias = []
        vistas = set()
        for especificacion in resultados:
            pareja = parejas.get(especificacion.ruta)
            if pareja is not None and especificacion.idioma == "INGLES":
                especificacion = pareja
            if especificacion.ruta not in vistas:
                panel.sugerencias.append(especificacion)
                vistas.add(especificacion.ruta)
        if clave == "nuevas":
            panel.sugerencias.sort(key=lambda e: not (e.en_progreso or (e.ruta in parejas and parejas[e.ruta].en_progreso)))
        panel.lista.delete(0, tk.END)
        for especificacion in panel.sugerencias:
            marca = "☑" if especificacion.ruta in panel.documentos_marcados else "☐"
            idioma = " · English" if especificacion.idioma == "INGLES" else ""
            if especificacion.ruta in parejas:
                idioma = " · ESP / ENG"
            en_progreso = clave == "nuevas" and (especificacion.en_progreso or (especificacion.ruta in parejas and parejas[especificacion.ruta].en_progreso))
            estado = " · WIP / En progreso" if en_progreso else ""
            panel.lista.insert(tk.END, f"{marca}  {especificacion.etiqueta}{idioma}{estado}")
            if en_progreso:
                panel.lista.itemconfigure(tk.END, foreground="#d4a000")
            elif especificacion.version_alineador == VERSION_ALINEADOR:
                panel.lista.itemconfigure(tk.END, foreground=ui.paleta()["alineado"])
        mensaje = (
            "No se encontraron especificaciones."
            if panel.entrada_busqueda.get().strip() and not panel.sugerencias
            else ""
        )
        panel.mensaje_busqueda.configure(text=mensaje)

    def _clic_en_lista(self, clave: str, evento: tk.Event) -> None:
        panel = self.paneles[clave]
        if evento.x > 24:
            return
        indice = panel.lista.nearest(evento.y)
        if not 0 <= indice < len(panel.sugerencias):
            return
        especificacion = panel.sugerencias[indice]
        if especificacion.ruta in panel.documentos_marcados:
            panel.documentos_marcados.remove(especificacion.ruta)
        else:
            panel.documentos_marcados.add(especificacion.ruta)
        # Sustituye solo el texto de esta fila; conserva seleccion, color y desplazamiento.
        texto = panel.lista.get(indice)
        color = panel.lista.itemcget(indice, "foreground")
        seleccion = indice in panel.lista.curselection()
        desplazamiento = panel.lista.yview()[0]
        marca = "\u2611" if especificacion.ruta in panel.documentos_marcados else "\u2610"
        panel.lista.delete(indice)
        panel.lista.insert(indice, marca + texto[1:])
        if color:
            panel.lista.itemconfigure(indice, foreground=color)
        if seleccion:
            panel.lista.selection_set(indice)
        panel.lista.activate(indice)
        panel.lista.yview_moveto(desplazamiento)
        self._actualizar_estado_accion(clave)

    def _actualizar_estado_accion(self, clave: str) -> None:
        panel = self.paneles[clave]
        if clave == "anteriores" and panel.boton_alinear is not None:
            panel.boton_alinear.configure(
                state="normal" if panel.documentos_marcados else "disabled"
            )
        if clave == "nuevas" and panel.boton_realinear is not None:
            seleccionadas = [
                especificacion
                for especificacion in panel.catalogo.especificaciones
                if especificacion.ruta in panel.documentos_marcados
            ]
            pendientes = especificaciones_pendientes_realineacion(seleccionadas)
            panel.boton_realinear.configure(
                state="normal" if pendientes else "disabled"
            )
        if clave == "nuevas" and panel.boton_comparar is not None:
            cantidad = len(panel.documentos_marcados)
            panel.boton_comparar.configure(state="normal" if 2 <= cantidad <= 10 else "disabled")

    def _seleccionar_primera_sugerencia(self, clave: str) -> str:
        panel = self.paneles[clave]
        if panel.sugerencias:
            self._mostrar_previa(clave, panel.sugerencias[0])
        return "break"

    def _enfocar_sugerencias(self, clave: str) -> str:
        panel = self.paneles[clave]
        if panel.sugerencias:
            panel.lista.focus_set()
            panel.lista.selection_set(0)
            panel.lista.activate(0)
        return "break"

    def _seleccionar_sugerencia_actual(self, clave: str) -> str:
        self._al_seleccionar(clave)
        return "break"

    def _al_seleccionar(self, clave: str) -> None:
        panel = self.paneles[clave]
        seleccion = panel.lista.curselection()
        if seleccion:
            self._mostrar_previa(clave, panel.sugerencias[seleccion[0]])

    @staticmethod
    def _firma_archivo(ruta):
        if ruta is None:
            return None
        try:
            info = ruta.stat()
            return (ruta, info.st_mtime_ns, info.st_ctime_ns, info.st_size)
        except FileNotFoundError:
            return (ruta, None)

    def _cargar_datos_previa(self, clave, especificacion):
        panel = self.paneles[clave]
        contraparte = (buscar_contraparte(especificacion, self.paneles["anteriores"].catalogo.especificaciones)
                       if clave == "nuevas" else None)
        firma = (self._firma_archivo(especificacion.ruta),
                 self._firma_archivo(contraparte.ruta if contraparte else None),
                 self._firma_archivo(ruta_maestro(BASE)))
        if firma in panel.cache_previa:
            datos = panel.cache_previa.pop(firma)
            panel.cache_previa[firma] = datos
            return firma, datos
        datos = leer_datos_json(especificacion.ruta)
        preparar = preparar_vista_previa if clave == "anteriores" else preparar_vista_previa_alineada
        previa = preparar(especificacion.ruta, datos=datos)
        original = leer_datos_json(contraparte.ruta) if contraparte else None
        cambios = cambios_documento(datos, previa.pruebas, original) if clave == "nuevas" else []
        resultado = (previa, cambios)
        panel.cache_previa[firma] = resultado
        while len(panel.cache_previa) > 12:
            panel.cache_previa.pop(next(iter(panel.cache_previa)))
        return firma, resultado

    def _crear_widgets_previa(self, clave):
        panel = self.paneles[clave]
        for widget in panel.marco_previa.winfo_children():
            widget.destroy()
        separador = crear_divisor_redimensionable(panel.marco_previa, tk.VERTICAL)
        separador.pack(fill="both", expand=True)
        area_tabla = ttk.Frame(separador, padding=10, cursor="arrow")
        area_detalles = ttk.Frame(separador, cursor="arrow")
        separador.add(area_tabla, minsize=120, stretch="always")
        separador.add(area_detalles, minsize=80, stretch="always")
        panel.separador_previa, panel.area_detalles = separador, area_detalles
        separador.paneconfigure(area_detalles, hide=panel.previa_compacta)
        inicializado = False

        def posicionar(evento):
            nonlocal inicializado
            if not inicializado and evento.height > 100:
                inicializado = True
                if not panel.previa_compacta:
                    separador.sash_place(0, 0, round(evento.height * 0.68))

        separador.bind("<Configure>", posicionar)
        titulo = ttk.Label(area_tabla, font=("Segoe UI", 19, "bold"),
                           foreground=ui.paleta()["alineado"], wraplength=780)
        titulo.pack(anchor="w", pady=(0, 4))
        contador = ttk.Label(area_tabla, font=("Segoe UI", 12, "bold"), foreground=ui.paleta()["alineado"])
        contador.pack(anchor="w", pady=(0, 7))
        marco = ttk.Frame(area_tabla)
        marco.pack(fill="both", expand=True)
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)
        columnas = tuple(campo for _, campo in CAMPOS_PRUEBA)
        tabla = ttk.Treeview(marco, columns=columnas, show="headings", style=f"{clave}.Preview.Treeview", cursor="arrow")
        for campo, ancho in zip(columnas, (240, 95, 155, 210, 130)):
            tabla.column(campo, width=ancho, minwidth=75, stretch=True)
        vertical = ttk.Scrollbar(marco, orient="vertical", command=tabla.yview)
        horizontal = ttk.Scrollbar(marco, orient="horizontal", command=tabla.xview)
        tabla.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        ordenador = OrdenadorColumnasTreeview(tabla, {})
        tabla._resaltador_alineacion = ResaltadorCambios(tabla, {}, vertical.set, horizontal.set)
        canvas = tk.Canvas(area_detalles, background=ui.paleta()["fondo"], highlightthickness=0, cursor="arrow")
        barra = ttk.Scrollbar(area_detalles, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=barra.set)
        canvas.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        pagina = _DetallesSelector(canvas)
        ventana = canvas.create_window((0, 0), window=pagina, anchor="nw")
        pagina.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(ventana, width=e.width))
        panel.canvas_previa, panel.tabla_pruebas = canvas, tabla
        panel.widgets_previa = dict(titulo=titulo, contador=contador, pagina=pagina,
                                    ordenador=ordenador, vertical=vertical)
        panel.marco_previa._ui_preservar_fuentes = True

    def _mostrar_previa(self, clave: str, especificacion: Especificacion) -> None:
        panel = self.paneles[clave]
        panel.especificacion_actual = especificacion
        try:
            firma_datos, (previa, cambios) = self._cargar_datos_previa(clave, especificacion)
        except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
            self._mostrar_error(clave, str(exc))
            return
        firma = (firma_datos, panel.zoom_previa, self.zoom_aplicacion)
        if firma == panel.firma_vista:
            return
        self._actualizar_boton_idioma(clave)
        self._actualizar_boton_contraparte(clave)
        if not panel.widgets_previa:
            self._crear_widgets_previa(clave)
        widgets = panel.widgets_previa
        tabla = panel.tabla_pruebas
        escala = panel.zoom_previa * 0.7
        estilo = ttk.Style(self.raiz)
        nombre = f"{clave}.Preview.Treeview"
        estilo.configure(nombre, font=("Segoe UI", max(7, round(9 * escala))),
                         rowheight=max(18, round(25 * escala)))
        estilo.configure(f"{nombre}.Heading", font=("Segoe UI", max(7, round(9 * escala)), "bold"))
        datos_nuevos = panel.firma_vista is None or panel.firma_vista[0] != firma_datos
        if datos_nuevos:
            widgets["titulo"].configure(text=previa.nombre_producto)
            widgets["contador"].configure(text=f"{'Tests' if previa.idioma == 'INGLES' else 'Pruebas'} ({len(previa.pruebas)})")
            columnas = tuple(campo for _, campo in CAMPOS_PRUEBA)
            for texto, campo in encabezados_tabla(previa.idioma):
                if campo not in {"prueba", "metodo"}:
                    tabla.heading(campo, text=texto)
            existentes = set(tabla.get_children())
            cambios_celdas = {}
            for i, prueba in enumerate(previa.pruebas):
                iid = f"prueba:{i}"
                valores = tuple(_texto(prueba.get(campo)) for campo in columnas)
                if iid in existentes:
                    tabla.item(iid, values=valores)
                    existentes.remove(iid)
                else:
                    tabla.insert("", "end", iid=iid, values=valores)
                tabla.move(iid, "", i)
                if i < len(cambios):
                    cambios_celdas[iid] = cambios[i]
            if existentes:
                tabla.delete(*existentes)
            if tabla.selection():
                tabla.selection_remove(*tabla.selection())
            ordenador = widgets["ordenador"]
            ordenador.columna_ordenada = None
            ordenador.registrar_orden_original()
            ordenador.establecer_encabezados({c: t for t, c in encabezados_tabla(previa.idioma) if c in {"prueba", "metodo"}})
            tabla._resaltador_alineacion.establecer_cambios(cambios_celdas)
            tabla.yview_moveto(0)
            panel.canvas_previa.yview_moveto(0)
        else:
            widgets["ordenador"]._actualizar_encabezados()
            tabla._resaltador_alineacion.programar()
        widgets["pagina"].actualizar(previa, especificacion.ruta.relative_to(self.carpetas_json[clave]),
                                     alineada=clave == "nuevas", escala=escala)
        self._vincular_scroll_preview(panel, widgets["pagina"], tabla, widgets["vertical"])
        panel.boton_editar.configure(state="normal")
        panel.boton_compactar.configure(state="normal")
        panel.firma_vista = firma

    def _actualizar_boton_idioma(self, clave: str) -> None:
        panel = self.paneles[clave]
        if panel.boton_idioma is None:
            return
        actual = panel.especificacion_actual
        pareja = _parejas_panel(panel).get(actual.ruta) if actual else None
        if pareja is None:
            panel.boton_idioma.pack_forget()
        else:
            panel.boton_idioma.configure(text="ENG" if actual.idioma == "ESPANOL" else "ESP", state="normal")
            panel.boton_idioma.pack(side="left", padx=(8, 0))

    def _cambiar_idioma(self) -> None:
        panel = self.paneles["nuevas"]
        actual = panel.especificacion_actual
        pareja = _parejas_panel(panel).get(actual.ruta) if actual else None
        if pareja is not None:
            self._mostrar_previa("nuevas", pareja)

    def _actualizar_boton_contraparte(self, clave: str) -> None:
        panel = self.paneles[clave]
        if panel.boton_contraparte is None:
            return
        especificacion = panel.especificacion_actual
        otra_clave = "nuevas" if clave == "anteriores" else "anteriores"
        contraparte = (
            buscar_contraparte(
                especificacion,
                self.paneles[otra_clave].catalogo.especificaciones,
            )
            if especificacion is not None
            else None
        )
        panel.boton_contraparte.configure(
            state="normal" if contraparte is not None else "disabled"
        )

    def _ir_a_contraparte(self, clave: str) -> None:
        panel = self.paneles[clave]
        especificacion = panel.especificacion_actual
        if especificacion is None:
            return
        otra_clave = "nuevas" if clave == "anteriores" else "anteriores"
        otro_panel = self.paneles[otra_clave]
        contraparte = buscar_contraparte(
            especificacion,
            otro_panel.catalogo.especificaciones,
        )
        if contraparte is None:
            self._actualizar_boton_contraparte(clave)
            return

        entrada_lista = contraparte
        if otra_clave == "nuevas" and contraparte.idioma == "INGLES":
            entrada_lista = _parejas_panel(otro_panel).get(
                contraparte.ruta, contraparte
            )

        otro_panel.entrada_busqueda.delete(0, tk.END)
        otro_panel.entrada_busqueda.insert(
            0,
            contraparte.codigo or contraparte.nombre_producto,
        )
        self._actualizar_sugerencias(otra_clave)
        indice = next(
            (
                indice
                for indice, sugerencia in enumerate(otro_panel.sugerencias)
                if sugerencia.ruta == entrada_lista.ruta
            ),
            None,
        )
        if indice is None:
            otro_panel.entrada_busqueda.delete(0, tk.END)
            self._actualizar_sugerencias(otra_clave)
            indice = next(
                (
                    indice
                    for indice, sugerencia in enumerate(otro_panel.sugerencias)
                    if sugerencia.ruta == entrada_lista.ruta
                ),
                None,
            )
        if indice is None:
            raise RuntimeError(
                f"No se pudo mostrar la especificación relacionada: {contraparte.ruta}"
            )

        otro_panel.lista.selection_clear(0, tk.END)
        otro_panel.lista.selection_set(indice)
        otro_panel.lista.activate(indice)
        otro_panel.lista.see(indice)
        self.pestanas.select(0 if otra_clave == "anteriores" else 1)
        self._mostrar_previa(otra_clave, contraparte)

    def _cambiar_zoom_aplicacion(self, cambio: float) -> None:
        self.zoom_aplicacion = min(1.4, max(0.8, self.zoom_aplicacion + cambio))
        ui.aplicar_preferencias({**ui.PREFERENCIAS, "escala": self.zoom_aplicacion})
        for clave, panel in self.paneles.items():
            if panel.especificacion_actual is not None:
                self._mostrar_previa(clave, panel.especificacion_actual)

    def _cambiar_zoom_previa(self, clave: str, cambio: float) -> None:
        panel = self.paneles[clave]
        panel.zoom_previa = min(1.8, max(0.7, panel.zoom_previa + cambio))
        if panel.especificacion_actual is not None:
            self._mostrar_previa(clave, panel.especificacion_actual)

    def _alternar_previa_compacta(self, clave: str) -> None:
        panel = self.paneles[clave]
        separador = panel.separador_previa
        if separador is None or not separador.winfo_exists():
            return
        if not panel.previa_compacta:
            panel.posicion_detalles = separador.sash_coord(0)
        panel.previa_compacta = not panel.previa_compacta
        separador.paneconfigure(panel.area_detalles, hide=panel.previa_compacta)
        if not panel.previa_compacta:
            posicion = panel.posicion_detalles

            def restaurar():
                if (panel.separador_previa is separador and separador.winfo_exists()
                        and not panel.previa_compacta):
                    x, y = posicion or (0, round(separador.winfo_height() * 0.68))
                    separador.sash_place(0, x, y)

            separador.after_idle(restaurar)

    def _cerrar_vista_previa(self, clave: str) -> None:
        panel = self.paneles[clave]
        panel.widgets_previa.clear()
        panel.firma_vista = None
        panel.separador_previa = None
        panel.area_detalles = None
        panel.posicion_detalles = None
        if panel.boton_compactar is not None:
            panel.boton_compactar.configure(state="disabled")
        panel.especificacion_actual = None
        self._actualizar_boton_idioma(clave)
        self._actualizar_boton_contraparte(clave)
        panel.canvas_previa = None
        panel.tabla_pruebas = None
        if panel.boton_editar is not None:
            panel.boton_editar.configure(state="disabled")
        for widget in panel.marco_previa.winfo_children():
            widget.destroy()
        ttk.Label(
            panel.marco_previa,
            text="Selecciona una especificación para ver cómo se interpreta en el formato.",
            justify="center",
            anchor="center",
            wraplength=650,
        ).pack(fill="both", expand=True)

    def _actualizar_especificacion(self, clave: str) -> None:
        panel = self.paneles[clave]
        seleccionadas = [
            especificacion.ruta
            for especificacion in panel.catalogo.especificaciones
            if especificacion.ruta in panel.documentos_marcados
        ]
        if not seleccionadas and panel.especificacion_actual is not None:
            seleccionadas = [panel.especificacion_actual.ruta]
        if len(seleccionadas) != 1:
            messagebox.showinfo(
                "Selecciona una especificación",
                "Marca exactamente una especificación para abrirla en el editor.",
                parent=self.raiz,
            )
            return
        if self.al_actualizar is None:
            aplicacion = BASE / "interfaz_aplicacion_especificaciones.py"
            try:
                subprocess.Popen(
                    [
                        sys.executable,
                        *([] if getattr(sys, "frozen", False) else [str(aplicacion)]),
                        "--editar-json",
                        str(seleccionadas[0]),
                    ],
                    cwd=BASE,
                )
            except OSError as exc:
                messagebox.showerror(
                    "No se pudo abrir el editor",
                    str(exc),
                    parent=self.raiz,
                )
            return
        self.al_actualizar(seleccionadas[0])

    def abrir_resumen_alineador(self) -> None:
        ruta = BASE / "RESUMEN_ALINEADOR_ESPECIFICACIONES.md"
        try:
            abrir_en_windows = getattr(os, "startfile", None)
            if callable(abrir_en_windows):
                abrir_en_windows(str(ruta))
            elif not webbrowser.open(ruta.resolve().as_uri()):
                raise OSError("No hay una aplicación disponible para abrir el resumen.")
        except OSError as exc:
            messagebox.showerror(
                "No se pudo abrir el resumen del alineador",
                str(exc),
                parent=self.raiz,
            )

    def _alinear_seleccionadas(self, clave: str) -> None:
        if clave != "anteriores":
            return
        panel = self.paneles[clave]
        archivos = [
            especificacion.ruta
            for especificacion in panel.catalogo.especificaciones
            if especificacion.ruta in panel.documentos_marcados
        ]
        if not archivos:
            messagebox.showinfo(
                "Selecciona una especificación",
                "Marca una o más especificaciones de origen para alinearlas.",
                parent=self.raiz,
            )
            return
        self._alinear_especificaciones(archivos)

    def _realinear_seleccionadas(self, clave: str) -> None:
        if clave != "nuevas":
            return
        panel = self.paneles[clave]
        seleccionadas = [
            especificacion
            for especificacion in panel.catalogo.especificaciones
            if especificacion.ruta in panel.documentos_marcados
        ]
        pendientes = especificaciones_pendientes_realineacion(seleccionadas)
        if not pendientes:
            return

        originales = self.paneles["anteriores"].catalogo.especificaciones
        archivos_origen: list[Path] = []
        archivos_destino: list[Path] = []
        sin_original = []
        for alineada in pendientes:
            original = buscar_contraparte(alineada, originales)
            if original is None:
                sin_original.append(alineada.ruta.name)
                continue
            archivos_origen.append(original.ruta)
            archivos_destino.append(alineada.ruta)

        if sin_original:
            messagebox.showwarning(
                "No se encontró el origen",
                "No se pueden realinear estos JSON porque no se encontró su "
                "especificación de origen:\n" + "\n".join(sin_original),
                parent=self.raiz,
            )
        if archivos_origen:
            self._alinear_especificaciones(
                archivos_origen,
                destinos_seleccionados=archivos_destino,
            )

    def _comparar_seleccionadas(self, clave: str) -> None:
        if clave != "nuevas" or self.al_comparar is None:
            return
        rutas = [e.ruta for e in self.paneles[clave].catalogo.especificaciones if e.ruta in self.paneles[clave].documentos_marcados]
        if 2 <= len(rutas) <= 10:
            self.al_comparar(rutas)

    def _alinear_especificaciones(
        self,
        archivos: list[Path],
        destinos_seleccionados: list[Path] | None = None,
    ) -> None:
        if destinos_seleccionados is not None and len(archivos) != len(
            destinos_seleccionados
        ):
            raise ValueError("Cada origen debe tener su destino alineado.")
        if destinos_seleccionados is None:
            pares = [(ruta, None) for ruta in archivos]
        else:
            pares = list(zip(archivos, destinos_seleccionados))
        archivos_en_espera = {
            ruta
            for lote, _destinos in self._lotes_pendientes_alineacion
            for ruta in lote
        }
        pares_nuevos = []
        vistos: set[Path] = set()
        for ruta, destino in pares:
            if (
                ruta in vistos
                or ruta in self._archivos_actuales_alineacion
                or ruta in archivos_en_espera
            ):
                continue
            vistos.add(ruta)
            pares_nuevos.append((ruta, destino))
        archivos_nuevos = [ruta for ruta, _destino in pares_nuevos]
        destinos_nuevos = (
            [destino for _ruta, destino in pares_nuevos if destino is not None]
            if destinos_seleccionados is not None
            else None
        )
        if not archivos_nuevos:
            return
        if self._alineacion_en_proceso:
            self._lotes_pendientes_alineacion.append(
                (archivos_nuevos, destinos_nuevos)
            )
            return

        self._alineacion_en_proceso = True
        self._archivos_actuales_alineacion = set(archivos_nuevos)
        self._lotes_pendientes_alineacion.clear()
        self._resumen_alineacion = {
            "archivos": 0,
            "alineados": 0,
            "revision": 0,
            "sin_coincidencia": 0,
            "errores": 0,
        }
        self._cola_alineacion = queue.Queue()
        ventana = ui.Ventana(self.raiz)
        self._ventana_progreso_alineacion = ventana
        ventana.title("Alineando especificaciones")
        ventana.transient(self.raiz)
        ventana.resizable(False, False)
        ventana.geometry("430x130")
        ventana.protocol("WM_DELETE_WINDOW", lambda: None)
        marco = ttk.Frame(ventana, padding=16)
        marco.pack(fill="both", expand=True)
        self._etiqueta_progreso_alineacion = ttk.Label(
            marco, text="Preparando alineación…", wraplength=390
        )
        self._etiqueta_progreso_alineacion.pack(anchor="w", pady=(0, 10))
        self._barra_progreso_alineacion = ttk.Progressbar(
            marco, mode="determinate", maximum=len(archivos_nuevos)
        )
        self._barra_progreso_alineacion.pack(fill="x")
        self._ejecutar_tarea_alineacion(archivos_nuevos, destinos_nuevos)
        self.raiz.after(100, self._revisar_progreso_alineacion)

    def _ejecutar_tarea_alineacion(
        self,
        archivos: list[Path],
        destinos_seleccionados: list[Path] | None = None,
    ) -> None:
        if self._etiqueta_progreso_alineacion is not None:
            self._etiqueta_progreso_alineacion.configure(
                text=f"Alineando {len(archivos)} especificación(es)…"
            )
        if self._barra_progreso_alineacion is not None:
            self._barra_progreso_alineacion.configure(
                maximum=max(1, len(archivos)), value=0
            )
        cola = self._cola_alineacion
        if cola is None:
            raise RuntimeError("No está disponible la cola de progreso del alineador.")

        def trabajar() -> None:
            try:
                resumen = ejecutar_alineacion(
                    BASE,
                    archivos_seleccionados=archivos,
                    destinos_seleccionados=destinos_seleccionados,
                    al_progreso=lambda completados, total, nombre: cola.put(
                        ("progreso", (completados, total, nombre))
                    ),
                    carpeta_entrada=self.carpetas_json["anteriores"],
                    carpeta_salida=self.carpetas_json["nuevas"],
                )
            except Exception as exc:
                cola.put(("error", str(exc)))
            else:
                cola.put(("resultado", resumen))

        threading.Thread(
            target=trabajar,
            name="alineacion-especificaciones",
            daemon=True,
        ).start()

    def _revisar_progreso_alineacion(self) -> None:
        cola = self._cola_alineacion
        if cola is None:
            return
        while True:
            try:
                tipo, dato = cola.get_nowait()
            except queue.Empty:
                break
            if tipo == "progreso":
                completados, total, nombre = dato
                if self._barra_progreso_alineacion is not None:
                    self._barra_progreso_alineacion.configure(
                        maximum=max(1, total), value=completados
                    )
                if self._etiqueta_progreso_alineacion is not None:
                    self._etiqueta_progreso_alineacion.configure(
                        text=f"Alineando {completados} de {total}: {Path(nombre).name}"
                    )
            elif tipo == "error":
                self._cerrar_progreso_alineacion()
                messagebox.showerror(
                    "No se pudo alinear la especificación",
                    str(dato),
                    parent=self.raiz,
                )
                return
            elif tipo == "resultado":
                for campo, cantidad in dato.items():
                    self._resumen_alineacion[campo] = (
                        self._resumen_alineacion.get(campo, 0) + cantidad
                    )
                if self._lotes_pendientes_alineacion:
                    lote_pendiente = self._lotes_pendientes_alineacion.pop(0)
                    archivos_pendientes, destinos_pendientes = lote_pendiente
                    self._archivos_actuales_alineacion = set(archivos_pendientes)
                    self._ejecutar_tarea_alineacion(
                        archivos_pendientes,
                        destinos_pendientes,
                    )
                    continue
                resumen = self._resumen_alineacion
                self._cerrar_progreso_alineacion()
                self.recargar_catalogos()
                mostrar = (
                    messagebox.showwarning
                    if resumen.get("errores", 0)
                    else messagebox.showinfo
                )
                mostrar(
                    "Alineación terminada con errores"
                    if resumen.get("errores", 0)
                    else "Alineación terminada",
                    (
                        f"Archivos procesados: {resumen.get('archivos', 0)}\n"
                        f"Pruebas alineadas: {resumen.get('alineados', 0)}\n"
                        f"Para revisión: {resumen.get('revision', 0)}\n"
                        f"Sin coincidencia: {resumen.get('sin_coincidencia', 0)}\n"
                        f"Errores: {resumen.get('errores', 0)}"
                    ),
                    parent=self.raiz,
                )
                return
        if self._alineacion_en_proceso:
            self.raiz.after(100, self._revisar_progreso_alineacion)

    def _cerrar_progreso_alineacion(self) -> None:
        ventana = self._ventana_progreso_alineacion
        self._ventana_progreso_alineacion = None
        self._etiqueta_progreso_alineacion = None
        self._barra_progreso_alineacion = None
        self._cola_alineacion = None
        self._archivos_actuales_alineacion.clear()
        self._lotes_pendientes_alineacion.clear()
        self._alineacion_en_proceso = False
        if ventana is not None and ventana.winfo_exists():
            ventana.destroy()

    def _editar_especificacion(self, clave: str) -> None:
        especificacion = self.paneles[clave].especificacion_actual
        if especificacion is None:
            return
        if self.al_editar is None:
            aplicacion = BASE / "interfaz_aplicacion_especificaciones.py"
            try:
                subprocess.Popen(
                    [
                        sys.executable,
                        *([] if getattr(sys, "frozen", False) else [str(aplicacion)]),
                        "--generar-json",
                        str(especificacion.ruta),
                    ],
                    cwd=BASE,
                )
            except OSError as exc:
                messagebox.showerror(
                    "No se pudo abrir el generador",
                    str(exc),
                    parent=self.raiz,
                )
            return
        self.al_editar(especificacion.ruta)

    def _mostrar_error(self, clave: str, mensaje: str) -> None:
        panel = self.paneles[clave]
        panel.widgets_previa.clear()
        panel.firma_vista = None
        panel.separador_previa = None
        panel.area_detalles = None
        panel.posicion_detalles = None
        if panel.boton_compactar is not None:
            panel.boton_compactar.configure(state="disabled")
        panel.canvas_previa = None
        for widget in panel.marco_previa.winfo_children():
            widget.destroy()
        ttk.Label(
            panel.marco_previa,
            text=f"No se puede mostrar la vista previa:\n\n{mensaje}",
            foreground="#a12b2b",
            justify="left",
            wraplength=700,
        ).pack(anchor="nw", padx=18, pady=18)

    def _vincular_scroll_preview(
        self,
        panel: PanelEspecificaciones,
        pagina: ttk.Frame,
        tabla: ttk.Treeview,
        barra_tabla: ttk.Scrollbar,
    ) -> None:
        def vincular(widget: tk.Widget) -> None:
            if not getattr(widget, "_rueda_selector_vinculada", False):
                widget._rueda_selector_vinculada = True
                widget.bind(
                    "<MouseWheel>",
                    lambda evento, origen=widget: self._desplazar_previa(
                        evento, panel, tabla, barra_tabla, origen
                    ),
                )
            for hijo in widget.winfo_children():
                vincular(hijo)

        vincular(pagina)
        vincular(tabla.master)

    @staticmethod
    def _es_descendiente(widget: tk.Widget, ancestro: tk.Widget) -> bool:
        ruta_ancestro = str(ancestro)
        actual = widget
        while True:
            if str(actual) == ruta_ancestro:
                return True
            padre = actual.winfo_parent()
            if not padre:
                return False
            actual = actual.nametowidget(padre)

    @staticmethod
    def _unidades_rueda(evento: tk.Event) -> int:
        unidades = int(-evento.delta / 120)
        if unidades == 0:
            return -1 if evento.delta > 0 else 1
        return unidades

    def _desplazar_previa(
        self,
        evento: tk.Event,
        panel: PanelEspecificaciones,
        tabla: ttk.Treeview,
        barra_tabla: ttk.Scrollbar,
        origen: tk.Widget,
    ) -> str:
        unidades = self._unidades_rueda(evento)
        if self._es_descendiente(origen, tabla) or self._es_descendiente(
            origen, barra_tabla
        ):
            tabla.yview_scroll(unidades, "units")
        elif panel.canvas_previa is not None:
            panel.canvas_previa.yview_scroll(unidades, "units")
        return "break"


def main() -> None:
    raiz = ui.Raiz()
    SelectorEspecificaciones(raiz)
    raiz.mainloop()


if __name__ == "__main__":
    main()
