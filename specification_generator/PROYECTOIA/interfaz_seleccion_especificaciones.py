#!/usr/bin/env python3
"""Selector con busqueda y vista previa de especificaciones en espanol."""
from __future__ import annotations

import json
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Callable

from alinear_especificaciones_json import (
    MAESTRO,
    cargar_maestro,
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


BASE = Path(__file__).resolve().parent
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


def etiqueta_control(campo: str, etiqueta: str, idioma: str) -> str:
    if idioma == "INGLES":
        return CAMPOS_CONTROL_INGLES.get(campo, etiqueta)
    return etiqueta


def crear_divisor_redimensionable(contenedor: tk.Misc, orientacion: str) -> tk.PanedWindow:
    cursor = "sb_v_double_arrow" if orientacion == tk.HORIZONTAL else "sb_h_double_arrow"
    return tk.PanedWindow(
        contenedor,
        orient=orientacion,
        sashwidth=5,
        handlesize=0,
        showhandle=False,
        sashrelief=tk.FLAT,
        background="#f0f0f0",
        borderwidth=0,
        relief=tk.FLAT,
        cursor=cursor,
        opaqueresize=False,
    )


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
    zoom_previa: float = 1.0
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


@lru_cache(maxsize=1)
def _referencias_maestras_astm() -> frozenset[str]:
    registros = cargar_maestro(BASE / MAESTRO)
    return frozenset(
        referencia
        for registro in registros
        for referencia in referencias_astm(registro.get("METODO", ""))
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


def cargar_catalogo(carpeta: Path) -> Catalogo:
    """Indexa JSON en español e inglés; reporta archivos inválidos."""
    if not carpeta.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de especificaciones: {carpeta}")

    especificaciones = []
    errores = []
    for ruta in detectar_json(carpeta):
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
                )
            )
        except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
            errores.append(f"{ruta.name}: {exc}")

    return Catalogo(
        especificaciones=tuple(especificaciones),
        errores=tuple(errores),
    )


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


def preparar_vista_previa(ruta: Path) -> VistaPrevia:
    """Interpreta un JSON con las mismas transformaciones del generador Word."""
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


def preparar_vista_previa_alineada(ruta: Path) -> VistaPrevia:
    """Muestra exclusivamente la instantanea de valores guardada en la auditoria alineada."""
    datos = leer_datos_json(ruta)
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
        raiz: tk.Tk,
        carpeta_json: Path = JSON_DIR_NUEVAS,
        carpeta_json_anterior: Path = JSON_DIR_ANTERIORES,
        al_actualizar: Callable[[Path], None] | None = None,
    ):
        self.raiz = raiz
        self.al_actualizar = al_actualizar
        self.carpetas_json = {
            "anteriores": carpeta_json_anterior,
            "nuevas": carpeta_json,
        }
        self.paneles: dict[str, PanelEspecificaciones] = {}
        self.zoom_aplicacion = 1.0
        self._escala_tk_base = float(self.raiz.tk.call("tk", "scaling"))
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
        estilo.configure("Seccion.TLabel", font=("Segoe UI", 12, "bold"), foreground="#163b70")
        estilo.configure("Treeview", rowheight=25, font=("Segoe UI", 9))
        estilo.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        encabezado = tk.Frame(self.raiz, background="#0752b9", padx=24, pady=15)
        encabezado.pack(fill="x")
        controles_zoom = ttk.Frame(encabezado)
        controles_zoom.pack(side="right", padx=(12, 0))
        ttk.Button(
            controles_zoom, text="A−", width=4, command=lambda: self._cambiar_zoom_aplicacion(-0.1)
        ).pack(side="left", padx=2)
        ttk.Button(
            controles_zoom, text="A+", width=4, command=lambda: self._cambiar_zoom_aplicacion(0.1)
        ).pack(side="left", padx=2)
        ttk.Label(
            encabezado,
            text="SELECCIONAR ESPECIFICACIÓN",
            style="Titulo.TLabel",
            background="#0752b9",
        ).pack(anchor="w")
        ttk.Label(
            encabezado,
            text="Compara las especificaciones de origen con las nuevas versiones alineadas.",
            style="Subtitulo.TLabel",
            background="#0752b9",
        ).pack(anchor="w", pady=(3, 0))

        self.pestanas = ttk.Notebook(self.raiz)
        self.pestanas.pack(fill="both", expand=True, padx=18, pady=16)
        for clave, titulo in (
            ("anteriores", "Anteriores · JSON extraídos"),
            ("nuevas", "Nuevas · JSON alineados"),
        ):
            self._crear_pestana(clave, titulo)
        self.pestanas.select(1)

    def _crear_pestana(self, clave: str, titulo: str) -> None:
        pestaña = ttk.Frame(self.pestanas)
        self.pestanas.add(pestaña, text=titulo)
        contenido = ttk.Panedwindow(pestaña, orient="horizontal")
        contenido.pack(fill="both", expand=True, padx=4, pady=8)
        panel_busqueda = ttk.Frame(contenido, padding=12)
        panel_previa = ttk.Frame(contenido, padding=12)
        contenido.add(panel_busqueda, weight=1)
        contenido.add(panel_previa, weight=4)

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
            command=self.recargar_catalogos,
        ).pack(fill="x", pady=(10, 0))
        estado = ttk.Label(panel_busqueda, text="", wraplength=280)
        estado.pack(anchor="w", pady=(8, 0))

        barra_vista = ttk.Frame(panel_previa)
        barra_vista.pack(fill="x", pady=(0, 9))
        ttk.Label(
            barra_vista,
            text=(
                "Vista previa · Datos extraídos"
                if clave == "anteriores"
                else "Vista previa · JSON alineado (campos actualizados)"
            ),
            style="Seccion.TLabel",
        ).pack(side="left", anchor="w")
        boton_contraparte = ttk.Button(
            barra_vista,
            text="Ver alineada →" if clave == "anteriores" else "← Ver original",
            width=16,
            command=lambda c=clave: self._ir_a_contraparte(c),
            state="disabled",
        )
        boton_contraparte.pack(side="left", padx=(10, 0))
        ttk.Button(
            barra_vista,
            text="−",
            width=3,
            command=lambda c=clave: self._cambiar_zoom_previa(c, -0.1),
        ).pack(side="right", padx=(3, 0))
        ttk.Button(
            barra_vista,
            text="+",
            width=3,
            command=lambda c=clave: self._cambiar_zoom_previa(c, 0.1),
        ).pack(side="right")
        ttk.Button(
            barra_vista,
            text="×",
            width=3,
            command=lambda c=clave: self._cerrar_vista_previa(c),
        ).pack(side="right", padx=(0, 6))
        marco_previa = ttk.Frame(panel_previa)
        marco_previa.pack(fill="both", expand=True)
        ttk.Label(
            marco_previa,
            text="Selecciona una especificación para ver cómo se interpreta en el formato.",
            justify="center",
            anchor="center",
            wraplength=650,
        ).pack(fill="both", expand=True)

        ttk.Button(
            panel_previa,
            text="Actualizar especificación",
            command=lambda c=clave: self._actualizar_especificacion(c),
        ).pack(anchor="e", pady=(10, 0))

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
        )
        entrada.bind("<KeyRelease>", lambda _evento, c=clave: self._actualizar_sugerencias(c))
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
            self._actualizar_boton_contraparte(clave)

    def _actualizar_sugerencias(self, clave: str) -> None:
        panel = self.paneles[clave]
        panel.sugerencias = buscar_especificaciones(
            panel.catalogo.especificaciones,
            panel.entrada_busqueda.get(),
        )
        panel.lista.delete(0, tk.END)
        for especificacion in panel.sugerencias:
            marca = "☑" if especificacion.ruta in panel.documentos_marcados else "☐"
            idioma = " · English" if especificacion.idioma == "INGLES" else ""
            panel.lista.insert(tk.END, f"{marca}  {especificacion.etiqueta}{idioma}")
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
        desplazamiento = panel.lista.yview()[0]
        self._actualizar_sugerencias(clave)
        panel.lista.yview_moveto(desplazamiento)
        panel.lista.selection_set(indice)
        panel.lista.activate(indice)
        self._mostrar_previa(clave, especificacion)

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

    def _mostrar_previa(self, clave: str, especificacion: Especificacion) -> None:
        panel = self.paneles[clave]
        panel.especificacion_actual = especificacion
        self._actualizar_boton_contraparte(clave)
        try:
            preparar = (
                preparar_vista_previa
                if clave == "anteriores"
                else preparar_vista_previa_alineada
            )
            previa = preparar(especificacion.ruta)
        except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
            self._mostrar_error(clave, str(exc))
            return

        for widget in panel.marco_previa.winfo_children():
            widget.destroy()
        separador = crear_divisor_redimensionable(panel.marco_previa, tk.VERTICAL)
        separador.pack(fill="both", expand=True)
        area_tabla = ttk.Frame(separador, padding=10)
        area_detalles = ttk.Frame(separador)
        separador.add(area_tabla, minsize=120, stretch="always")
        separador.add(area_detalles, minsize=80, stretch="always")
        separador_inicializado = False

        def posicionar_separador(evento: tk.Event) -> None:
            nonlocal separador_inicializado
            if not separador_inicializado and evento.height > 100:
                separador_inicializado = True
                separador.sash_place(0, 0, round(evento.height * 0.68))

        separador.bind("<Configure>", posicionar_separador)
        canvas = tk.Canvas(area_detalles, background="#edf2f7", highlightthickness=0)
        barra = ttk.Scrollbar(area_detalles, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=barra.set)
        canvas.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        panel.canvas_previa = canvas
        pagina = ttk.Frame(canvas, padding=20)
        ventana = canvas.create_window((0, 0), window=pagina, anchor="nw")
        pagina.bind(
            "<Configure>",
            lambda _evento: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda evento: canvas.itemconfigure(ventana, width=evento.width),
        )

        escala_fuente = panel.zoom_previa
        ttk.Label(
            area_tabla,
            text=previa.nombre_producto,
            font=("Segoe UI", 19, "bold"),
            foreground="#163b70",
            wraplength=780,
        ).pack(anchor="w", pady=(0, 4))
        ttk.Label(
            area_tabla,
            text=(
                f"Tests ({len(previa.pruebas)})"
                if previa.idioma == "INGLES"
                else f"Pruebas ({len(previa.pruebas)})"
            ),
            font=("Segoe UI", 12, "bold"),
            foreground="#163b70",
        ).pack(anchor="w", pady=(0, 7))
        marco_tabla = ttk.Frame(area_tabla)
        marco_tabla.pack(fill="both", expand=True)
        marco_tabla.rowconfigure(0, weight=1)
        marco_tabla.columnconfigure(0, weight=1)
        columnas = tuple(campo for _, campo in CAMPOS_PRUEBA)
        estilo_tabla = ttk.Style(self.raiz)
        estilo_nombre = f"{clave}.Preview.Treeview"
        estilo_nombre_encabezado = f"{estilo_nombre}.Heading"
        estilo_tabla.configure(
            estilo_nombre,
            font=("Segoe UI", max(7, round(9 * escala_fuente))),
            rowheight=max(18, round(25 * escala_fuente)),
        )
        estilo_tabla.configure(
            estilo_nombre_encabezado,
            font=("Segoe UI", max(7, round(9 * escala_fuente)), "bold"),
        )
        tabla = ttk.Treeview(
            marco_tabla,
            columns=columnas,
            show="headings",
            height=min(max(len(previa.pruebas), 5), 18),
            style=estilo_nombre,
        )
        panel.tabla_pruebas = tabla
        encabezados = {
            campo: etiqueta for etiqueta, campo in encabezados_tabla(previa.idioma)
        }
        anchos = {
            "prueba": 240,
            "unidades": 95,
            "metodo": 155,
            "especificacion": 210,
            "valor_tipico": 130,
        }
        for campo in columnas:
            tabla.heading(campo, text=encabezados[campo])
            tabla.column(campo, width=anchos[campo], minwidth=75, stretch=True)
        barra_tabla = ttk.Scrollbar(marco_tabla, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=barra_tabla.set)
        barra_tabla_x = ttk.Scrollbar(marco_tabla, orient="horizontal", command=tabla.xview)
        tabla.configure(xscrollcommand=barra_tabla_x.set)
        tabla.grid(row=0, column=0, sticky="nsew")
        barra_tabla.grid(row=0, column=1, sticky="ns")
        barra_tabla_x.grid(row=1, column=0, sticky="ew")
        for prueba in previa.pruebas:
            tabla.insert(
                "",
                tk.END,
                values=tuple(_texto(prueba.get(campo)) for campo in columnas),
            )

        if previa.idioma == "INGLES":
            ttk.Label(
                pagina,
                text="English version",
                font=("Segoe UI", 9, "italic"),
                foreground="#526174",
            ).pack(anchor="w")
        ttk.Label(
            pagina,
            text=f"Origen: {especificacion.ruta.relative_to(self.carpetas_json[clave])}",
            foreground="#526174",
        ).pack(anchor="w", pady=(0, 12))

        if previa.mensajes_control:
            marco_control_error = ttk.LabelFrame(
                pagina, text="Avisos del control documental", padding=10
            )
            marco_control_error.pack(fill="x", pady=(0, 14))
            for mensaje in previa.mensajes_control:
                ttk.Label(
                    marco_control_error,
                    text=f"• {mensaje}",
                    foreground="#9a3412",
                    wraplength=780,
                    justify="left",
                ).pack(anchor="w", pady=2)

        if previa.mensajes_astm_no_encontradas:
            marco_astm = ttk.LabelFrame(
                pagina,
                text="Pruebas ASTM no encontradas en la lista maestra",
                padding=10,
            )
            marco_astm.pack(fill="x", pady=(0, 14))
            for mensaje in previa.mensajes_astm_no_encontradas:
                ttk.Label(
                    marco_astm,
                    text=f"• {mensaje}",
                    foreground="#9a3412",
                    wraplength=780,
                    justify="left",
                ).pack(anchor="w", pady=2)

        marco_control = ttk.LabelFrame(
            pagina,
            text="Document control" if previa.idioma == "INGLES" else "Control documental",
            padding=10,
        )
        marco_control.pack(fill="x", pady=(0, 14))
        for indice, (etiqueta, valor) in enumerate(previa.control):
            fila, columna = divmod(indice, 2)
            ttk.Label(marco_control, text=f"{etiqueta}:", font=("Segoe UI", 9, "bold")).grid(
                row=fila, column=columna * 2, sticky="w", padx=(4, 5), pady=4
            )
            ttk.Label(
                marco_control,
                text=valor,
                wraplength=280,
            ).grid(row=fila, column=columna * 2 + 1, sticky="w", padx=(0, 18), pady=4)
        marco_control.columnconfigure(1, weight=1)
        marco_control.columnconfigure(3, weight=1)

        if clave == "nuevas":
            ttk.Label(
                pagina,
                text=(
                    "This table uses valores_alineados from "
                    "alineacion_metodos_lab.detalle."
                    if previa.idioma == "INGLES"
                    else "Esta tabla se construye con los valores_alineados guardados "
                    "en alineacion_metodos_lab.detalle del JSON alineado."
                ),
                foreground="#526174",
                wraplength=800,
            ).pack(anchor="w", pady=(0, 7))

        if previa.notas:
            ttk.Label(
                pagina,
                text="Notas",
                font=("Segoe UI", 11, "bold"),
                foreground="#163b70",
            ).pack(anchor="w", pady=(14, 4))
            for nota in previa.notas:
                ttk.Label(pagina, text=f"• {nota}", wraplength=800).pack(anchor="w", pady=2)
        if previa.pie_pagina:
            ttk.Separator(pagina).pack(fill="x", pady=(16, 8))
            ttk.Label(
                pagina,
                text=previa.pie_pagina,
                wraplength=800,
                foreground="#526174",
            ).pack(anchor="w")
        _aplicar_zoom_vista_previa(pagina, escala_fuente)
        self._vincular_scroll_preview(panel, pagina, tabla, barra_tabla)

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
                if sugerencia.ruta == contraparte.ruta
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
                    if sugerencia.ruta == contraparte.ruta
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
        self.zoom_aplicacion = min(1.6, max(0.8, self.zoom_aplicacion + cambio))
        self.raiz.tk.call(
            "tk",
            "scaling",
            self._escala_tk_base * self.zoom_aplicacion,
        )
        for clave, panel in self.paneles.items():
            if panel.especificacion_actual is not None:
                self._mostrar_previa(clave, panel.especificacion_actual)

    def _cambiar_zoom_previa(self, clave: str, cambio: float) -> None:
        panel = self.paneles[clave]
        panel.zoom_previa = min(1.8, max(0.7, panel.zoom_previa + cambio))
        if panel.especificacion_actual is not None:
            self._mostrar_previa(clave, panel.especificacion_actual)

    def _cerrar_vista_previa(self, clave: str) -> None:
        panel = self.paneles[clave]
        panel.especificacion_actual = None
        self._actualizar_boton_contraparte(clave)
        panel.canvas_previa = None
        panel.tabla_pruebas = None
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
                        str(aplicacion),
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

    def _mostrar_error(self, clave: str, mensaje: str) -> None:
        panel = self.paneles[clave]
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
        for widget in (tabla, barra_tabla):
            widget.bind(
                "<MouseWheel>",
                lambda evento, origen=widget: self._desplazar_previa(
                    evento, panel, tabla, barra_tabla, origen
                ),
            )

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
    raiz = tk.Tk()
    SelectorEspecificaciones(raiz)
    raiz.mainloop()


if __name__ == "__main__":
    main()
