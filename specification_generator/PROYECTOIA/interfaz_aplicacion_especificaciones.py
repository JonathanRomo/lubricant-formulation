#!/usr/bin/env python3
"""Interfaz principal para cargar, editar, seleccionar y comparar especificaciones."""
from __future__ import annotations

import json
import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from typing import Any
from zipfile import BadZipFile

from docx.opc.exceptions import PackageNotFoundError
from interfaz_seleccion_especificaciones import (
    crear_divisor_redimensionable,
    encabezados_tabla,
    preparar_vista_previa,
    preparar_vista_previa_alineada,
)


BASE = Path(__file__).resolve().parent
RUTAS_PREDETERMINADAS = {
    "especificaciones_viejas": BASE / "ESPECIFICACIONES_VIEJAS",
    "json_anteriores": BASE / "ESPECIFICACIONES_JSON",
    "json_alineadas": BASE / "ESPECIFICACIONES_JSON_ALINEADAS",
    "formatos": BASE / "FORMATOS",
    "historico": BASE / "HISTORICO_ESPECIFICACIONES",
}
ARCHIVO_CONFIGURACION = BASE / "configuracion_interfaz.json"
CAMPOS_TABLA = ("prueba", "unidades", "metodo", "especificacion", "valor_tipico")
EXTENSIONES_WORD = (".doc", ".docm", ".docx")


def crear_json_temporal_comparacion(
    archivo_word: Path,
    carpeta_temporal: Path,
) -> Path:
    """Extrae Word con el gestor sin escribir en su registro ni en sus carpetas."""
    import gestor_especificaciones as gestor

    lectura, _ = gestor.preparar(archivo_word, carpeta_temporal)
    idioma = "en" if gestor.contiene_marcadores_ingles(lectura) else "es"
    datos = gestor.extraer(lectura, idioma)
    ruta_json = carpeta_temporal / "comparacion_temporal.json"
    ruta_json.write_text(
        json.dumps(datos, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return ruta_json


def crear_documento_alineado_manual(datos: dict[str, Any]) -> dict[str, Any]:
    """Valida un JSON de especificación y crea los valores que consume la vista alineada."""
    resultado = dict(datos)
    producto = resultado.get("producto")
    if (
        not isinstance(producto, dict)
        or not isinstance(producto.get("nombre"), str)
        or not producto["nombre"].strip()
    ):
        raise ValueError("Se requiere 'producto.nombre' como texto.")
    pruebas = resultado.get("pruebas")
    if not isinstance(pruebas, list) or any(not isinstance(fila, dict) for fila in pruebas):
        raise ValueError("'pruebas' debe ser una lista de objetos.")

    alineacion_existente = resultado.get("alineacion_metodos_lab")
    detalle_existente = (
        alineacion_existente.get("detalle")
        if isinstance(alineacion_existente, dict)
        else None
    )
    if isinstance(alineacion_existente, dict) and isinstance(detalle_existente, list):
        tiene_snapshot = (
            len(detalle_existente) == len(pruebas)
            and all(
                isinstance(detalle, dict)
                and isinstance(detalle.get("valores_alineados"), dict)
                and all(campo in detalle["valores_alineados"] for campo in CAMPOS_TABLA)
                for detalle in detalle_existente
            )
        )
        if tiene_snapshot:
            return resultado
        raise ValueError(
            "La auditoría de este JSON no tiene valores_alineados completos. "
            "Edítalos en alineacion_metodos_lab.detalle antes de guardar para "
            "no reemplazarlos con valores extraídos."
        )

    detalle = []
    for indice, fila in enumerate(pruebas, 1):
        valores = {}
        for campo in CAMPOS_TABLA:
            valor = fila.get(campo)
            if valor is None:
                valor = ""
            elif not isinstance(valor, (str, int, float, bool)):
                raise ValueError(f"pruebas[{indice}].{campo} debe ser un valor simple.")
            valores[campo] = str(valor)
        detalle.append(
            {
                "numero_prueba": indice,
                "estado": "EDICION_MANUAL",
                "valores_alineados": valores,
            }
        )

    resultado["alineacion_metodos_lab"] = {
        "version_programa": "EDICION_MANUAL",
        "detalle": detalle,
        "mensajes_astm_no_encontradas": [],
    }
    return resultado


def _cargar_configuracion() -> dict[str, Path]:
    rutas = dict(RUTAS_PREDETERMINADAS)
    if not ARCHIVO_CONFIGURACION.exists():
        return rutas
    datos = json.loads(ARCHIVO_CONFIGURACION.read_text(encoding="utf-8"))
    if not isinstance(datos, dict):
        raise ValueError("La configuración de carpetas debe ser un objeto JSON.")
    for clave, predeterminada in RUTAS_PREDETERMINADAS.items():
        valor = datos.get(clave)
        if isinstance(valor, str) and valor.strip():
            ruta = Path(valor).expanduser()
            rutas[clave] = ruta if ruta.is_absolute() else (BASE / ruta).resolve()
        else:
            rutas[clave] = predeterminada
    return rutas


def _guardar_configuracion(rutas: dict[str, Path]) -> None:
    temporal = ARCHIVO_CONFIGURACION.with_suffix(".tmp")
    temporal.write_text(
        json.dumps({clave: str(ruta) for clave, ruta in rutas.items()}, indent=2),
        encoding="utf-8",
    )
    temporal.replace(ARCHIVO_CONFIGURACION)


def _abrir_carpeta(ruta: Path) -> None:
    if not ruta.is_dir():
        raise FileNotFoundError(f"No existe la carpeta:\n{ruta}")
    abrir_en_windows = getattr(os, "startfile", None)
    if callable(abrir_en_windows):
        abrir_en_windows(str(ruta))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(ruta)])
    else:
        subprocess.Popen(["xdg-open", str(ruta)])


def mostrar_vista_json(contenedor: tk.Misc, ruta: Path, zoom: float = 1.0) -> None:
    """Renderiza un JSON antiguo o alineado en un panel con zoom propio."""
    for widget in contenedor.winfo_children():
        widget.destroy()
    marco = ttk.Frame(contenedor)
    marco.pack(fill="both", expand=True)
    try:
        datos = json.loads(ruta.read_text(encoding="utf-8-sig"))
        alineado = (
            ruta.parent.name.casefold() == "especificaciones_json_alineadas".casefold()
            or isinstance(datos, dict) and isinstance(datos.get("alineacion_metodos_lab"), dict)
        )
        previa = (
            preparar_vista_previa_alineada(ruta)
            if alineado
            else preparar_vista_previa(ruta)
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
        ttk.Label(
            marco,
            text=f"No se pudo mostrar la especificación:\n{exc}",
            foreground="#9a3412",
            wraplength=440,
            justify="left",
        ).pack(anchor="w", padx=8, pady=8)
        return

    separador = crear_divisor_redimensionable(marco, tk.VERTICAL)
    separador.pack(fill="both", expand=True)
    area_tabla = ttk.Frame(separador, padding=8)
    area_detalles = ttk.Frame(separador)
    separador.add(area_tabla, minsize=120, stretch="always")
    separador.add(area_detalles, minsize=80, stretch="always")
    area_tabla.rowconfigure(2, weight=1)
    area_tabla.columnconfigure(0, weight=1)
    separador_inicializado = False

    def posicionar_separador(evento: tk.Event) -> None:
        nonlocal separador_inicializado
        if not separador_inicializado and evento.height > 100:
            separador_inicializado = True
            separador.sash_place(0, 0, round(evento.height * 0.68))

    separador.bind("<Configure>", posicionar_separador)
    ttk.Label(
        area_tabla,
        text=previa.nombre_producto,
        font=("Segoe UI", max(11, round(17 * zoom)), "bold"),
        foreground="#163b70",
        wraplength=680,
    ).grid(row=0, column=0, sticky="w", pady=(0, 4))
    estilo = ttk.Style(contenedor)
    sufijo = str(id(contenedor))
    estilo_filas = f"Preview{sufijo}.Treeview"
    estilo.configure(
        estilo_filas,
        font=("Segoe UI", max(7, round(9 * zoom))),
        rowheight=max(18, round(24 * zoom)),
    )
    estilo.configure(
        f"{estilo_filas}.Heading",
        font=("Segoe UI", max(7, round(9 * zoom)), "bold"),
    )
    ttk.Label(
        area_tabla,
        text=(
            f"Tests ({len(previa.pruebas)})"
            if previa.idioma == "INGLES"
            else f"Pruebas ({len(previa.pruebas)})"
        ),
        font=("Segoe UI", max(9, round(11 * zoom)), "bold"),
        foreground="#163b70",
    ).grid(row=1, column=0, sticky="w", pady=(0, 5))
    tabla = ttk.Treeview(
        area_tabla,
        columns=CAMPOS_TABLA,
        show="headings",
        height=min(max(len(previa.pruebas), 4), 16),
        style=estilo_filas,
    )
    for titulo, campo in encabezados_tabla(previa.idioma):
        tabla.heading(campo, text=titulo)
        tabla.column(campo, width=max(100, round(145 * zoom)), minwidth=70, stretch=True)
    barra_tabla_y = ttk.Scrollbar(area_tabla, orient="vertical", command=tabla.yview)
    barra_tabla_x = ttk.Scrollbar(area_tabla, orient="horizontal", command=tabla.xview)
    tabla.configure(yscrollcommand=barra_tabla_y.set, xscrollcommand=barra_tabla_x.set)
    tabla.grid(row=2, column=0, sticky="nsew")
    barra_tabla_y.grid(row=2, column=1, sticky="ns")
    barra_tabla_x.grid(row=3, column=0, sticky="ew")
    for prueba in previa.pruebas:
        tabla.insert("", tk.END, values=tuple(str(prueba.get(campo, "")) for campo in CAMPOS_TABLA))

    canvas = tk.Canvas(area_detalles, highlightthickness=0, background="#f4f6f8")
    vertical = ttk.Scrollbar(area_detalles, orient="vertical", command=canvas.yview)
    horizontal = ttk.Scrollbar(area_detalles, orient="horizontal", command=canvas.xview)
    canvas.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    canvas.grid(row=0, column=0, sticky="nsew")
    vertical.grid(row=0, column=1, sticky="ns")
    horizontal.grid(row=1, column=0, sticky="ew")
    area_detalles.rowconfigure(0, weight=1)
    area_detalles.columnconfigure(0, weight=1)
    pagina = ttk.Frame(canvas, padding=12)
    ventana = canvas.create_window((0, 0), window=pagina, anchor="nw")
    pagina.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfigure(ventana, width=max(e.width, 640)))

    if previa.idioma == "INGLES":
        ttk.Label(
            pagina,
            text="English version",
            font=("Segoe UI", max(8, round(9 * zoom)), "italic"),
            foreground="#526174",
        ).pack(anchor="w", pady=(0, 4))
    ttk.Label(pagina, text=str(ruta), wraplength=680).pack(anchor="w", pady=(0, 8))
    for etiqueta, valor in previa.control:
        ttk.Label(
            pagina,
            text=f"{etiqueta}: {valor}",
            font=("Segoe UI", max(8, round(9 * zoom))),
            wraplength=680,
        ).pack(anchor="w")

    avisos = (*previa.mensajes_control, *previa.mensajes_astm_no_encontradas)
    for aviso in avisos:
        ttk.Label(
            pagina,
            text=aviso,
            foreground="#9a3412",
            font=("Segoe UI", max(8, round(9 * zoom))),
            wraplength=680,
        ).pack(anchor="w", pady=(5, 0))

    if previa.notas:
        ttk.Label(
            pagina,
            text=(
                "Notes: " if previa.idioma == "INGLES" else "Notas: "
            ) + " · ".join(previa.notas),
            wraplength=680,
            font=("Segoe UI", max(8, round(9 * zoom))),
        ).pack(anchor="w", pady=(8, 0))


class EditorJson:
    def __init__(
        self,
        app: "AplicacionEspecificaciones",
        archivo_inicial: Path | None = None,
    ):
        self.app = app
        self.ventana = tk.Toplevel(app.raiz)
        self.ventana.title("Generar o editar JSON alineado")
        self.ventana.geometry("1050x720")
        self.ventana.minsize(700, 480)
        barra = ttk.Frame(self.ventana, padding=8)
        barra.pack(fill="x")
        ttk.Button(barra, text="Nuevo", command=self.nuevo).pack(side="left", padx=3)
        ttk.Button(barra, text="Abrir JSON", command=self.abrir).pack(side="left", padx=3)
        ttk.Button(barra, text="Validar", command=self.validar).pack(side="left", padx=3)
        ttk.Button(barra, text="Guardar alineada", command=self.guardar).pack(side="left", padx=3)
        self.texto = ScrolledText(self.ventana, wrap="none", undo=True, font=("Consolas", 10))
        self.texto.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        if archivo_inicial is None:
            self.nuevo()
        else:
            self._abrir_ruta(archivo_inicial)

    def nuevo(self) -> None:
        plantilla = {
            "esquema": "especificacion_producto_es_v3",
            "producto": {"nombre": ""},
            "documento": {"tipo": "ESPECIFICACION", "idioma": "es"},
            "control_documental": {"codigo_unico_producto": "", "revision": "00", "fecha": ""},
            "pruebas": [
                {
                    "prueba": "",
                    "unidades": "",
                    "metodo": "",
                    "especificacion": "",
                    "valor_tipico": "",
                }
            ],
            "notas": [],
        }
        self._poner_texto(plantilla)

    def _poner_texto(self, datos: dict[str, Any]) -> None:
        self.texto.delete("1.0", tk.END)
        self.texto.insert("1.0", json.dumps(datos, ensure_ascii=False, indent=2))

    def _leer(self) -> dict[str, Any]:
        try:
            datos = json.loads(self.texto.get("1.0", "end-1c"))
        except json.JSONDecodeError as exc:
            raise ValueError(f"JSON inválido en línea {exc.lineno}, columna {exc.colno}: {exc.msg}") from exc
        if not isinstance(datos, dict):
            raise ValueError("La raíz del JSON debe ser un objeto.")
        return crear_documento_alineado_manual(datos)

    def validar(self) -> None:
        try:
            self._leer()
        except ValueError as exc:
            messagebox.showerror("JSON no válido", str(exc), parent=self.ventana)
            return
        messagebox.showinfo("Validación", "El JSON es válido y sus pruebas se incluirán en detalle.valores_alineados.", parent=self.ventana)

    def abrir(self) -> None:
        ruta = filedialog.askopenfilename(
            parent=self.ventana,
            initialdir=self.app.rutas["json_alineadas"],
            filetypes=(("Archivos JSON", "*.json"), ("Todos los archivos", "*.*")),
        )
        if not ruta:
            return
        self._abrir_ruta(Path(ruta))

    def _abrir_ruta(self, ruta: Path) -> None:
        try:
            datos = json.loads(ruta.read_text(encoding="utf-8-sig"))
            if not isinstance(datos, dict):
                raise ValueError("La raíz del JSON debe ser un objeto.")
            self._poner_texto(datos)
        except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
            messagebox.showerror("No se pudo abrir", str(exc), parent=self.ventana)

    def guardar(self) -> None:
        try:
            datos = self._leer()
        except ValueError as exc:
            messagebox.showerror("JSON no válido", str(exc), parent=self.ventana)
            return
        self.app.rutas["json_alineadas"].mkdir(parents=True, exist_ok=True)
        ruta = filedialog.asksaveasfilename(
            parent=self.ventana,
            initialdir=self.app.rutas["json_alineadas"],
            defaultextension=".json",
            filetypes=(("Archivos JSON", "*.json"),),
        )
        if not ruta:
            return
        destino = Path(ruta)
        temporal = destino.with_suffix(destino.suffix + ".tmp")
        try:
            temporal.write_text(json.dumps(datos, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            temporal.replace(destino)
        except OSError as exc:
            temporal.unlink(missing_ok=True)
            messagebox.showerror("No se pudo guardar", str(exc), parent=self.ventana)
            return
        messagebox.showinfo("Guardado", f"JSON alineado guardado en:\n{destino}", parent=self.ventana)


class PanelComparacion:
    def __init__(self, comparador: "Comparador", indice: int):
        self.comparador = comparador
        self.indice = indice
        self.ruta: Path | None = None
        self.archivo_origen: Path | None = None
        self.directorio_temporal: tempfile.TemporaryDirectory[str] | None = None
        self.zoom = 1.0
        self.marco = ttk.LabelFrame(comparador.contenedor, text=f"Especificación {indice}", padding=8)
        comparador.contenedor.add(
            self.marco,
            minsize=240,
            stretch="always",
        )
        self.marco.rowconfigure(1, weight=1)
        self.marco.columnconfigure(0, weight=1)
        comandos = ttk.Frame(self.marco)
        comandos.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(comandos, text="Seleccionar spec / Word", command=self.seleccionar).pack(side="left")
        ttk.Button(
            comandos,
            text="ⓘ",
            width=3,
            command=self.mostrar_ayuda_seleccion,
        ).pack(side="left", padx=(4, 0))
        ttk.Button(comandos, text="−", width=3, command=lambda: self.cambiar_zoom(-0.1)).pack(side="left", padx=(8, 2))
        ttk.Button(comandos, text="+", width=3, command=lambda: self.cambiar_zoom(0.1)).pack(side="left")
        ttk.Button(comandos, text="Emergente", command=self.comparador.abrir_emergente).pack(side="right")
        ttk.Button(comandos, text="×", width=3, command=self.limpiar).pack(side="right", padx=(0, 6))
        self.previa = ttk.Frame(self.marco)
        self.previa.grid(row=1, column=0, sticky="nsew")
        self.etiqueta = ttk.Label(self.previa, text="Selecciona un JSON para mostrarlo.", anchor="center")
        self.etiqueta.pack(fill="both", expand=True)

    def mostrar_ayuda_seleccion(self) -> None:
        messagebox.showinfo(
            "Archivos para comparar",
            "Puedes seleccionar:\n"
            "• JSON limpios o alineados.\n"
            "• JSON viejos o extraídos.\n"
            "• Tus propias especificaciones Word (.doc, .docx o .docm).\n\n"
            "Los documentos Word se usan solo para esta comparación y no se "
            "registran en el historial.",
            parent=self.comparador.ventana,
        )

    def seleccionar(self) -> None:
        ruta = filedialog.askopenfilename(
            parent=self.comparador.ventana,
            initialdir=self.comparador.app.rutas["json_alineadas"],
            filetypes=(
                ("Especificaciones y documentos", "*.json *.doc *.docx *.docm"),
                ("Archivos JSON", "*.json"),
                ("Documentos Word", "*.doc *.docx *.docm"),
            ),
        )
        if ruta:
            self.cargar_archivo(Path(ruta))

    def cargar_archivo(self, archivo: Path) -> None:
        self.limpiar()
        if archivo.suffix.casefold() == ".json":
            self.archivo_origen = archivo
            self.ruta = archivo
        elif archivo.suffix.casefold() in EXTENSIONES_WORD:
            temporal = tempfile.TemporaryDirectory(prefix="comparacion_especificacion_")
            self.directorio_temporal = temporal
            try:
                self.ruta = crear_json_temporal_comparacion(
                    archivo,
                    Path(temporal.name),
                )
            except (
                OSError,
                RuntimeError,
                ValueError,
                ImportError,
                subprocess.SubprocessError,
                BadZipFile,
                PackageNotFoundError,
            ) as exc:
                self.limpiar()
                messagebox.showerror(
                    "No se pudo preparar el documento Word",
                    f"No fue posible extraer la especificación para compararla:\n{exc}",
                    parent=self.comparador.ventana,
                )
                return
            self.archivo_origen = archivo
            self.marco.configure(text=f"Especificación {self.indice} · Temporal")
        else:
            messagebox.showerror(
                "Formato no compatible",
                "Selecciona un archivo JSON, DOC, DOCX o DOCM.",
                parent=self.comparador.ventana,
            )
            return
        mostrar_vista_json(self.previa, self.ruta, self.zoom)

    def limpiar(self) -> None:
        self.ruta = None
        self.archivo_origen = None
        self.marco.configure(text=f"Especificación {self.indice}")
        if self.directorio_temporal is not None:
            self.directorio_temporal.cleanup()
            self.directorio_temporal = None
        for widget in self.previa.winfo_children():
            widget.destroy()
        self.etiqueta = ttk.Label(
            self.previa,
            text="Selecciona un JSON para mostrarlo.",
            anchor="center",
        )
        self.etiqueta.pack(fill="both", expand=True)

    def cambiar_zoom(self, cambio: float) -> None:
        self.zoom = min(1.8, max(0.7, self.zoom + cambio))
        if self.ruta is not None:
            mostrar_vista_json(self.previa, self.ruta, self.zoom)


class Comparador:
    def __init__(self, app: "AplicacionEspecificaciones"):
        self.app = app
        self.ventana = tk.Toplevel(app.raiz)
        self.ventana.title("Comparar especificaciones")
        self.ventana.geometry("1350x820")
        self.ventana.minsize(850, 550)
        ttk.Label(
            self.ventana,
            text="Comparación de especificaciones",
            font=("Segoe UI", 15, "bold"),
            foreground="#163b70",
        ).pack(anchor="w", padx=14, pady=(12, 6))
        self.contenedor = crear_divisor_redimensionable(self.ventana, tk.HORIZONTAL)
        self.contenedor.pack(fill="both", expand=True, padx=8, pady=5)
        separador_inicializado = False

        def centrar_separador(evento: tk.Event) -> None:
            nonlocal separador_inicializado
            if not separador_inicializado and evento.width > 100:
                separador_inicializado = True
                self.contenedor.sash_place(0, evento.width // 2, 0)

        self.contenedor.bind("<Configure>", centrar_separador)
        self.paneles: list[PanelComparacion] = []
        self.agregar_panel()
        self.agregar_panel()
        self.ventana.protocol("WM_DELETE_WINDOW", self.cerrar)

    def agregar_panel(self) -> None:
        if len(self.paneles) >= 3:
            messagebox.showinfo("Límite alcanzado", "La comparación permite hasta 3 especificaciones.", parent=self.ventana)
            return
        panel = PanelComparacion(self, len(self.paneles) + 1)
        self.paneles.append(panel)

    def abrir_emergente(self) -> None:
        panel_seleccionado = self.ventana.focus_get()
        panel = next(
            (p for p in self.paneles if panel_seleccionado is not None and str(panel_seleccionado).startswith(str(p.marco))),
            None,
        )
        if panel is None:
            panel = next((p for p in self.paneles if p.ruta is not None), None)
        if panel is None or panel.ruta is None:
            messagebox.showinfo("Selecciona una especificación", "Primero carga un JSON en el panel que quieres abrir.", parent=self.ventana)
            return
        self.app.abrir_emergente(
            panel.ruta,
            panel.zoom,
            directorio_temporal=panel.directorio_temporal,
        )

    def cerrar(self) -> None:
        for panel in self.paneles:
            panel.limpiar()
        self.ventana.destroy()


class AplicacionEspecificaciones:
    def __init__(self, raiz: tk.Tk, zoom_aplicacion: float = 1.0):
        self.raiz = raiz
        self.zoom_aplicacion = zoom_aplicacion
        self._escala_tk_base = float(self.raiz.tk.call("tk", "scaling"))
        self.ventanas_flotantes: set[tk.Toplevel] = set()
        try:
            self.rutas = _cargar_configuracion()
        except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
            messagebox.showerror("Configuración inválida", str(exc), parent=raiz)
            self.rutas = dict(RUTAS_PREDETERMINADAS)
        self.raiz.title("Gestor de especificaciones")
        self.raiz.geometry("1380x820")
        self.raiz.minsize(900, 600)
        self.raiz.configure(background="#0752b9")
        self._crear_interfaz()

    def _crear_interfaz(self) -> None:
        barra = tk.Frame(self.raiz, background="#0752b9", padx=20, pady=12)
        barra.pack(fill="x")
        zoom = ttk.Frame(barra)
        zoom.pack(side="right", padx=(8, 0))
        ttk.Button(zoom, text="A−", width=4, command=lambda: self.cambiar_zoom(-0.1)).pack(side="left", padx=2)
        ttk.Button(zoom, text="A+", width=4, command=lambda: self.cambiar_zoom(0.1)).pack(side="left", padx=2)
        for texto, accion in (
            ("INSTRUCCIONES", self.abrir_instrucciones),
            ("VER HISTÓRICO", self.abrir_historico),
            ("ACTUALIZAR FORMATOS", self.actualizar_formatos),
            ("ACTUALIZAR MÉTODOS", self.actualizar_metodos),
            ("CONFIGURACIÓN", self.configuracion),
        ):
            tk.Button(
                barra,
                text=texto,
                command=accion,
                relief="flat",
                background="white",
                activebackground="#dbeafe",
                font=("Segoe UI", 10),
                padx=18,
                pady=5,
            ).pack(side="left", padx=5)

        acciones = tk.Frame(self.raiz, background="#0752b9", padx=45, pady=45)
        acciones.pack(anchor="nw")
        for texto, accion in (
            ("SUBIR ESPEC", self.subir_especificacion),
            ("GENERAR ESPEC", self.generar_especificacion),
            ("SELECCIONAR SPEC", self.seleccionar_especificacion),
            ("COMPARAR ESPECS", self.comparar_especificaciones),
        ):
            tk.Button(
                acciones,
                text=texto,
                command=accion,
                relief="flat",
                background="white",
                activebackground="#dbeafe",
                font=("Segoe UI", 11),
                padx=24,
                pady=8,
                width=22,
            ).pack(anchor="w", pady=9)

    def abrir_instrucciones(self) -> None:
        for nombre in ("README_ACTUALIZACION_CODIGOS_200.md", "README.md"):
            ruta = BASE / nombre
            if ruta.is_file():
                try:
                    abrir_en_windows = getattr(os, "startfile", None)
                    if callable(abrir_en_windows):
                        abrir_en_windows(str(ruta))
                    else:
                        _abrir_carpeta(BASE)
                except OSError as exc:
                    messagebox.showerror("No se pudo abrir", str(exc), parent=self.raiz)
                return
        messagebox.showerror("No encontrado", "No se encontró el archivo de instrucciones.", parent=self.raiz)

    def abrir_historico(self) -> None:
        try:
            self.rutas["historico"].mkdir(parents=True, exist_ok=True)
            _abrir_carpeta(self.rutas["historico"])
        except OSError as exc:
            messagebox.showerror("No se pudo abrir el histórico", str(exc), parent=self.raiz)

    def subir_especificacion(self) -> None:
        archivos = filedialog.askopenfilenames(
            parent=self.raiz,
            title="Selecciona especificaciones Word",
            filetypes=(("Documentos Word", "*.doc *.docm *.docx"),),
        )
        if not archivos:
            return
        destino = self.rutas["especificaciones_viejas"]
        try:
            destino.mkdir(parents=True, exist_ok=True)
            copiados = []
            for archivo in archivos:
                origen = Path(archivo)
                if origen.suffix.lower() not in EXTENSIONES_WORD:
                    raise ValueError(f"Extensión no compatible: {origen.name}")
                nombre = origen.name
                candidato = destino / nombre
                indice = 1
                while candidato.exists():
                    if candidato.read_bytes() == origen.read_bytes():
                        break
                    candidato = destino / f"{origen.stem}_{indice}{origen.suffix}"
                    indice += 1
                if not candidato.exists():
                    shutil.copy2(origen, candidato)
                    copiados.append(candidato.name)
        except (OSError, ValueError) as exc:
            messagebox.showerror("No se pudieron copiar los archivos", str(exc), parent=self.raiz)
            return
        resumen = f"Archivos recibidos: {len(archivos)}\nCopias nuevas: {len(copiados)}\nDestino: {destino}"
        messagebox.showinfo("Carga terminada", resumen, parent=self.raiz)

    def generar_especificacion(self) -> None:
        EditorJson(self)

    def seleccionar_especificacion(self) -> None:
        from interfaz_seleccion_especificaciones import SelectorEspecificaciones

        ventana = tk.Toplevel(self.raiz)
        SelectorEspecificaciones(
            ventana,
            carpeta_json=self.rutas["json_alineadas"],
            carpeta_json_anterior=self.rutas["json_anteriores"],
            al_actualizar=lambda ruta: EditorJson(self, ruta),
        )

    def comparar_especificaciones(self) -> None:
        Comparador(self)

    def actualizar_formatos(self) -> None:
        archivos = filedialog.askopenfilenames(
            parent=self.raiz,
            title="Selecciona plantillas Word para FORMATOS",
            filetypes=(("Plantillas Word", "*.doc *.docm *.docx"),),
        )
        if not archivos:
            return
        destino = self.rutas["formatos"]
        try:
            destino.mkdir(parents=True, exist_ok=True)
            nombres = []
            for archivo in archivos:
                origen = Path(archivo)
                if origen.suffix.lower() not in EXTENSIONES_WORD:
                    raise ValueError(f"Extensión no compatible: {origen.name}")
                shutil.copy2(origen, destino / origen.name)
                nombres.append(origen.name)
        except (OSError, ValueError) as exc:
            messagebox.showerror("No se pudieron actualizar los formatos", str(exc), parent=self.raiz)
            return
        messagebox.showinfo("Formatos actualizados", f"{len(nombres)} archivo(s) copiado(s) en:\n{destino}", parent=self.raiz)

    def actualizar_metodos(self) -> None:
        python = (
            BASE / ".venv" / "Scripts" / "python.exe"
            if os.name == "nt"
            else BASE / ".venv" / "bin" / "python"
        )
        if not python.is_file():
            python = Path(sys.executable)
        script = BASE / "unificar_metodos_lab_v3.py"
        if not script.is_file():
            messagebox.showerror("No encontrado", f"No se encontró {script.name}.", parent=self.raiz)
            return
        try:
            flags = subprocess.CREATE_NEW_CONSOLE if os.name == "nt" else 0
            subprocess.Popen(
                [str(python), str(script), "--base", str(BASE)],
                cwd=BASE,
                creationflags=flags,
            )
        except OSError as exc:
            messagebox.showerror("No se pudo iniciar la actualización", str(exc), parent=self.raiz)

    def configuracion(self) -> None:
        ventana = tk.Toplevel(self.raiz)
        ventana.title("Configuración de carpetas")
        ventana.transient(self.raiz)
        ventana.grab_set()
        campos: dict[str, tk.StringVar] = {}
        nombres = {
            "especificaciones_viejas": "Documentos Word recibidos",
            "json_anteriores": "JSON extraídos",
            "json_alineadas": "JSON alineados",
            "formatos": "Plantillas Word",
            "historico": "Documentos generados",
        }
        for fila, (clave, etiqueta) in enumerate(nombres.items()):
            ttk.Label(ventana, text=etiqueta).grid(row=fila, column=0, sticky="w", padx=10, pady=7)
            variable = tk.StringVar(value=str(self.rutas[clave]))
            campos[clave] = variable
            ttk.Entry(ventana, textvariable=variable, width=64).grid(row=fila, column=1, padx=5, pady=7)
            ttk.Button(
                ventana,
                text="Examinar",
                command=lambda v=variable: self._elegir_carpeta(v, ventana),
            ).grid(row=fila, column=2, padx=5, pady=7)

        def guardar() -> None:
            self.rutas = {clave: Path(variable.get()).expanduser().resolve() for clave, variable in campos.items()}
            try:
                _guardar_configuracion(self.rutas)
            except OSError as exc:
                messagebox.showerror("No se pudo guardar", str(exc), parent=ventana)
                return
            ventana.destroy()

        ttk.Button(ventana, text="Guardar", command=guardar).grid(row=len(nombres), column=1, sticky="e", padx=8, pady=12)
        ttk.Button(ventana, text="Cancelar", command=ventana.destroy).grid(row=len(nombres), column=2, padx=8, pady=12)

    @staticmethod
    def _elegir_carpeta(variable: tk.StringVar, parent: tk.Misc) -> None:
        ruta = filedialog.askdirectory(parent=parent, initialdir=variable.get() or str(BASE))
        if ruta:
            variable.set(ruta)

    def cambiar_zoom(self, cambio: float) -> None:
        self.zoom_aplicacion = min(1.6, max(0.8, self.zoom_aplicacion + cambio))
        self.raiz.tk.call(
            "tk",
            "scaling",
            self._escala_tk_base * self.zoom_aplicacion,
        )

    def abrir_emergente(
        self,
        ruta: Path,
        zoom_inicial: float,
        *,
        directorio_temporal: tempfile.TemporaryDirectory[str] | None = None,
    ) -> None:
        vivas = {ventana for ventana in self.ventanas_flotantes if ventana.winfo_exists()}
        self.ventanas_flotantes = vivas
        if len(vivas) >= 6:
            messagebox.showinfo("Límite alcanzado", "Se permiten hasta 6 vistas emergentes.", parent=self.raiz)
            return
        temporal_emergente = None
        if directorio_temporal is not None:
            temporal_emergente = tempfile.TemporaryDirectory(
                prefix="comparacion_emergente_"
            )
            ruta_emergente = Path(temporal_emergente.name) / ruta.name
            try:
                shutil.copy2(ruta, ruta_emergente)
            except OSError as exc:
                temporal_emergente.cleanup()
                messagebox.showerror(
                    "No se pudo abrir la vista emergente",
                    str(exc),
                    parent=self.raiz,
                )
                return
            ruta = ruta_emergente
        ventana = tk.Toplevel(self.raiz)
        ventana.title(f"Vista emergente · {ruta.name}")
        ventana.geometry("900x680")
        ventana.minsize(600, 400)
        self.ventanas_flotantes.add(ventana)
        if temporal_emergente is not None:
            ventana.bind(
                "<Destroy>",
                lambda evento, temporal=temporal_emergente: temporal.cleanup()
                if evento.widget is ventana
                else None,
                add="+",
            )
        zoom = [zoom_inicial]
        barra = ttk.Frame(ventana, padding=6)
        barra.pack(fill="x")
        marco = ttk.Frame(ventana)
        marco.pack(fill="both", expand=True)

        def actualizar(cambio: float) -> None:
            zoom[0] = min(1.8, max(0.7, zoom[0] + cambio))
            mostrar_vista_json(marco, ruta, zoom[0])

        ttk.Button(barra, text="−", width=3, command=lambda: actualizar(-0.1)).pack(side="left")
        ttk.Button(barra, text="+", width=3, command=lambda: actualizar(0.1)).pack(side="left", padx=4)
        ttk.Label(barra, text=str(ruta)).pack(side="left", padx=10)
        ventana.bind("<Destroy>", lambda evento: self.ventanas_flotantes.discard(ventana) if evento.widget is ventana else None)
        mostrar_vista_json(marco, ruta, zoom[0])


def main() -> None:
    parser = argparse.ArgumentParser(description="Gestor gráfico de especificaciones.")
    parser.add_argument(
        "--editar-json",
        type=Path,
        help="abre directamente el JSON indicado en el editor de especificaciones",
    )
    args = parser.parse_args()
    raiz = tk.Tk()
    app = AplicacionEspecificaciones(raiz)
    if args.editar_json is not None:
        EditorJson(app, args.editar_json)
    raiz.mainloop()


if __name__ == "__main__":
    main()
