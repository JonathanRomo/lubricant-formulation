#!/usr/bin/env python3
"""Genera HII bilingües preservando macros, objetos y formato del XLSM original."""
from __future__ import annotations
import estilo_interfaz as ui


import argparse
import copy
import io
import json
import re
import sqlite3
import math
import textwrap
import tkinter as tk
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox
from estilo_interfaz import ttk

from zipfile import ZipFile

from lxml import etree as ET
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from generador_especs_hii import clave_texto, leer_idioma_datos, _separar_notas_de_pruebas
from interfaz_seleccion_especificaciones import (
    preparar_vista_previa_alineada, cargar_catalogo, buscar_especificaciones,
    encabezados_tabla, crear_divisor_redimensionable,
)
import registro_hii

from rutas_proyecto import BASE, resolver_ruta, rutas_configuradas
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
FECHA = re.compile(r"DD\s*/\s*MM\s*/\s*YY", re.I)
REFERENCIA = re.compile(r"(?<![\w])([$]?[A-Z]{1,3}[$]?)(\d+)(?!\w)")
COLUMNAS = ("numero", "prueba", "metodo", "periodicidad", "unidades", "especificacion")


def texto(valor) -> str:
    return "" if valor is None else str(valor).strip()


@dataclass
class DatosHII:
    nombre: str
    codigo: str
    idioma: str
    pruebas: list[dict]
    comentarios: str
    reanalisis: str
    fecha: str


def comentarios_simples(notas: list[str]) -> str:
    # TODO: asignar aquí comportamientos especiales a TODO lo relacionado con
    # comentarios (referencias, superíndices, reglas por prueba y por idioma).
    # Por ahora únicamente conserva el texto, sin interpretar ni traducirlo.
    return "\n".join(dict.fromkeys(texto(nota) for nota in notas if texto(nota)))


def leer_reanalisis(datos: dict, notas: list[str]) -> str:
    for objeto in (datos, datos.get("control_documental", {}), datos.get("producto", {})):
        if isinstance(objeto, dict):
            for campo, valor in objeto.items():
                if clave_texto(campo) in {"FECHA DE REANALISIS", "FECHA REANALISIS",
                                         "REANALYSIS DATE", "REANALISIS", "REANALYSIS"}:
                    if texto(valor):
                        return texto(valor)
    for nota in notas:
        coincidencia = re.search(
            r"(?:fecha\s+(?:de\s+)?rean[aá]lisis|rean[aá]lisis|reanalysis\s+date)\s*[:：]\s*([^\n]+)",
            texto(nota), re.I)
        if coincidencia:
            return coincidencia.group(1).strip()
    return ""


def preparar_datos(ruta: Path, fecha: date | None = None) -> DatosHII:
    datos = json.loads(ruta.read_text(encoding="utf-8-sig"))
    vista = preparar_vista_previa_alineada(ruta)
    idioma = leer_idioma_datos(datos)
    if idioma not in {"ESPANOL", "INGLES"}:
        raise ValueError("La especificación debe indicar idioma español o inglés.")
    codigo = texto(datos.get("control_documental", {}).get("codigo_unico_producto"))
    if not codigo:
        raise ValueError("Falta control_documental.codigo_unico_producto.")
    pruebas = []
    originales = {texto(p.get("numero", i)): p for i, p in enumerate(datos.get("pruebas", []), 1)}
    for i, detalle in enumerate(datos["alineacion_metodos_lab"]["detalle"], 1):
        valores = dict(detalle["valores_alineados"])
        original = originales.get(texto(detalle.get("numero_prueba", i)), {})
        periodicidad = next((texto(objeto.get(campo))
                            for objeto in (valores, original)
                            for campo in ("periodicidad", "PERIODIDICAD DE ANALISIS", "periodicidad_analisis")
                            if texto(objeto.get(campo))), "")
        valores["periodicidad"] = periodicidad
        pruebas.append(valores)
    pruebas, _ = _separar_notas_de_pruebas(pruebas)
    for i, prueba in enumerate(pruebas, 1):
        prueba["numero"] = str(i)
    if not pruebas:
        raise ValueError("La especificación alineada no contiene pruebas.")
    return DatosHII(vista.nombre_producto, codigo, idioma, pruebas,
                    comentarios_simples(list(vista.notas)),
                    leer_reanalisis(datos, list(vista.notas)),
                    (fecha or date.today()).strftime("%d/%m/%y"))


def seleccionar_plantilla(carpeta: Path, idioma: str) -> Path:
    codigo = "91400019" if idioma == "INGLES" else "91400005"
    candidatas = sorted(p for p in carpeta.glob("*.xlsm") if p.name.startswith(codigo))
    if len(candidatas) != 1:
        raise ValueError(f"Se requiere una sola plantilla {codigo} XLSM en {carpeta}.")
    return candidatas[0]


def desplazar_referencia(valor: str, extra: int) -> str:
    return REFERENCIA.sub(lambda m: m[1] + str(int(m[2]) + (extra if int(m[2]) >= 21 else 0)), valor)


def escribir_celda(hoja, referencia: str, valor: str) -> None:
    fila_numero = int(re.search(r"\d+", referencia)[0])
    filas = hoja.find("s:sheetData", NS)
    fila = next((f for f in filas if int(f.get("r")) == fila_numero), None)
    if fila is None:
        fila = ET.SubElement(filas, f"{{{NS['s']}}}row", r=str(fila_numero))
        filas[:] = sorted(filas, key=lambda f: int(f.get("r")))
    celda = next((c for c in fila if c.get("r") == referencia), None)
    if celda is None:
        celda = ET.SubElement(fila, f"{{{NS['s']}}}c", r=referencia)
        fila[:] = sorted(fila, key=lambda c: c.get("r").rstrip("0123456789"))
    for hijo in list(celda):
        celda.remove(hijo)
    celda.set("t", "inlineStr")
    cadena = ET.SubElement(celda, f"{{{NS['s']}}}is")
    t = ET.SubElement(cadena, f"{{{NS['s']}}}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = valor


def construir_excel(plantilla: Path, datos: DatosHII) -> bytes:
    """Edita únicamente las partes necesarias del paquete; conserva los demás bytes."""
    extra = max(0, len(datos.pruebas) - 6)
    modelo_libro = load_workbook(plantilla, read_only=False)
    modelo_hoja = modelo_libro.worksheets[0]
    anchos = [modelo_hoja.column_dimensions[get_column_letter(c)].width for c in range(1, 7)]
    nombre_hoja = modelo_hoja.title
    modelo_libro.close()
    with ZipFile(plantilla) as archivo:
        partes = {info.filename: archivo.read(info.filename) for info in archivo.infolist()}
        infos = archivo.infolist()
    hoja_nombre = "xl/worksheets/sheet1.xml"
    hoja = ET.fromstring(partes[hoja_nombre])
    estilos = ET.fromstring(partes["xl/styles.xml"])
    formatos_celda = estilos.find("s:cellXfs", NS)
    estilos_envueltos = {}
    if extra:
        filas = hoja.find("s:sheetData", NS)
        modelo = copy.deepcopy(next(f for f in filas if f.get("r") == "20"))
        for elemento in hoja.iter():
            for atributo in ("r", "ref", "sqref"):
                if elemento.get(atributo):
                    valor = elemento.get(atributo)
                    elemento.set(atributo, str(int(valor) + extra) if atributo == "r" and valor.isdigit() and int(valor) >= 21
                                 else desplazar_referencia(valor, extra))
            if ET.QName(elemento).localname in {"f", "formula1", "formula2"} and elemento.text:
                elemento.text = desplazar_referencia(elemento.text, extra)
        for numero in range(21, 21 + extra):
            fila = copy.deepcopy(modelo)
            fila.set("r", str(numero))
            for celda in fila:
                celda.set("r", re.sub(r"\d+", str(numero), celda.get("r")))
            filas.append(fila)
        filas[:] = sorted(filas, key=lambda f: int(f.get("r")))
    for indice in range(max(6, len(datos.pruebas))):
        prueba = datos.pruebas[indice] if indice < len(datos.pruebas) else {}
        for col, campo in enumerate(COLUMNAS, 1):
            escribir_celda(hoja, f"{get_column_letter(col)}{15 + indice}", texto(prueba.get(campo)))
        fila = next(f for f in hoja.find("s:sheetData", NS) if f.get("r") == str(15 + indice))
        for celda in fila:
            if re.match(r"^[A-F]\d+$", celda.get("r", "")):
                estilo = int(celda.get("s", "0"))
                if estilo not in estilos_envueltos:
                    nuevo = copy.deepcopy(formatos_celda[estilo])
                    alineacion = nuevo.find("s:alignment", NS)
                    if alineacion is None:
                        alineacion = ET.SubElement(nuevo, f"{{{NS['s']}}}alignment")
                    alineacion.set("wrapText", "1")
                    nuevo.set("applyAlignment", "1")
                    estilos_envueltos[estilo] = len(formatos_celda)
                    formatos_celda.append(nuevo)
                celda.set("s", str(estilos_envueltos[estilo]))
        lineas = max(max(1, sum(max(1, math.ceil(len(linea) / max(4, ancho - 2)))
                                   for linea in texto(prueba.get(campo)).split("\n")))
                     for ancho, campo in zip(anchos, COLUMNAS))
        fila.set("ht", str(max(float(fila.get("ht", "15")), lineas * 12 + 6)))
        fila.set("customHeight", "1")
    comentario = "D24" if datos.idioma == "INGLES" else "D25"
    # Comentarios: mantener toda lógica especial centralizada en comentarios_simples.
    if datos.comentarios:
        escribir_celda(hoja, desplazar_referencia(comentario, extra), datos.comentarios)
    escribir_celda(hoja, f"D{33 + extra}", datos.reanalisis)
    partes[hoja_nombre] = ET.tostring(hoja, xml_declaration=True, encoding="UTF-8")
    formatos_celda.set("count", str(len(formatos_celda)))
    partes["xl/styles.xml"] = ET.tostring(estilos, xml_declaration=True, encoding="UTF-8")
    for nombre, contenido in list(partes.items()):
        if nombre.startswith("xl/drawings/") and nombre.endswith(".xml") and "/_rels/" not in nombre:
            raiz = ET.fromstring(contenido)
            for ancla in raiz:
                nodos = ancla.findall(".//a:t", NS)
                for nodo in nodos:
                    nodo.text = FECHA.sub(datos.fecha, nodo.text or "")
                origen = ancla.find("xdr:from", NS)
                fila = int(origen.find("xdr:row", NS).text) if origen is not None else -1
                columna = int(origen.find("xdr:col", NS).text) if origen is not None else -1
                reemplazo = datos.nombre if (fila == 6 and columna == 1) else datos.codigo if (fila == 9 and columna == 1) else None
                if reemplazo is not None and nodos:
                    nodos[0].text = reemplazo
                    for nodo in nodos[1:]:
                        nodo.text = ""
                for nodo in ancla.findall(".//xdr:row", NS):
                    if int(nodo.text) >= 20:
                        nodo.text = str(int(nodo.text) + extra)
            partes[nombre] = ET.tostring(raiz, xml_declaration=True, encoding="UTF-8")
        elif extra and (nombre.startswith("xl/tables/") or nombre == "xl/workbook.xml") and nombre.endswith(".xml"):
            raiz = ET.fromstring(contenido)
            for elemento in raiz.iter():
                if elemento.get("ref"):
                    elemento.set("ref", desplazar_referencia(elemento.get("ref"), extra))
                if ET.QName(elemento).localname == "definedName" and elemento.text:
                    elemento.text = desplazar_referencia(elemento.text, extra)
            partes[nombre] = ET.tostring(raiz, xml_declaration=True, encoding="UTF-8")
    # Limita la impresión a la HII, sin incluir el catálogo auxiliar de la plantilla inglesa.
    libro = ET.fromstring(partes["xl/workbook.xml"])
    definidos = libro.find("s:definedNames", NS)
    if definidos is None:
        definidos = ET.Element(f"{{{NS['s']}}}definedNames")
        posicion = list(libro).index(libro.find("s:sheets", NS)) + 1
        libro.insert(posicion, definidos)
    area = next((n for n in definidos if n.get("name") == "_xlnm.Print_Area" and n.get("localSheetId") == "0"), None)
    if area is None:
        area = ET.SubElement(definidos, f"{{{NS['s']}}}definedName", name="_xlnm.Print_Area", localSheetId="0")
    area.text = "'" + nombre_hoja.replace("'", "''") + f"'!$A$1:$F${43 + extra}"
    partes["xl/workbook.xml"] = ET.tostring(libro, xml_declaration=True, encoding="UTF-8")
    # Reemplazar fechas también en cadenas compartidas, encabezados y otros objetos XML.
    for nombre, contenido in list(partes.items()):
        if nombre.endswith((".xml", ".vml")) and b"DD/MM/YY" in contenido:
            partes[nombre] = contenido.replace(b"DD/MM/YY", datos.fecha.encode())
    salida = io.BytesIO()
    with ZipFile(salida, "w") as archivo:
        for info in infos:
            archivo.writestr(info, partes[info.filename])
    return salida.getvalue()


class GeneradorHII:
    def __init__(self, padre, carpeta_formatos=None, carpeta_json=None, app=None, base_datos=BASE / "especificaciones.db"):
        self.app = app
        rutas = rutas_configuradas(BASE)
        self.base_datos = resolver_ruta(base_datos, BASE)
        self.idioma_objetivo = None
        self.origen_actual = None
        self.editores_traduccion = {}
        self.formatos = resolver_ruta(carpeta_formatos if carpeta_formatos is not None else rutas["formatos"], BASE)
        self.carpeta_json = resolver_ruta(carpeta_json if carpeta_json is not None else rutas["json_alineadas"], BASE)
        self.excel = None
        self.datos = None
        self.marcadas = set()
        self.compacta = False
        self.actual = None
        self.ventana = ui.Ventana(padre)
        self.ventana.title("Generar HII — especificación alineada y Excel")
        self.ventana.geometry("1400x850")
        self.ventana.minsize(1000, 650)
        self.ventana.configure(background="#edf2f7")
        cabecera = tk.Frame(self.ventana, background="#0752b9", padx=24, pady=16)
        cabecera.pack(fill="x")
        tk.Label(cabecera, text="GENERAR HII", background="#0752b9", foreground="white",
                 font=("Segoe UI", 17, "bold")).pack(anchor="w")
        tk.Label(cabecera, text="Marca las especificaciones para revisar su resumen y la hoja de instrucción e inspección.",
                 background="#0752b9", foreground="#dbeafe", font=("Segoe UI", 10)).pack(anchor="w", pady=(4, 0))
        barra = ttk.Frame(self.ventana, padding=8)
        barra.pack(fill="x")
        self.editar_seleccion = ttk.Button(barra, text="Ampliar lista", command=self.alternar_lista)
        self.editar_seleccion.pack(side="left")
        self.lote_boton = ttk.Button(barra, text="Guardar seleccionadas", command=self.guardar_seleccionadas, state="disabled")
        self.lote_boton.pack(side="left", padx=8)
        self.guardar_boton = ttk.Button(barra, text="Guardar HII (.xlsm)", command=self.guardar, state="disabled")
        self.guardar_boton.pack(side="left", padx=8)
        self.estado = ttk.Label(barra, text="Marca una especificación de la lista.")
        self.estado.pack(side="left")
        idiomas = ttk.Frame(self.ventana, padding=(8, 0, 8, 8))
        idiomas.pack(fill="x")
        ttk.Label(idiomas, text="Generar versión:").pack(side="left", padx=(0, 8))
        ttk.Button(idiomas, text="Español", command=lambda: self.generar_idioma("ESPANOL")).pack(side="left", padx=3)
        ttk.Button(idiomas, text="English", command=lambda: self.generar_idioma("INGLES")).pack(side="left", padx=3)
        ttk.Button(idiomas, text="Idioma original", command=lambda: self.generar_idioma(None)).pack(side="left", padx=3)
        self.etiqueta_idioma = ttk.Label(idiomas, text="Se usa el idioma de la especificación.")
        self.etiqueta_idioma.pack(side="left", padx=12)
        self.contenido = ttk.Panedwindow(self.ventana, orient="horizontal")
        self.contenido.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        self.busqueda_panel = ttk.Frame(self.contenido, padding=12)
        self.previa_panel = ttk.Frame(self.contenido)
        self.contenido.add(self.busqueda_panel, weight=2)
        self.contenido.add(self.previa_panel, weight=3)
        ttk.Label(self.busqueda_panel, text="Buscar especificación", foreground="#163b70",
                  font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))
        self.consulta = tk.StringVar()
        ttk.Entry(self.busqueda_panel, textvariable=self.consulta).pack(fill="x", pady=(0, 8))
        self.consulta.trace_add("write", lambda *_: self.actualizar_lista())
        marco_lista = ttk.Frame(self.busqueda_panel)
        marco_lista.pack(fill="both", expand=True)
        self.lista = ttk.Treeview(marco_lista, columns=("marca", "nombre"), show="headings", selectmode="browse")
        self.lista.heading("marca", text="✓")
        self.lista.heading("nombre", text="Especificación alineada")
        self.lista.column("marca", width=32, minwidth=32, stretch=False, anchor="center")
        self.lista.column("nombre", width=430, minwidth=140)
        scroll = ttk.Scrollbar(marco_lista, orient="vertical", command=self.lista.yview)
        scroll_x = ttk.Scrollbar(marco_lista, orient="horizontal", command=self.lista.xview)
        self.lista.configure(yscrollcommand=scroll.set, xscrollcommand=scroll_x.set)
        marco_lista.rowconfigure(0, weight=1)
        marco_lista.columnconfigure(0, weight=1)
        self.lista.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.lista.bind("<ButtonRelease-1>", self.clic_lista)
        self.lista.bind("<space>", self.marcar_actual)
        self.lista.bind("<Return>", self.marcar_actual)
        self.lista.bind("<<TreeviewSelect>>", self.seleccionar_fila)
        ttk.Button(self.busqueda_panel, text="Actualizar lista", command=self.recargar_catalogo).pack(fill="x", pady=(10, 0))
        self.contador = ttk.Label(self.busqueda_panel, wraplength=250)
        self.contador.pack(anchor="w", pady=(8, 0))
        self.paneles = ttk.Panedwindow(self.previa_panel, orient="horizontal")
        self.paneles.pack(fill="both", expand=True)
        self.izquierda = ttk.Frame(self.paneles)
        self.derecha = ttk.Frame(self.paneles)
        self.paneles.add(self.izquierda, weight=1)
        self.paneles.add(self.derecha, weight=1)
        self.vaciar_previa()
        self.recargar_catalogo()
        self.ventana.after(100, self.posicionar_paneles)

    def posicionar_paneles(self):
        ancho = self.contenido.winfo_width()
        self.contenido.sashpos(0, 250 if self.compacta else round(ancho * 0.42))
        self.ventana.after_idle(lambda: self.paneles.sashpos(0, self.paneles.winfo_width() // 2))

    def alternar_lista(self):
        self.compacta = not self.compacta
        self.editar_seleccion.configure(text="Ampliar lista" if self.compacta else "Compactar lista")
        self.actualizar_lista()
        self.posicionar_paneles()

    def recargar_catalogo(self):
        try:
            self.catalogo = cargar_catalogo(self.carpeta_json)
            with registro_hii.conectar(self.base_datos) as con:
                registro_hii.sincronizar_idiomas(con, self.carpeta_json)
        except (OSError, ValueError, sqlite3.Error) as exc:
            messagebox.showerror("No se pudo cargar la lista", str(exc), parent=self.ventana)
            return
        self.marcadas.intersection_update(e.ruta for e in self.catalogo.especificaciones)
        self.actualizar_lista()
        if self.origen_actual and self.origen_actual not in self.marcadas:
            self.vaciar_previa()
        elif self.actual:
            self.cargar_previa(self.origen_actual or self.actual)

    def actualizar_lista(self):
        if not hasattr(self, "catalogo"):
            return
        self.lista.delete(*self.lista.get_children())
        candidatas = self.catalogo.especificaciones
        if self.compacta and self.marcadas:
            candidatas = [e for e in candidatas if e.ruta in self.marcadas]
        self.visibles = buscar_especificaciones(candidatas, self.consulta.get())
        for i, especificacion in enumerate(self.visibles):
            self.lista.insert("", "end", iid=str(i), values=("☑" if especificacion.ruta in self.marcadas else "☐",
                especificacion.etiqueta + (" · English" if especificacion.idioma == "INGLES" else " · Español")))
        ingles = sum(e.idioma == "INGLES" for e in self.catalogo.especificaciones)
        errores = f"\n{len(self.catalogo.errores)} archivos inválidos" if self.catalogo.errores else ""
        self.contador.configure(text=f"{len(self.catalogo.especificaciones)} especificaciones ({ingles} en inglés)\n"
                               f"{len(self.marcadas)} seleccionadas · {len(self.visibles)} visibles{errores}")
        self.lote_boton.configure(state="normal" if self.marcadas else "disabled")

    def clic_lista(self, evento):
        fila = self.lista.identify_row(evento.y)
        if fila and self.lista.identify_column(evento.x) == "#1":
            self.marcar(int(fila))

    def marcar_actual(self, _evento=None):
        seleccion = self.lista.selection()
        if seleccion:
            self.marcar(int(seleccion[0]))
        return "break"

    def marcar(self, indice):
        especificacion = self.visibles[indice]
        if especificacion.ruta in self.marcadas:
            self.marcadas.remove(especificacion.ruta)
            if self.origen_actual == especificacion.ruta:
                siguiente = next((e for e in self.catalogo.especificaciones if e.ruta in self.marcadas), None)
                if siguiente:
                    self.cargar_previa(siguiente.ruta)
                else:
                    self.vaciar_previa()
        else:
            self.marcadas.add(especificacion.ruta)
            self.cargar_previa(especificacion.ruta)
            self.compacta = True
            self.consulta.set("")
        if not self.marcadas:
            self.compacta = False
        self.editar_seleccion.configure(text="Ampliar lista" if self.compacta else "Compactar lista")
        self.actualizar_lista()
        self.posicionar_paneles()

    def seleccionar_fila(self, _evento=None):
        seleccion = self.lista.selection()
        if seleccion:
            especificacion = self.visibles[int(seleccion[0])]
            if especificacion.ruta in self.marcadas and especificacion.ruta != self.origen_actual:
                self.cargar_previa(especificacion.ruta)

    def vaciar_previa(self):
        self.excel = self.datos = self.actual = self.origen_actual = None
        self.guardar_boton.configure(state="disabled")
        for panel, mensaje in ((self.izquierda, "Marca una especificación para ver su resumen."),
                               (self.derecha, "Aquí verás la hoja Excel que se generará.")):
            for widget in panel.winfo_children():
                widget.destroy()
            ttk.Label(panel, text=mensaje, anchor="center", wraplength=350).pack(fill="both", expand=True, padx=12)
        self.estado.configure(text="Marca una especificación de la lista.")

    def abrir(self):
        ruta = filedialog.askopenfilename(parent=self.ventana, initialdir=self.carpeta_json,
                                          filetypes=[("JSON alineado", "*.json *.JSON")])
        if not ruta:
            return
        self.cargar_previa(Path(ruta))

    def cargar_previa(self, ruta):
        try:
            origen = Path(ruta)
            ruta = self.resolver_idioma(origen)
            if ruta is None:
                return
            datos = preparar_datos(Path(ruta))
            plantilla = seleccionar_plantilla(self.formatos, datos.idioma)
            excel = construir_excel(plantilla, datos)
            self.mostrar_especificacion(Path(ruta))
            self.mostrar_excel(excel, datos)
            self.datos, self.excel = datos, excel
            self.actual = Path(ruta)
            self.origen_actual = origen
            self.guardar_boton.configure(state="normal")
            aviso = " | Sin fecha de reanálisis en el JSON" if not datos.reanalisis else ""
            self.estado.configure(text=f"{datos.codigo} | {datos.idioma}{aviso}")
        except Exception as exc:
            self.vaciar_previa()
            messagebox.showerror("No se pudo generar HII", str(exc), parent=self.ventana)

    def generar_idioma(self, idioma):
        self.idioma_objetivo = idioma
        self.etiqueta_idioma.configure(text={None: "Se usa el idioma de la especificación.",
            "ESPANOL": "Versión solicitada: Español", "INGLES": "Versión solicitada: English"}[idioma])
        origen = self.origen_actual or self.actual
        if origen is None:
            origen = next((e.ruta for e in self.catalogo.especificaciones if e.ruta in self.marcadas), None)
        if origen:
            self.cargar_previa(origen)
        else:
            self.estado.configure(text="Marca una especificación para generar el idioma elegido.")

    def resolver_idioma(self, origen):
        datos = preparar_datos(origen)
        if self.idioma_objetivo is None or datos.idioma == self.idioma_objetivo:
            return origen
        with registro_hii.conectar(self.base_datos) as con:
            opciones = registro_hii.candidatas(con, origen, self.idioma_objetivo,
                                              [e.ruta for e in self.catalogo.especificaciones])
        if len(opciones) > 1:
            dialogo = ui.Ventana(self.ventana)
            dialogo.title("Elegir especificación traducida")
            dialogo.transient(self.ventana)
            ttk.Label(dialogo, text="Hay varias versiones con el mismo código, revisión y cliente. Elige la correspondiente.",
                      wraplength=650, padding=12).pack()
            lista = tk.Listbox(dialogo, width=90, height=min(12, len(opciones)), exportselection=False)
            for opcion in opciones:
                lista.insert("end", str(opcion.relative_to(self.carpeta_json.resolve())))
            lista.pack(fill="both", expand=True, padx=12, pady=8)
            elegida = []
            def aceptar():
                if lista.curselection():
                    elegida.append(opciones[lista.curselection()[0]])
                    dialogo.destroy()
            ttk.Button(dialogo, text="Usar esta especificación", command=aceptar).pack(pady=10)
            dialogo.grab_set()
            self.ventana.wait_window(dialogo)
            if not elegida:
                self.vaciar_previa()
                self.origen_actual = origen
                return None
            opciones = elegida
        if opciones:
            destino = opciones[0]
            if preparar_datos(destino).idioma != self.idioma_objetivo:
                raise ValueError("La especificación vinculada no está en el idioma solicitado.")
            with registro_hii.conectar(self.base_datos) as con:
                registro_hii.vincular(con, origen, destino, permitir_revision_distinta=True)
            return destino
        self.vaciar_previa()
        self.origen_actual = origen
        self.estado.configure(text="Falta especificación traducida. Completa y guarda la traducción en el editor.")
        self.abrir_traduccion(origen, self.idioma_objetivo)
        return None

    def abrir_traduccion(self, origen, idioma):
        clave = (origen.resolve(), idioma)
        existente = self.editores_traduccion.get(clave)
        if existente is not None and existente.ventana.winfo_exists():
            existente.ventana.lift()
            return
        from interfaz_aplicacion_especificaciones import GeneradorEspecificacion
        from types import SimpleNamespace
        app = self.app or SimpleNamespace(raiz=self.ventana,
            rutas={"json_alineadas": self.carpeta_json, "formatos": self.formatos})
        sufijo = "ENG" if idioma == "INGLES" else "ESP"
        destino = self.carpeta_json / f"{origen.stem} {sufijo}.json"
        indice = 1
        while destino.exists():
            destino = self.carpeta_json / f"{origen.stem} {sufijo}_{indice}.json"
            indice += 1
        def guardada(ruta):
            datos = preparar_datos(ruta)
            if datos.idioma != idioma:
                raise ValueError("La traducción no tiene el idioma solicitado.")
            with registro_hii.conectar(self.base_datos) as con:
                registro_hii.vincular(con, origen, ruta, permitir_revision_distinta=True)
            self.recargar_catalogo()
            self.cargar_previa(origen)
        editor = GeneradorEspecificacion(app, archivo_inicial=origen, destino_inicial=destino,
            idioma_destino="en" if idioma == "INGLES" else "es", al_guardar=guardada)
        self.editores_traduccion[clave] = editor

    def registrar_hii(self, destino, especificacion, datos):
        with registro_hii.conectar(self.base_datos) as con:
            registro_hii.registrar(con, destino, especificacion, datos.idioma, datos.fecha)

    def mostrar_especificacion(self, ruta):
        from tkinter.scrolledtext import ScrolledText
        for widget in self.izquierda.winfo_children():
            widget.destroy()
        vista = preparar_vista_previa_alineada(ruta)
        ttk.Label(self.izquierda, text="Resumen · Especificación alineada", padding=8,
                  foreground="#163b70", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(self.izquierda, text=vista.nombre_producto, foreground="#163b70",
                  font=("Segoe UI", 16, "bold"), wraplength=450).pack(anchor="w", padx=10, pady=5)
        separador = crear_divisor_redimensionable(self.izquierda, tk.VERTICAL)
        separador.pack(fill="both", expand=True, padx=8, pady=8)
        tabla_panel = ttk.Frame(separador)
        detalles_panel = ttk.Frame(separador)
        separador.add(tabla_panel, minsize=140, stretch="always")
        separador.add(detalles_panel, minsize=100, stretch="always")
        campos = encabezados_tabla(vista.idioma)
        tabla = ttk.Treeview(tabla_panel, columns=[c for _, c in campos], show="headings")
        for etiqueta, campo in campos:
            tabla.heading(campo, text=etiqueta)
            tabla.column(campo, width=210 if campo == "prueba" else 120, minwidth=80)
        for prueba in vista.pruebas:
            tabla.insert("", "end", values=[texto(prueba.get(c)) for _, c in campos])
        vertical = ttk.Scrollbar(tabla_panel, orient="vertical", command=tabla.yview)
        horizontal = ttk.Scrollbar(tabla_panel, orient="horizontal", command=tabla.xview)
        tabla.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tabla_panel.rowconfigure(0, weight=1)
        tabla_panel.columnconfigure(0, weight=1)
        tabla.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        def fijar_divisor(evento):
            if evento.height > 100 and not getattr(separador, "iniciado", False):
                separador.sash_place(0, 0, round(evento.height * .6))
                separador.iniciado = True
        separador.bind("<Configure>", fijar_divisor)
        visor = ScrolledText(detalles_panel, wrap="word", font=("Segoe UI", 10))
        visor.pack(fill="both", expand=True, padx=8, pady=8)
        lineas = [f"{len(vista.pruebas)} pruebas · {vista.idioma}",
                  *[f"{k}: {v}" for k, v in vista.control], ""]
        lineas.extend(vista.notas)
        lineas.extend(vista.mensajes_astm_no_encontradas)
        visor.insert("1.0", "\n".join(lineas))
        visor.configure(state="disabled")

    def guardar_seleccionadas(self):
        carpeta = filedialog.askdirectory(parent=self.ventana, title="Guardar las HII seleccionadas")
        if not carpeta:
            return
        seleccionadas = [e for e in self.catalogo.especificaciones if e.ruta in self.marcadas]
        resueltas = []
        for especificacion in seleccionadas:
            ruta = self.resolver_idioma(especificacion.ruta)
            if ruta is None:
                return
            if ruta not in resueltas:
                resueltas.append(ruta)
        salidas = [(ruta, Path(carpeta) / f"{ruta.stem} HII {preparar_datos(ruta).idioma}.xlsm") for ruta in resueltas]
        existentes = [p for _, p in salidas if p.exists()]
        if existentes and not messagebox.askyesno("Archivos existentes", f"Se reemplazarán {len(existentes)} archivos HII. ¿Continuar?", parent=self.ventana):
            return
        errores, guardadas = [], 0
        for especificacion, destino in salidas:
            try:
                datos = preparar_datos(especificacion)
                plantilla = seleccionar_plantilla(self.formatos, datos.idioma)
                if destino.resolve() in {p.resolve() for p in self.formatos.glob("*.xlsm")}:
                    raise ValueError("El destino coincide con una plantilla.")
                destino.write_bytes(construir_excel(plantilla, datos))
                self.registrar_hii(destino, especificacion, datos)
                guardadas += 1
            except Exception as exc:
                errores.append(f"{especificacion.name}: {exc}")
        self.estado.configure(text=f"{guardadas} HII guardadas")
        if errores:
            messagebox.showerror("Algunas HII no se generaron", "\n".join(errores), parent=self.ventana)

    def mostrar_excel(self, contenido, datos):
        """Dibuja celdas, merges y textos de objetos del MISMO XLSM que se guarda."""
        for widget in self.derecha.winfo_children():
            widget.destroy()
        ttk.Label(self.derecha, text="VISTA PREVIA HII — EXCEL", padding=8).pack()
        marco = ttk.Frame(self.derecha)
        marco.pack(fill="both", expand=True)
        canvas = tk.Canvas(marco, background="#e5e7eb")
        vertical = ttk.Scrollbar(marco, orient="vertical", command=canvas.yview)
        horizontal = ttk.Scrollbar(marco, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)
        canvas.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        libro = load_workbook(io.BytesIO(contenido), data_only=True)
        hoja = libro.worksheets[0]
        filas = 43 + max(0, len(datos.pruebas) - 6)
        xs, ys = [20], [20]
        for col in range(1, 7):
            xs.append(xs[-1] + max(45, hoja.column_dimensions[get_column_letter(col)].width * 7))
        for fila in range(1, filas + 1):
            alto = (hoja.row_dimensions[fila].height or 15) * 1.33
            if 15 <= fila < 15 + len(datos.pruebas):
                alto = max(alto, max(len(textwrap.wrap(texto(hoja.cell(fila, col).value),
                                       max(5, int((xs[col] - xs[col-1]) / 6)))) for col in range(1, 7)) * 14 + 6)
            ys.append(ys[-1] + alto)
        canvas.create_rectangle(20, 20, xs[-1], ys[-1], fill="white", outline="white")
        cubiertas = set()
        for combinado in hoja.merged_cells.ranges:
            for fila in range(combinado.min_row, combinado.max_row + 1):
                for col in range(combinado.min_col, combinado.max_col + 1):
                    if (fila, col) != (combinado.min_row, combinado.min_col):
                        cubiertas.add((fila, col))
        for fila in range(1, filas + 1):
            for col in range(1, 7):
                if (fila, col) in cubiertas:
                    continue
                celda = hoja.cell(fila, col)
                combinado = next((m for m in hoja.merged_cells.ranges if celda.coordinate in m), None)
                ultima_col = min(6, combinado.max_col) if combinado else col
                ultima_fila = min(filas, combinado.max_row) if combinado else fila
                x, y, x2, y2 = xs[col-1], ys[fila-1], xs[ultima_col], ys[ultima_fila]
                relleno = "white"
                if celda.fill.fgColor.type == "rgb" and celda.fill.patternType == "solid":
                    relleno = "#" + celda.fill.fgColor.rgb[-6:]
                borde = any(getattr(celda.border, lado).style for lado in ("left", "right", "top", "bottom"))
                canvas.create_rectangle(x, y, x2, y2, fill=relleno, outline="#666" if borde else relleno)
                if celda.value is not None:
                    canvas.create_text(x + 4, y + 3, text=texto(celda.value), anchor="nw", width=max(10, x2-x-8),
                                       font=("Arial", 9, "bold" if celda.font.bold else "normal"))
        self.imagenes = []
        with ZipFile(io.BytesIO(contenido)) as z:
            for nombre in z.namelist():
                if not (nombre.startswith("xl/drawings/") and nombre.endswith(".xml") and "/_rels/" not in nombre):
                    continue
                for ancla in ET.fromstring(z.read(nombre)):
                    origen = ancla.find("xdr:from", NS)
                    if origen is None:
                        continue
                    col, fila = int(origen.find("xdr:col", NS).text), int(origen.find("xdr:row", NS).text)
                    textos = ancla.findall(".//a:t", NS)
                    if col >= 6 or fila >= filas:
                        continue
                    xoff = int(origen.find("xdr:colOff", NS).text) / 9525
                    yoff = int(origen.find("xdr:rowOff", NS).text) / 9525
                    destino = ancla.find("xdr:to", NS)
                    ancho = max(60, xs[-1]-xs[col]-xoff)
                    if destino is not None:
                        col_fin = int(destino.find("xdr:col", NS).text)
                        if col_fin <= 6:
                            ancho = max(30, xs[col_fin] + int(destino.find("xdr:colOff", NS).text) / 9525 - xs[col] - xoff)
                    if textos:
                        valor = "\n".join(t.text for t in textos if t.text)
                        canvas.create_text(xs[col] + xoff, ys[fila] + yoff, text=valor, anchor="nw",
                                           width=ancho, font=("Arial", 10, "bold"))
                    imagen = ancla.find(".//a:blip", NS)
                    if imagen is not None:
                        import posixpath
                        relacion = imagen.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                        ruta_rels = posixpath.join(posixpath.dirname(nombre), "_rels", posixpath.basename(nombre) + ".rels")
                        rels = ET.fromstring(z.read(ruta_rels))
                        destino_imagen = next(r.get("Target") for r in rels if r.get("Id") == relacion)
                        ruta_imagen = (destino_imagen.lstrip("/") if destino_imagen.startswith("/") else
                                       posixpath.normpath(posixpath.join(posixpath.dirname(nombre), destino_imagen)))
                        foto = tk.PhotoImage(master=canvas, data=z.read(ruta_imagen))
                        factor = max(1, math.ceil(foto.width() / ancho))
                        foto = foto.subsample(factor, factor)
                        self.imagenes.append(foto)
                        canvas.create_image(xs[col] + xoff, ys[fila] + yoff, image=foto, anchor="nw")
        canvas.configure(scrollregion=(0, 0, xs[-1] + 20, ys[-1] + 20))
        libro.close()

    def guardar(self):
        if not self.excel:
            return
        ruta = filedialog.asksaveasfilename(parent=self.ventana, defaultextension=".xlsm",
            initialfile=f"{self.datos.codigo} HII {self.datos.idioma}.xlsm", filetypes=[("Excel con macros", "*.xlsm")])
        if ruta:
            try:
                destino = Path(ruta)
                if destino.resolve() in {p.resolve() for p in self.formatos.glob("*.xlsm")}:
                    raise ValueError("Guarda la HII con otro nombre para conservar la plantilla.")
                destino.write_bytes(self.excel)
                self.registrar_hii(destino, self.actual, self.datos)
                self.estado.configure(text=f"Guardado: {destino.name}")
            except (OSError, sqlite3.Error) as exc:
                messagebox.showerror("No se pudo guardar", str(exc), parent=self.ventana)
            except ValueError as exc:
                messagebox.showerror("No se pudo guardar", str(exc), parent=self.ventana)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--formatos", type=lambda valor: resolver_ruta(valor, BASE))
    parser.add_argument("--json", type=lambda valor: resolver_ruta(valor, BASE))
    args = parser.parse_args()
    raiz = ui.Raiz()
    raiz.withdraw()
    generador = GeneradorHII(raiz, args.formatos, args.json)
    generador.ventana.protocol("WM_DELETE_WINDOW", raiz.destroy)
    raiz.mainloop()


if __name__ == "__main__":
    main()
