"""CustomTkinter, preferencias visuales y adaptación de todas las ventanas."""
from __future__ import annotations

import json
import os
import re
import tempfile
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk as ttk_original
from pathlib import Path
from types import SimpleNamespace
import weakref

import customtkinter as ctk

BASE = Path(__file__).resolve().parent
ARCHIVO = BASE / "preferencias_interfaz.json"
PREDETERMINADAS = {"apariencia": "Claro", "acento": "Azul", "escala": 1.0,
                  "densidad": "Cómoda", "fuente": "Segoe UI", "maximizar": False,
                  "control_compacto": False, "adaptacion_auto": True}
ACENTOS = {"Azul": "#1765c1", "Petróleo": "#087f8c", "Índigo": "#6152c7", "Verde": "#16826a"}
VENTANAS = weakref.WeakSet()


def validar_preferencias(datos):
    resultado = dict(PREDETERMINADAS)
    if not isinstance(datos, dict):
        return resultado
    for clave, opciones in (("apariencia", ("Claro", "Oscuro", "Sistema")),
                            ("acento", tuple(ACENTOS)), ("densidad", ("Compacta", "Cómoda")),
                            ("fuente", ("Segoe UI", "Arial", "Calibri"))):
        if datos.get(clave) in opciones:
            resultado[clave] = datos[clave]
    escala = datos.get("escala")
    if isinstance(escala, (int, float)) and not isinstance(escala, bool) and .8 <= escala <= 1.4:
        resultado["escala"] = float(escala)
    for clave in ("maximizar", "control_compacto", "adaptacion_auto"):
        if isinstance(datos.get(clave), bool):
            resultado[clave] = datos[clave]
    return resultado


def leer_preferencias():
    try:
        return validar_preferencias(json.loads(ARCHIVO.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        return dict(PREDETERMINADAS)


PREFERENCIAS = leer_preferencias()


def guardar_preferencias(datos):
    datos = validar_preferencias(datos)
    temporal = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=BASE,
                                        prefix=".apariencia_", suffix=".tmp", delete=False) as archivo:
            temporal = Path(archivo.name)
            json.dump(datos, archivo, ensure_ascii=False, indent=2)
        os.replace(temporal, ARCHIVO)
    finally:
        if temporal is not None:
            temporal.unlink(missing_ok=True)
    aplicar_preferencias(datos)


def paleta():
    oscuro = ctk.get_appearance_mode() == "Dark"
    return {"fondo": "#141c2a" if oscuro else "#eef3f9",
            "panel": "#202c3d" if oscuro else "#ffffff",
            "texto": "#e8eef8" if oscuro else "#172b45",
            "secundario": "#a5b5ca" if oscuro else "#5d718a",
            "borde": "#34465d" if oscuro else "#d8e2ef",
            "seleccion": "#29476a" if oscuro else "#dceafe",
            "alineado": "#93c5fd" if oscuro else "#163b70",
            "acento": ACENTOS[PREFERENCIAS["acento"]]}


def aplicar_preferencias(datos):
    validadas = validar_preferencias(datos)
    PREFERENCIAS.clear()
    PREFERENCIAS.update(validadas)
    ctk.set_appearance_mode({"Claro": "Light", "Oscuro": "Dark", "Sistema": "System"}[PREFERENCIAS["apariencia"]])
    ctk.set_widget_scaling(PREFERENCIAS["escala"])
    for ventana in list(VENTANAS):
        if ventana.winfo_exists():
            ventana.aplicar_estilo()


def dimensiones_seguras(ancho, alto, pantalla_ancho, pantalla_alto, escala=1.0):
    disponible_ancho = max(320, int((pantalla_ancho - 32) / escala))
    disponible_alto = max(240, int((pantalla_alto - 80) / escala))
    return min(ancho, disponible_ancho), min(alto, disponible_alto)


class VentanaAdaptable:
    def _iniciar_estilo(self):
        VENTANAS.add(self)
        self._escala_tk_original = getattr(self._root(), "_escala_tk_original", float(self.tk.call("tk", "scaling")))
        self.option_add("*Font", (PREFERENCIAS["fuente"], 10))
        self.after_idle(self.aplicar_estilo)

    def geometry(self, valor=None):
        if valor:
            coincidencia = re.match(r"^(\d+)x(\d+)(.*)$", valor)
            if coincidencia:
                ancho, alto = dimensiones_seguras(int(coincidencia[1]), int(coincidencia[2]),
                    self.winfo_screenwidth(), self.winfo_screenheight(), self._get_window_scaling())
                valor = f"{ancho}x{alto}{coincidencia[3]}"
        return super().geometry(valor)

    def minsize(self, ancho=None, alto=None):
        if ancho is not None and alto is not None:
            ancho, alto = dimensiones_seguras(ancho, alto, self.winfo_screenwidth(),
                                              self.winfo_screenheight(), self._get_window_scaling())
        return super().minsize(ancho, alto)

    def configure(self, **kwargs):
        if "background" in kwargs:
            kwargs["fg_color"] = kwargs.pop("background")
        if "bg" in kwargs:
            kwargs["fg_color"] = kwargs.pop("bg")
        return super().configure(**kwargs)

    config = configure

    def _set_appearance_mode(self, mode_string):
        super()._set_appearance_mode(mode_string)
        if hasattr(self, "_escala_tk_original"):
            self.after_idle(self.aplicar_estilo)

    def aplicar_estilo(self):
        colores = paleta()
        self.configure(fg_color=colores["fondo"])
        self.tk.call("tk", "scaling", self._escala_tk_original * PREFERENCIAS["escala"])
        estilo = ttk_original.Style(self)
        estilo.theme_use("clam")
        fuente = PREFERENCIAS["fuente"]
        pixeles = max(10, round(12 * self._get_window_scaling() * PREFERENCIAS["escala"]))
        fuente_normal = (fuente, -pixeles)
        compacto = PREFERENCIAS["densidad"] == "Compacta"
        estilo.configure(".", background=colores["fondo"], foreground=colores["texto"], font=fuente_normal)
        estilo.configure("TFrame", background=colores["fondo"])
        estilo.configure("Borde.TFrame", borderwidth=1, relief="solid", bordercolor=colores["borde"],
                         lightcolor=colores["borde"], darkcolor=colores["borde"])
        estilo.configure("TLabel", background=colores["fondo"], foreground=colores["texto"])
        estilo.configure("Seccion.TLabel", foreground=colores["alineado"])
        estilo.configure("TLabelframe", background=colores["fondo"], bordercolor=colores["borde"],
                         lightcolor=colores["borde"], darkcolor=colores["borde"], borderwidth=1, relief="solid")
        estilo.configure("TLabelframe.Label", background=colores["fondo"], foreground=colores["texto"], font=(fuente, -pixeles, "bold"))
        for nombre in ("TEntry", "TCombobox"):
            estilo.configure(nombre, fieldbackground=colores["panel"], foreground=colores["texto"],
                             background=colores["panel"], arrowcolor=colores["texto"],
                             arrowsize=max(12, round(pixeles * .8)),
                             lightcolor=colores["borde"], darkcolor=colores["borde"],
                             bordercolor=colores["borde"], padding=2 if compacto else 3)
            estilo.map(nombre, fieldbackground=[("disabled", colores["fondo"]), ("!disabled", colores["panel"])],
                       background=[("active", colores["seleccion"]), ("!disabled", colores["panel"])],
                       foreground=[("!disabled", colores["texto"])], arrowcolor=[("!disabled", colores["texto"])],
                       selectbackground=[("!disabled", colores["acento"])], selectforeground=[("!disabled", "white")])
        for opcion, valor in (("background", colores["panel"]), ("foreground", colores["texto"]),
                             ("selectBackground", colores["acento"]), ("selectForeground", "white"),
                             ("font", fuente_normal)):
            self.option_add(f"*TCombobox*Listbox.{opcion}", valor)
        for nombre in ("Treeview", "Generador.Treeview"):
            estilo.configure(nombre, background=colores["panel"], fieldbackground=colores["panel"],
                             foreground=colores["texto"], rowheight=max(24 if compacto else 29, pixeles + 10),
                             bordercolor=colores["borde"], lightcolor=colores["borde"],
                             darkcolor=colores["borde"], borderwidth=1, font=fuente_normal)
            estilo.map(nombre, background=[("selected", colores["acento"])], foreground=[("selected", "white")])
            estilo.configure(nombre + ".Heading", background=colores["fondo"], foreground=colores["texto"],
                             bordercolor=colores["borde"], lightcolor=colores["borde"], darkcolor=colores["borde"],
                             font=(fuente, -pixeles, "bold"), padding=(6, 4 if compacto else 7))
            estilo.map(nombre + ".Heading", background=[("active", colores["seleccion"]), ("!disabled", colores["fondo"])],
                       foreground=[("!disabled", colores["texto"])])
        estilo.configure("TNotebook", background=colores["fondo"], borderwidth=0,
                         bordercolor=colores["borde"], lightcolor=colores["borde"], darkcolor=colores["borde"])
        estilo.configure("TNotebook.Tab", padding=(16, 7), background=colores["panel"], foreground=colores["texto"],
                         bordercolor=colores["borde"], lightcolor=colores["borde"], darkcolor=colores["borde"])
        estilo.map("TNotebook.Tab", background=[("selected", colores["acento"])], foreground=[("selected", "white")])
        estilo.configure("Horizontal.TProgressbar", background=colores["acento"], troughcolor=colores["borde"])
        for nombre in ("Horizontal.TScrollbar", "Vertical.TScrollbar"):
            estilo.configure(nombre, background=colores["borde"], troughcolor=colores["fondo"],
                             bordercolor=colores["fondo"], lightcolor=colores["borde"],
                             darkcolor=colores["borde"], arrowcolor=colores["texto"])
        self._estilizar_hijos(self)

    def _estilizar_hijos(self, padre, ajustar_fuentes=True):
        ajustar_fuentes = ajustar_fuentes and not getattr(padre, "_ui_preservar_fuentes", False)
        colores = paleta()
        pixeles = max(10, round(12 * self._get_window_scaling() * PREFERENCIAS["escala"]))
        for widget in padre.winfo_children():
            if isinstance(widget, ctk.CTkScrollableFrame):
                self._estilizar_hijos(widget, ajustar_fuentes)
                continue
            if isinstance(widget, (ctk.CTkBaseClass, ctk.CTkToplevel)):
                if isinstance(widget, Boton):
                    secundario = widget._variante == "secundario"
                    widget.configure(bg_color=widget._fondo_control or colores["fondo"],
                                     fg_color=colores["seleccion"] if secundario else colores["acento"],
                                     hover_color=colores["borde"] if secundario else "#10446f",
                                     text_color=colores["texto"] if secundario else "white",
                                     height=widget._altura_fija if widget._altura_fija is not None else
                                     ((18 if PREFERENCIAS["densidad"] == "Compacta" else 20) if widget._boton_pequeno else
                                      (26 if PREFERENCIAS["densidad"] == "Compacta" else 30)))
                elif isinstance(widget, ctk.CTkOptionMenu):
                    widget.configure(fg_color=colores["acento"], button_color=colores["acento"])
                elif isinstance(widget, ctk.CTkSlider):
                    widget.configure(button_color=colores["acento"], progress_color=colores["acento"])
                elif isinstance(widget, ctk.CTkSwitch):
                    widget.configure(progress_color=colores["acento"])
                if ajustar_fuentes and isinstance(widget, (ctk.CTkButton, ctk.CTkLabel, ctk.CTkEntry, ctk.CTkOptionMenu, ctk.CTkTextbox)):
                    fuente = widget.cget("font")
                    if isinstance(fuente, ctk.CTkFont):
                        fuente.configure(family=PREFERENCIAS["fuente"], size=12 if abs(fuente.cget("size")) < 18 else fuente.cget("size"))
                    elif isinstance(fuente, tuple):
                        tamano = 12 if abs(fuente[1]) < 18 else fuente[1]
                        widget.configure(font=(PREFERENCIAS["fuente"], tamano, *fuente[2:]))
                if isinstance(widget, ctk.CTkFrame):
                    self._estilizar_hijos(widget, ajustar_fuentes)
                continue
            if isinstance(widget, tk.Listbox):
                widget.configure(background=colores["panel"], foreground=colores["texto"],
                                 selectbackground=colores["seleccion"], selectforeground=colores["texto"],
                                 highlightbackground=colores["borde"], highlightcolor=colores["acento"], borderwidth=0)
                for indice in range(widget.size()):
                    if str(widget.itemcget(indice, "foreground")) in {"#163b70", "#93c5fd"}:
                        widget.itemconfigure(indice, foreground=colores["alineado"])
            elif isinstance(widget, tk.Text):
                widget.configure(background=colores["panel"], foreground=colores["texto"], insertbackground=colores["texto"])
            elif isinstance(widget, ttk_original.Treeview):
                widget.tag_configure("fila_alterna", background=colores["fondo"])
            elif isinstance(widget, tk.PanedWindow) and not isinstance(widget, ttk_original.Widget):
                widget.configure(background=colores["fondo"], sashrelief="flat", borderwidth=0)
            elif isinstance(widget, tk.Canvas):
                widget.configure(background=colores["fondo"], highlightbackground=colores["borde"])
            elif isinstance(widget, (tk.Frame, tk.Label)) and not isinstance(widget, ttk_original.Widget):
                fondo = str(widget.cget("background"))
                if fondo in {"#edf2f7", "#f0f0f0", "#d9d9d9", "#eef3f9", "#141c2a", "SystemButtonFace"}:
                    widget.configure(background=colores["fondo"])
                    if isinstance(widget, tk.Label):
                        widget.configure(foreground=colores["texto"])
            elif isinstance(widget, ttk_original.Label):
                originales = {"#163b70": "alineado", "#526174": "secundario", "#9a3412": "aviso"}
                color = str(widget.cget("foreground"))
                if not hasattr(widget, "_ui_rol_color") and color in originales:
                    widget._ui_rol_color = originales[color]
                rol = getattr(widget, "_ui_rol_color", None)
                if rol:
                    widget.configure(foreground=("#fdba74" if ctk.get_appearance_mode() == "Dark" else "#9a3412")
                                     if rol == "aviso" else colores[rol])
            if ajustar_fuentes and isinstance(widget, (ttk_original.Entry, ttk_original.Combobox, tk.Listbox, tk.Text, tk.Label, ttk_original.Label)):
                actual = tkfont.Font(root=widget, font=widget.cget("font") or "TkDefaultFont")
                if not hasattr(widget, "_ui_fuente_control"):
                    tamano = actual.actual("size")
                    widget._ui_fuente_control = 0 <= tamano < 14 or tamano < 0 and abs(tamano) <= pixeles * 1.3
                if widget._ui_fuente_control:
                    familia = actual.actual("family") if actual.actual("family") in {"Consolas", "Courier New"} else PREFERENCIAS["fuente"]
                    widget.configure(font=(familia, -pixeles, actual.actual("weight")))
                if isinstance(widget, ttk_original.Combobox):
                    popup = self.tk.call("ttk::combobox::PopdownWindow", str(widget))
                    self.tk.call(str(popup) + ".f.l", "configure", "-background", colores["panel"],
                                 "-foreground", colores["texto"], "-selectbackground", colores["acento"],
                                 "-selectforeground", "white", "-font", (PREFERENCIAS["fuente"], -pixeles))
            self._estilizar_hijos(widget, ajustar_fuentes)


class Raiz(VentanaAdaptable, ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._iniciar_estilo()


class Ventana(VentanaAdaptable, ctk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._iniciar_estilo()
        self._presentacion_inicial_hecha = False
        self.bind("<Map>", self._programar_presentacion, add="+")

    def _programar_presentacion(self, evento):
        if evento.widget is self and not self._presentacion_inicial_hecha:
            # CTk restaura el foco previo después de dibujar la barra del título.
            self.after(50, self._mostrar_al_frente)

    def _mostrar_al_frente(self):
        if self._presentacion_inicial_hecha or not self.winfo_exists() or not self.winfo_ismapped():
            return
        self.lift()
        if 0 <= self.winfo_rootx() < self.winfo_screenwidth() and 0 <= self.winfo_rooty() < self.winfo_screenheight():
            self.focus_set()
        self._presentacion_inicial_hecha = True


class Boton(ctk.CTkButton):
    """Botón CTk compatible con las llamadas existentes de ttk.Button."""
    def __init__(self, master=None, **kwargs):
        self._variante = kwargs.pop("variante", "principal")
        self._fondo_control = kwargs.pop("fondo", None)
        ancho = kwargs.pop("width", None)
        self._ancho_automatico = ancho is None
        self._boton_pequeno = ancho is not None and ancho <= 4
        self._altura_fija = kwargs.get("height")
        kwargs.pop("style", None)
        kwargs.pop("padding", None)
        kwargs.setdefault("width", (max(22, ancho * 6 + 6) if self._boton_pequeno else ancho * 8 + 16) if ancho is not None else
                          max(110, len(str(kwargs.get("text", ""))) * 7 + 24))
        kwargs.setdefault("height", (18 if PREFERENCIAS["densidad"] == "Compacta" else 20) if self._boton_pequeno else
                          (26 if PREFERENCIAS["densidad"] == "Compacta" else 30))
        kwargs.setdefault("corner_radius", 6)
        kwargs.setdefault("bg_color", self._fondo_control or paleta()["fondo"])
        kwargs.setdefault("fg_color", paleta()["seleccion"] if self._variante == "secundario" else paleta()["acento"])
        kwargs.setdefault("hover_color", paleta()["borde"] if self._variante == "secundario" else "#10446f")
        kwargs.setdefault("text_color", paleta()["texto"] if self._variante == "secundario" else "white")
        kwargs.setdefault("font", (PREFERENCIAS["fuente"], 12))
        super().__init__(master, **kwargs)

    def configure(self, require_redraw=False, **kwargs):
        if "text" in kwargs and self._ancho_automatico and "width" not in kwargs:
            kwargs["width"] = max(110, len(str(kwargs["text"])) * 7 + 24)
        return super().configure(require_redraw=require_redraw, **kwargs)


ttk = SimpleNamespace(**{nombre: getattr(ttk_original, nombre) for nombre in dir(ttk_original) if not nombre.startswith("_")})
ttk.Button = Boton


class AyudaHover:
    def __init__(self, widget, texto):
        self.widget, self.texto = widget, texto
        widget._ayuda_hover = self
        self.pendiente = self.ventana = None
        widget.bind("<Enter>", self.programar, add="+")
        widget.bind("<Leave>", self.ocultar, add="+")
        widget.bind("<Button-1>", self.ocultar, add="+")
        widget.bind("<Destroy>", self.ocultar, add="+")

    def programar(self, _evento=None):
        self.ocultar()
        self.pendiente = self.widget.after(400, self.mostrar)

    def mostrar(self):
        self.pendiente = None
        if not self.widget.winfo_exists():
            return
        ventana = self.ventana = ctk.CTkToplevel(self.widget)
        ventana.withdraw()
        ventana.overrideredirect(True)
        ventana.configure(fg_color=paleta()["panel"])
        ctk.CTkLabel(ventana, text=self.texto, wraplength=260, justify="left",
                     font=(PREFERENCIAS["fuente"], 12), text_color=paleta()["texto"]).pack(padx=12, pady=8)
        ventana.update_idletasks()
        x = min(self.widget.winfo_rootx() + self.widget.winfo_width() + 8,
                self.widget.winfo_screenwidth() - ventana.winfo_reqwidth() - 12)
        y = min(self.widget.winfo_rooty(), self.widget.winfo_screenheight() - ventana.winfo_reqheight() - 12)
        if self.widget.winfo_rootx() >= self.widget.winfo_screenwidth():
            x, y = self.widget.winfo_rootx() + self.widget.winfo_width() + 8, self.widget.winfo_rooty()
        ventana.geometry(f"+{max(0, x)}+{max(0, y)}")
        ventana.deiconify()

    def ocultar(self, _evento=None):
        if self.pendiente:
            self.widget.after_cancel(self.pendiente)
            self.pendiente = None
        if self.ventana:
            self.ventana.destroy()
            self.ventana = None
ctk.set_appearance_mode({"Claro": "Light", "Oscuro": "Dark", "Sistema": "System"}[PREFERENCIAS["apariencia"]])
ctk.set_widget_scaling(PREFERENCIAS["escala"])
