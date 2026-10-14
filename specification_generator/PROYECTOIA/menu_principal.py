"""Menú principal por pestañas y personalización del espacio de trabajo."""
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import estilo_interfaz as ui
from estilo_interfaz import ctk


class BarraPestanas(ctk.CTkFrame):
    """Pestañas repartidas en el ancho disponible, sin barra de desplazamiento."""
    def __init__(self, padre, command=None, **kwargs):
        self.command = command
        self.seleccion = None
        self.paginas, self.botones = {}, {}
        self.color_seleccion = kwargs.pop("segmented_button_selected_color", ui.paleta()["acento"])
        self.color_inactivo = kwargs.pop("segmented_button_unselected_color", ("#425a77", "#29384d"))
        kwargs.pop("anchor", None)
        kwargs.pop("text_color", None)
        super().__init__(padre, **kwargs)
        self.barra = ctk.CTkFrame(self, fg_color="transparent")
        self.barra.pack(fill="x", padx=8, pady=(8, 0))
        self.cuerpo = ctk.CTkFrame(self, fg_color="transparent")
        self.cuerpo.pack(fill="both", expand=True)

    def add(self, nombre):
        pagina = ctk.CTkFrame(self.cuerpo, fg_color="transparent")
        self.paginas[nombre] = pagina
        columna = len(self.botones)
        self.barra.grid_columnconfigure(columna, weight=1, uniform="pestanas")
        boton = ctk.CTkButton(self.barra, text=nombre, width=1, height=36,
            fg_color=self.color_inactivo, hover_color="#10446f", text_color="white",
            command=lambda: self.set(nombre))
        boton.grid(row=0, column=columna, sticky="ew", padx=4, pady=2)
        self.botones[nombre] = boton
        return pagina

    def set(self, nombre):
        if self.seleccion:
            self.paginas[self.seleccion].pack_forget()
            self.botones[self.seleccion].configure(fg_color=self.color_inactivo)
        self.seleccion = nombre
        self.paginas[nombre].pack(fill="both", expand=True)
        self.botones[nombre].configure(fg_color=self.color_seleccion)
        if self.command:
            self.command()

    def get(self):
        return self.seleccion

    def tab(self, nombre):
        return self.paginas[nombre]

    def configure(self, require_redraw=False, **kwargs):
        if "segmented_button_selected_color" in kwargs:
            self.color_seleccion = kwargs.pop("segmented_button_selected_color")
            if self.seleccion:
                self.botones[self.seleccion].configure(fg_color=self.color_seleccion)
        return super().configure(require_redraw=require_redraw, **kwargs)


class MenuPrincipal:
    def __init__(self, app):
        self.app = app
        self.raiz = app.raiz
        self.creadas = set()
        self.tarjetas = []
        colores = ui.paleta()
        self.crear_pie()
        self.pestanas = BarraPestanas(self.raiz, command=self.cambiar_pestana,
            corner_radius=12, anchor="w", fg_color=("#ffffff", "#202c3d"),
            segmented_button_selected_color=colores["acento"],
            segmented_button_unselected_color=("#425a77", "#29384d"), text_color="white")
        self.pestanas.pack(fill="both", expand=True, padx=18, pady=(8, 12))
        for titulo in ("Inicio", "Instrucciones", "Histórico", "Formatos", "Métodos", "Configuración", "Personalizar"):
            self.pestanas.add(titulo)
        self.pestanas.set("Inicio")
        self.cambiar_pestana()
        self.raiz.bind("<Configure>", self.adaptar, add="+")
        self.raiz.after(150, lambda: self.adaptar(None))

    def crear_pie(self):
        from interfaz_aplicacion_especificaciones import VERSION, VERSION_ALINEADOR
        app = self.app
        pie = ctk.CTkFrame(self.raiz, corner_radius=0, fg_color=("#e5edf7", "#182235"))
        pie.pack(side="bottom", fill="x")
        pie.grid_columnconfigure(1, weight=1)
        app._barra_progreso_carga = pie
        app._estado_progreso_carga = ctk.CTkLabel(pie, text="Listo.", width=90, anchor="w")
        app._estado_progreso_carga.grid(row=0, column=0, padx=(18, 8), pady=7)
        app._indicador_progreso_carga = ui.ttk.Progressbar(pie, mode="determinate", maximum=100)
        app._indicador_progreso_carga.grid(row=0, column=1, sticky="ew", pady=7)
        ui.Boton(pie, text="Ver más", width=8, command=app._mostrar_detalles_carga).grid(row=0, column=2, padx=12, pady=7)
        ctk.CTkLabel(pie, text=f"Interfaz v{VERSION}  ·  Alineador v{VERSION_ALINEADOR}",
                     font=(ui.PREFERENCIAS["fuente"], 11)).grid(row=0, column=3, padx=(4, 18), pady=7)

    def pagina(self, nombre, titulo, subtitulo):
        pagina = ctk.CTkScrollableFrame(self.pestanas.tab(nombre), fg_color="transparent", corner_radius=0)
        pagina.pack(fill="both", expand=True, padx=8, pady=8)
        pagina.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(pagina, text=titulo, font=(ui.PREFERENCIAS["fuente"], 24, "bold"),
                     anchor="w").pack(fill="x", padx=14, pady=(8, 2))
        ctk.CTkLabel(pagina, text=subtitulo, anchor="w", justify="left", wraplength=850,
                     text_color=("#5d718a", "#a5b5ca")).pack(fill="x", padx=14, pady=(0, 18))
        return pagina

    def cambiar_pestana(self):
        nombre = self.pestanas.get()
        if nombre in self.creadas:
            if nombre == "Histórico" and hasattr(self, "historico"):
                self.historico._cargar_pagina()
            return
        self.creadas.add(nombre)
        creadores = {"Inicio": self.inicio, "Instrucciones": self.instrucciones,
                     "Histórico": self.historial, "Formatos": self.formatos, "Métodos": self.metodos,
                     "Configuración": self.configuracion, "Personalizar": self.personalizar}
        creadores[nombre]()
        self.raiz.after_idle(self.raiz.aplicar_estilo)

    def inicio(self):
        pagina = ctk.CTkFrame(self.pestanas.tab("Inicio"), fg_color="transparent")
        pagina.pack(fill="both", expand=True, padx=18, pady=18)
        pagina.grid_columnconfigure(0, weight=1, uniform="inicio")
        pagina.grid_columnconfigure(1, weight=2, uniform="inicio")
        pagina.grid_rowconfigure(0, weight=1)
        self.rejilla = ctk.CTkFrame(pagina, fg_color="transparent")
        self.rejilla.grid(row=0, column=0, sticky="new", padx=(0, 14))
        self.rejilla.grid_columnconfigure(0, weight=1)
        self.espacio_reservado = ctk.CTkFrame(pagina, fg_color="transparent")
        self.espacio_reservado.grid(row=0, column=1, sticky="nsew")
        for numero, (titulo, detalle, comando) in enumerate((
            ("SUBIR ESPEC", "Extrae documentos Word y revisa su alineación.", self.app.subir_especificacion),
            ("GENERAR ESPEC", "Crea una especificación o completa sus pruebas.", self.app.generar_especificacion),
            ("SELECCIONAR ESPEC", "Busca por producto, código, cliente o revisión.", self.app.seleccionar_especificacion),
            ("COMPARAR ESPECS", "Revisa versiones y diferencias en paralelo.", self.app.comparar_especificaciones),
            ("GENERAR HII", "Prepara hojas Excel en español o inglés.", self.app.generar_hii),
        )):
            boton = ui.Boton(self.rejilla, text=titulo, command=comando, height=48, font=(ui.PREFERENCIAS["fuente"], 12, "bold"))
            boton.grid(row=numero, column=0, sticky="ew", pady=(0, 12))
            ui.AyudaHover(boton, detalle)
            self.tarjetas.append(boton)
        self.adaptar(None)

    def adaptar(self, evento):
        if evento is not None and evento.widget is not self.raiz:
            return
        if not hasattr(self, "rejilla"):
            return
        alto = self.raiz.winfo_height() / self.raiz._get_window_scaling() / ui.PREFERENCIAS["escala"]
        altura_boton = max(30, min(48, int((alto - 155) / 5) - 12))
        for boton in self.tarjetas:
            boton._altura_fija = altura_boton
            boton.configure(height=altura_boton)

    def instrucciones(self):
        from interfaz_aplicacion_especificaciones import BASE
        pagina = self.pagina("Instrucciones", "Guía del gestor", "Consulta las instrucciones sin salir de la aplicación.")
        selector = ctk.CTkOptionMenu(pagina, values=["README_ACTUALIZACION_CODIGOS_200.md", "RESUMEN_ALINEADOR_ESPECIFICACIONES.md"],
                                    command=lambda nombre: cargar(nombre))
        selector.pack(anchor="w", padx=14, pady=(0, 12))
        texto = ctk.CTkTextbox(pagina, height=360, wrap="word")
        texto.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        def cargar(nombre):
            try:
                contenido = (BASE / nombre).read_text(encoding="utf-8-sig")
            except OSError as exc:
                contenido = str(exc)
            texto.configure(state="normal")
            texto.delete("1.0", "end")
            texto.insert("1.0", contenido)
            texto.configure(state="disabled")
        cargar(selector.get())

    def historial(self):
        from interfaz_aplicacion_especificaciones import VisorHistorico, BASE
        contenedor = self.pestanas.tab("Histórico")
        try:
            self.historico = VisorHistorico(self.raiz, BASE / "especificaciones.db", self.app.rutas["historico"], contenedor=contenedor)
        except (OSError, ValueError) as exc:
            ctk.CTkLabel(contenedor, text=str(exc), wraplength=700).pack(pady=30)
            ui.Boton(contenedor, text="Volver a intentar", command=lambda: self.reintentar_historial(contenedor)).pack()

    def reintentar_historial(self, contenedor):
        for widget in contenedor.winfo_children():
            widget.destroy()
        self.historial()

    def formatos(self):
        pagina = self.pagina("Formatos", "Plantillas de documentos", "Gestiona los formatos Word y Excel que se usan para especificaciones y HII.")
        ui.Boton(pagina, text="Añadir o actualizar plantillas", command=self.app.actualizar_formatos).pack(anchor="w", padx=14)
        lista = ctk.CTkTextbox(pagina, height=300)
        lista.pack(fill="x", padx=14, pady=16)
        def actualizar():
            carpeta = self.app.rutas["formatos"]
            archivos = sorted(p.name for p in carpeta.iterdir() if p.is_file()) if carpeta.is_dir() else []
            lista.configure(state="normal")
            lista.delete("1.0", "end")
            lista.insert("1.0", str(carpeta) + "\n\n" + "\n".join(archivos))
            lista.configure(state="disabled")
        ui.Boton(pagina, text="Actualizar lista", command=actualizar).pack(anchor="w", padx=14)
        actualizar()

    def metodos(self):
        from interfaz_aplicacion_especificaciones import EditorListaMaestra, BASE, MAESTRO
        pagina = self.pagina("Métodos", "Lista maestra de métodos", "Consulta y edita los métodos del laboratorio.")
        def editar():
            try:
                EditorListaMaestra(self.raiz, BASE / MAESTRO)
            except (OSError, ValueError) as exc:
                messagebox.showerror("Lista maestra", str(exc), parent=self.raiz)
        ui.Boton(pagina, text="Editar lista maestra", command=editar).pack(anchor="w", padx=14, pady=6)
        ui.Boton(pagina, text="Ubicación y otras opciones", command=self.app.actualizar_metodos).pack(anchor="w", padx=14, pady=6)

    def configuracion(self):
        from interfaz_aplicacion_especificaciones import _guardar_configuracion
        pagina = self.pagina("Configuración", "Carpetas del proyecto", "Elige dónde leer documentos y guardar los resultados.")
        campos = {}
        for clave, etiqueta in (("especificaciones_viejas", "Documentos Word recibidos"), ("json_anteriores", "JSON extraídos"),
            ("json_alineadas", "JSON alineados"), ("formatos", "Plantillas Word y Excel"), ("historico", "Documentos generados")):
            ctk.CTkLabel(pagina, text=etiqueta, anchor="w", font=(ui.PREFERENCIAS["fuente"], 13, "bold")).pack(fill="x", padx=14, pady=(9, 2))
            fila = ctk.CTkFrame(pagina, fg_color="transparent")
            fila.pack(fill="x", padx=14)
            variable = tk.StringVar(value=str(self.app.rutas[clave]))
            campos[clave] = variable
            ctk.CTkEntry(fila, textvariable=variable).pack(side="left", fill="x", expand=True, padx=(0, 8))
            ui.Boton(fila, text="Examinar", width=10, command=lambda v=variable: self.app._elegir_carpeta(v, self.raiz)).pack(side="right")
        def guardar():
            nuevas = {clave: Path(variable.get()).expanduser().resolve() for clave, variable in campos.items()}
            try:
                _guardar_configuracion(nuevas)
                self.app.rutas = nuevas
                if hasattr(self, "historico"):
                    self.historico.carpeta_historico = nuevas["historico"]
                messagebox.showinfo("Configuración", "Carpetas guardadas.", parent=self.raiz)
            except OSError as exc:
                messagebox.showerror("Configuración", str(exc), parent=self.raiz)
        ui.Boton(pagina, text="Guardar carpetas", command=guardar).pack(anchor="w", padx=14, pady=18)

    def personalizar(self):
        pagina = ctk.CTkScrollableFrame(self.pestanas.tab("Personalizar"), fg_color="transparent", corner_radius=0)
        pagina.pack(fill="both", expand=True, padx=8, pady=8)
        variables = {}
        for clave, etiqueta, opciones in (("apariencia", "Apariencia", ["Claro", "Oscuro", "Sistema"]),
            ("acento", "Color de acento", list(ui.ACENTOS)), ("densidad", "Espaciado y altura de filas", ["Cómoda", "Compacta"]),
            ("fuente", "Tipografía", ["Segoe UI", "Arial", "Calibri"])):
            fila = ctk.CTkFrame(pagina, fg_color="transparent")
            fila.pack(fill="x", padx=14, pady=7)
            ctk.CTkLabel(fila, text=etiqueta, width=230, anchor="w").pack(side="left")
            variable = tk.StringVar(value=ui.PREFERENCIAS[clave])
            variables[clave] = variable
            ctk.CTkOptionMenu(fila, variable=variable, values=opciones).pack(side="left")
        escala = tk.DoubleVar(value=ui.PREFERENCIAS["escala"])
        ctk.CTkLabel(pagina, text="Tamaño de controles y texto (80–140 %)", anchor="w").pack(fill="x", padx=14, pady=(14, 2))
        porcentaje = ctk.CTkLabel(pagina, text=f"{round(escala.get()*100)} %")
        porcentaje.pack(anchor="w", padx=14)
        ctk.CTkSlider(pagina, from_=.8, to=1.4, number_of_steps=12, variable=escala,
                      command=lambda valor: porcentaje.configure(text=f"{round(valor*100)} %")).pack(fill="x", padx=14, pady=(0, 14))
        for clave, etiqueta in (("maximizar", "Abrir el menú principal maximizado"),
                                ("adaptacion_auto", "Optimizar automáticamente el editor en pantallas pequeñas"),
                                ("control_compacto", "Abrir el editor con control documental colapsado")):
            variable = tk.BooleanVar(value=ui.PREFERENCIAS[clave])
            variables[clave] = variable
            ctk.CTkSwitch(pagina, text=etiqueta, variable=variable).pack(anchor="w", padx=14, pady=7)
        def aplicar(restaurar=False):
            datos = dict(ui.PREDETERMINADAS) if restaurar else {**{k: v.get() for k, v in variables.items()}, "escala": escala.get()}
            try:
                ui.guardar_preferencias(datos)
            except OSError as exc:
                messagebox.showerror("Personalización", str(exc), parent=self.raiz)
                return
            for clave, variable in variables.items():
                variable.set(datos[clave])
            escala.set(datos["escala"])
            porcentaje.configure(text=f"{round(datos['escala']*100)} %")
            self.pestanas.configure(segmented_button_selected_color=ui.paleta()["acento"])
            self.app.zoom_aplicacion = datos["escala"]
            self.adaptar(None)
            if datos["maximizar"]:
                self.raiz.state("zoomed")
            else:
                self.raiz.state("normal")
        acciones = ctk.CTkFrame(pagina, fg_color="transparent")
        acciones.pack(fill="x", padx=14, pady=20)
        ui.Boton(acciones, text="Aplicar y guardar", command=aplicar).pack(side="left", padx=(0, 10))
        ui.Boton(acciones, text="Restaurar valores", command=lambda: aplicar(True)).pack(side="left")


def crear_menu_principal(app):
    app.menu_principal = MenuPrincipal(app)
