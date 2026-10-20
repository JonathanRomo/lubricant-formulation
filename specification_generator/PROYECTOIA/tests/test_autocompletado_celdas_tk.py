"""Regresiones de edición con los controles reales de Tk, sin abrir documentos."""
import tkinter as tk
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import interfaz_aplicacion_especificaciones as modulo


class AutocompletadoCeldasTkTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.raiz = tk.Tk()
        cls.raiz.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.raiz.destroy()

    def setUp(self):
        self.errores = []
        self.raiz.report_callback_exception = lambda *error: self.errores.append(error)
        g = self.generador = modulo.GeneradorEspecificacion.__new__(modulo.GeneradorEspecificacion)
        g.ventana = tk.Toplevel(self.raiz)
        g.ventana.withdraw()
        g.ventana.attributes("-alpha", 0)
        g.ventana.geometry("900x240+0+0")
        for nombre in (
            "_editor_celda", "_celda_editada", "_cierre_editor_pendiente",
            "_texto_editor", "_traza_texto_editor", "_sugerencias_editor_pendientes",
            "_comentario_marco", "_comentario_boton_mas", "_periodicidad_marco",
            "_periodicidad_boton_mas", "_popup_sugerencias", "_lista_sugerencias",
            "_entrada_sugerencias", "_al_seleccionar_sugerencia", "_sugerencia_actual",
        ):
            setattr(g, nombre, None)
        g._ultimo_texto_filtrado = ""
        g._periodicidad_editores = []
        g._nombres_sugeridos_por_fila = {}
        g._idioma = tk.StringVar(master=g.ventana, value="Español")
        g.maestro = []
        g._traductor = modulo.SERVICIO_TRADUCTOR.diccionario
        g.pruebas = [{campo: "" for campo in (*modulo.CAMPOS_TABLA, "periodicidad")}]
        g.pruebas[0]["periodicidad"] = modulo.PERIODICIDAD_PREDETERMINADA
        g.tabla_pruebas = modulo.ttk.Treeview(
            g.ventana, columns=(*modulo.CAMPOS_TABLA, "periodicidad"), show="headings", height=4,
        )
        for campo in (*modulo.CAMPOS_TABLA, "periodicidad"):
            g.tabla_pruebas.column(campo, width=140)
        g.tabla_pruebas.pack(fill="both", expand=True)
        g.tabla_pruebas.insert("", "end", iid="0", values=tuple(g.pruebas[0].values()))
        g.tabla_pruebas.bind("<Button-1>", g._editar_celda)
        g.etiqueta_sugerencia = modulo.ttk.Label(g.ventana)
        self.exterior = modulo.ttk.Label(g.ventana, text="Fuera de la celda")
        self.exterior.pack()
        g.notas = tk.Text(g.ventana)
        g._actualizar_tabla = Mock()
        g._actualizar_sugerencia_contextual = Mock()
        g._mostrar_referencia_bilingue = Mock()
        g._sugerencias = Mock(return_value=("ASTM D445", "ASTM D446", "ASTM D447"))
        g.ventana.bind("<Button-1>", g._cerrar_editor_al_clic_fuera, add="+")
        g.ventana.bind("<Destroy>", g._al_destruir_ventana_editor, add="+")
        g.ventana.deiconify()
        self.raiz.update()
        inicializar_toplevel = tk.Toplevel.__init__

        def auxiliar_invisible(ventana, *args, **kwargs):
            inicializar_toplevel(ventana, *args, **kwargs)
            ventana.attributes("-alpha", 0)

        parche = patch.object(tk.Toplevel, "__init__", auxiliar_invisible)
        parche.start()
        self.addCleanup(parche.stop)

    def tearDown(self):
        if self.generador.ventana.winfo_exists():
            self.generador.ventana.destroy()
        self.raiz.update()
        self.assertFalse(self.errores, self.errores)

    def abrir(self, campo="metodo"):
        columna = (*modulo.CAMPOS_TABLA, "periodicidad").index(campo) + 1
        self.generador._editar_celda(SimpleNamespace(iid="0", columna=f"#{columna}"))
        self.raiz.update()
        self.generador._editor_celda.focus_force()
        self.raiz.update()
        return self.generador._editor_celda

    def escribir(self, entrada, texto):
        entrada.delete(0, "end")
        entrada.insert(0, texto)
        self.raiz.update()

    def procesar_cierres(self):
        self.raiz.after(200, self.raiz.quit)
        self.raiz.mainloop()

    def test_campos_con_texto_libre_y_periodicidad_con_desplegable(self):
        for campo in ("prueba", "metodo", "unidades"):
            with self.subTest(campo=campo):
                entrada = self.abrir(campo)
                self.assertIsInstance(entrada, modulo.ttk.Entry)
                self.assertNotIsInstance(entrada, modulo.ttk.Combobox)
                self.generador._cerrar_editor_celda(guardar=False)
        self.abrir("periodicidad")
        self.assertEqual(str(self.generador._periodicidad_editores[0].cget("state")), "readonly")

    def test_escritura_navegacion_y_aceptacion_no_cierran_ni_reenfocan(self):
        g = self.generador
        entrada = self.abrir()
        with patch.object(entrada, "focus_set", wraps=entrada.focus_set) as enfocar:
            self.escribir(entrada, "ASTM")
            lista = g._lista_sugerencias
            entrada.event_generate("<Down>")
            entrada.event_generate("<Down>")
            entrada.event_generate("<Up>")
            entrada.event_generate("<KeyRelease-Down>")
            self.raiz.update()
            self.procesar_cierres()
            self.assertEqual(str(self.raiz.tk.call("focus")), str(entrada))
            self.assertEqual(lista.curselection(), (0,))
            self.assertEqual(entrada.get(), "ASTM")
            self.assertIs(g._editor_celda, entrada)
            enfocar.assert_not_called()
        entrada.event_generate("<Return>")
        self.raiz.update()
        self.assertEqual(entrada.get(), "ASTM D445")
        self.assertIs(g._editor_celda, entrada)
        self.assertIsNone(g._popup_sugerencias)
        entrada.insert("end", " revisado")
        self.raiz.update()
        self.assertEqual(entrada.get(), "ASTM D445 revisado")
        self.assertEqual(g.pruebas[0]["metodo"], "")

    def test_escape_no_reabre_la_lista_y_conserva_el_texto(self):
        entrada = self.abrir()
        self.escribir(entrada, "Texto libre")
        entrada.event_generate("<Escape>")
        entrada.event_generate("<KeyRelease-Escape>")
        self.raiz.update()
        self.assertIsNone(self.generador._popup_sugerencias)
        self.assertEqual(entrada.get(), "Texto libre")
        self.assertIs(self.generador._editor_celda, entrada)
        entrada.event_generate("<Down>")
        self.assertIsNotNone(self.generador._popup_sugerencias)

    def test_clic_en_sugerencia_no_guarda_y_permite_continuar(self):
        g = self.generador
        entrada = self.abrir()
        self.escribir(entrada, "ASTM")
        lista = g._lista_sugerencias
        x, y, ancho, alto = lista.bbox(1)
        lista.event_generate("<ButtonPress-1>", x=x + 3, y=y + alto // 2)
        lista.event_generate("<ButtonRelease-1>", x=x + 3, y=y + alto // 2)
        self.raiz.update()
        self.assertEqual(entrada.get(), "ASTM D446")
        self.assertIs(g._editor_celda, entrada)
        self.assertEqual(g.pruebas[0]["metodo"], "")

    def test_popup_y_boton_de_comentario_pertenecen_a_la_edicion(self):
        g = self.generador
        entrada = self.abrir()
        self.escribir(entrada, "ASTM")
        self.assertTrue(g._foco_en_editor(str(g._lista_sugerencias)))
        self.assertTrue(g._foco_en_editor(str(g._comentario_boton_mas)))
        # La ruta de otro control con un prefijo parecido no es un descendiente.
        self.assertFalse(g._foco_en_editor(str(g._comentario_marco) + "2"))
        self.assertFalse(g._foco_en_editor(str(self.exterior)))
        g._lista_sugerencias.focus_force()
        self.raiz.update()
        g._programar_cierre_editor()
        self.procesar_cierres()
        self.assertIs(g._editor_celda, entrada)
        self.exterior.focus_force()
        self.raiz.update()
        self.assertEqual(str(self.raiz.tk.call("focus")), str(self.exterior))
        g._programar_cierre_editor()
        self.procesar_cierres()
        self.assertIsNone(g._editor_celda)
        self.assertEqual(g.pruebas[0]["metodo"], "ASTM")

    def test_insertar_borrar_y_texto_sin_cambios_actualizan_solo_lo_necesario(self):
        g = self.generador
        entrada = self.abrir()
        entrada.insert(0, "ASTM D445")
        self.raiz.update()
        self.assertEqual(entrada.get(), "ASTM D445")
        entrada.icursor(0)
        entrada.event_generate("<Delete>")
        self.raiz.update()
        self.assertEqual(entrada.get(), "STM D445")
        g._sugerencias.reset_mock()
        entrada.event_generate("<Left>")
        entrada.event_generate("<KeyRelease-Left>")
        g._texto_editor.set(entrada.get())
        self.raiz.update()
        g._sugerencias.assert_not_called()

    def test_enter_no_acepta_una_opcion_del_texto_anterior(self):
        g = self.generador
        entrada = self.abrir()
        self.escribir(entrada, "ASTM")
        entrada.event_generate("<Down>")
        entrada.insert("end", " libre")
        entrada.event_generate("<Return>")
        self.raiz.update()
        self.assertIsNone(g._editor_celda)
        self.assertEqual(g.pruebas[0]["metodo"], "ASTM libre")

    def test_clic_fuera_y_shift_tab_guardan_texto_libre_sin_aceptar_opciones(self):
        for accion in ("clic", "shift-tab", "enter"):
            with self.subTest(accion=accion):
                g = self.generador
                entrada = self.abrir()
                self.escribir(entrada, "Texto libre " + accion)
                if accion == "clic":
                    self.exterior.event_generate("<Button-1>", x=2, y=2)
                else:
                    entrada.event_generate({
                        "tab": "<Tab>", "shift-tab": "<Shift-Tab>", "enter": "<Return>",
                    }[accion])
                self.raiz.update()
                self.assertIsNone(g._editor_celda)
                self.assertIsNone(g._popup_sugerencias)
                self.assertEqual(g.pruebas[0]["metodo"], "Texto libre " + accion)

    def test_tab_aplica_primera_sugerencia_o_la_resaltada_sin_cerrar(self):
        g = self.generador
        entrada = self.abrir()
        self.escribir(entrada, "ASTM")
        entrada.event_generate("<Tab>")
        self.raiz.update()
        self.assertEqual(entrada.get(), "ASTM D445")
        self.assertIs(g._editor_celda, entrada)
        self.escribir(entrada, "ASTM")
        entrada.event_generate("<Down>")
        entrada.event_generate("<Down>")
        entrada.event_generate("<Tab>")
        self.raiz.update()
        self.assertEqual(entrada.get(), "ASTM D446")
        self.assertIs(g._editor_celda, entrada)
        entrada.event_generate("<Tab>")
        self.raiz.update()
        self.assertIsNone(g._editor_celda)
        self.assertEqual(g.pruebas[0]["metodo"], "ASTM D446")

    def test_tab_resuelve_termino_pendiente_y_conserva_el_resto(self):
        from traductor_especificaciones import DiccionarioTraductor
        g = self.generador
        g._traductor = DiccionarioTraductor([
            {"es": "MÁXIMO", "en": "MAXIMUM", "activo": True},
        ])
        entrada = self.abrir("especificacion")
        entrada.insert(0, "30 MAX; mm²/s ¹")
        entrada.icursor(len("30 MAX"))
        entrada.event_generate("<Tab>")
        self.raiz.update()
        self.assertEqual(entrada.get(), "30 MÁXIMO; mm²/s ¹")
        self.assertIs(g._editor_celda, entrada)

    def test_tab_sin_coincidencias_guarda_y_sale(self):
        g = self.generador
        g._sugerencias.return_value = ()
        entrada = self.abrir()
        self.escribir(entrada, "Texto libre")
        entrada.event_generate("<Tab>")
        self.raiz.update()
        self.assertIsNone(g._editor_celda)
        self.assertEqual(g.pruebas[0]["metodo"], "Texto libre")

    def test_doble_clic_en_marca_alineada_edita_y_quita_solo_su_azul(self):
        g = self.generador
        g._bilingue = None
        g._idioma_anterior = "es"
        g._celda_activa = None
        g._marcas_pendientes = []
        g._cambios_alineacion = {"es": {"0": {"especificacion", "unidades"}}}
        g._programar_resaltado = Mock()
        g.boton_confirmar_traduccion = Mock()
        g.boton_confirmar_notas = Mock()
        g.referencia_notas = Mock()
        g.pruebas[0].update(especificacion="30", unidades="cSt")
        g.tabla_pruebas.item("0", values=tuple(g.pruebas[0][c] for c in (*modulo.CAMPOS_TABLA, "periodicidad")))
        g._resaltar_pendientes()
        self.raiz.update()
        marca = next(m for m in g._marcas_pendientes if m.cget("text") == "30")
        marca.event_generate("<ButtonPress-1>", x=4, y=4, time=1000)
        marca.event_generate("<ButtonRelease-1>", x=4, y=4, time=1010)
        self.assertIsNone(g._editor_celda)
        marca.event_generate("<ButtonPress-1>", x=4, y=4, time=1100)
        marca.event_generate("<ButtonRelease-1>", x=4, y=4, time=1110)
        self.raiz.update()
        entrada = g._editor_celda
        self.assertIsNotNone(entrada)
        self.escribir(entrada, "40")
        g._cerrar_editor_celda(guardar=True)
        self.assertEqual(g.pruebas[0]["especificacion"], "40")
        self.assertEqual(g._cambios_alineacion["es"]["0"], {"unidades"})
        # Abrir y guardar sin cambios conserva el azul.
        g._abrir_celda_marcada("0", 2)
        g._cerrar_editor_celda(guardar=True)
        self.assertEqual(g._cambios_alineacion["es"]["0"], {"unidades"})

    def test_metodo_primero_completa_primera_prueba_y_unidad_del_idioma(self):
        g = self.generador
        g.maestro = [
            {"METODO": "ASTM D445", "NOMBRE IQM": "Viscosidad", "NOMBRE IQM INGLES": "Viscosity",
             "UNIDADES": "cSt", "UNITS": "mm²/s"},
            {"METODO": "ASTM D445", "NOMBRE IQM": "Otra prueba", "UNIDADES": "Otra unidad"},
        ]
        for idioma, nombre, unidad in (("Español", "Viscosidad", "cSt"), ("English", "Viscosity", "mm²/s")):
            with self.subTest(idioma=idioma):
                g._idioma.set(idioma)
                g.pruebas[0].update(prueba="", metodo="", unidades="")
                self.escribir(self.abrir(), "ASTM D445")
                g._cerrar_editor_celda(guardar=True)
                self.assertEqual(g.pruebas[0]["prueba"], nombre)
                self.assertEqual(g.pruebas[0]["unidades"], unidad)
        g.pruebas[0].update(prueba="Personalizada", unidades="Unidad propia")
        self.escribir(self.abrir(), "ASTM D445")
        g._cerrar_editor_celda(guardar=True)
        self.assertEqual(g.pruebas[0]["prueba"], "Personalizada")
        self.assertEqual(g.pruebas[0]["unidades"], "Unidad propia")
        g.pruebas[0].update(prueba="", unidades="")
        entrada = self.abrir()
        self.escribir(entrada, "ASTM")
        entrada.event_generate("<Tab>")
        self.raiz.update()
        self.assertIs(g._editor_celda, entrada)
        self.assertEqual(g.pruebas[0]["prueba"], "Viscosity")
        self.assertEqual(g.pruebas[0]["unidades"], "mm²/s")

    def test_cambiar_de_celda_con_clic_no_cierra_el_editor_nuevo(self):
        g = self.generador
        self.escribir(self.abrir(), "Método escrito")
        x, y, ancho, alto = g.tabla_pruebas.bbox("0", "#2")
        g.tabla_pruebas.event_generate("<Button-1>", x=x + ancho // 2, y=y + alto // 2)
        self.raiz.update()
        self.assertEqual(g._celda_editada, ("0", "unidades"))
        self.assertEqual(g.pruebas[0]["metodo"], "Método escrito")

    def test_comentario_conserva_el_texto_y_actualiza_las_notas(self):
        g = self.generador
        for campo, valor in (("metodo", "ASTM D445"), ("prueba", "Viscosidad"), ("especificacion", "30 MÁXIMO")):
            with self.subTest(campo=campo):
                entrada = self.abrir(campo)
                self.escribir(entrada, valor)
                self.assertIsNotNone(g._comentario_boton_mas)
                self.assertTrue(g._foco_en_editor(str(g._comentario_boton_mas)))
                g._comentario_boton_mas.invoke()
                self.assertIn(valor, entrada.get())
                self.assertNotEqual(entrada.get(), valor)
                self.assertTrue(g.notas.get("1.0", "end-1c").strip())
                texto = entrada.get()
                g._cerrar_editor_celda(guardar=True)
                self.assertEqual(g.pruebas[0][campo], texto)

    def test_borrar_especificacion_con_comentario_conserva_las_otras_notas(self):
        g = self.generador
        g.pruebas[0].update(metodo="Método ¹", especificacion="30 MÁXIMO ²")
        g.notas.insert("1.0", "Nota base\n¹ Conservar\n² Quitar")
        g._celda_activa = ("0", "especificacion")
        g._borrar_celda_activa()
        self.assertEqual(g.pruebas[0]["especificacion"], "")
        self.assertEqual(g.pruebas[0]["metodo"], "Método ¹")
        self.assertEqual(g.notas.get("1.0", "end-1c"), "Nota base\n¹ Conservar")

    def test_editor_nuevo_abre_maximizado(self):
        if self.raiz.tk.call("tk", "windowingsystem") != "win32":
            self.skipTest("El estado zoomed corresponde a Windows")
        with (
            patch.object(modulo, "cargar_maestro", return_value=[]),
            patch.object(modulo, "_cargar_catalogo_part_ids", return_value=[]),
            patch.object(modulo.GeneradorEspecificacion, "_actualizar_traductor_editor"),
            patch.object(modulo.GeneradorEspecificacion, "_iniciar_bilingue"),
            patch.object(modulo.GeneradorEspecificacion, "_vigilar_revision"),
        ):
            g = modulo.GeneradorEspecificacion(SimpleNamespace(raiz=self.raiz))
            try:
                self.raiz.update()
                # CTk vuelve a dibujar la barra de título a los 200 ms.
                self.raiz.after(300, self.raiz.quit)
                self.raiz.mainloop()
                self.assertEqual(g.ventana.state(), "zoomed")
                g.notas.insert("end", "\n" + "\n".join(f"Nota {i}" for i in range(40)))
                self.raiz.update_idletasks()
                self.assertEqual(str(g.barra_notas.cget("orient")), "vertical")
                g.barra_notas.tk.call(g.barra_notas.cget("command"), "moveto", 1)
                self.assertGreater(g.notas.yview()[0], 0)
                g.variables_control["codigo_unico_producto"].set("P1")
                g._revision_calculada = ("P1", "08")
                g.variables_control["revision"].set("08")
                self.assertEqual(str(g.entrada_revision.cget("state")), "readonly")
                self.assertEqual(g.entrada_revision.cget("style"), "RevisionCalculada.TEntry")
                g.entrada_revision.insert(0, "9")
                self.assertEqual(g.variables_control["revision"].get(), "08")
                g.entrada_revision.event_generate("<ButtonPress-1>", x=3, y=3, time=2000)
                g.entrada_revision.event_generate("<ButtonRelease-1>", x=3, y=3, time=2010)
                g.entrada_revision.event_generate("<ButtonPress-1>", x=3, y=3, time=2100)
                g.entrada_revision.event_generate("<ButtonRelease-1>", x=3, y=3, time=2110)
                self.assertEqual(str(g.entrada_revision.cget("state")), "normal")
                g.variables_control["revision"].set("09")
                self.assertEqual(g.entrada_revision.cget("style"), "RevisionManual.TEntry")
                g.variables_control["revision"].set("08")
                self.assertEqual(g.entrada_revision.cget("style"), "RevisionCalculada.TEntry")
            finally:
                g.ventana.destroy()

    def test_destruir_con_actualizacion_pendiente_cancela_callbacks(self):
        g = self.generador
        entrada = self.abrir()
        entrada.insert(0, "ASTM")
        g._programar_cierre_editor()
        self.assertIsNotNone(g._sugerencias_editor_pendientes)
        self.assertIsNotNone(g._cierre_editor_pendiente)
        g.ventana.destroy()
        self.raiz.update()
        self.assertIsNone(g._sugerencias_editor_pendientes)
        self.assertIsNone(g._cierre_editor_pendiente)

    def test_autocompletar_termino_conserva_numeros_unidades_y_comentario(self):
        from traductor_especificaciones import DiccionarioTraductor
        g = self.generador
        g._traductor = DiccionarioTraductor([
            {"es": "MÁXIMO", "en": "MAXIMUM", "activo": True},
        ])
        for campo in ("especificacion", "valor_tipico"):
            with self.subTest(campo=campo):
                entrada = self.abrir(campo)
                self.escribir(entrada, "30 MAX; 10 mm²/s ¹")
                entrada.icursor(len("30 MAX"))
                self.procesar_cierres()
                self.assertEqual(g._lista_sugerencias.get(0), "MÁXIMO")
                entrada.event_generate("<Down>")
                entrada.event_generate("<Return>")
                self.raiz.update()
                self.assertEqual(entrada.get(), "30 MÁXIMO; 10 mm²/s ¹")
                self.assertEqual(entrada.index(tk.INSERT), len("30 MÁXIMO"))
                self.assertIs(g._editor_celda, entrada)
                g._cerrar_editor_celda(guardar=False)

    def test_terminos_escape_y_escritura_rapida_cancelan_lista_pendiente(self):
        g = self.generador
        entrada = self.abrir("valor_tipico")
        self.escribir(entrada, "MAX")
        entrada.event_generate("<Escape>")
        self.procesar_cierres()
        self.assertIsNone(g._popup_sugerencias)
        self.escribir(entrada, "MIN")
        self.escribir(entrada, "30")
        self.procesar_cierres()
        self.assertIsNone(g._popup_sugerencias)
        self.assertEqual(entrada.get(), "30")


if __name__ == "__main__":
    unittest.main()
