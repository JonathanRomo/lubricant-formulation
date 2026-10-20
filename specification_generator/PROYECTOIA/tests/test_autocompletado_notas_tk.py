"""Sugerencias de notas con Text real, protección de formato y navegación."""
import tkinter as tk
import unittest
from unittest.mock import patch

from editor_comportamientos import NotasProtegidas
from interfaz_aplicacion_especificaciones import GeneradorEspecificacion, NOTAS_PLANTILLA
from traductor_especificaciones import DiccionarioTraductor


class AutocompletadoNotasTkTests(unittest.TestCase):
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
        g = self.editor = GeneradorEspecificacion.__new__(GeneradorEspecificacion)
        g.ventana = tk.Toplevel(self.raiz)
        g.ventana.attributes("-alpha", 0)
        g.ventana.geometry("650x240+0+0")
        for nombre in (
            "_editor_celda", "_celda_editada", "_cierre_editor_pendiente", "_texto_editor",
            "_traza_texto_editor", "_sugerencias_editor_pendientes", "_popup_sugerencias",
            "_lista_sugerencias", "_entrada_sugerencias", "_al_seleccionar_sugerencia",
        ):
            setattr(g, nombre, None)
        g._idioma = tk.StringVar(master=g.ventana, value="Español")
        g._traductor = DiccionarioTraductor([
            {"es": "MÁXIMO", "en": "MAXIMUM", "activo": True},
            {"es": "MÍNIMO", "en": "MINIMUM", "activo": True},
            {"es": "ÁMBAR CLARO", "en": "LIGHT AMBER", "activo": True},
            {"es": "AMBAR CLARA", "en": "LIGHT AMBER", "activo": True},
            {"es": "LA VERSIÓN", "en": "THE VERSION", "activo": True},
            {"es": "CRÍTICO", "en": "CRITICAL", "activo": False},
        ])
        g.notas = NotasProtegidas(g.ventana, height=4, wrap="word")
        g.notas.pack(fill="both", expand=True)
        self.otro = tk.Entry(g.ventana)
        self.otro.pack()
        g._iniciar_autocompletado_notas()
        g.ventana.bind("<Destroy>", g._al_destruir_ventana_editor, add="+")
        self.raiz.update()
        g.notas.focus_force()
        self.raiz.update()
        constructor = tk.Toplevel.__init__

        def invisible(ventana, *args, **kwargs):
            constructor(ventana, *args, **kwargs)
            ventana.attributes("-alpha", 0)

        parche = patch.object(tk.Toplevel, "__init__", invisible)
        parche.start()
        self.addCleanup(parche.stop)

    def tearDown(self):
        if self.editor.ventana.winfo_exists():
            self.editor.ventana.destroy()
        self.raiz.update()
        self.assertFalse(self.errores, self.errores)

    def escribir(self, texto, cursor="end-1c"):
        notas = self.editor.notas
        notas.delete("1.0", "end")
        notas.insert("1.0", texto)
        notas.mark_set("insert", cursor)
        notas.see("insert")
        self.raiz.update()

    def esperar_sugerencias(self):
        self.raiz.after(160, self.raiz.quit)
        self.raiz.mainloop()

    def test_pegar_reemplaza_seleccion_con_texto_multilinea(self):
        self.escribir("Nota anterior\nConservar")
        notas = self.editor.notas
        notas.tag_add("sel", "1.0", "1.end")
        notas.mark_set("insert", "1.0")
        notas.clipboard_clear()
        notas.clipboard_append("Nueva nota\nSegunda nota")
        notas.event_generate("<<Paste>>")
        self.raiz.update()
        self.assertEqual(notas.get("1.0", "end-1c"), "Nueva nota\nSegunda nota\nConservar")

    def test_control_c_copia_seleccion_sin_modificar_notas(self):
        original = "Ámbar ≤ 30 °C\nViscosidad mm²/s"
        for evento in ("<Control-c>", "<Control-C>", "<<Copy>>"):
            with self.subTest(evento=evento):
                self.escribir(original)
                notas = self.editor.notas
                notas.tag_add("sel", "1.0", "end-1c")
                notas.event_generate(evento)
                self.raiz.update()
                self.assertEqual(notas.clipboard_get(), original)
                self.assertEqual(notas.get("1.0", "end-1c"), original)

    def test_control_c_sin_seleccion_conserva_portapapeles(self):
        self.escribir("Sin selección")
        notas = self.editor.notas
        notas.clipboard_clear()
        notas.clipboard_append("Conservar")
        notas.event_generate("<Control-c>")
        self.raiz.update()
        self.assertEqual(notas.clipboard_get(), "Conservar")

    def test_control_c_permite_copiar_nota_protegida(self):
        original = "- La versión vigente del método aplicable corresponderá al laboratorio."
        self.escribir(original)
        notas = self.editor.notas
        notas.tag_add("sel", "1.0", "end-1c")
        notas.event_generate("<Control-c>")
        self.raiz.update()
        self.assertEqual(notas.clipboard_get(), original)
        self.assertEqual(notas.get("1.0", "end-1c"), original)

    def test_pegar_normaliza_saltos_y_caracteres_especiales(self):
        self.escribir("")
        notas = self.editor.notas
        notas.clipboard_clear()
        notas.clipboard_append("Ámbar ≤ 30 °C\r\nViscosidad mm²/s\rOtra nota")
        notas.event_generate("<<Paste>>")
        self.raiz.update()
        self.assertEqual(notas.get("1.0", "end-1c"), "Ámbar ≤ 30 °C\nViscosidad mm²/s\nOtra nota")

    def test_pegar_sobre_nota_protegida_conserva_toda_la_seleccion(self):
        original = "Nota libre\n- La versión vigente del método aplicable corresponderá al laboratorio."
        self.escribir(original)
        notas = self.editor.notas
        notas.tag_add("sel", "1.0", "end-1c")
        notas.clipboard_clear()
        notas.clipboard_append("Reemplazo")
        notas.event_generate("<<Paste>>")
        self.raiz.update()
        self.assertEqual(notas.get("1.0", "end-1c"), original)

    def test_completa_palabra_en_medio_y_conserva_otras_notas_y_comentarios(self):
        g = self.editor
        self.escribir("¹ Conservar esta nota.\n² 30 MAX; 10 mm²/s\n³ Conservar", "2.8")
        self.esperar_sugerencias()
        self.assertEqual(g._lista_sugerencias.get(0), "MÁXIMO")
        self.assertIs(self.raiz.focus_get(), g.notas)
        g.notas.event_generate("<Down>")
        g.notas.event_generate("<Return>")
        self.raiz.update()
        self.assertEqual(g.notas.get("1.0", "end-1c"),
                         "¹ Conservar esta nota.\n² 30 MÁXIMO; 10 mm²/s\n³ Conservar")
        self.assertEqual(g.notas.index("insert"), "2.11")
        self.assertIsNone(g._popup_sugerencias)
        self.esperar_sugerencias()
        self.assertIsNone(g._popup_sugerencias)

    def test_clic_completa_frase_sin_quitar_el_foco(self):
        g = self.editor
        self.escribir("Color: AMBAR C; conservar", "1.14")
        self.esperar_sugerencias()
        lista = g._lista_sugerencias
        self.assertEqual(lista.get(0, "end"), ("ÁMBAR CLARO", "AMBAR CLARA"))
        x, y, ancho, alto = lista.bbox(1)
        lista.event_generate("<ButtonPress-1>", x=x + 3, y=y + alto // 2)
        lista.event_generate("<ButtonRelease-1>", x=x + 3, y=y + alto // 2)
        self.raiz.update()
        self.assertEqual(g.notas.get("1.0", "end-1c"), "Color: AMBAR CLARA; conservar")
        self.assertIs(self.raiz.focus_get(), g.notas)

    def test_tab_resuelve_filtrado_pendiente_y_respeta_idioma_y_activos(self):
        g = self.editor
        for idioma, prefijo, esperado in (("Español", "MAX", "MÁXIMO"), ("English", "MAX", "MAXIMUM")):
            with self.subTest(idioma=idioma):
                g._idioma.set(idioma)
                self.escribir(prefijo)
                self.assertIsNotNone(g._sugerencias_notas_pendientes)
                g.notas.event_generate("<Tab>")
                self.raiz.update()
                self.assertEqual(g.notas.get("1.0", "end-1c"), esperado)
        self.escribir("CRI")
        self.esperar_sugerencias()
        self.assertIsNone(g._popup_sugerencias)

    def test_enter_sin_seleccion_conserva_salto_de_linea_y_flechas_normales(self):
        g = self.editor
        self.escribir("MAX")
        self.esperar_sugerencias()
        self.assertIsNotNone(g._popup_sugerencias)
        g.notas.event_generate("<Return>")
        self.raiz.update()
        self.assertEqual(g.notas.get("1.0", "end-1c"), "MAX\n")
        self.assertIsNone(g._popup_sugerencias)
        g.notas.event_generate("<Up>")
        self.assertEqual(g.notas.index("insert"), "1.0")

    def test_escape_y_cambio_de_foco_cancelan_lista_pendiente(self):
        g = self.editor
        self.escribir("MAX")
        g.notas.event_generate("<Escape>")
        self.esperar_sugerencias()
        self.assertIsNone(g._popup_sugerencias)
        self.assertEqual(g.notas.get("1.0", "end-1c"), "MAX")
        self.escribir("MIN")
        self.otro.focus_force()
        self.raiz.update()
        self.esperar_sugerencias()
        self.assertIsNone(g._popup_sugerencias)
        self.assertIsNone(g._sugerencias_notas_pendientes)

    def test_cursor_movido_o_texto_cambiado_no_acepta_sugerencia_anterior(self):
        g = self.editor
        self.escribir("MAX; conservar", "1.3")
        self.esperar_sugerencias()
        callback = g._al_seleccionar_sugerencia
        g.notas.mark_set("insert", "end-1c")
        callback("MÁXIMO")
        self.assertEqual(g.notas.get("1.0", "end-1c"), "MAX; conservar")
        g.notas.mark_set("insert", "1.3")
        g._filtrar_sugerencias_notas()
        callback = g._al_seleccionar_sugerencia
        self.escribir("MIN; conservar", "1.3")
        callback("MÁXIMO")
        self.assertEqual(g.notas.get("1.0", "end-1c"), "MIN; conservar")

    def test_respeta_proteccion_de_nota_del_formato_y_desbloqueo_explicito(self):
        g = self.editor
        original = NOTAS_PLANTILLA["es"]
        self.escribir(original, "1.6")
        self.esperar_sugerencias()
        g.notas.event_generate("<Control-space>")
        g.notas.event_generate("<Tab>")
        self.raiz.update()
        self.assertIsNone(g._popup_sugerencias)
        self.assertEqual(g.notas.get("1.0", "end-1c"), original)
        g.notas._desbloquear()
        g.notas.mark_set("insert", "1.6")
        g.notas.event_generate("<Control-space>")
        g.notas.event_generate("<Tab>")
        self.raiz.update()
        self.assertEqual(g.notas.get("1.0", "end-1c"), "LA VERSIÓN" + original[len("La versión"):])

    def test_no_une_frases_de_dos_notas_y_consulta_diccionario_actual(self):
        g = self.editor
        self.escribir("AMBAR\nCL")
        self.esperar_sugerencias()
        self.assertIsNone(g._popup_sugerencias)
        self.escribir("MAX")
        self.esperar_sugerencias()
        callback = g._al_seleccionar_sugerencia
        g._traductor = DiccionarioTraductor([{"es": "MAXIMIZAR", "en": "MAXIMIZE", "activo": True}])
        callback("MÁXIMO")
        self.assertEqual(g.notas.get("1.0", "end-1c"), "MAX")
        g.notas.event_generate("<Control-space>")
        g.notas.event_generate("<Tab>")
        self.raiz.update()
        self.assertEqual(g.notas.get("1.0", "end-1c"), "MAXIMIZAR")

    def test_destruir_ventana_cancela_actualizaciones_pendientes(self):
        g = self.editor
        self.escribir("MAX")
        self.assertIsNotNone(g._sugerencias_notas_pendientes)
        g.ventana.destroy()
        self.raiz.update()
        self.assertIsNone(g._sugerencias_notas_pendientes)


if __name__ == "__main__":
    unittest.main()
