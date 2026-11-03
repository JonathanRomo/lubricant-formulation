"""Comprueba los controles reales de Tk, incluidas sus opciones nativas."""
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import interfaz_seleccion_especificaciones as modulo
from test_interfaz_seleccion_especificaciones import json_especificacion


class VistaCompactaTkTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = modulo.ui.Raiz()
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()

    def test_selector_compacta_restaura_y_conserva_la_tabla(self):
        root = self.root
        errores = []
        root.report_callback_exception = lambda *error: errores.append(error)
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            datos = json_especificacion()
            datos["alineacion_metodos_lab"] = {
                "detalle": [{"valores_alineados": datos["pruebas"][0]}]
            }
            (base / "producto.json").write_text(json.dumps(datos), encoding="utf-8")
            selector = modulo.SelectorEspecificaciones(root, carpeta_json=base, carpeta_json_anterior=base)
            root.update_idletasks()
            self.assertEqual(set(selector.paneles), {"anteriores", "nuevas"})
            for clave in selector.paneles:
                panel = selector.paneles[clave]
                selector._mostrar_previa(clave, panel.catalogo.especificaciones[0])
                root.update_idletasks()
                divisor = panel.separador_previa
                self.assertIsNotNone(divisor, (clave, [w.cget("text") for w in panel.marco_previa.winfo_children() if "text" in w.keys()]))
                self.assertFalse(divisor.cget("showhandle"))
                self.assertEqual(str(divisor.cget("sashrelief")), "flat")
                self.assertFalse(divisor.cget("opaqueresize"))
                self.assertEqual(panel.boton_compactar.cget("text"), "\u00d8")
                self.assertEqual(panel.boton_compactar.cget("width"), panel.boton_compactar.cget("height"))
                tabla = panel.tabla_pruebas
                seleccionadas = set(panel.documentos_marcados)
                selector._alternar_previa_compacta(clave)
                root.update_idletasks()
                posicion = panel.posicion_detalles
                self.assertTrue(divisor.panecget(panel.area_detalles, "hide"))
                selector._alternar_previa_compacta(clave)
                root.update_idletasks()
                self.assertFalse(divisor.panecget(panel.area_detalles, "hide"))
                self.assertEqual(divisor.sash_coord(0), posicion)
                self.assertIs(panel.tabla_pruebas, tabla)
                self.assertEqual(panel.documentos_marcados, seleccionadas)
                selector._alternar_previa_compacta(clave)
                selector._cambiar_zoom_previa(clave, 0.1)
                root.update_idletasks()
                self.assertTrue(panel.separador_previa.panecget(panel.area_detalles, "hide"))
                selector._mostrar_previa(clave, panel.catalogo.especificaciones[0])
                root.update_idletasks()
                self.assertTrue(panel.separador_previa.panecget(panel.area_detalles, "hide"))
                # A pending restore must not touch the next preview's widgets.
                selector._alternar_previa_compacta(clave)
                selector._cerrar_vista_previa(clave)
                selector._alternar_previa_compacta(clave)
                root.update_idletasks()
                self.assertIsNone(panel.separador_previa)
                self.assertEqual(panel.boton_compactar.cget("state"), "disabled")
            self.assertFalse(errores, errores)

    def test_rayitas_arrastran_solo_al_soltar(self):
        root = self.root
        ventana = modulo.tk.Toplevel(root)
        self.addCleanup(ventana.destroy)
        ventana.withdraw()
        ventana.attributes("-alpha", 0)
        ventana.geometry("600x400+0+0")
        divisor = modulo.crear_divisor_redimensionable(
            ventana, modulo.tk.HORIZONTAL, mostrar_tirador=True)
        divisor.pack(fill="both", expand=True)
        izquierda = modulo.tk.Frame(divisor, width=250)
        derecha = modulo.tk.Frame(divisor, width=350)
        divisor.add(izquierda, minsize=100, stretch="always")
        divisor.add(derecha, minsize=100, stretch="always")
        ventana.deiconify()
        root.update()
        root._estilizar_hijos(ventana)
        root.update_idletasks()
        agarre = divisor._tirador_agarre
        self.assertTrue(agarre.winfo_ismapped(), (ventana.state(), divisor.winfo_width(), divisor.panes(), agarre.place_info()))
        self.assertEqual(len(agarre.find_all()), 3)
        self.assertEqual(str(divisor.cget("sashrelief")), "flat")
        self.assertFalse(divisor.cget("showhandle"))
        self.assertFalse(divisor.cget("opaqueresize"))
        self.assertEqual(agarre.cget("background"), divisor.cget("background"))
        antes = divisor.sash_coord(0)
        ancho = izquierda.winfo_width()
        agarre.event_generate("<ButtonPress-1>", x=4, y=13)
        agarre.event_generate("<B1-Motion>", x=64, y=13, state=256)
        root.update_idletasks()
        self.assertEqual(divisor.sash_coord(0), antes)
        self.assertEqual(izquierda.winfo_width(), ancho)
        self.assertEqual(divisor.proxy_coord()[0], antes[0] + 60)
        agarre.event_generate("<ButtonRelease-1>", x=64, y=13, state=256)
        root.update_idletasks()
        self.assertEqual(divisor.sash_coord(0)[0], antes[0] + 60)
        self.assertEqual(agarre.winfo_x(), divisor.sash_coord(0)[0])

    def test_reutiliza_widgets_cache_e_invalida_archivos_editados(self):
        root = self.root
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            fuente, alineados = base / "fuente", base / "alineados"
            fuente.mkdir()
            alineados.mkdir()
            for nombre in ("A", "B"):
                datos = json_especificacion(nombre=f"Producto {nombre}", codigo=f"P{nombre}",
                                             idioma="en" if nombre == "B" else "es")
                if nombre == "B":
                    datos["notas"] = []
                    datos["pruebas"] = [dict(datos["pruebas"][0]) for _ in range(3)]
                (fuente / f"{nombre}.json").write_text(json.dumps(datos), encoding="utf-8")
                datos["alineacion_metodos_lab"] = {"detalle": [{"valores_alineados": p} for p in datos["pruebas"] if "metodo" in p]}
                (alineados / f"{nombre}.json").write_text(json.dumps(datos), encoding="utf-8")
            selector = modulo.SelectorEspecificaciones(root, carpeta_json=alineados, carpeta_json_anterior=fuente)
            panel = selector.paneles["nuevas"]
            a, b = panel.catalogo.especificaciones
            with patch.object(modulo, "leer_datos_json", wraps=modulo.leer_datos_json) as leer:
                selector._mostrar_previa("nuevas", a)
                root.update_idletasks()
                tabla, separador = panel.tabla_pruebas, panel.separador_previa
                titulo = panel.widgets_previa["titulo"]
                detalle = panel.widgets_previa["pagina"]._widgets[("valor", 0)]
                lecturas = leer.call_count
                selector._mostrar_previa("nuevas", a)
                self.assertEqual(leer.call_count, lecturas)
                selector._mostrar_previa("nuevas", b)
                self.assertIn("Producto B", titulo.cget("text"))
                self.assertEqual(len(tabla.get_children()), 3)
                self.assertEqual(panel.widgets_previa["pagina"]._widgets[("nota", 0)].winfo_manager(), "")
                self.assertIn("Test", tabla.heading("prueba", "text"))
                lecturas = leer.call_count
                selector._mostrar_previa("nuevas", a)
                self.assertEqual(leer.call_count, lecturas)
                self.assertIn("Producto A", titulo.cget("text"))
                self.assertEqual(len(tabla.get_children()), 1)
                self.assertEqual(panel.widgets_previa["pagina"]._widgets["idioma"].winfo_manager(), "")
                self.assertIn("\u25bc", tabla.heading("prueba", "text"))
                datos = json.loads(a.ruta.read_text(encoding="utf-8"))
                datos["producto"]["nombre"] = "Producto A modificado externamente"
                a.ruta.write_text(json.dumps(datos), encoding="utf-8")
                selector._mostrar_previa("nuevas", a)
                self.assertGreater(leer.call_count, lecturas)
                self.assertIn("modificado externamente", titulo.cget("text"))
                lecturas = leer.call_count
                original = fuente / a.ruta.name
                original.write_text(original.read_text(encoding="utf-8") + " ", encoding="utf-8")
                selector._mostrar_previa("nuevas", a)
                self.assertGreater(leer.call_count, lecturas)
                selector._cambiar_zoom_previa("nuevas", 0.1)
                root.update_idletasks()
                self.assertIs(panel.tabla_pruebas, tabla)
                self.assertIs(panel.separador_previa, separador)
                self.assertIs(panel.widgets_previa["pagina"]._widgets[("valor", 0)], detalle)
                self.assertEqual(str(tabla.cget("cursor")), "arrow")
                self.assertEqual(str(separador.cget("cursor")), "arrow")
                selector._actualizar_sugerencias("nuevas")
                panel.lista.selection_set(0)
                filas = panel.lista.size()
                vista = panel.firma_vista
                with patch.object(selector, "_actualizar_sugerencias") as reconstruir:
                    selector._clic_en_lista("nuevas", type("Evento", (), {"x": 0, "y": 1})())
                    reconstruir.assert_not_called()
                self.assertEqual(panel.lista.size(), filas)
                self.assertTrue(panel.lista.get(0).startswith("\u2611"))
                self.assertEqual(panel.firma_vista, vista)
                self.assertIs(panel.tabla_pruebas, tabla)
                contenido = a.ruta.read_text(encoding="utf-8")
                for i in range(15):
                    a.ruta.write_text(contenido + " " * i, encoding="utf-8")
                    selector._mostrar_previa("nuevas", a)
                self.assertEqual(len(panel.cache_previa), 12)
                a.ruta.unlink()
                selector._mostrar_previa("nuevas", a)
                self.assertIsNone(panel.separador_previa)
                self.assertIsNone(panel.firma_vista)
                selector._cerrar_vista_previa("nuevas")
            for hijo in list(root.winfo_children()):
                hijo.destroy()

    def test_encabezados_no_recalculan_si_solo_cambia_la_altura(self):
        tabla = modulo.ttk.Treeview(self.root, columns=("prueba",), show="headings")
        self.addCleanup(tabla.destroy)
        ordenador = modulo.OrdenadorColumnasTreeview(tabla, {"prueba": "Prueba"})
        with patch.object(modulo.tkfont, "Font", wraps=modulo.tkfont.Font) as fuente:
            for _ in range(10):
                ordenador._al_cambiar_tamano()
            fuente.assert_not_called()
            tabla.column("prueba", width=450)
            ordenador._al_cambiar_tamano()
            self.assertEqual(fuente.call_count, 1)
