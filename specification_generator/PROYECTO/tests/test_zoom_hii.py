import unittest
from unittest.mock import Mock, patch

from generador_hii import aplicar_zoom_canvas, fuente_escalada, ZoomHII


class Canvas:
    def __init__(self):
        self._hii_region = (0, 0, 600, 1000)
        self._hii_imagenes_base = [(2, Mock())]
        self.coordenadas = [20, 30]
        self.opciones = {}
        self.region = None

    def find_all(self): return (1, 2)
    def type(self, item): return "text" if item == 1 else "image"
    def itemcget(self, item, campo): return "Arial 10" if campo == "font" else "100"
    def xview(self): return (.2, .5)
    def yview(self): return (.3, .5)
    def xview_moveto(self, valor): self.x = valor
    def yview_moveto(self, valor): self.y = valor
    def configure(self, *, scrollregion): self.region = scrollregion
    def itemconfigure(self, item, **valores): self.opciones.setdefault(item, {}).update(valores)
    def scale(self, tag, x, y, sx, sy): self.coordenadas = [self.coordenadas[0] * sx, self.coordenadas[1] * sy]


class ZoomTests(unittest.TestCase):
    def test_fuentes_conservan_unidades_y_vuelven_al_tamano_original(self):
        self.assertEqual(fuente_escalada(("Arial", -12, "bold"), 1.5), ("Arial", -18, "bold"))
        self.assertEqual(fuente_escalada(("Arial", 10, "normal"), .6), ("Arial", 6, "normal"))

    def test_canvas_escala_textos_imagenes_y_region_sin_perder_desplazamiento(self):
        canvas = Canvas()
        with patch("generador_hii._fuente", return_value=("Arial", 10, "bold")):
            aplicar_zoom_canvas(canvas, 1.5)
            self.assertEqual(canvas.coordenadas, [30, 45])
            self.assertEqual(canvas.opciones[1], {"font": ("Arial", 15, "bold"), "width": 150})
            self.assertEqual(canvas.region, (0, 0, 900, 1500))
            canvas._hii_imagenes_base[0][1].zoom.assert_called_once_with(3)
            self.assertEqual((canvas.x, canvas.y), (.2, .3))
            aplicar_zoom_canvas(canvas, 1)
            self.assertEqual(canvas.coordenadas, [20, 30])
            self.assertEqual(canvas.opciones[1]["font"], ("Arial", 10, "bold"))
            self.assertIs(canvas.opciones[2]["image"], canvas._hii_imagenes_base[0][1])
            self.assertEqual(canvas.region, canvas._hii_region)

    def test_solo_recorre_el_contenido_indicado_y_no_escala_internos_ctk(self):
        class Contenedor:
            def __init__(self, hijos=()): self.hijos = hijos
            def winfo_children(self): return self.hijos
        class Boton(Contenedor):
            _ancho_automatico = True
            def __init__(self):
                self.opciones = {"font": ("Arial", 12), "width": 110, "height": 30, "text": "Guardar"}
            def cget(self, clave): return self.opciones[clave]
            def configure(self, **opciones): self.opciones.update(opciones)
            def winfo_children(self): raise AssertionError("No escalar los internos CTk")
        boton, cabecera = Boton(), Boton()
        with patch("generador_hii.ui.Boton", Boton):
            zoom = ZoomHII(Contenedor((boton,)))
            zoom.aplicar(1.5)
            self.assertEqual(boton.opciones["height"], 45)
            self.assertEqual(boton.opciones["font"], ("Arial", 18))
            self.assertEqual(cabecera.opciones["font"], ("Arial", 12))
            zoom.aplicar(1)
            self.assertEqual(boton.opciones["font"], ("Arial", 12))
            self.assertEqual(boton.opciones["height"], 30)

    def test_tablas_usan_estilos_locales_y_escalan_columnas_y_filas(self):
        class Tabla:
            def __init__(self): self.estilo, self.medidas = "", {"width": 100, "minwidth": 20}
            def __getitem__(self, clave): return ("prueba",)
            def cget(self, clave): return self.estilo
            def configure(self, *, style): self.estilo = style
            def column(self, nombre, opcion=None, **valores):
                if opcion: return self.medidas[opcion]
                self.medidas.update(valores)
            def winfo_children(self): return ()
        tabla = Tabla()
        with patch("generador_hii.ttk_nativo.Treeview", Tabla), patch("generador_hii.ttk_nativo.Style") as estilo, \
             patch("generador_hii._fuente", return_value=("Arial", 10, "normal")):
            estilo.return_value.lookup.side_effect = lambda nombre, campo: 24 if campo == "rowheight" else "TkDefaultFont"
            zoom = ZoomHII(tabla)
            zoom.aplicar(1.5)
            self.assertEqual(tabla.medidas, {"width": 150, "minwidth": 30})
            self.assertTrue(tabla.estilo.startswith("HII"))
            self.assertTrue(all(c.args[0].startswith("HII") for c in estilo.return_value.configure.call_args_list))
            zoom.aplicar(1)
            self.assertEqual(tabla.medidas, {"width": 100, "minwidth": 20})


if __name__ == "__main__":
    unittest.main()
