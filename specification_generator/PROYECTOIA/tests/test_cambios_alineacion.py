import copy
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from cambios_alineacion import cambios_documento, diferencias_filas, original_disponible, ResaltadorCambios, CambiosFila


class CambiosAlineacionTests(unittest.TestCase):
    def test_metodo_no_encontrado_se_pinta_rojo_aunque_no_haya_cambios(self):
        tabla = Mock()
        tabla.winfo_exists.return_value = True
        tabla.cget.return_value = "Treeview"
        tabla.selection.return_value = ()
        tabla.get_children.return_value = ["f1"]
        tabla.yview.return_value = (0.0, 1.0)
        tabla.item.return_value = ("NO PATOGENOS B", "UFC", "NOM-092-SSA1 1994", "< 1000", "")
        tabla.bbox.return_value = (5, 25, 100, 25)
        tabla.winfo_width.return_value = 800
        with patch("cambios_alineacion.tk.Label") as etiqueta, patch("cambios_alineacion.ttk.Style") as estilo:
            estilo.return_value.lookup.side_effect = lambda _, campo, *a: {"font": ("Segoe UI", 12), "fieldbackground": "white"}.get(campo, "")
            resaltador = ResaltadorCambios(tabla, {"f1": CambiosFila((), "Método no encontrado en lista maestra")}, Mock(), Mock())
            resaltador.dibujar()
            self.assertEqual(etiqueta.call_count, 5)
            colores = [llamada.kwargs["foreground"] for llamada in etiqueta.return_value.configure.call_args_list]
            self.assertTrue(all(c in {"#c62828", "#ff6b6b"} for c in colores))
            tabla.selection.return_value = ("f1",)
            resaltador.dibujar()
            self.assertEqual(etiqueta.return_value.configure.call_args.kwargs["foreground"], "#ffb3b3")

    def test_auditoria_existente_sin_mutar_json(self):
        original = {"prueba": "Densidad anterior", "unidades": "g/ml", "metodo": "ASTM D 1298", "especificacion": "0.8", "valor_tipico": "0.85"}
        alineada = {**original, "prueba": "DENSIDAD", "metodo": "ASTM D1298"}
        datos = {"pruebas": [alineada], "alineacion_metodos_lab": {"detalle": [{
            "nombre_original": original["prueba"], "metodo_original": original["metodo"],
            "unidades_originales": original["unidades"], "valores_alineados": alineada}]}}
        antes = copy.deepcopy(datos)
        self.assertEqual(cambios_documento(datos, [alineada]), [{"prueba", "metodo"}])
        self.assertEqual(datos, antes)

    def test_original_externo_detecta_formatos_y_no_confunde_reordenamiento(self):
        fuente = {"prueba": "PRUEBA A", "valor_tipico": "1.000", "especificacion": "150 MAXIMO"}
        alineada = {**fuente, "valor_tipico": "1,000", "especificacion": "150 MÁXIMO"}
        segunda = {"prueba": "PRUEBA B"}
        datos = {"pruebas": [alineada, segunda], "alineacion_metodos_lab": {"detalle": [{"valores_alineados": alineada}, {"valores_alineados": segunda}]}}
        original = {"pruebas": [fuente, segunda]}
        self.assertEqual(cambios_documento(datos, [segunda, alineada], original), [set(), {"valor_tipico", "especificacion"}])
        self.assertEqual(diferencias_filas([{"metodo": "A"}], [{"metodo": "B", "unidades": "m"}]), [{"metodo"}])

    def test_original_de_revision_o_idioma_distinto_no_se_usa(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "producto.json"
            datos = {"control_documental": {"codigo_unico_producto": "P1", "revision": "01", "nombre_cliente": "Cliente"}, "documento": {"idioma": "es"}, "pruebas": []}
            ruta.write_text(json.dumps(datos), encoding="utf-8")
            self.assertEqual(original_disponible(datos, ruta, carpeta), datos)
            otro = copy.deepcopy(datos)
            otro["control_documental"]["revision"] = "02"
            self.assertIsNone(original_disponible(otro, ruta, carpeta))
            otro = copy.deepcopy(datos)
            otro["documento"]["idioma"] = "en"
            self.assertIsNone(original_disponible(otro, ruta, carpeta))

    def test_celda_azul_del_selector_respeta_fondo_fuente_y_fila(self):
        tabla = Mock()
        tabla.winfo_exists.return_value = True
        tabla.cget.return_value = "Treeview"
        tabla.selection.return_value = ()
        tabla.get_children.return_value = ["fila2", "fila1"]
        tabla.yview.return_value = (0.0, 1.0)
        tabla.item.return_value = ("Prueba", "g/ml", "ASTM", "150", "160")
        tabla.bbox.return_value = (5, 25, 100, 25)
        tabla.winfo_width.return_value = 800
        with patch("cambios_alineacion.tk.Label") as etiqueta, patch("cambios_alineacion.ttk.Style") as estilo:
            estilo.return_value.lookup.side_effect = lambda _, campo, *a: {"font": ("Segoe UI", 12), "fieldbackground": "white"}.get(campo, "")
            resaltador = ResaltadorCambios(tabla, {"fila1": {"metodo"}}, Mock(), Mock())
            resaltador.dibujar()
            self.assertEqual(etiqueta.call_count, 1)
            self.assertEqual(etiqueta.return_value.configure.call_args.kwargs["foreground"], "#3b82f6")
            self.assertEqual(etiqueta.return_value.configure.call_args.kwargs["font"], ("Segoe UI", 12))
            tabla.bbox.assert_called_once_with("fila1", "#3")

    def test_resaltado_reutiliza_etiquetas_y_solo_consulta_filas_visibles(self):
        tabla = Mock()
        tabla.winfo_exists.return_value = True
        tabla.cget.return_value = "Treeview"
        tabla.selection.return_value = ()
        tabla.get_children.return_value = [f"f{i}" for i in range(200)]
        tabla.yview.return_value = (0.5, 0.55)
        tabla.item.return_value = ("Prueba", "g/ml", "ASTM", "150", "160")
        tabla.bbox.return_value = (5, 25, 100, 25)
        tabla.winfo_width.return_value = 800
        with patch("cambios_alineacion.tk.Label") as etiqueta, patch("cambios_alineacion.ttk.Style") as estilo:
            estilo.return_value.lookup.side_effect = lambda _, campo, *a: {"font": ("Segoe UI", 12), "fieldbackground": "white"}.get(campo, "")
            resaltador = ResaltadorCambios(tabla, {f"f{i}": {"metodo"} for i in range(200)}, Mock(), Mock())
            resaltador.dibujar()
            creadas = etiqueta.call_count
            self.assertGreater(creadas, 0)
            self.assertLess(creadas, 15)
            for llamada in tabla.bbox.call_args_list:
                self.assertGreaterEqual(int(llamada.args[0][1:]), 100)
            resaltador.dibujar()
            self.assertEqual(etiqueta.call_count, creadas)
            etiqueta.return_value.destroy.assert_not_called()



if __name__ == "__main__":
    unittest.main()
