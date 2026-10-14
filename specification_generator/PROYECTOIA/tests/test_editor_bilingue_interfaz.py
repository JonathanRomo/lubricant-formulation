import copy
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from edicion_bilingue import BorradoresBilingues
from interfaz_aplicacion_especificaciones import GeneradorEspecificacion, NOTAS_PLANTILLA, CAMPOS_CONTROL_GENERADOR
from interfaz_seleccion_especificaciones import cargar_catalogo, parejas_idioma


class EditorBilingueTests(unittest.TestCase):
    def editor(self, carpeta):
        editor = GeneradorEspecificacion.__new__(GeneradorEspecificacion)
        editor.maestro = [{"NOMBRE IQM": "DENSIDAD", "NOMBRE IQM INGLES": "DENSITY",
                           "UNIDADES": "g/ml", "UNITS": "g/ml", "METODO": "ASTM D445", "MÉTODO INGLÉS": "ASTM D445"}]
        contenido = {"idioma": "es", "nombre_producto": "Producto", "desarrollador": "Nombre",
                     "control": {"codigo_unico_producto": "P1", "revision": "01", "nombre_cliente": "Cliente", "elaboro_y_aprobo": "Persona"},
                     "notas": [NOTAS_PLANTILLA["es"]], "pruebas": [
                         {"prueba": "DENSIDAD", "metodo": "ASTM D445", "unidades": "g/ml", "especificacion": "149 MÍNIMO", "valor_tipico": "150", "periodicidad": "▶"}]}
        editor._bilingue = BorradoresBilingues(contenido, editor._traducir_campo_bilingue)
        editor._idioma_anterior = "es"
        editor.pruebas = copy.deepcopy(editor._bilingue.borradores["es"]["pruebas"])
        editor._capturar_bilingue = Mock()
        editor._rutas_idioma = {"es": Path(carpeta) / "es.json", "en": Path(carpeta) / "en.json"}
        editor._originales_idioma = {}
        editor._cambios_alineacion = {}
        editor.app = SimpleNamespace(rutas={"json_alineadas": Path(carpeta)})
        editor.nombre_producto = Mock(get=Mock(return_value="Producto"))
        editor.variables_control = {"revision": Mock(get=Mock(return_value="01"))}
        editor.ventana = Mock()
        editor.al_guardar = Mock()
        editor._contenido_para_revision = Mock(return_value=())
        editor._consultas_revision = set()
        editor._revisiones_historial = {}
        return editor

    def test_maestro_traduce_con_fuente_intacta_y_no_usa_fallback_como_traduccion(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            en = editor._bilingue.borradores["en"]
            self.assertEqual(en["pruebas"][0]["prueba"], "DENSITY")
            self.assertEqual(en["pruebas"][0]["metodo"], "ASTM D445")
            self.assertEqual(en["notas"], [NOTAS_PLANTILLA["en"]])
            self.assertFalse(editor._bilingue.pendientes["en"])
            editor.maestro[0]["NOMBRE IQM INGLES"] = ""
            _, seguro = editor._traducir_campo_bilingue("prueba", "DENSIDAD", "es", "en", {"prueba": "DENSIDAD"})
            self.assertFalse(seguro)

    def test_selector_de_idioma_traduce_la_pareja_guardada_automaticamente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            pareja = copy.deepcopy(editor._bilingue.borradores["en"])
            pareja["pruebas"][0]["especificacion"] = "149 MÍNIMO"
            editor._bilingue.cargar_pareja(pareja)
            editor._idioma = Mock(get=Mock(return_value="English"))
            editor._nombres_sugeridos_por_fila = {}
            editor._limpiar_sugerencia = Mock()
            editor._cerrar_editor_celda = Mock()
            editor._seleccion = Mock(return_value=None)
            editor._actualizar_tabla = Mock()
            editor._programar_resaltado = Mock()
            editor.variables_control = {c: Mock(get=Mock(return_value="Persona")) for c, _, _ in CAMPOS_CONTROL_GENERADOR}
            editor._etiquetas = {c: Mock() for c, _, _ in CAMPOS_CONTROL_GENERADOR}
            editor.desarrollador = Mock()
            editor.notas = Mock(get=Mock(return_value=NOTAS_PLANTILLA["en"]))
            editor.etiqueta_codigo_invalido = None
            for nombre in ("etiqueta_idioma", "selector_idioma", "etiqueta_producto", "etiqueta_desarrollador", "etiqueta_titulo_control", "marco_pruebas", "tabla_pruebas", "_ordenador_pruebas", "boton_eliminar", "boton_alinear", "etiqueta_notas", "boton_guardar", "boton_cancelar"):
                setattr(editor, nombre, Mock())
            editor._actualizar_idioma()
            self.assertEqual(editor.pruebas[0]["especificacion"], "149 MINIMUM")
            self.assertEqual(editor._idioma_anterior, "en")
            self.assertEqual(editor._bilingue.borradores["es"]["pruebas"][0]["especificacion"], "149 MÍNIMO")

    def test_guarda_dos_json_y_referencias_correctas_sin_ids_internos(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch("registro_hii.conectar"), patch("registro_hii.vincular"):
                mensajes.askyesno.return_value = True
                editor._guardar_bilingue()
                mensajes.showerror.assert_not_called()
            es = json.loads((Path(carpeta) / "es.json").read_text(encoding="utf-8"))
            en = json.loads((Path(carpeta) / "en.json").read_text(encoding="utf-8"))
            self.assertEqual(es["documento"]["idioma"], "es")
            self.assertEqual(en["documento"]["idioma"], "en")
            self.assertEqual(en["pruebas"][0]["especificacion"], "149 MINIMUM")
            self.assertEqual(es["versiones_idioma"]["INGLES"][0]["archivo_json"], "en.json")
            self.assertNotIn("_id_bilingue", en["pruebas"][0])
            self.assertEqual(editor.al_guardar.call_count, 2)
            catalogo = cargar_catalogo(Path(carpeta))
            self.assertEqual(len(parejas_idioma(catalogo.especificaciones)), 2)

    def test_cancelar_y_rutas_iguales_no_guardan_archivos(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                mensajes.askyesno.return_value = False
                editor._guardar_bilingue()
                self.assertFalse(list(Path(carpeta).iterdir()))
                editor._rutas_idioma["en"] = editor._rutas_idioma["es"]
                editor._guardar_bilingue()
                mensajes.showerror.assert_called_once()
                self.assertFalse(list(Path(carpeta).iterdir()))

    def test_pendiente_rojo_conserva_fondo_fuente_y_asterisco_solo_visual(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            editor._idioma_anterior = "en"
            editor.pruebas = editor._bilingue.borradores["en"]["pruebas"]
            editor._bilingue.pendientes["en"].add((editor.pruebas[0]["_id_bilingue"], "especificacion"))
            editor._cambios_alineacion["en"] = {editor.pruebas[0]["_id_bilingue"]: {"especificacion"}}
            editor._marcas_pendientes = []
            editor._celda_editada = None
            editor._celda_activa = None
            editor.tabla_pruebas = Mock()
            editor.tabla_pruebas.bbox.return_value = (5, 25, 200, 25)
            editor.tabla_pruebas.winfo_width.return_value = 800
            editor.tabla_pruebas.cget.return_value = "Treeview"
            editor.tabla_pruebas.selection.return_value = ()
            for nombre in ("notas", "boton_confirmar_traduccion", "boton_confirmar_notas", "referencia_notas"):
                setattr(editor, nombre, Mock())
            with patch("interfaz_aplicacion_especificaciones.tk.Label") as etiqueta, patch("interfaz_aplicacion_especificaciones.ttk.Style") as estilo, patch("interfaz_aplicacion_especificaciones.ui.paleta", return_value={"panel": "#ffffff", "texto": "#102030", "fondo": "#eef3f9", "acento": "#1261a0"}):
                estilo.return_value.lookup.side_effect = lambda _nombre, campo, *args: {"fieldbackground": "#ffffff", "font": ("Segoe UI", 12)}.get(campo, "")
                editor._resaltar_pendientes()
                self.assertEqual(etiqueta.call_count, 1)
                self.assertEqual(etiqueta.call_args.kwargs["background"], "#ffffff")
                self.assertEqual(etiqueta.call_args.kwargs["foreground"], "#c62828")
                self.assertEqual(etiqueta.call_args.kwargs["font"], ("Segoe UI", 12))
                editor.pruebas[0]["especificacion"] = ""
                etiqueta.reset_mock()
                editor._resaltar_pendientes()
                self.assertEqual(etiqueta.call_args.kwargs["text"], "*")
                self.assertEqual(editor.pruebas[0]["especificacion"], "")
                editor._bilingue.pendientes["en"].clear()
                etiqueta.reset_mock()
                editor._resaltar_pendientes()
                self.assertEqual(etiqueta.call_args.kwargs["foreground"], "#3b82f6")

    def test_guardado_con_pendientes_enfoca_celda_y_no_escribe(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            editor._idioma_anterior = "en"
            editor.pruebas = editor._bilingue.borradores["en"]["pruebas"]
            editor._bilingue.pendientes["en"].add((editor.pruebas[0]["_id_bilingue"], "especificacion"))
            editor.tabla_pruebas = Mock()
            editor._mostrar_referencia_bilingue = Mock()
            editor._programar_resaltado = Mock()
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                editor._guardar_bilingue()
                mensajes.showwarning.assert_called_once()
                mensajes.askyesno.assert_not_called()
            editor.tabla_pruebas.see.assert_called_once_with("0")
            self.assertEqual(editor._celda_activa, ("0", "especificacion"))
            self.assertFalse(list(Path(carpeta).iterdir()))


if __name__ == "__main__":
    unittest.main()
