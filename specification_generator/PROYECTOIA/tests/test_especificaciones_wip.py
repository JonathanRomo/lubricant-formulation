import copy
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import test_editor_bilingue_interfaz as fixtures
from edicion_bilingue import BorradoresBilingues
from especificaciones_wip import es_wip, restaurar_estado, ruta_estado
from interfaz_seleccion_especificaciones import SelectorEspecificaciones, cargar_catalogo, parejas_idioma


class WorkInProgressTests(unittest.TestCase):
    def guardar_borrador(self, carpeta):
        editor = fixtures.EditorBilingueTests().editor(carpeta)
        for ruta in editor._rutas_idioma.values():
            ruta.write_text("especificación terminada intacta", encoding="utf-8")
        en = editor._bilingue.borradores["en"]
        en["pruebas"][0]["especificacion"] = ""
        for contenido in editor._bilingue.borradores.values():
            contenido["fecha_reanalisis"] = ""
            contenido["control"].update(nombre_cliente="", revision_cliente="")
        editor.fecha_reanalisis.get.return_value = ""
        clave = (en["pruebas"][0]["_id_bilingue"], "especificacion")
        editor._bilingue.pendientes["en"].add(clave)
        rutas = [Path(carpeta) / "trabajo_es.json", Path(carpeta) / "trabajo_en.json"]
        with patch("interfaz_aplicacion_especificaciones.filedialog.asksaveasfilename", side_effect=[str(r) for r in rutas]), patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
            mensajes.askyesno.return_value = True
            editor._guardar_bilingue(en_progreso=True)
            mensajes.showerror.assert_not_called()
        return editor, rutas, clave

    def test_borrador_con_pendientes_se_guarda_fuera_del_json_y_preserva_original(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, rutas, clave = self.guardar_borrador(carpeta)
            self.assertTrue(all(es_wip(r) for r in rutas))
            self.assertEqual((Path(carpeta) / "es.json").read_text(encoding="utf-8"), "especificación terminada intacta")
            datos = json.loads(rutas[1].read_text(encoding="utf-8"))
            self.assertEqual(datos["pruebas"][0]["especificacion"], "")
            self.assertEqual(datos["FECHA DE REANALISIS"], "")
            self.assertEqual(datos["Desarrollador"], "Nombre")
            self.assertNotIn("edicion_bilingue", datos["documento"])
            self.assertNotIn("wip", datos)
            self.assertNotIn("pendientes", datos)
            modelo = BorradoresBilingues(copy.deepcopy(editor._bilingue.borradores["es"]), editor._traducir_campo_bilingue)
            recuperadas = restaurar_estado(modelo, rutas[1])
            self.assertEqual(recuperadas, {"es": rutas[0].resolve(), "en": rutas[1].resolve()})
            self.assertIn(clave, modelo.pendientes["en"])
            self.assertEqual(modelo.borradores, editor._bilingue.borradores)
            catalogo = cargar_catalogo(Path(carpeta))
            self.assertEqual(len(catalogo.especificaciones), 2)
            self.assertTrue(all(e.en_progreso for e in catalogo.especificaciones))
            self.assertEqual(len(parejas_idioma(catalogo.especificaciones)), 2)

    def test_wip_recupera_confirmaciones_y_admite_estado_antiguo(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, rutas, _ = self.guardar_borrador(carpeta)
            fila = editor._bilingue.borradores["en"]["pruebas"][0]
            marca = (fila["_id_bilingue"], "valor_tipico")
            editor._bilingue.confirmar("en", marca)
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                mensajes.askyesno.return_value = True
                editor._guardar_bilingue(en_progreso=True)
            modelo = BorradoresBilingues(copy.deepcopy(editor._bilingue.borradores["es"]), editor._traducir_campo_bilingue)
            restaurar_estado(modelo, rutas[0])
            self.assertIn(marca, modelo.manuales["en"])
            archivo = ruta_estado(rutas[0])
            datos = json.loads(archivo.read_text(encoding="utf-8"))
            datos.pop("manuales")
            archivo.write_text(json.dumps(datos), encoding="utf-8")
            restaurar_estado(modelo, rutas[0])
            self.assertEqual(modelo.manuales, {"es": set(), "en": set()})

    def test_borrador_admite_filas_vacias_y_periodicidad_incompleta(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = fixtures.EditorBilingueTests().editor(carpeta)
            contenido = copy.deepcopy(editor._bilingue.borradores["es"])
            contenido["pruebas"] = [{"prueba": "", "periodicidad": ""}]
            datos = editor._datos_en_progreso(contenido, "es")
            self.assertEqual(datos["pruebas"][0]["prueba"], "")
            self.assertEqual(datos["alineacion_metodos_lab"]["detalle"][0]["valores_alineados"]["prueba"], "")
            contenido["pruebas"] = []
            self.assertEqual(editor._datos_en_progreso(contenido, "es")["pruebas"], [])

    def test_editor_reabre_estado_bilingue_con_pendientes(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, rutas, clave = self.guardar_borrador(carpeta)
            editor.archivo_inicial = rutas[1]
            editor.destino_inicial = rutas[1]
            editor.datos_originales = json.loads(rutas[1].read_text(encoding="utf-8"))
            editor._idioma_anterior = "en"
            editor._contenido_bilingue = Mock(return_value=copy.deepcopy(editor._bilingue.borradores["en"]))
            editor.variables_control = {c: Mock() for c in editor._bilingue.borradores["en"]["control"]}
            editor.desarrollador = Mock()
            editor.notas = Mock()
            editor.boton_guardar_wip = Mock()
            editor._programar_resaltado = Mock()
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                editor._iniciar_bilingue()
                mensajes.showerror.assert_not_called()
            self.assertIn(clave, editor._bilingue.pendientes["en"])
            self.assertEqual(editor.fecha_reanalisis.get(), "")
            self.assertEqual(editor._bilingue.borradores["en"]["control"]["nombre_cliente"], "")
            self.assertEqual(editor.pruebas[0]["especificacion"], "")
            self.assertEqual(editor._bilingue.borradores["es"]["pruebas"][0]["especificacion"], "149 MÍNIMO")

    def test_finalizar_quita_estado_wip(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, rutas, clave = self.guardar_borrador(carpeta)
            editor._bilingue.borradores["en"]["pruebas"][0]["especificacion"] = "149 MINIMUM"
            editor._bilingue.confirmar("en", clave)
            for contenido in editor._bilingue.borradores.values():
                contenido["fecha_reanalisis"] = "06"
                contenido["control"].update(nombre_cliente="Cliente", revision_cliente="-")
            editor.fecha_reanalisis.get.return_value = "06"
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch("generador_hii.conectar"), patch("generador_hii.vincular"):
                mensajes.askyesno.return_value = True
                editor._guardar_bilingue()
                mensajes.showerror.assert_not_called()
            self.assertTrue(all(not es_wip(r) for r in rutas))
            for ruta in rutas:
                datos = json.loads(ruta.read_text(encoding="utf-8"))
                self.assertEqual(datos["FECHA DE REANALISIS"],
                                 "06 months" if datos["documento"]["idioma"] == "en" else "06 meses")
            self.assertTrue(all(not e.en_progreso for e in cargar_catalogo(Path(carpeta)).especificaciones))

    def test_borrador_antiguo_recupera_reanalisis_desde_el_campo_json(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, rutas, _ = self.guardar_borrador(carpeta)
            for ruta in rutas:
                documento = json.loads(ruta.read_text(encoding="utf-8"))
                documento["FECHA DE REANALISIS"] = "06 meses"
                ruta.write_text(json.dumps(documento), encoding="utf-8")
                estado = json.loads(ruta_estado(ruta).read_text(encoding="utf-8"))
                for borrador in estado["borradores"].values():
                    borrador.pop("fecha_reanalisis")
                ruta_estado(ruta).write_text(json.dumps(estado), encoding="utf-8")
            editor.archivo_inicial = editor.destino_inicial = rutas[0]
            editor.datos_originales = json.loads(rutas[0].read_text(encoding="utf-8"))
            editor._contenido_bilingue = Mock(return_value=copy.deepcopy(editor._bilingue.borradores["es"]))
            editor.variables_control = {c: Mock() for c in editor._bilingue.borradores["es"]["control"]}
            editor.notas = Mock()
            editor.boton_guardar_wip = Mock()
            editor._programar_resaltado = Mock()
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                editor._iniciar_bilingue()
                mensajes.showerror.assert_not_called()
            self.assertEqual(editor.fecha_reanalisis.get(), "06")
            self.assertTrue(all(b["fecha_reanalisis"] == "06" for b in editor._bilingue.borradores.values()))

    def test_selector_pone_wip_primero_y_amarillo(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, rutas, _ = self.guardar_borrador(carpeta)
            documento = json.loads(rutas[0].read_text(encoding="utf-8"))
            documento["producto"]["nombre"] = "Terminada"
            documento["versiones_idioma"] = {}
            (Path(carpeta) / "terminada.json").write_text(json.dumps(documento), encoding="utf-8")
            catalogo = cargar_catalogo(Path(carpeta))
            panel = SimpleNamespace(catalogo=catalogo, entrada_busqueda=Mock(get=Mock(return_value="")),
                                    lista=Mock(), documentos_marcados=set(), mensaje_busqueda=Mock())
            selector = SimpleNamespace(paneles={"nuevas": panel})
            SelectorEspecificaciones._actualizar_sugerencias(selector, "nuevas")
            self.assertTrue(panel.sugerencias[0].en_progreso)
            self.assertFalse(panel.sugerencias[-1].en_progreso)
            self.assertEqual(len(panel.sugerencias), 2)
            self.assertIn("WIP / En progreso", panel.lista.insert.call_args_list[0].args[1])
            self.assertEqual(panel.lista.itemconfigure.call_args_list[0].kwargs["foreground"], "#d4a000")


if __name__ == "__main__":
    unittest.main()
