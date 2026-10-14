import json
import copy
import queue
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook
from traductor_especificaciones import (ServicioTraductor, DiccionarioTraductor, ARCHIVO_EXCEL,
                                       ARCHIVO_ESTADO, ARCHIVO_HISTORIAL)


class TraductorTests(unittest.TestCase):
    def escribir(self, carpeta, filas):
        libro = Workbook()
        hoja = libro.active
        hoja.title = "Traducciones"
        hoja.append(["ESPAÑOL", "INGLÉS", "ACTIVO", "OBSERVACIONES"])
        for fila in filas:
            hoja.append(fila)
        libro.save(Path(carpeta) / ARCHIVO_EXCEL)
        libro.close()

    def test_frases_largas_primero_cache_y_no_hay_sustituciones_encadenadas(self):
        registros = [{"es": es, "en": en, "activo": True} for es, en in
                     (("VERDE", "GREEN"), ("VERDE CLARO", "LIGHT GREEN"), ("ROJO", "RED"), ("RED", "BLUE"))]
        traductor = DiccionarioTraductor(registros)
        with patch.object(Path, "read_bytes", side_effect=AssertionError("No leer archivos al traducir")):
            self.assertEqual(traductor.traducir("VERDE CLARO 149", "es", "en"), ("LIGHT GREEN 149", True))
            self.assertEqual(traductor.traducir("ROJO", "es", "en"), ("RED", True))
            traductor.traducir("ROJO", "es", "en")
            self.assertEqual(traductor.traducir.cache_info().hits, 1)

    def test_carga_estado_y_solo_registra_cambios_reales(self):
        with tempfile.TemporaryDirectory() as carpeta:
            self.escribir(carpeta, [("VERDE", "GREEN", "SI", "")])
            servicio = ServicioTraductor(carpeta)
            self.assertEqual(servicio.actualizar()["cambios"], 1)
            registro = Path(carpeta) / ARCHIVO_HISTORIAL
            antes = registro.read_bytes()
            self.assertEqual(servicio.actualizar()["cambios"], 0)
            self.assertEqual(antes, registro.read_bytes())
            # La última versión válida permite iniciar sin leer Excel.
            with patch.object(Path, "read_bytes", side_effect=AssertionError("No abrir Excel")):
                recargado = ServicioTraductor(carpeta)
                self.assertEqual(recargado.diccionario.traducir("VERDE", "es", "en"), ("GREEN", True))
            self.escribir(carpeta, [("VERDE", "GREENISH", "SI", "Cambio"), ("AZUL", "BLUE", "SI", "")])
            self.assertEqual(servicio.actualizar()["cambios"], 2)
            self.assertIn("MODIFICADA", registro.read_text(encoding="utf-8"))
            self.assertEqual(servicio.diccionario.traducir.cache_info().currsize, 0)
            estado = json.loads((Path(carpeta) / ARCHIVO_ESTADO).read_text(encoding="utf-8"))
            self.assertEqual(len(estado["ultimos_cambios"]), 2)

    def test_excel_invalido_conserva_diccionario_estado_e_historial(self):
        with tempfile.TemporaryDirectory() as carpeta:
            self.escribir(carpeta, [("VERDE", "GREEN", "SI", "")])
            servicio = ServicioTraductor(carpeta)
            servicio.actualizar()
            diccionario = servicio.diccionario
            estado = (Path(carpeta) / ARCHIVO_ESTADO).read_bytes()
            historial = (Path(carpeta) / ARCHIVO_HISTORIAL).read_bytes()
            for filas in ([('VERDE', 'GREEN', 'SI', ''), ('verde', 'BLUE', 'SI', '')],
                          [('VERDE', '', 'SI', '')], [('VERDE', 'GREEN', 'TALVEZ', '')],
                          [('VERDE', '=A1', 'SI', '')]):
                self.escribir(carpeta, filas)
                with self.assertRaises(ValueError):
                    servicio.actualizar()
                self.assertIs(diccionario, servicio.diccionario)
                self.assertEqual(estado, (Path(carpeta) / ARCHIVO_ESTADO).read_bytes())
                self.assertEqual(historial, (Path(carpeta) / ARCHIVO_HISTORIAL).read_bytes())

    def test_alias_inverso_primera_fila_y_desactivacion(self):
        with tempfile.TemporaryDirectory() as carpeta:
            self.escribir(carpeta, [("ÁMBAR", "AMBER", "SI", ""), ("AMBAR", "AMBER", "SI", "Variante")])
            servicio = ServicioTraductor(carpeta)
            servicio.actualizar()
            self.assertEqual(servicio.diccionario.traducir("AMBER", "en", "es"), ("ÁMBAR", True))
            self.escribir(carpeta, [("ÁMBAR", "AMBER", "NO", ""), ("AMBAR", "AMBER", "SI", "Variante")])
            servicio.actualizar()
            self.assertEqual(servicio.diccionario.traducir("AMBER", "en", "es"), ("AMBAR", True))
            self.assertFalse(servicio.diccionario.traducir("ÁMBAR", "es", "en")[1])

    def test_estado_corrupto_no_impide_recuperar_desde_excel(self):
        with tempfile.TemporaryDirectory() as carpeta:
            (Path(carpeta) / ARCHIVO_ESTADO).write_text("[]", encoding="utf-8")
            self.escribir(carpeta, [("VERDE", "GREEN", "SI", "")])
            servicio = ServicioTraductor(carpeta)
            servicio.actualizar()
            self.assertEqual(servicio.diccionario.traducir("VERDE", "es", "en"), ("GREEN", True))

    def test_edicion_desde_metodos_actualiza_excel_e_impide_sobrescribir_cambios_externos(self):
        with tempfile.TemporaryDirectory() as carpeta:
            self.escribir(carpeta, [("VERDE", "GREEN", "SI", "")])
            servicio = ServicioTraductor(carpeta)
            servicio.actualizar()
            huella = servicio.estado["sha256_excel"]
            filas = [{"es": "AZUL", "en": "BLUE", "activo": True, "observaciones": "Nuevo"}]
            resultado = servicio.guardar_registros(filas, huella)
            self.assertEqual(resultado["cambios"], 2)
            self.assertEqual(resultado["sha256_excel"], servicio.estado["sha256_excel"])
            self.assertEqual(servicio.diccionario.traducir("AZUL", "es", "en"), ("BLUE", True))
            with self.assertRaisesRegex(ValueError, "cambió"):
                servicio.guardar_registros(filas, huella)
            with self.assertRaises(ValueError):
                servicio.guardar_registros([{"es": "", "en": "BLUE", "activo": True}], resultado["sha256_excel"])
            self.assertEqual(servicio.diccionario.traducir("AZUL", "es", "en"), ("BLUE", True))

    def test_editor_recarga_pendientes_sin_reemplazar_correcciones_manuales(self):
        import test_editor_bilingue_interfaz as fixtures
        from edicion_bilingue import BorradoresBilingues
        from traductor_especificaciones import SERVICIO
        with tempfile.TemporaryDirectory() as carpeta:
            editor = fixtures.EditorBilingueTests().editor(carpeta)
            editor._traductor = SERVICIO.diccionario
            contenido = copy.deepcopy(editor._bilingue.borradores["es"])
            contenido["pruebas"][0]["especificacion"] = "FRASE NUEVA"
            editor._bilingue = BorradoresBilingues(contenido, editor._traducir_campo_bilingue)
            en = copy.deepcopy(editor._bilingue.borradores["en"])
            en["pruebas"][0]["valor_tipico"] = "MANUAL VALUE"
            editor._bilingue.actualizar("en", en)
            nuevo = DiccionarioTraductor([{"es": "FRASE NUEVA", "en": "NEW PHRASE", "activo": True}])
            editor._cola_traductor = queue.Queue()
            editor._cola_traductor.put((True, {"cambios": 1, "equivalencias": 1}, nuevo, True))
            editor.ventana.winfo_exists.return_value = True
            editor.boton_actualizar_traductor = Mock()
            editor._cerrar_editor_celda = Mock()
            editor._seleccion = Mock(return_value=None)
            editor._aplicar_borrador = Mock()
            editor._actualizar_tabla = Mock()
            with patch("interfaz_aplicacion_especificaciones.messagebox"):
                editor._recibir_actualizacion_traductor()
            en = editor._bilingue.borradores["en"]["pruebas"][0]
            self.assertEqual(en["especificacion"], "NEW PHRASE")
            self.assertEqual(en["valor_tipico"], "MANUAL VALUE")
            self.assertNotIn((en["_id_bilingue"], "especificacion"), editor._bilingue.pendientes["en"])

    def test_hilo_de_carga_solo_entrega_resultado_por_cola(self):
        import test_editor_bilingue_interfaz as fixtures
        with tempfile.TemporaryDirectory() as carpeta:
            editor = fixtures.EditorBilingueTests().editor(carpeta)
            editor._cola_traductor = queue.Queue()
            editor._cargando_traductor = False
            editor.boton_actualizar_traductor = Mock()
            with patch("interfaz_aplicacion_especificaciones.threading.Thread") as hilo, patch("interfaz_aplicacion_especificaciones.actualizar_traductor", return_value={"cambios": 0, "equivalencias": 335}):
                editor._actualizar_traductor_editor()
                llamadas = editor.ventana.mock_calls[:]
                hilo.call_args.kwargs["target"]()
                self.assertEqual(editor.ventana.mock_calls, llamadas)
                self.assertTrue(editor._cola_traductor.get_nowait()[0])


if __name__ == "__main__":
    unittest.main()
