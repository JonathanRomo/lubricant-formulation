import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
from rutas_proyecto import carpeta_listas, ruta_maestro
from traductor_especificaciones import ServicioTraductor, ARCHIVO_EXCEL
from interfaz_aplicacion_especificaciones import _guardar_configuracion, _cargar_configuracion
from alinear_especificaciones_json import ejecutar


class RutasProyectoTests(unittest.TestCase):
    def test_default_relativa_y_cambio_sin_releer_configuracion(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal).resolve()
            config = base / "configuracion_interfaz.json"
            self.assertEqual(carpeta_listas(base), base / "METODOS_LAB")
            config.write_text(json.dumps({"metodos_lab": "listas"}), encoding="utf-8")
            self.assertEqual(ruta_maestro(base), base / "listas" / "METODOS_LAB_UNIFICADO.xlsx")
            with patch.object(Path, "read_text", side_effect=AssertionError("No releer sin cambios")):
                self.assertEqual(carpeta_listas(base), base / "listas")
            config.write_text(json.dumps({"metodos_lab": "otras_listas"}), encoding="utf-8")
            self.assertEqual(carpeta_listas(base), base / "otras_listas")

    def test_guarda_y_recupera_carpeta_y_rechaza_listas_ausentes(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal).resolve()
            listas = base / "listas"
            listas.mkdir()
            config = base / "configuracion_interfaz.json"
            with patch("interfaz_aplicacion_especificaciones.ARCHIVO_CONFIGURACION", config):
                with self.assertRaises(ValueError):
                    _guardar_configuracion({"metodos_lab": listas})
                self.assertFalse(config.exists())
                (listas / "METODOS_LAB_UNIFICADO.xlsx").write_bytes(b"maestro")
                (listas / ARCHIVO_EXCEL).write_bytes(b"traducciones")
                _guardar_configuracion({"metodos_lab": listas})
                self.assertEqual(_cargar_configuracion()["metodos_lab"], listas)
            self.assertEqual(ruta_maestro(base), listas / "METODOS_LAB_UNIFICADO.xlsx")

    def test_traductor_cambia_de_carpeta_sin_perder_version_valida(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal).resolve()
            carpetas = [base / "a", base / "b"]
            for carpeta, traduccion in zip(carpetas, ("GREEN", "GREENISH")):
                carpeta.mkdir()
                libro = Workbook()
                hoja = libro.active
                hoja.title = "Traducciones"
                hoja.append(["ESPAÑOL", "INGLÉS", "ACTIVO", "OBSERVACIONES"])
                hoja.append(["VERDE", traduccion, "SI", ""])
                libro.save(carpeta / ARCHIVO_EXCEL)
                libro.close()
            servicio = ServicioTraductor(carpetas[0])
            servicio.actualizar()
            anterior = servicio.diccionario
            servicio.configurar_carpeta(carpetas[1])
            self.assertIs(servicio.diccionario, anterior)
            servicio.actualizar()
            self.assertEqual(servicio.diccionario.traducir("VERDE", "es", "en"), ("GREENISH", True))
            self.assertTrue((carpetas[1] / ".estado_traductor.json").is_file())

    def test_alineador_lee_maestro_desde_carpeta_configurada(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal).resolve()
            listas = base / "listas"
            entrada = base / "origen"
            listas.mkdir()
            entrada.mkdir()
            maestro = listas / "METODOS_LAB_UNIFICADO.xlsx"
            maestro.write_bytes(b"maestro de prueba para hash")
            (base / "configuracion_interfaz.json").write_text(json.dumps({"metodos_lab": str(listas)}), encoding="utf-8")
            (entrada / "producto.json").write_text(json.dumps({"producto": {"nombre": "Producto"}, "documento": {"idioma": "es"},
                "control_documental": {"codigo_unico_producto": "P1", "revision": "00", "nombre_cliente": "Cliente"}, "pruebas": []}), encoding="utf-8")
            with patch("alinear_especificaciones_json.cargar_maestro", return_value=[]) as cargar, patch("builtins.print"):
                resultado = ejecutar(base, carpeta_entrada=entrada, carpeta_salida=base / "salida")
            cargar.assert_called_once_with(maestro)
            self.assertEqual(resultado["errores"], 0)
            self.assertEqual(resultado["archivos"], 1)


if __name__ == "__main__":
    unittest.main()


class ControlesSeguridadTests(unittest.TestCase):
    def test_rutas_derivadas_no_salen_ni_admiten_streams(self):
        from rutas_proyecto import ruta_contenida, referencia_documento
        with tempfile.TemporaryDirectory() as td:
            base=Path(td)/'permitida';base.mkdir()
            self.assertEqual(ruta_contenida('correcto.json',base),(base/'correcto.json').resolve())
            for nombre in ('../fuera.json', '../fuera.docx', 'archivo:oculto.json', 'CON.json'):
                with self.subTest(nombre=nombre),self.assertRaises(ValueError):
                    ruta_contenida(nombre,base)
            with self.assertRaises(ValueError):referencia_documento('../fuera.json',base)
            with self.assertRaises(ValueError):referencia_documento('archivo.txt',base)

    def test_paquetes_rechazan_traversal_y_expansion_desproporcionada(self):
        import io,zipfile
        from rutas_proyecto import validar_archivo_office
        for nombre,contenido in [('../fuera.xml',b'x'),('word/document.xml',b'0'*(9*1024**2))]:
            salida=io.BytesIO()
            with zipfile.ZipFile(salida,'w',compression=zipfile.ZIP_DEFLATED) as z:z.writestr(nombre,contenido)
            with self.subTest(nombre=nombre),self.assertRaises(ValueError):validar_archivo_office(salida.getvalue())
        salida=io.BytesIO()
        with zipfile.ZipFile(salida,'w') as z:z.writestr('word/document.xml',b'<x/>')
        validar_archivo_office(salida.getvalue())

    def test_limite_fisico_antes_de_interpretar_office(self):
        from rutas_proyecto import validar_archivo_office
        with tempfile.TemporaryDirectory() as td:
            ruta=Path(td)/'grande.docx'
            with ruta.open('wb') as f:f.truncate(129*1024**2)
            with self.assertRaises(ValueError):validar_archivo_office(ruta)

    def test_no_ejecuta_programa_colocado_en_cwd(self):
        import os
        from rutas_proyecto import ejecutable_conocido
        with tempfile.TemporaryDirectory() as td:
            ruta=Path(td)/'auxiliar_seguridad.exe';ruta.write_bytes(b'no ejecutar')
            with patch('pathlib.Path.cwd',return_value=Path(td)),patch.dict(os.environ,{'PATH':td}):
                self.assertIsNone(ejecutable_conocido('auxiliar_seguridad.exe'))

    def test_ambas_rutas_word_desactivan_macros_y_vinculos_antes_de_abrir(self):
        import queue
        import conversion_word,gestor_especificaciones
        from unittest.mock import Mock
        for reutilizable in (True,False):
            word=Mock();resultados=queue.Queue()
            def abrir(*args,**kwargs):
                self.assertEqual(word.AutomationSecurity,3)
                self.assertIs(word.Options.UpdateLinksAtOpen,False)
                self.assertTrue(kwargs['ReadOnly'])
                return Mock()
            word.Documents.Open.side_effect=abrir
            with patch('win32com.client.DispatchEx',return_value=word),patch('pythoncom.CoInitialize'),patch('pythoncom.CoUninitialize'):
                if reutilizable:
                    ordenes=queue.Queue();ordenes.put(('origen.doc','destino.docx'));ordenes.put(None)
                    conversion_word.trabajador_word(ordenes,resultados)
                else:gestor_especificaciones._convertir_word_com_worker('origen.doc','destino.docx',resultados)
            word.Documents.Open.assert_called_once()
            mensajes=[]
            while not resultados.empty():mensajes.append(resultados.get())
            self.assertTrue(any(m[0]=='ok' for m in mensajes),mensajes)

    def test_revalida_archivo_modificado_sin_confiar_en_cache_anterior(self):
        import zipfile
        from rutas_proyecto import validar_archivo_office
        with tempfile.TemporaryDirectory() as td:
            ruta=Path(td)/'documento.docx'
            with zipfile.ZipFile(ruta,'w') as z:z.writestr('word/document.xml','<x/>')
            validar_archivo_office(ruta)
            with zipfile.ZipFile(ruta,'w') as z:z.writestr('../fuera.xml','<x/>')
            with self.assertRaises(ValueError):validar_archivo_office(ruta)

    def test_no_sigue_enlace_fuera_de_la_carpeta_autorizada(self):
        from rutas_proyecto import ruta_contenida
        with tempfile.TemporaryDirectory() as td:
            base=Path(td)/'permitida';base.mkdir()
            fuera=Path(td)/'fuera';fuera.mkdir()
            try:(base/'enlace').symlink_to(fuera,target_is_directory=True)
            except OSError:self.skipTest('El equipo no permite crear enlaces simbolicos')
            with self.assertRaises(ValueError):ruta_contenida(base/'enlace'/'archivo.json',base)
