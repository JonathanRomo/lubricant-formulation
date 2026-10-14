import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
from rutas_listas import carpeta_listas, ruta_maestro
from traductor_especificaciones import ServicioTraductor, ARCHIVO_EXCEL
from interfaz_aplicacion_especificaciones import _guardar_configuracion, _cargar_configuracion
from alinear_especificaciones_json import ejecutar


class RutasListasTests(unittest.TestCase):
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
