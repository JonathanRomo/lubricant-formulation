import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from openpyxl import Workbook
from menu_principal import fecha_actualizacion, _fecha_registrada, MenuPrincipal
from interfaz_aplicacion_especificaciones import _guardar_cambios_maestro, COLUMNAS_MAESTRO
from unificar_metodos_lab_v3 import ejecutar, SALIDA, ESTADO


class FechasListasTests(unittest.TestCase):
    def test_fecha_estado_cache_y_formato_local(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "estado.json"
            ruta.write_text(json.dumps({"ultima_modificacion": "2026-10-14T16:35:00+00:00"}), encoding="utf-8")
            self.assertEqual(fecha_actualizacion(carpeta, "estado.json", "historial.txt"), "Última actualización: 14/10/2026 · 10:35")
            with patch.object(Path, "read_text", side_effect=AssertionError("No releer estado sin cambios")):
                self.assertIn("10:35", fecha_actualizacion(carpeta, "estado.json", "historial.txt"))
            ruta.write_text(json.dumps({"ultima_modificacion": "2026-10-15T11:00:00-06:00"}), encoding="utf-8")
            self.assertIn("15/10/2026 · 11:00", fecha_actualizacion(carpeta, "estado.json", "historial.txt"))

    def test_historial_heredado_y_registro_ausente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            self.assertEqual(fecha_actualizacion(carpeta, "estado.json", "historial.txt"), "Sin actualización registrada")
            (Path(carpeta) / "historial.txt").write_text("[2026-10-09T14:23:35]\nCambio\n[2026-10-11T22:19:32]\nOtro\n", encoding="utf-8")
            self.assertIn("11/10/2026 · 22:19", fecha_actualizacion(carpeta, "estado.json", "historial.txt"))

    def test_edicion_maestro_registra_cambios_y_no_avanza_si_son_iguales(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / SALIDA
            libro = Workbook()
            hoja = libro.active
            hoja.title = "METODOS"
            hoja.append(list(COLUMNAS_MAESTRO))
            valores = tuple("A" for _ in COLUMNAS_MAESTRO)
            hoja.append(valores)
            libro.save(ruta)
            libro.close()
            estado = Path(carpeta) / ESTADO
            estado.write_text(json.dumps({"archivos": {"fuente": {"hash": "h"}}}), encoding="utf-8")
            _guardar_cambios_maestro(ruta, {2: valores})
            self.assertNotIn("ultima_modificacion", json.loads(estado.read_text(encoding="utf-8")))
            nuevos = ("B", *valores[1:])
            _guardar_cambios_maestro(ruta, {2: nuevos})
            registrado = estado.read_bytes()
            self.assertIn("ultima_modificacion", json.loads(registrado))
            self.assertIn("fuente", json.loads(registrado)["archivos"])
            _guardar_cambios_maestro(ruta, {2: nuevos})
            self.assertEqual(estado.read_bytes(), registrado)

    def test_reconstruccion_sin_cambio_conserva_fecha(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            (base / "METODOS_LAB").mkdir()
            estado = base / "METODOS_LAB" / ESTADO
            estado.write_text(json.dumps({"archivos": {}, "ultima_modificacion": "2026-10-01T10:00:00-06:00"}), encoding="utf-8")
            with patch("unificar_metodos_lab_v3.archivos", return_value=[]), patch("unificar_metodos_lab_v3.metodos_ingles_existentes", return_value=({}, {})), patch("unificar_metodos_lab_v3.filas_maestro_actual", return_value=[]), patch("unificar_metodos_lab_v3.crear"), patch("builtins.print"):
                ejecutar(base)
            self.assertEqual(json.loads(estado.read_text(encoding="utf-8"))["ultima_modificacion"], "2026-10-01T10:00:00-06:00")

    def test_tarjetas_se_apilan_y_no_reorganizan_sin_cambio_de_ancho(self):
        menu = MenuPrincipal.__new__(MenuPrincipal)
        menu._pagina_metodos = SimpleNamespace(_get_widget_scaling=lambda: 1)
        menu._rejilla_metodos = Mock()
        menu._encabezado_metodos = Mock()
        menu._tarjetas_metodos = [(Mock(), Mock()), (Mock(), Mock())]
        menu._disposicion_metodos = menu._margen_metodos = menu._wrap_metodos = None
        menu._adaptar_tarjetas_metodos(SimpleNamespace(width=1200))
        self.assertEqual(menu._tarjetas_metodos[1][0].grid.call_args.kwargs["column"], 1)
        menu._adaptar_tarjetas_metodos(SimpleNamespace(width=600))
        self.assertEqual(menu._tarjetas_metodos[1][0].grid.call_args.kwargs["row"], 1)
        llamadas = menu._rejilla_metodos.mock_calls[:]
        menu._adaptar_tarjetas_metodos(SimpleNamespace(width=600))
        self.assertEqual(menu._rejilla_metodos.mock_calls, llamadas)


if __name__ == "__main__":
    unittest.main()
