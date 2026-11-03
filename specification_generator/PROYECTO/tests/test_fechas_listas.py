import json
import os
import tempfile
import unittest
from pathlib import Path
from datetime import datetime
from types import SimpleNamespace
from unittest.mock import Mock, patch

from openpyxl import Workbook
from menu_principal import fecha_actualizacion, fecha_actualizacion_maestro, MenuPrincipal
from interfaz_aplicacion_especificaciones import _guardar_cambios_maestro, _leer_filas_maestro_editable, COLUMNAS_MAESTRO


class FechasListasTests(unittest.TestCase):
    def test_fecha_estado_cache_y_formato_local(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "estado.json"
            ruta.write_text(json.dumps({"ultima_modificacion": "2026-10-14T16:35:00+00:00"}), encoding="utf-8")
            self.assertEqual(fecha_actualizacion(carpeta, "estado.json", "historial.txt"), "Última actualización: 14/10/2026 · 10:35")
            with patch.object(Path, "read_text", side_effect=AssertionError("No releer estado sin cambios")):
                self.assertIn("10:35", fecha_actualizacion(carpeta, "estado.json", "historial.txt"))
            estado_anterior = ruta.stat()
            ruta.write_text(json.dumps({"ultima_modificacion": "2026-10-15T11:00:00-06:00"}), encoding="utf-8")
            # Garantiza una modificación detectable incluso con marcas de tiempo poco precisas.
            os.utime(ruta, ns=(estado_anterior.st_atime_ns, estado_anterior.st_mtime_ns + 1_000_000_000))
            self.assertIn("15/10/2026 · 11:00", fecha_actualizacion(carpeta, "estado.json", "historial.txt"))

    def test_historial_heredado_y_registro_ausente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            self.assertEqual(fecha_actualizacion(carpeta, "estado.json", "historial.txt"), "Sin actualización registrada")
            (Path(carpeta) / "historial.txt").write_text("[2026-10-09T14:23:35]\nCambio\n[2026-10-11T22:19:32]\nOtro\n", encoding="utf-8")
            self.assertIn("11/10/2026 · 22:19", fecha_actualizacion(carpeta, "estado.json", "historial.txt"))

    def test_edicion_manual_conserva_datos_y_no_reescribe_sin_cambios(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "METODOS_LAB_UNIFICADO.xlsx"
            libro = Workbook()
            hoja = libro.active
            hoja.title = "METODOS"
            hoja.append(list(COLUMNAS_MAESTRO))
            valores = tuple("A" for _ in COLUMNAS_MAESTRO)
            hoja.append(valores)
            libro.save(ruta)
            libro.close()
            anterior = ruta.read_bytes(), ruta.stat().st_mtime_ns
            _guardar_cambios_maestro(ruta, {2: valores})
            self.assertEqual((ruta.read_bytes(), ruta.stat().st_mtime_ns), anterior)
            nuevos = ("B", *valores[1:])
            _guardar_cambios_maestro(ruta, {2: nuevos})
            self.assertEqual(_leer_filas_maestro_editable(ruta), [(2, nuevos)])
            actualizado = ruta.read_bytes(), ruta.stat().st_mtime_ns
            _guardar_cambios_maestro(ruta, {2: nuevos})
            self.assertEqual((ruta.read_bytes(), ruta.stat().st_mtime_ns), actualizado)
            self.assertEqual(list(Path(carpeta).iterdir()), [ruta])

    def test_fecha_maestro_refleja_ediciones_externas_sin_leer_el_excel(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "METODOS_LAB_UNIFICADO.xlsx"
            self.assertEqual(fecha_actualizacion_maestro(ruta), "Lista maestra no disponible")
            ruta.write_bytes(b"No es necesario interpretar este archivo")
            for fecha, esperado in (("2026-10-14T16:35:00+00:00", "14/10/2026 · 10:35"),
                                    ("2026-10-15T17:00:00+00:00", "15/10/2026 · 11:00")):
                instante = datetime.fromisoformat(fecha).timestamp()
                os.utime(ruta, (instante, instante))
                with patch.object(Path, "open", side_effect=AssertionError("No abrir el Excel")):
                    self.assertEqual(fecha_actualizacion_maestro(ruta), "Última actualización: " + esperado)

    def test_menu_consulta_el_excel_configurado_y_conserva_fechas_de_traducciones(self):
        menu = MenuPrincipal.__new__(MenuPrincipal)
        menu.app = SimpleNamespace(rutas={"metodos_lab": Path("listas_manuales")})
        menu._fechas_metodos = {"metodos": Mock(), "traducciones": Mock()}
        with patch("menu_principal.fecha_actualizacion_maestro", return_value="Fecha del Excel") as maestro, \
             patch("menu_principal.fecha_actualizacion", return_value="Fecha de traducciones") as traducciones:
            menu.refrescar_fechas_metodos()
        maestro.assert_called_once_with(Path("listas_manuales/METODOS_LAB_UNIFICADO.xlsx"))
        traducciones.assert_called_once_with(Path("listas_manuales"), ".estado_traductor.json", "CAMBIOS_TRADUCTOR.txt")
        menu._fechas_metodos["metodos"].configure.assert_called_once_with(text="Fecha del Excel")
        menu._fechas_metodos["traducciones"].configure.assert_called_once_with(text="Fecha de traducciones")

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
