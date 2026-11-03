from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from interfaz_aplicacion_especificaciones import GeneradorEspecificacion, AplicacionEspecificaciones


class TransicionEditorHiiTests(unittest.TestCase):
    def test_selector_pasa_sus_marcas_al_generador_del_menu_principal(self):
        app = AplicacionEspecificaciones.__new__(AplicacionEspecificaciones)
        app.raiz = Mock()
        app.rutas = {"formatos": Path("formatos"), "json_alineadas": Path("json"),
                     "json_anteriores": Path("originales")}
        generador = Mock()
        generador._cerrada = False
        app._generador_hii = generador
        rutas = (Path("primera.json"), Path("segunda.json"))
        with patch("interfaz_aplicacion_especificaciones.ui.Ventana"), patch(
            "interfaz_seleccion_especificaciones.SelectorEspecificaciones"
        ) as selector:
            app.seleccionar_especificacion()
            selector.call_args.kwargs["al_generar_hii"](rutas)
        generador.seleccionar_json_guardados.assert_called_once_with(rutas)
        generador.ventana.lift.assert_called()

    def test_cierra_editor_despues_de_entregar_todos_los_json(self):
        g = GeneradorEspecificacion.__new__(GeneradorEspecificacion)
        rutas = (Path("es.json"), Path("en.json"))
        acciones = []
        g.app = SimpleNamespace(generar_hii=Mock(side_effect=lambda **kw: acciones.append(kw["rutas_json"])))
        g.ventana = Mock()
        g.ventana.destroy.side_effect = lambda: acciones.append("cerrado")
        g._finalizar_guardado_json(rutas)
        self.assertEqual(acciones, [rutas, "cerrado"])

    def test_fallo_al_abrir_hii_conserva_editor_y_reporta_que_json_ya_se_guardo(self):
        g = GeneradorEspecificacion.__new__(GeneradorEspecificacion)
        g.app = SimpleNamespace(generar_hii=Mock(side_effect=RuntimeError("Error de ventana")))
        g.ventana = Mock()
        with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
            g._finalizar_guardado_json((Path("es.json"),))
            mensajes.showwarning.assert_called_once()
            self.assertIn("se guardaron", mensajes.showwarning.call_args.args[1])
        g.ventana.destroy.assert_not_called()

    def test_traduccion_regresa_al_generador_que_la_abrio(self):
        g = GeneradorEspecificacion.__new__(GeneradorEspecificacion)
        g.app = Mock()
        g.ventana = Mock()
        g.al_finalizar = Mock()
        rutas = (Path("en.json"),)
        g._finalizar_guardado_json(rutas)
        g.al_finalizar.assert_called_once_with(rutas)
        g.app.generar_hii.assert_not_called()
        g.ventana.destroy.assert_called_once()

    def test_menu_reutiliza_hii_abierto_y_recrea_uno_cerrado(self):
        app = AplicacionEspecificaciones.__new__(AplicacionEspecificaciones)
        app.raiz = Mock()
        app.rutas = {"formatos": Path("formatos"), "json_alineadas": Path("json")}
        abierto, reemplazo = Mock(), Mock()
        abierto._cerrada = reemplazo._cerrada = False
        app._generador_hii = abierto
        rutas = (Path("es.json"),)
        with patch("generador_hii.GeneradorHII", return_value=reemplazo) as crear:
            self.assertIs(app.generar_hii(rutas_json=rutas), abierto)
            abierto.seleccionar_json_guardados.assert_called_once_with(rutas)
            crear.assert_not_called()
            abierto._cerrada = True
            self.assertIs(app.generar_hii(rutas_json=rutas), reemplazo)
            reemplazo.seleccionar_json_guardados.assert_called_once_with(rutas)
            self.assertEqual(crear.call_count, 1)
