import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from interfaz_aplicacion_especificaciones import AplicacionEspecificaciones


class RecuperacionInterfazTests(unittest.TestCase):
    def ejecutar(self, resultado=None, error=None):
        raiz = Mock()
        raiz.winfo_exists.return_value = True
        accion = Mock()
        app = SimpleNamespace(raiz=raiz, rutas={"json_anteriores": "originales", "json_alineadas": "alineadas"},
            _reconstruccion_en_proceso=False, _carga_en_proceso=False,
            _acciones_tras_recuperacion=[accion], menu_principal=SimpleNamespace(creadas=set()))
        with patch("interfaz_aplicacion_especificaciones.ui.Ventana") as ventana, \
             patch("interfaz_aplicacion_especificaciones.ttk.Label"), \
             patch("interfaz_aplicacion_especificaciones.ttk.Progressbar"), \
             patch("interfaz_aplicacion_especificaciones.threading.Thread") as hilo, \
             patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, \
             patch("reconstruir_base_datos.asegurar_base", return_value=resultado, side_effect=error) as recuperar:
            AplicacionEspecificaciones.reconstruir_base_datos(app, automatica=True)
            self.assertTrue(app._reconstruccion_en_proceso)
            self.assertFalse(recuperar.called)
            llamadas_tk = len(ventana.return_value.mock_calls)
            hilo.call_args.kwargs["target"]()
            self.assertEqual(len(ventana.return_value.mock_calls), llamadas_tk, "El trabajador no debe tocar Tk.")
            self.assertFalse(accion.called)
            raiz.after.call_args.args[1]()
            self.assertFalse(app._reconstruccion_en_proceso)
            ventana.return_value.grab_release.assert_called_once()
            ventana.return_value.destroy.assert_called_once()
            recuperar.assert_called_once()
            return accion, mensajes.showinfo.called, mensajes.showerror.called

    def test_recuperacion_automatica_termina_en_hilo_principal_y_reanuda_accion(self):
        resultado = dict(documentos=20, versiones=20, relaciones=2, wip=1, omitidos=[], respaldo=None)
        accion, informacion, error = self.ejecutar(resultado=resultado)
        accion.assert_called_once()
        self.assertTrue(informacion)
        self.assertFalse(error)

    def test_error_desbloquea_interfaz_y_no_abre_editor_pendiente(self):
        accion, informacion, error = self.ejecutar(error=OSError("disco lleno"))
        accion.assert_not_called()
        self.assertFalse(informacion)
        self.assertTrue(error)


if __name__ == "__main__":
    unittest.main()
