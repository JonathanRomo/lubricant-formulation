import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from interfaz_aplicacion_especificaciones import AplicacionEspecificaciones


class ErroresCallbackTests(unittest.TestCase):
    def app(self):
        raiz = Mock()
        raiz.tk.call.return_value = 1.0
        with patch("interfaz_aplicacion_especificaciones._cargar_configuracion", return_value={}), \
             patch.object(AplicacionEspecificaciones, "_crear_interfaz"), \
             patch("reconstruir_base_datos.necesita_recuperacion", return_value=False):
            app = AplicacionEspecificaciones(raiz)
        app._mostrar_error_callback = Mock()
        return app

    def fallo(self, texto="fallo de prueba"):
        try:
            raise ValueError(texto)
        except ValueError as error:
            return type(error), error, error.__traceback__

    def test_registra_manejador_y_conserva_solo_ultimo_diagnostico_acotado(self):
        app = self.app()
        self.assertEqual(app.raiz.report_callback_exception, app._informar_error_callback)
        app.raiz.report_callback_exception(*self.fallo())
        self.assertIn("ValueError: fallo de prueba", app._ultimo_error_callback)
        self.assertIn("Aplicación de especificaciones v", app._ultimo_error_callback)
        app.raiz.report_callback_exception(*self.fallo("x" * 40000))
        self.assertLess(len(app._ultimo_error_callback), 16500)
        self.assertNotIn("fallo de prueba", app._ultimo_error_callback)
        self.assertFalse(app._notificando_error_callback)

    def test_no_repite_avisos_y_permite_otro_error_o_reintento_posterior(self):
        app = self.app()
        error = self.fallo()
        with patch("interfaz_aplicacion_especificaciones.time.monotonic", side_effect=[10, 11, 12, 20]):
            app._informar_error_callback(*error)
            app._informar_error_callback(*error)
            self.assertEqual(app._mostrar_error_callback.call_count, 1)
            app._informar_error_callback(*self.fallo("otro problema"))
            self.assertEqual(app._mostrar_error_callback.call_count, 2)
            app._informar_error_callback(*error)
            self.assertEqual(app._mostrar_error_callback.call_count, 3)

    def test_reutiliza_aviso_abierto_y_actualiza_detalles(self):
        app = self.app()
        ventana = app._ventana_error_callback = Mock()
        app._informar_error_callback(*self.fallo("nuevo problema"))
        app._mostrar_error_callback.assert_not_called()
        self.assertIs(app._ventana_error_callback, ventana)
        self.assertIn("nuevo problema", app._ultimo_error_callback)

    def test_error_del_aviso_no_causa_recursion_ni_impide_errores_posteriores(self):
        app = self.app()
        def fallar():
            app._informar_error_callback(*self.fallo("anidado"))
            raise RuntimeError("fallo al dibujar")
        app._mostrar_error_callback.side_effect = fallar
        with patch("interfaz_aplicacion_especificaciones.messagebox.showerror") as aviso:
            app._informar_error_callback(*self.fallo())
            aviso.assert_called_once()
        self.assertFalse(app._notificando_error_callback)
        self.assertNotIn("anidado", app._ultimo_error_callback)

    def test_no_crea_ventana_si_la_raiz_ya_se_cerro(self):
        app = self.app()
        app.raiz.winfo_exists.return_value = False
        app._informar_error_callback(*self.fallo())
        app._mostrar_error_callback.assert_not_called()


class RecuperacionInterfazTests(unittest.TestCase):
    def test_actualizar_listas_notifica_a_ventanas_sin_reconstruir_base(self):
        app = SimpleNamespace(raiz=Mock(), _estado_progreso_carga=Mock())
        with patch("interfaz_aplicacion_especificaciones.notificar_catalogo") as notificar:
            AplicacionEspecificaciones.actualizar_listas_json(app)
        notificar.assert_called_once_with(app.raiz)
        app._estado_progreso_carga.configure.assert_called_once_with(text="Listas actualizadas.")

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
        resultado = dict(documentos=20, versiones=20, relaciones=2, wip=1, omitidos=[], respaldo=None,
                         duplicados=dict(json_duplicados=0, word_duplicados=0, advertencias=[]))
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
