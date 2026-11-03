"""Verifica el botón real y el flujo de revisión en segundo plano con archivos temporales."""
import hashlib
import json
from pathlib import Path
import tempfile
import time
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import interfaz_aplicacion_especificaciones as aplicacion
from menu_principal import MenuPrincipal
from rutas_proyecto import rutas_configuradas
from reconstruir_base_datos import EVENTO_CATALOGO
from test_duplicados_especificaciones import datos


class RevisionDuplicadosTkTests(unittest.TestCase):
    def setUp(self):
        try:
            self.raiz = aplicacion.ui.Raiz()
        except aplicacion.tk.TclError as exc:
            self.skipTest(f"Tcl/Tk no disponible: {exc}")
        self.raiz.withdraw()
        self.addCleanup(self.raiz.destroy)

    def test_boton_junto_a_reconstruir_con_mismo_estilo_y_hover(self):
        menu = MenuPrincipal.__new__(MenuPrincipal)
        menu.raiz = self.raiz
        pagina = aplicacion.ui.ctk.CTkFrame(self.raiz)
        menu.pestanas = SimpleNamespace(tab=lambda _titulo: pagina)
        menu.app = SimpleNamespace(revisar_duplicados=Mock(), reconstruir_base_datos=Mock(), actualizar_listas_json=Mock())
        menu.personalizar()
        def botones(widget):
            resultado = []
            if isinstance(widget, aplicacion.ui.ctk.CTkButton):
                resultado.append(widget)
            for hijo in widget.winfo_children():
                resultado.extend(botones(hijo))
            return resultado
        encontrados = {b.cget("text"): b for b in botones(pagina)}
        refrescar = encontrados["⭮"]
        self.assertEqual(refrescar.cget("width"), refrescar.cget("height"))
        self.assertEqual(refrescar.cget("fg_color"), aplicacion.ui.paleta()["acento"])
        self.assertIn("JSON añadidos manualmente", refrescar._ayuda_hover.texto)
        refrescar.invoke()
        menu.app.actualizar_listas_json.assert_called_once_with()
        self.raiz.aplicar_estilo()
        self.assertEqual(refrescar.cget("fg_color"), aplicacion.ui.paleta()["acento"])
        self.assertEqual(refrescar.cget("font")[0], "Segoe UI Symbol")
        revisar = encontrados["Revisar duplicados"]
        reconstruir = encontrados["Reconstruir base desde JSON"]
        self.assertEqual(refrescar.cget("height"), reconstruir.cget("height"))
        self.assertEqual(refrescar.cget("corner_radius"), reconstruir.cget("corner_radius"))
        densidad = aplicacion.ui.PREFERENCIAS["densidad"]
        try:
            for valor in ("Compacta", "Cómoda"):
                aplicacion.ui.PREFERENCIAS["densidad"] = valor
                self.raiz.aplicar_estilo()
                self.assertEqual(refrescar.cget("width"), refrescar.cget("height"))
                self.assertEqual(refrescar.cget("height"), reconstruir.cget("height"))
        finally:
            aplicacion.ui.PREFERENCIAS["densidad"] = densidad
        self.assertNotIn("Restaurar valores", encontrados)
        self.assertEqual(revisar.master, reconstruir.master)
        self.assertEqual(revisar.cget("fg_color"), reconstruir.cget("fg_color"))
        reconstruir.invoke()
        menu.app.reconstruir_base_datos.assert_called_once_with()
        self.assertIn("REVISAR/DUPLICADOS", revisar._ayuda_hover.texto)
        revisar.invoke()
        menu.app.revisar_duplicados.assert_called_once_with()

    def test_recarga_detecta_json_anadido_manualmente_en_ventana_abierta(self):
        from reconstruir_base_datos import suscribir_catalogo
        from interfaz_seleccion_especificaciones import cargar_catalogo
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            catalogos = []
            self.raiz.update()
            suscribir_catalogo(self.raiz, lambda: catalogos.append(cargar_catalogo(base)))
            self.assertFalse(cargar_catalogo(base).especificaciones)
            nuevo = base / "nuevo.json"
            nuevo.write_text(json.dumps(datos()), encoding="utf-8")
            app = SimpleNamespace(raiz=self.raiz, _estado_progreso_carga=Mock())
            aplicacion.AplicacionEspecificaciones.actualizar_listas_json(app)
            self.raiz.update()
            self.assertEqual(len(catalogos), 1)
            self.assertEqual(len(catalogos[0].especificaciones), 1)
            self.assertTrue(catalogos[0].especificaciones[0].ruta.samefile(nuevo))

    def test_revision_desde_interfaz_actualiza_resultado_y_cierra_progreso(self):
        errores_tk = []
        self.raiz.report_callback_exception = lambda *error: errores_tk.append(error)
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta).resolve()
            rutas = rutas_configuradas(base)
            rutas["json_anteriores"].mkdir()
            rutas["especificaciones_viejas"].mkdir()
            for nombre in ("a", "b"):
                word = rutas["especificaciones_viejas"] / (nombre + ".docx")
                word.write_bytes(nombre.encode())
                documento = datos(archivo_origen=word.relative_to(base).as_posix(),
                                  hash_sha256_origen=hashlib.sha256(word.read_bytes()).hexdigest())
                (rutas["json_anteriores"] / (nombre + ".json")).write_text(json.dumps(documento), encoding="utf-8")
            app = aplicacion.AplicacionEspecificaciones.__new__(aplicacion.AplicacionEspecificaciones)
            app.raiz, app.rutas = self.raiz, rutas
            app._reconstruccion_en_proceso = app._carga_en_proceso = False
            app._acciones_tras_recuperacion = []
            app.menu_principal = SimpleNamespace(creadas=set())
            with patch.object(aplicacion, "BASE", base), patch.object(aplicacion.messagebox, "showinfo") as aviso, patch.object(aplicacion.messagebox, "showerror") as error:
                app.revisar_duplicados()
                for ventana in self.raiz.winfo_children():
                    if isinstance(ventana, aplicacion.tk.Toplevel):
                        ventana.withdraw()
                limite = time.monotonic() + 10
                while app._reconstruccion_en_proceso and time.monotonic() < limite:
                    self.raiz.update()
                    time.sleep(.02)
                self.assertFalse(app._reconstruccion_en_proceso)
                error.assert_not_called()
                self.assertEqual(aviso.call_args.args[0], "Revisión de duplicados terminada")
                self.assertIn("1 JSON y 1 Word", aviso.call_args.args[1])
            self.assertTrue((base / "REVISAR" / "DUPLICADOS" / "b.json").is_file())
            self.assertEqual(errores_tk, [])


class RevisionDuplicadosFlujoTests(unittest.TestCase):
    def test_accion_del_menu_ejecuta_revision_y_muestra_resumen(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta).resolve()
            rutas = rutas_configuradas(base)
            rutas["json_anteriores"].mkdir()
            for nombre in ("a.json", "b.json"):
                (rutas["json_anteriores"] / nombre).write_text(json.dumps(datos()), encoding="utf-8")
            app = aplicacion.AplicacionEspecificaciones.__new__(aplicacion.AplicacionEspecificaciones)
            app.raiz, app.rutas = Mock(), rutas
            app._reconstruccion_en_proceso = app._carga_en_proceso = False
            app._acciones_tras_recuperacion = []
            app.menu_principal = SimpleNamespace(creadas=set())
            pendientes = []
            app.raiz.after.side_effect = lambda _tiempo, accion: pendientes.append(accion)
            ventana = Mock()
            with (patch.object(aplicacion, "BASE", base),
                  patch.object(aplicacion.ui, "Ventana", return_value=ventana),
                  patch.object(aplicacion.ttk, "Label"), patch.object(aplicacion.ttk, "Progressbar"),
                  patch.object(aplicacion.threading, "Thread", side_effect=lambda target, daemon: SimpleNamespace(start=target)),
                  patch.object(aplicacion.messagebox, "showinfo") as aviso,
                  patch.object(aplicacion.messagebox, "showerror") as error):
                app.revisar_duplicados()
                while pendientes:
                    pendientes.pop(0)()
                self.assertFalse(app._reconstruccion_en_proceso)
                error.assert_not_called()
                self.assertIn("1 JSON y 0 Word", aviso.call_args.args[1])
                ventana.destroy.assert_called_once()
                app.raiz.event_generate.assert_called_once_with(EVENTO_CATALOGO, when="tail")
            self.assertTrue((base / "REVISAR" / "DUPLICADOS" / "b.json").is_file())


if __name__ == "__main__":
    unittest.main()
