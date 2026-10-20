import sqlite3
import tempfile
import unittest
from threading import Thread as HiloReal, get_ident
from dataclasses import replace
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from generador_hii import GeneradorHII, FILAS_POR_BLOQUE, _trabajar_excel, _procesar_trabajos_excel
from interfaz_seleccion_especificaciones import Catalogo, Especificacion


class RelojTk:
    def __init__(self):
        self.pendientes = {}
        self.contador = 0
        self.ventana = Mock()
        self.ventana.after.side_effect = self.after
        self.ventana.after_cancel.side_effect = lambda clave: self.pendientes.pop(clave, None)

    def after(self, demora, accion):
        self.contador += 1
        clave = str(self.contador)
        self.pendientes[clave] = (demora, accion)
        return clave

    def ejecutar(self, clave):
        _, accion = self.pendientes.pop(clave)
        accion()


class Lista:
    def __init__(self):
        self.filas = {}
        self.seleccion = ()

    def insert(self, _padre, _posicion, *, iid, values):
        if iid in self.filas:
            raise AssertionError("Fila duplicada")
        self.filas[iid] = values

    def delete(self, *filas):
        for fila in filas:
            self.filas.pop(fila)

    def get_children(self):
        return tuple(self.filas)

    def item(self, fila, opcion=None, *, values=None):
        if values is not None:
            self.filas[fila] = tuple(values)
        return self.filas[fila] if opcion == "values" else None

    def selection(self):
        return self.seleccion


class Divisor:
    def __init__(self):
        self.paneles = []
        self.pack = Mock()
        self.pack_forget = Mock()
        self.sashpos = Mock()

    def panes(self):
        return tuple(self.paneles)

    def forget(self, panel):
        self.paneles.remove(panel)

    def add(self, panel, **_):
        self.paneles.append(panel)

    def winfo_width(self):
        return 1000


class CargaGeneradorTests(unittest.TestCase):
    def setUp(self):
        temporal = tempfile.TemporaryDirectory()
        self.addCleanup(temporal.cleanup)
        self.base = Path(temporal.name).resolve()
        self.reloj = RelojTk()
        with patch("generador_hii.ui.Ventana", return_value=self.reloj.ventana), \
             patch("generador_hii.tk.Frame"), patch("generador_hii.tk.Label"), \
             patch("generador_hii.tk.StringVar"), patch("generador_hii.tk.BooleanVar") as ambos, patch("generador_hii.ttk"), \
             patch("generador_hii.cargar_catalogo_hii", side_effect=AssertionError("El constructor no debe cargar")), \
             patch("generador_hii.conectar", side_effect=AssertionError("No abrir SQLite")):
            ambos.return_value.get.return_value = False
            self.g = GeneradorHII(None, carpeta_json=self.base, carpeta_formatos=self.base)
        self.g.lista = Lista()
        self.g.consulta = Mock(get=Mock(return_value=""))
        self.g.estado = Mock()
        self.g.contador = Mock()
        self.g.actualizar_boton = Mock()
        self.g.lote_boton = Mock()

    def catalogo(self, cantidad):
        return Catalogo(tuple(Especificacion(self.base / f"{i}.json", f"Producto {i}", f"P{i}", "Cliente", "01",
                                            f"PRODUCTO {i} P{i}") for i in range(cantidad)), ())

    def preparar_vistas(self):
        self.g.paneles = Divisor()
        self.g.izquierda, self.g.derecha = Mock(), Mock()
        for panel in (self.g.izquierda, self.g.derecha):
            panel.winfo_children.return_value = []
        self.g.botones_ampliar_vista = {c: Mock() for c in ("tabla", "excel")}
        self.g.etiquetas_vista = {c: (Mock(), c) for c in ("tabla", "excel")}
        self.g.mostrar_especificacion = Mock()
        self.g.mostrar_excel = Mock()
        self.g._actualizar_idiomas = Mock()
        self.g.actual = self.base / "es.json"
        self.g.datos = SimpleNamespace(codigo="P1", idioma="ESPANOL", reanalisis="12 meses")
        self.g.plantilla_actual = self.base / "formato.xlsm"
        self.g.plantilla_actual.write_bytes(b"plantilla")
        self.g._vistas_renderizadas.clear()
        self.hilos_excel = self.enterContext(patch("generador_hii.threading.Thread"))
        self.enterContext(patch("generador_hii.ttk.Label"))
        self.enterContext(patch("generador_hii.preparar_dibujo_excel", return_value=((0, 0, 600, 1000), ())))

    def ejecutar_trabajador_excel(self):
        trabajo = self.g._pendientes_excel.get_nowait()
        _trabajar_excel(trabajo, self.g._cola_excel)

    def terminar_excel(self):
        self.ejecutar_trabajador_excel()
        self.g._recibir_excel()

    def test_excel_inicia_oculto_y_no_se_construye_al_seleccionar(self):
        self.assertEqual(self.g.vistas_visibles, {"tabla": True, "excel": False})
        self.preparar_vistas()
        with patch("generador_hii.preparar_seleccion", return_value=(self.g.datos, Mock())), patch(
            "generador_hii.seleccionar_plantilla", return_value=self.g.plantilla_actual
        ), patch("generador_hii.construir_excel") as construir:
            self.g.cargar_previa(self.g.actual)
            construir.assert_not_called()
        self.g.mostrar_especificacion.assert_called_once()
        self.g.mostrar_excel.assert_not_called()
        self.assertIsNone(self.g.excel)

    def test_abrir_excel_construye_al_ocultarlo_libera_y_al_reabrir_reconstruye(self):
        self.preparar_vistas()
        with patch("generador_hii.construir_excel", return_value=b"excel") as construir:
            self.g.alternar_vista("excel")
            construir.assert_not_called()
            self.terminar_excel()
            self.assertEqual(construir.call_count, 1)
            self.g.alternar_vista("tabla")
            self.assertEqual(construir.call_count, 1)
            self.g.imagenes = [Mock()]
            self.g.alternar_vista("excel")
            self.assertIsNone(self.g.excel)
            self.assertEqual(self.g.imagenes, [])
            self.g.alternar_vista("excel")
            self.terminar_excel()
            self.assertEqual(construir.call_count, 2)

    def test_ampliar_cualquiera_y_revertir_recupera_disposicion(self):
        self.preparar_vistas()
        self.g.vistas_visibles = {"tabla": True, "excel": True}
        with patch("generador_hii.construir_excel", return_value=b"excel"):
            for clave in ("tabla", "excel"):
                self.g.ampliar_vista(clave)
                self.assertEqual(self.g.vistas_visibles, {c: c == clave for c in ("tabla", "excel")})
                self.assertEqual(len(self.g.paneles.panes()), 1)
                self.g.botones_ampliar_vista[clave].configure.assert_called_with(text="−")
                self.g.ampliar_vista(clave)
                self.assertTrue(all(self.g.vistas_visibles.values()))
                self.assertEqual(len(self.g.paneles.panes()), 2)
                self.g.botones_ampliar_vista[clave].configure.assert_called_with(text="+")

    def test_puede_ocultar_ambas_y_abrir_cualquiera_sin_sash_inexistente(self):
        self.preparar_vistas()
        self.g.alternar_vista("tabla")
        self.assertFalse(any(self.g.vistas_visibles.values()))
        self.assertEqual(self.g.paneles.panes(), ())
        self.g.paneles.pack_forget.assert_called_once()
        self.g._centrar_vistas()
        self.g.paneles.sashpos.assert_not_called()
        self.g.alternar_vista("tabla")
        self.assertEqual(self.g.paneles.panes(), (self.g.izquierda,))
        self.g.mostrar_especificacion.assert_called_once()
        self.g.mostrar_excel.assert_not_called()

    def test_ampliar_excel_y_restaurar_recupera_excel_oculto_inicial(self):
        self.preparar_vistas()
        with patch("generador_hii.construir_excel", return_value=b"excel"):
            self.g.ampliar_vista("excel")
            self.assertEqual(self.g.vistas_visibles, {"tabla": False, "excel": True})
            self.g.ampliar_vista("excel")
            self.assertEqual(self.g.vistas_visibles, {"tabla": True, "excel": False})
            self.g.ampliar_vista("tabla")
            self.g.alternar_vista("tabla")
            self.assertFalse(any(self.g.vistas_visibles.values()))
            self.assertIsNone(self.g._vista_ampliada)

    def test_ambas_ocultas_no_dibujan_al_cambiar_especificacion(self):
        self.preparar_vistas()
        self.g.vistas_visibles = {"tabla": False, "excel": False}
        with patch("generador_hii.preparar_seleccion", return_value=(self.g.datos, Mock())), patch(
            "generador_hii.seleccionar_plantilla", return_value=self.g.plantilla_actual
        ), patch("generador_hii.construir_excel") as construir:
            self.g.cargar_previa(self.g.actual)
            construir.assert_not_called()
        self.g.mostrar_especificacion.assert_not_called()
        self.g.mostrar_excel.assert_not_called()

    def test_guardar_con_excel_oculto_genera_archivo_sin_dibujarlo(self):
        self.preparar_vistas()
        self.g.guardar_archivo = Mock(return_value=None)
        with patch("generador_hii.filedialog.asksaveasfilename", return_value=str(self.base / "salida.xlsm")), patch(
            "generador_hii.construir_excel", return_value=b"excel"
        ) as construir, patch("generador_hii.messagebox") as mensajes:
            self.g.guardar()
            construir.assert_not_called()
            self.terminar_excel()
            construir.assert_called_once_with(self.g.plantilla_actual, self.g.datos)
            mensajes.showerror.assert_not_called()
        self.assertEqual(self.g.guardar_archivo.call_args.args[-1], b"excel")
        self.g.mostrar_excel.assert_not_called()

    def test_fallo_de_generacion_diferida_al_guardar_muestra_error(self):
        self.preparar_vistas()
        self.g.guardar_archivo = Mock()
        with patch("generador_hii.filedialog.asksaveasfilename", return_value=str(self.base / "salida.xlsm")), patch(
            "generador_hii.construir_excel", side_effect=RuntimeError("Plantilla inválida")
        ), patch("generador_hii.messagebox") as mensajes:
            self.g.guardar()
            self.terminar_excel()
            mensajes.showerror.assert_called_once()
        self.g.guardar_archivo.assert_not_called()

    def test_trabajador_no_toca_tk_y_comparte_generacion_entre_vista_y_guardado(self):
        self.preparar_vistas()
        callback = Mock()
        self.g.vistas_visibles["excel"] = True
        with patch("generador_hii.construir_excel", return_value=b"excel") as construir:
            self.g._solicitar_excel()
            self.g._solicitar_excel(callback)
            self.assertEqual(self.hilos_excel.call_count, 1)
            llamadas = list(self.g.ventana.mock_calls)
            self.ejecutar_trabajador_excel()
            self.assertEqual(self.g.ventana.mock_calls, llamadas)
            self.g.mostrar_excel.assert_not_called()
            callback.assert_not_called()
            self.g._recibir_excel()
            construir.assert_called_once()
            callback.assert_called_once_with(b"excel")
            self.g.mostrar_excel.assert_called_once()

    def test_trabajador_real_procesa_cola_y_termina_sin_tocar_interfaz(self):
        self.preparar_vistas()
        hilos_usados = []
        callback = Mock()
        with patch("generador_hii.construir_excel", side_effect=lambda *_: hilos_usados.append(get_ident()) or b"excel"):
            self.g._solicitar_excel(callback)
            llamadas = list(self.g.ventana.mock_calls)
            self.g._pendientes_excel.put(None)
            trabajador = HiloReal(target=_procesar_trabajos_excel,
                args=(self.g._pendientes_excel, self.g._cola_excel), daemon=True)
            trabajador.start()
            trabajador.join(timeout=5)
            self.assertFalse(trabajador.is_alive())
            self.assertEqual(self.g.ventana.mock_calls, llamadas)
            self.assertNotEqual(hilos_usados, [get_ident()])
            callback.assert_not_called()
            self.g._recibir_excel()
        callback.assert_called_once_with(b"excel")

    def test_resultado_anterior_no_reabre_excel_oculto(self):
        self.preparar_vistas()
        self.g.vistas_visibles["excel"] = True
        with patch("generador_hii.construir_excel", return_value=b"anterior"):
            self.g._solicitar_excel()
            self.ejecutar_trabajador_excel()
            self.g.alternar_vista("excel")
            self.g._recibir_excel()
        self.assertIsNone(self.g.excel)
        self.assertIsNone(self.g._dibujo_excel)
        self.g.mostrar_excel.assert_not_called()

    def test_guardado_pendiente_conserva_especificacion_elegida_aunque_cambie_seleccion(self):
        self.preparar_vistas()
        origen = self.g.actual
        self.g.guardar_archivo = Mock(return_value=None)
        with patch("generador_hii.filedialog.asksaveasfilename", return_value=str(self.base / "salida.xlsm")), patch(
            "generador_hii.construir_excel", return_value=b"archivo P1"
        ):
            self.g.guardar()
            self.g._invalidar_excel()
            self.g.actual = self.base / "otro.json"
            self.g.datos = SimpleNamespace(codigo="P2", idioma="ESPANOL")
            self.terminar_excel()
        _, guardado, datos, contenido = self.g.guardar_archivo.call_args.args
        self.assertEqual(guardado, origen)
        self.assertEqual(datos.codigo, "P1")
        self.assertEqual(contenido, b"archivo P1")
        self.assertIsNone(self.g.excel)

    def test_cerrar_cancela_trabajos_y_no_ejecuta_callbacks_pendientes(self):
        self.preparar_vistas()
        callback = Mock()
        self.g._solicitar_excel(callback)
        self.g._al_destruir(SimpleNamespace(widget=self.g.ventana))
        with patch("generador_hii.construir_excel") as construir:
            self.terminar_excel()
            construir.assert_not_called()
        callback.assert_not_called()

    def test_dibujo_por_bloques_respeta_zoom_y_puede_cancelarse(self):
        self.preparar_vistas()
        self.g.vistas_visibles["excel"] = True
        self.g.zoom_factor = 1.5
        canvas = SimpleNamespace(_hii_factor=1.5, _hii_textos={}, winfo_exists=lambda: True,
                                 create_text=Mock(side_effect=range(100)), create_rectangle=Mock())
        ordenes = [("text", (20, 30), {"text": str(i), "font": ("Arial", 10), "width": 100}) for i in range(90)]
        revision = self.g._revision_dibujo
        self.g._dibujar_bloque_excel(canvas, ordenes, 0, revision)
        self.assertLessEqual(canvas.create_text.call_count, 40)
        canvas.create_text.assert_any_call(30.0, 45.0, text="0", font=("Arial", 15), width=150.0)
        while "dibujo" in self.g._tareas:
            self.reloj.ejecutar(self.g._tareas["dibujo"])
        self.assertEqual(canvas.create_text.call_count, 90)
        self.g._invalidar_excel()
        self.g._dibujar_bloque_excel(canvas, ordenes, 0, revision)
        self.assertEqual(canvas.create_text.call_count, 90)

    def test_constructor_difiere_carga_y_trabajador_no_toca_tk_ni_sqlite(self):
        self.assertEqual(self.g.catalogo.especificaciones, ())
        self.assertFalse(self.g._cargando_catalogo)
        with patch("generador_hii.threading.Thread") as hilo, \
             patch("generador_hii.cargar_catalogo_hii", return_value=self.catalogo(3)) as cargar, \
             patch("generador_hii.conectar", side_effect=AssertionError("No SQLite")), \
             patch("generador_hii.sincronizar_idiomas", side_effect=AssertionError("No sincronizar")):
            self.reloj.ejecutar(self.g._tareas["inicio"])
            self.assertTrue(self.g._cargando_catalogo)
            cargar.assert_not_called()
            llamadas_tk = len(self.reloj.ventana.mock_calls)
            llamadas_estado = len(self.g.estado.mock_calls)
            hilo.call_args.kwargs["target"]()
            self.assertEqual(len(self.reloj.ventana.mock_calls), llamadas_tk)
            self.assertEqual(len(self.g.estado.mock_calls), llamadas_estado)
            self.assertEqual(len(self.g.lista.filas), 0)
            self.reloj.ejecutar(self.g._tareas["catalogo"])
        self.assertFalse(self.g._cargando_catalogo)
        self.assertEqual(len(self.g.lista.filas), 3)

    def test_actualizaciones_durante_carga_se_agrupan_y_reanudan_al_final(self):
        primera, segunda = Mock(), Mock()
        with patch("generador_hii.threading.Thread") as hilo, \
             patch("generador_hii.cargar_catalogo_hii", side_effect=[self.catalogo(1), self.catalogo(2)]):
            self.g.recargar_catalogo(al_terminar=primera)
            self.g.recargar_catalogo(al_terminar=segunda)
            self.assertEqual(hilo.call_count, 1)
            hilo.call_args.kwargs["target"]()
            self.reloj.ejecutar(self.g._tareas["catalogo"])
            primera.assert_not_called()
            segunda.assert_not_called()
            self.assertEqual(hilo.call_count, 2)
            hilo.call_args.kwargs["target"]()
            self.reloj.ejecutar(self.g._tareas["catalogo"])
        self.assertEqual(len(self.g.catalogo.especificaciones), 2)
        primera.assert_called_once()
        segunda.assert_called_once()

    def test_error_conserva_catalogo_anterior_y_habilita_actualizacion(self):
        anterior = self.catalogo(2)
        self.g.catalogo = anterior
        accion = Mock()
        with patch("generador_hii.threading.Thread") as hilo, \
             patch("generador_hii.cargar_catalogo_hii", side_effect=OSError("carpeta no disponible")), \
             patch("generador_hii.messagebox") as mensajes:
            self.g.recargar_catalogo(al_terminar=accion)
            hilo.call_args.kwargs["target"]()
            self.reloj.ejecutar(self.g._tareas["catalogo"])
            mensajes.showerror.assert_called_once()
        self.assertEqual(self.g.catalogo, anterior)
        self.g.actualizar_boton.configure.assert_called_with(state="normal")
        accion.assert_not_called()

    def test_cerrar_cancela_callbacks_y_resultados_tardios_no_actualizan_widgets(self):
        with patch("generador_hii.threading.Thread") as hilo, \
             patch("generador_hii.cargar_catalogo_hii", return_value=self.catalogo(3)):
            self.g.recargar_catalogo()
            self.g._al_destruir(SimpleNamespace(widget=object()))
            self.assertFalse(self.g._cerrada)
            self.g._al_destruir(SimpleNamespace(widget=self.g.ventana))
            self.assertTrue(self.g._cancelar_catalogo.is_set())
            self.assertFalse(self.reloj.pendientes)
            llamadas = len(self.reloj.ventana.mock_calls)
            hilo.call_args.kwargs["target"]()
            self.g._consultar_catalogo()
            self.g._programar("tarde", 0, Mock())
            self.assertEqual(len(self.reloj.ventana.mock_calls), llamadas)
            self.assertEqual(self.g.catalogo.especificaciones, ())

    def test_marcar_varias_conserva_lista_busqueda_y_previa(self):
        self.g.catalogo = self.catalogo(3)
        self.g._actualizar_idiomas = Mock()
        self.g.cargar_previa = Mock()
        self.g.vaciar_previa = Mock()
        self.g.actualizar_lista()
        filas = self.g.lista.get_children()
        self.g.lista.seleccion = (filas[0],)
        self.g.origen_actual = self.g.visibles[0].ruta
        self.g.marcar(0)
        self.g.marcar(1)
        self.assertEqual(len(self.g.marcadas), 2)
        self.assertFalse(self.g.compacta)
        self.assertEqual(self.g.lista.get_children(), filas)
        self.assertEqual(self.g.lista.selection(), (filas[0],))
        self.g.consulta.set.assert_not_called()
        self.g.lote_boton.configure.assert_called_with(text="Guardar seleccionadas (2)", state="normal")
        self.g.marcar(0)
        self.g.cargar_previa.assert_not_called()
        self.g.vaciar_previa.assert_not_called()

    def test_previsualizar_fila_sin_marcar_y_filtrar_manualmente(self):
        self.g.catalogo = self.catalogo(3)
        self.g.cargar_previa = Mock()
        self.g.posicionar_paneles = Mock()
        self.g.actualizar_lista()
        self.g.lista.seleccion = (self.g.lista.get_children()[1],)
        self.g.seleccionar_fila()
        self.g.cargar_previa.assert_called_once_with(self.g.visibles[1].ruta)
        self.assertFalse(self.g.marcadas)
        self.g.marcadas.add(self.g.visibles[0].ruta)
        self.g.alternar_lista()
        self.assertEqual(len(self.g.visibles), 1)
        self.g.alternar_lista()
        self.assertEqual(len(self.g.visibles), 3)

    def test_json_guardados_se_marcan_tras_cargar_y_se_muestra_su_previa(self):
        g = self.g
        catalogo = self.catalogo(3)
        rutas = (catalogo.especificaciones[1].ruta, catalogo.especificaciones[0].ruta)
        g._actualizar_idiomas = Mock()
        g.cargar_previa = Mock()
        g.posicionar_paneles = Mock()
        with patch("generador_hii.threading.Thread") as hilo, patch(
            "generador_hii.cargar_catalogo_hii", return_value=catalogo
        ):
            g.seleccionar_json_guardados(rutas)
            self.assertNotIn("inicio", g._tareas)
            self.assertFalse(g.marcadas)
            hilo.call_args.kwargs["target"]()
            self.reloj.ejecutar(g._tareas["catalogo"])
        self.assertEqual(g.marcadas, set(rutas))
        self.assertEqual({e.ruta for e in g.visibles}, set(rutas))
        self.assertTrue(all(v[0] == "☑" for v in g.lista.filas.values()))
        g.cargar_previa.assert_called_once_with(rutas[0])
        g.lote_boton.configure.assert_called_with(text="Guardar seleccionadas (2)", state="normal")

    def test_json_externo_guardado_se_incluye_y_se_conserva_al_recargar(self):
        g = self.g
        externo = replace(self.catalogo(1).especificaciones[0], ruta=self.base.parent / "externo.json")
        g._actualizar_idiomas = Mock()
        g.cargar_previa = Mock()
        g.posicionar_paneles = Mock()
        with patch("generador_hii.threading.Thread") as hilo, patch(
            "generador_hii.cargar_catalogo_hii", return_value=self.catalogo(1)
        ), patch.object(g, "_leer_entrada", return_value=externo) as leer:
            g.seleccionar_json_guardados((externo.ruta,))
            hilo.call_args.kwargs["target"]()
            self.reloj.ejecutar(g._tareas["catalogo"])
            self.assertEqual(g.marcadas, {externo.ruta})
            self.assertEqual([e.ruta for e in g.visibles], [externo.ruta])
            g.recargar_catalogo()
            hilo.call_args.kwargs["target"]()
            self.reloj.ejecutar(g._tareas["catalogo"])
            self.assertEqual(g.marcadas, {externo.ruta})
            self.assertEqual(leer.call_count, 2)

    def test_traduccion_terminada_selecciona_la_pareja_en_el_mismo_hii(self):
        g = self.g
        origen, traducida = self.base / "es.json", self.base / "en.json"
        rutas = (origen, traducida)
        g.seleccionar_json_guardados = Mock()
        with patch("interfaz_aplicacion_especificaciones.GeneradorEspecificacion") as editor, patch(
            "generador_hii.preparar_datos",
            side_effect=lambda ruta: SimpleNamespace(idioma="INGLES" if ruta == traducida else "ESPANOL"),
        ), patch("generador_hii.conectar"), patch("generador_hii.vincular") as vincular:
            g.abrir_traduccion(origen, "INGLES")
            callback = editor.call_args.kwargs["al_finalizar"]
            callback(rutas)
            self.assertNotIn("al_guardar", editor.call_args.kwargs)
            self.assertEqual(vincular.call_args.args[1:], (origen, traducida))
            g.seleccionar_json_guardados.assert_called_once_with(rutas)

    def test_busqueda_espera_250_ms_y_solo_ejecuta_la_ultima(self):
        self.g.actualizar_lista = Mock()
        for _ in range(3):
            self.g.programar_busqueda()
        clave = self.g._tareas["busqueda"]
        self.assertEqual(self.reloj.pendientes[clave][0], 250)
        self.g.actualizar_lista.assert_not_called()
        self.reloj.ejecutar(clave)
        self.g.actualizar_lista.assert_called_once()

    def test_inserta_por_bloques_y_descarta_filas_de_busquedas_anteriores(self):
        self.g.catalogo = self.catalogo(250)
        self.g.actualizar_lista()
        revision = self.g._revision_lista
        self.assertEqual(len(self.g.lista.filas), FILAS_POR_BLOQUE)
        pendiente = self.g._tareas["lista"]
        self.g.consulta.get.return_value = "P249"
        self.g.actualizar_lista()
        self.assertNotIn(pendiente, self.reloj.pendientes)
        self.assertEqual(len(self.g.lista.filas), 1)
        self.g._insertar_bloque(revision, 100)
        self.assertEqual(len(self.g.lista.filas), 1)
        self.g.lista.seleccion = (f"{revision}:0",)
        self.g.cargar_previa = Mock()
        self.g.seleccionar_fila()
        self.g.cargar_previa.assert_not_called()

    def test_todos_los_bloques_terminan_sin_duplicar_y_conservan_marcas(self):
        self.g.catalogo = self.catalogo(250)
        self.g.marcadas.add(self.g.catalogo.especificaciones[-1].ruta)
        self.g.actualizar_lista()
        while "lista" in self.g._tareas:
            self.reloj.ejecutar(self.g._tareas["lista"])
        self.assertEqual(len(self.g.lista.filas), 250)
        self.assertEqual(list(self.g.lista.filas.values())[-1][0], "☑")

    def test_idioma_original_no_relee_documento_para_resolver(self):
        with patch("generador_hii.preparar_datos", side_effect=AssertionError("No preparar todavía")), \
             patch("generador_hii.cargar_catalogo", side_effect=AssertionError("No releer")):
            self.assertEqual(self.g.resolver_idioma(self.base / "es.json"), self.base / "es.json")

    def test_traduccion_disponible_con_sqlite_danado_no_se_pierde(self):
        origen = self.catalogo(1).especificaciones[0]
        destino = replace(origen, ruta=self.base / "en.json", idioma="INGLES")
        self.g.catalogo = Catalogo((origen, destino), ())
        self.g.idioma_objetivo = "INGLES"
        with patch.object(self.g, "_leer_entrada", side_effect=[origen, destino]), \
             patch("generador_hii.conectar", side_effect=sqlite3.DatabaseError("base dañada")), \
             patch("generador_hii.messagebox") as mensajes:
            self.assertEqual(self.g.resolver_idioma(origen.ruta), destino.ruta)
            mensajes.showwarning.assert_called_once()

    def test_varias_traducciones_requieren_eleccion_y_validacion_actual(self):
        origen = self.catalogo(1).especificaciones[0]
        una = replace(origen, ruta=self.base / "en1.json", idioma="INGLES")
        otra = replace(una, ruta=self.base / "en2.json")
        self.g.catalogo = Catalogo((origen, una, otra), ())
        self.g.idioma_objetivo = "INGLES"
        with patch.object(self.g, "_leer_entrada", side_effect=[origen, otra]), \
             patch("generador_hii.ui.Ventana"), patch("generador_hii.tk.Listbox") as lista, \
             patch("generador_hii.ttk") as widgets, patch("generador_hii.conectar"), patch("generador_hii.vincular") as vincular:
            lista.return_value.curselection.return_value = (1,)
            self.g.ventana.wait_window.side_effect = lambda _: widgets.Button.call_args.kwargs["command"]()
            self.assertEqual(self.g.resolver_idioma(origen.ruta), otra.ruta)
            vincular.assert_called_once()
            self.assertEqual(vincular.call_args.args[1:3], (origen.ruta, otra.ruta))

    def test_idiomas_y_ambos_se_habilitan_segun_todas_las_selecciones(self):
        origen = self.catalogo(1).especificaciones[0]
        destino = replace(origen, ruta=self.base / "en.json", idioma="INGLES")
        self.g.catalogo = Catalogo((origen, destino), ())
        self.g.marcadas = {origen.ruta}
        self.g.origen_actual = origen.ruta
        self.g.botones_idioma = {i: Mock() for i in ("ESPANOL", "INGLES")}
        self.g.ambos_check = Mock()
        self.g.ambos_idiomas = Mock(get=Mock(return_value=False))
        self.g._actualizar_idiomas()
        for boton in self.g.botones_idioma.values():
            self.assertEqual(boton.configure.call_args.kwargs["state"], "normal")
        self.g.ambos_check.configure.assert_called_with(state="normal")
        unica = replace(origen, ruta=self.base / "sin_traduccion.json", codigo="P2")
        self.g.catalogo = Catalogo((origen, destino, unica), ())
        self.g.marcadas.add(unica.ruta)
        self.g._actualizar_idiomas()
        self.assertEqual(self.g.botones_idioma["INGLES"].configure.call_args.kwargs["state"], "disabled")
        self.g.ambos_check.configure.assert_called_with(state="disabled")
        self.g.ambos_idiomas.set.assert_called_with(False)

    def test_no_abre_traductor_al_solicitar_idioma_ausente(self):
        origen = self.catalogo(1).especificaciones[0]
        self.g.catalogo = Catalogo((origen,), ())
        self.g.origen_actual = origen.ruta
        self.g.cargar_previa = Mock()
        self.g.abrir_traduccion = Mock()
        self.g.generar_idioma("INGLES")
        self.g.cargar_previa.assert_not_called()
        self.g.abrir_traduccion.assert_not_called()
        self.g.idioma_objetivo = "INGLES"
        with patch.object(self.g, "_leer_entrada", return_value=origen):
            with self.assertRaisesRegex(ValueError, "No hay una versión"):
                self.g.resolver_idioma(origen.ruta)

    def test_rutas_de_ambos_restaura_idioma_incluso_si_falla_una_version(self):
        self.g.idioma_objetivo = "ESPANOL"
        with patch.object(self.g, "_disponibles", return_value={"ESPANOL", "INGLES"}), \
             patch.object(self.g, "resolver_idioma", side_effect=[self.base / "es.json", ValueError("JSON cambió")]):
            with self.assertRaises(ValueError):
                self.g._rutas_guardado(self.base / "es.json", True)
        self.assertEqual(self.g.idioma_objetivo, "ESPANOL")

    def test_zoom_local_con_limites_no_regenera_el_excel(self):
        self.g.zoom_porcentaje = Mock()
        self.g.zoom_mas, self.g.zoom_menos = Mock(), Mock()
        self.g.zoom_contenido = Mock()
        with patch("generador_hii.load_workbook", side_effect=AssertionError("No releer Excel")):
            for _ in range(20):
                self.g.cambiar_zoom(.1)
            self.assertEqual(self.g.zoom_factor, 1.6)
            self.reloj.ejecutar(self.g._tareas["zoom"])
            self.g.zoom_contenido.aplicar.assert_called_with(1.6)
            for _ in range(20):
                self.g.cambiar_zoom(-.1)
            self.assertEqual(self.g.zoom_factor, .6)
            self.reloj.ejecutar(self.g._tareas["zoom"])
            self.g.zoom_contenido.aplicar.assert_called_with(.6)


if __name__ == "__main__":
    unittest.main()
