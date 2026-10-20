import json
from pathlib import Path
import tempfile
import tkinter as tk
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import interfaz_aplicacion_especificaciones as modulo


def variable(valor):
    resultado = Mock()
    resultado.get.return_value = valor
    resultado.set.side_effect = lambda nuevo: setattr(resultado.get, "return_value", nuevo)
    return resultado


class ProductosNuevosEditorTests(unittest.TestCase):
    def editor(self, nombre="Producto nuevo", codigo="20999999"):
        g = modulo.GeneradorEspecificacion.__new__(modulo.GeneradorEspecificacion)
        g.nombre_producto = variable(nombre)
        g.variables_control = {campo: variable("-") for campo, _, _ in modulo.CAMPOS_CONTROL_GENERADOR}
        g.variables_control.update(codigo_unico_producto=variable(codigo), revision=variable("00"))
        g.pruebas = [{"prueba": "Color", "unidades": "N/A", "metodo": "Visual",
                      "especificacion": "ÁMBAR", "valor_tipico": "ÁMBAR", "periodicidad": "▶"}]
        g.codigos_part_id = [
            {"part_id": "20100001", "descripcion": "Producto conocido"},
            {"part_id": "20100002", "descripcion": "Producto compartido"},
            {"part_id": "20100003", "descripcion": "Producto compartido"},
        ]
        g._idioma = variable("Español")
        g._codigo_validacion_activada = True
        g.etiqueta_codigo_asterisco = Mock()
        g.etiqueta_codigo_invalido = Mock()
        g.ventana = Mock()
        return g

    def test_producto_nuevo_se_acepta_y_aviso_pide_actualizar_lista(self):
        g = self.editor()
        self.assertTrue(g._actualizar_validacion_codigo())
        self.assertEqual(g.nombre_producto.get(), "PRODUCTO NUEVO")
        self.assertEqual(g.variables_control["codigo_unico_producto"].get(), "20999999")
        g.etiqueta_codigo_asterisco.configure.assert_called_with(text="")
        aviso = g.etiqueta_codigo_invalido.configure.call_args.kwargs
        self.assertEqual(aviso["text"],
            "Producto y Part ID no encontrados en el catálogo actual. Se guardarán los datos capturados. "
            "Favor de actualizar la lista de códigos.")
        self.assertNotEqual(aviso["foreground"], "#c00000")
        g.etiqueta_codigo_invalido.grid.assert_called_once()

    def test_codigo_conocido_conserva_su_nombre_del_catalogo(self):
        g = self.editor(nombre="Otro texto", codigo=" 20100001 ")
        self.assertTrue(g._actualizar_validacion_codigo())
        self.assertEqual(g.nombre_producto.get(), "PRODUCTO CONOCIDO")
        self.assertEqual(g.variables_control["codigo_unico_producto"].get(), "20100001")
        g.etiqueta_codigo_invalido.grid_remove.assert_called_once()

    def test_excepcion_no_aplica_a_vacios_ni_nombres_ya_conocidos(self):
        for nombre, codigo in (
            ("", "20999999"), ("Producto nuevo", " "),
            (" PRODUCTO CONOCIDO ", "20999999"),
            ("Producto compartido", "20999999"),
        ):
            with self.subTest(nombre=nombre, codigo=codigo):
                self.assertFalse(self.editor(nombre, codigo)._actualizar_validacion_codigo())

    def test_actualizar_catalogo_reconoce_producto_sin_excepcion(self):
        g = self.editor()
        self.assertTrue(g._actualizar_validacion_codigo())
        g.codigos_part_id.append({"part_id": "20999999", "descripcion": "Producto nuevo"})
        self.assertTrue(g._actualizar_validacion_codigo())
        g.etiqueta_codigo_invalido.configure.assert_called_with(text="", foreground="#c00000")
        g.etiqueta_codigo_invalido.grid_remove.assert_called_once()

    def test_normaliza_nombre_y_conserva_la_busqueda_en_catalogo(self):
        g = self.editor()
        for nombre in ("aceite nuevo", "Mezcla ácida ñ 20", "", "Producto conocido"):
            with self.subTest(nombre=nombre):
                g.nombre_producto.set(nombre)
                g._sincronizar_codigo_desde_nombre()
                self.assertEqual(g.nombre_producto.get(), nombre.upper())
        self.assertEqual(g.variables_control["codigo_unico_producto"].get(), "20100001")

    def test_nombre_se_convierte_con_traza_real_sin_perder_acentos(self):
        # StringVar con Tcl permite verificar la traza real sin abrir ventanas.
        try:
            interprete = tk.Tcl()
        except tk.TclError as exc:
            self.skipTest(f"Tcl no está disponible: {exc}")
        g = self.editor()
        g.nombre_producto = tk.StringVar(master=interprete)
        g.nombre_producto.trace_add("write", g._sincronizar_codigo_desde_nombre)
        for nombre in ("aceite nuevo", "Mezcla ácida ñ 20", "", "Producto conocido"):
            with self.subTest(nombre=nombre):
                g.nombre_producto.set(nombre)
                self.assertEqual(g.nombre_producto.get(), nombre.upper())
        self.assertEqual(g.variables_control["codigo_unico_producto"].get(), "20100001")

    def test_estructura_actual_solo_advierte_para_codigos_nuevos(self):
        for idioma in ("Español", "English"):
            for codigo in ("20999999", "30999999", " 30999999 ", "19999999", "40999999",
                           "2999999", "299999999", "2ABC9999", "２９９９９９９９", "2 999999"):
                with self.subTest(idioma=idioma, codigo=codigo):
                    g = self.editor(codigo=codigo)
                    g._idioma.set(idioma)
                    self.assertTrue(g._actualizar_validacion_codigo())
                    aviso = g.etiqueta_codigo_invalido.configure.call_args.kwargs
                    if codigo.strip() in ("20999999", "30999999"):
                        self.assertNotEqual(aviso["foreground"], "#c00000")
                        self.assertNotIn("8", aviso["text"])
                    else:
                        self.assertEqual(aviso["foreground"], "#c00000")
                        self.assertIn("8", aviso["text"])
                        self.assertIn("2 o 3" if idioma == "Español" else "2 or 3", aviso["text"])
                        self.assertTrue(aviso["text"].endswith("con 2 o 3." if idioma == "Español" else "with 2 or 3."))
                        self.assertNotIn("guardar" if idioma == "Español" else "save", aviso["text"])

    def test_aviso_espera_500_ms_desde_la_ultima_edicion(self):
        g = self.editor(codigo="200000")
        g.ventana.after.side_effect = ("primero", "segundo")
        g._programar_validacion_codigo()
        g.etiqueta_codigo_invalido.configure.assert_not_called()
        g.etiqueta_codigo_invalido.grid_remove.assert_called_once()
        g.ventana.after.assert_called_once()
        self.assertEqual(g.ventana.after.call_args.args[0], 500)

        g.variables_control["codigo_unico_producto"].set("30999999")
        g._programar_validacion_codigo()
        g.ventana.after_cancel.assert_called_once_with("primero")
        self.assertEqual(g.ventana.after.call_args.args[0], 500)
        g.etiqueta_codigo_invalido.configure.assert_not_called()

        g.ventana.after.call_args.args[1]()
        aviso = g.etiqueta_codigo_invalido.configure.call_args.kwargs
        self.assertNotEqual(aviso["foreground"], "#c00000")
        self.assertNotIn("estructura", aviso["text"])
        self.assertIsNone(g._validacion_codigo_pendiente)

    def test_editar_oculta_aviso_anterior_y_cerrar_cancela_la_espera(self):
        g = self.editor(codigo="CODIGO-NUEVO")
        g._actualizar_validacion_codigo()
        self.assertEqual(g.etiqueta_codigo_invalido.configure.call_args.kwargs["foreground"], "#c00000")
        g._programar_validacion_codigo()
        g.etiqueta_codigo_invalido.grid_remove.assert_called_once()
        g._cancelar_validacion_codigo_al_cerrar(SimpleNamespace(widget=g.ventana))
        g.ventana.after_cancel.assert_called_once()
        self.assertIsNone(g._validacion_codigo_pendiente)

    def test_codigo_del_formato_restringido_incluso_si_esta_en_catalogo(self):
        for en_catalogo in (True, False):
            with self.subTest(en_catalogo=en_catalogo):
                g = self.editor(codigo=" 20000000 ")
                if en_catalogo:
                    g.codigos_part_id.append({"part_id": "20000000", "descripcion": "Formato"})
                self.assertFalse(g._actualizar_validacion_codigo())
                g.etiqueta_codigo_invalido.configure.assert_called_with(
                    text="El código 20000000 corresponde al formato, está restringido para su uso",
                    foreground="#c00000",
                )

    def test_codigo_restringido_bloquea_guardado_json_y_en_progreso_antes_del_delay(self):
        for metodo in ("guardar", "guardar_en_progreso"):
            with self.subTest(metodo=metodo), tempfile.TemporaryDirectory() as carpeta:
                g = self.editor(codigo="20000000")
                g.idioma_destino = None
                g._bilingue = SimpleNamespace(borradores={})
                g._cerrar_editor_celda = Mock()
                g._capturar_bilingue = Mock()
                g.entradas_control = {"codigo_unico_producto": Mock()}
                g.app = SimpleNamespace(rutas={"json_alineadas": Path(carpeta)})
                g._programar_validacion_codigo()
                with patch.object(modulo, "messagebox") as mensajes, patch.object(
                    modulo, "filedialog"
                ) as dialogos, patch.object(modulo, "guardar_pareja_json") as escribir:
                    getattr(g, metodo)()
                    mensajes.showerror.assert_called_once()
                    self.assertEqual(mensajes.showerror.call_args.args[1], modulo.MENSAJE_CODIGO_RESTRINGIDO)
                    dialogos.asksaveasfilename.assert_not_called()
                    escribir.assert_not_called()
                g._capturar_bilingue.assert_not_called()
                self.assertIsNone(g._validacion_codigo_pendiente)
                self.assertFalse(list(Path(carpeta).iterdir()))

    def test_codigo_historico_del_catalogo_no_advierte_sobre_estructura_actual(self):
        g = self.editor(codigo="ANT-01")
        g.codigos_part_id.append({"part_id": "ANT-01", "descripcion": "Producto histórico"})
        self.assertTrue(g._actualizar_validacion_codigo())
        self.assertEqual(g.nombre_producto.get(), "PRODUCTO HISTÓRICO")
        g.etiqueta_codigo_invalido.configure.assert_called_with(text="", foreground="#c00000")

    def test_guardado_final_conserva_nombre_y_codigo_sin_confirmacion_extra(self):
        with tempfile.TemporaryDirectory() as carpeta:
            destino = Path(carpeta) / "producto.json"
            g = self.editor(codigo="CODIGO-NUEVO")
            g.idioma_destino = None
            g._cerrar_editor_celda = Mock()
            g._actualizar_revision_automatica = Mock(return_value=True)
            g._bilingue = None
            g.desarrollador = variable("Autor")
            g.fecha_reanalisis = variable("12")
            g.notas = Mock(get=Mock(return_value="Nota de prueba"))
            g.datos_originales = None
            g.archivo_inicial = None
            g.destino_inicial = None
            g.al_guardar = Mock()
            g.app = SimpleNamespace(rutas={"json_alineadas": Path(carpeta)}, generar_hii=Mock())
            with patch.object(modulo, "messagebox") as mensajes, patch.object(modulo, "filedialog") as dialogos:
                dialogos.asksaveasfilename.return_value = str(destino)
                g.guardar()
                mensajes.showerror.assert_not_called()
                mensajes.askyesno.assert_not_called()
            datos = json.loads(destino.read_text(encoding="utf-8"))
            self.assertEqual(datos["producto"]["nombre"], "PRODUCTO NUEVO")
            self.assertEqual(datos["control_documental"]["codigo_unico_producto"], "CODIGO-NUEVO")
            self.assertEqual(g.etiqueta_codigo_invalido.configure.call_args.kwargs["foreground"], "#c00000")
            self.assertIn("alineacion_metodos_lab", datos)
            g.al_guardar.assert_called_once_with(destino)
            g.app.generar_hii.assert_called_once_with(rutas_json=(destino,))
            g.ventana.destroy.assert_called_once()

    def test_guardado_no_permite_nombre_conocido_con_codigo_nuevo(self):
        g = self.editor(nombre="Producto conocido")
        g.idioma_destino = None
        g._cerrar_editor_celda = Mock()
        g._actualizar_revision_automatica = Mock(return_value=True)
        g.desarrollador = variable("Autor")
        g.fecha_reanalisis = variable("12")
        g.entradas_control = {"codigo_unico_producto": Mock()}
        with patch.object(modulo, "messagebox") as mensajes, patch.object(modulo, "filedialog") as dialogos:
            g.guardar()
            mensajes.showerror.assert_called_once()
            dialogos.asksaveasfilename.assert_not_called()
