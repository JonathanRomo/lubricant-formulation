import copy
import json
import queue
import tempfile
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from edicion_bilingue import BorradoresBilingues
from interfaz_aplicacion_especificaciones import (
    GeneradorEspecificacion, NOTAS_PLANTILLA, CAMPOS_CONTROL_GENERADOR,
    crear_datos_especificacion_nueva,
)
from interfaz_seleccion_especificaciones import cargar_catalogo, parejas_idioma


class EditorBilingueTests(unittest.TestCase):
    def editor(self, carpeta):
        editor = GeneradorEspecificacion.__new__(GeneradorEspecificacion)
        editor.maestro = [{"NOMBRE IQM": "DENSIDAD", "NOMBRE IQM INGLES": "DENSITY",
                           "UNIDADES": "g/ml", "UNITS": "g/ml", "METODO": "ASTM D445", "MÉTODO INGLÉS": "ASTM D445"}]
        contenido = {"idioma": "es", "nombre_producto": "Producto", "desarrollador": "Nombre", "fecha_reanalisis": "12",
                     "control": {"codigo_unico_producto": "P1", "revision": "01", "nombre_cliente": "Cliente", "elaboro_y_aprobo": "Persona",
                                 "folio_apqp": "-", "folio_vap": "-", "fecha": "13/10/26", "revision_cliente": "-"},
                     "notas": [NOTAS_PLANTILLA["es"]], "pruebas": [
                         {"prueba": "DENSIDAD", "metodo": "ASTM D445", "unidades": "g/ml", "especificacion": "149 MÍNIMO", "valor_tipico": "150", "periodicidad": "▶"}]}
        editor._bilingue = BorradoresBilingues(contenido, editor._traducir_campo_bilingue)
        editor._idioma_anterior = "es"
        editor._idioma = Mock(get=Mock(return_value="Español"))
        editor.pruebas = copy.deepcopy(editor._bilingue.borradores["es"]["pruebas"])
        editor._capturar_bilingue = Mock()
        editor._rutas_idioma = {"es": Path(carpeta) / "es.json", "en": Path(carpeta) / "en.json"}
        editor._originales_idioma = {}
        editor._cambios_alineacion = {}
        editor.app = SimpleNamespace(rutas={"json_alineadas": Path(carpeta)}, generar_hii=Mock())
        editor.nombre_producto = Mock(get=Mock(return_value="Producto"))
        editor.variables_control = {campo: Mock(get=Mock(return_value=valor)) for campo, valor in contenido["control"].items()}
        editor.ventana = Mock()
        editor.al_guardar = Mock()
        editor._contenido_para_revision = Mock(return_value=())
        editor._consultas_revision = set()
        editor._revisiones_historial = {}
        editor.desarrollador = Mock(get=Mock(return_value="Nombre"))
        editor.fecha_reanalisis = Mock(get=Mock(return_value="12"))
        editor.fecha_reanalisis.set.side_effect = lambda valor: setattr(editor.fecha_reanalisis.get, "return_value", valor)
        for nombre in ("entrada_desarrollador", "entrada_reanalisis", "etiqueta_reanalisis", "unidad_reanalisis"):
            setattr(editor, nombre, Mock())
        return editor

    def editor_con_vinculo(self, carpeta, revision="07", idioma="es"):
        editor = self.editor(carpeta)
        documentos = {}
        for lengua, contenido in editor._bilingue.borradores.items():
            documentos[lengua] = crear_datos_especificacion_nueva(
                contenido["nombre_producto"], lengua, contenido["control"],
                contenido["pruebas"], contenido["notas"], contenido["desarrollador"], contenido["fecha_reanalisis"])
        otro = "en" if idioma == "es" else "es"
        ruta = Path(carpeta) / f"Producto_{'ENG' if otro == 'en' else 'ESP'}.json"
        documentos[otro]["control_documental"]["revision"] = revision
        ruta.write_text(json.dumps(documentos[otro]), encoding="utf-8")
        documentos[idioma]["versiones_idioma"] = {
            "INGLES" if otro == "en" else "ESPANOL": [{"archivo_json": ruta.name}]}
        editor.archivo_inicial = Path(carpeta) / f"{idioma}.json"
        editor.archivo_inicial.write_text(json.dumps(documentos[idioma]), encoding="utf-8")
        editor.destino_inicial = editor.archivo_inicial
        editor.datos_originales = documentos[idioma]
        editor._idioma_anterior = idioma
        editor._contenido_bilingue = Mock(return_value=copy.deepcopy(editor._bilingue.borradores[idioma]))
        editor._originales_idioma = {}
        editor._rutas_idioma = {}
        for nombre in ("boton_guardar", "selector_idioma", "boton_guardar_wip", "aviso_pareja", "_programar_resaltado"):
            setattr(editor, nombre, Mock())
        return editor, ruta

    def test_revision_distinta_abre_borrador_sin_modal_ni_modificar_originales(self):
        for idioma in ("es", "en"):
            with self.subTest(idioma=idioma), tempfile.TemporaryDirectory() as carpeta:
                editor, ruta = self.editor_con_vinculo(carpeta, idioma=idioma)
                originales = {r: r.read_bytes() for r in (editor.archivo_inicial, ruta)}
                with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                    editor._iniciar_bilingue()
                    mensajes.assert_not_called()
                    self.assertFalse(mensajes.method_calls)
                otro = "en" if idioma == "es" else "es"
                self.assertEqual(editor._bilingue.borradores[otro]["control"]["revision"], "01")
                self.assertEqual(set(editor._rutas_idioma), {idioma})
                self.assertEqual(set(editor._originales_idioma), {idioma})
                self.assertIn(ruta.resolve(), editor._rutas_pareja_descartadas)
                self.assertIn("revisión (01 / 07)", editor.aviso_pareja.configure.call_args.kwargs["text"])
                editor.selector_idioma.configure.assert_not_called()
                editor.boton_guardar.configure.assert_not_called()
                self.assertEqual(originales, {r: r.read_bytes() for r in originales})

    def test_vinculos_no_utilizables_permiten_crear_traduccion(self):
        for caso in ("ausente", "json_invalido", "varios", "filas_distintas", "cliente_distinto", "mismo_idioma"):
            with self.subTest(caso=caso), tempfile.TemporaryDirectory() as carpeta:
                editor, ruta = self.editor_con_vinculo(carpeta, revision="01")
                datos = json.loads(ruta.read_text(encoding="utf-8"))
                if caso == "ausente":
                    ruta.unlink()
                elif caso == "json_invalido":
                    ruta.write_text("{", encoding="utf-8")
                elif caso == "varios":
                    editor.datos_originales["versiones_idioma"]["INGLES"].append({"archivo_json": "otro.json"})
                else:
                    if caso == "filas_distintas":
                        datos["pruebas"] = []
                    elif caso == "cliente_distinto":
                        datos["control_documental"]["nombre_cliente"] = "Otro cliente"
                    else:
                        datos["documento"]["idioma"] = "es"
                    ruta.write_text(json.dumps(datos), encoding="utf-8")
                with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                    editor._iniciar_bilingue()
                    self.assertFalse(mensajes.method_calls)
                self.assertEqual(editor._bilingue.borradores["en"]["pruebas"][0]["prueba"], "DENSITY")
                self.assertNotIn("en", editor._rutas_idioma)
                editor.aviso_pareja.pack.assert_called_once()

    def test_pareja_compatible_carga_version_existente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, ruta = self.editor_con_vinculo(carpeta, revision="01")
            datos = json.loads(ruta.read_text(encoding="utf-8"))
            datos["notas"] = ["Manually translated note"]
            ruta.write_text(json.dumps(datos), encoding="utf-8")
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                editor._iniciar_bilingue()
                self.assertFalse(mensajes.method_calls)
            self.assertEqual(editor._rutas_idioma["en"], ruta.resolve())
            self.assertEqual(editor._bilingue.borradores["en"]["notas"], datos["notas"])
            editor.aviso_pareja.pack.assert_not_called()

    def test_guarda_traduccion_nueva_y_conserva_revision_inglesa_anterior(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, anterior = self.editor_con_vinculo(carpeta)
            original = anterior.read_bytes()
            editor._iniciar_bilingue()
            nueva = Path(carpeta) / "Producto_ENG_1.json"
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch(
                "interfaz_aplicacion_especificaciones.filedialog.asksaveasfilename", return_value=str(nueva)
            ) as dialogo, patch("generador_hii.conectar"), patch("generador_hii.vincular"):
                mensajes.askyesno.return_value = True
                editor._guardar_bilingue()
                mensajes.showerror.assert_not_called()
                mensajes.showwarning.assert_not_called()
            self.assertEqual(dialogo.call_args.kwargs["initialfile"], nueva.name)
            self.assertEqual(anterior.read_bytes(), original)
            datos = json.loads(nueva.read_text(encoding="utf-8"))
            self.assertEqual(datos["control_documental"]["revision"], "01")
            self.assertEqual(datos["versiones_idioma"]["ESPANOL"][0]["archivo_json"], "es.json")
            fuente = json.loads(editor.archivo_inicial.read_text(encoding="utf-8"))
            self.assertEqual(fuente["versiones_idioma"]["INGLES"][0]["archivo_json"], nueva.name)

    def test_guardar_no_sobrescribe_pareja_descartada(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, anterior = self.editor_con_vinculo(carpeta)
            editor._iniciar_bilingue()
            originales = {r: r.read_bytes() for r in (editor.archivo_inicial, anterior)}
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch(
                "interfaz_aplicacion_especificaciones.filedialog.asksaveasfilename", return_value=str(anterior)
            ), patch("interfaz_aplicacion_especificaciones.guardar_pareja_json") as guardar:
                editor._guardar_bilingue()
                mensajes.showerror.assert_called_once()
                mensajes.askyesno.assert_not_called()
                guardar.assert_not_called()
            self.assertEqual(originales, {r: r.read_bytes() for r in originales})

    def test_guardar_valida_identidad_e_idiomas_antes_de_escribir(self):
        for campo, valor in (("revision", "07"), ("codigo_unico_producto", "P2"), ("nombre_cliente", "Otro"), ("idioma", "es")):
            with self.subTest(campo=campo), tempfile.TemporaryDirectory() as carpeta:
                editor = self.editor(carpeta)
                destino = editor._bilingue.borradores["en"]
                (destino if campo == "idioma" else destino["control"])[campo] = valor
                with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch(
                    "interfaz_aplicacion_especificaciones.filedialog.asksaveasfilename"
                ) as dialogo:
                    editor._guardar_bilingue()
                    mensajes.showerror.assert_called_once()
                    mensajes.askyesno.assert_not_called()
                    dialogo.assert_not_called()
                self.assertFalse(list(Path(carpeta).iterdir()))

    def test_revision_y_fecha_nuevas_se_guardan_en_ambos_idiomas(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, _ = self.editor_con_vinculo(carpeta, revision="01")
            editor._iniciar_bilingue()
            valores = dict(editor._bilingue.borradores["es"]["control"], fecha="24-02-12")
            editor.variables_control = {
                campo: SimpleNamespace(get=lambda c=campo: valores[c], set=lambda valor, c=campo: valores.__setitem__(c, valor))
                for campo in valores}
            editor._contenido_bilingue.side_effect = lambda idioma: dict(
                copy.deepcopy(editor._bilingue.borradores[idioma]),
                control=copy.deepcopy(valores), pruebas=copy.deepcopy(editor.pruebas))
            editor._capturar_bilingue = GeneradorEspecificacion._capturar_bilingue.__get__(editor)
            editor._contenido_para_revision = GeneradorEspecificacion._contenido_para_revision.__get__(editor)
            editor._contenido_inicial_revision = editor._contenido_para_revision()
            editor._revision_inicial = "01"
            editor._revision_aplicada = False
            editor._cambio_alineacion_revision = False
            editor._resultados_revision = queue.Queue()
            editor._consultas_revision = {"P1"}
            editor._revisiones_historial = {"P1": "08"}
            editor.pruebas[0]["valor_tipico"] = "151"
            with patch("interfaz_aplicacion_especificaciones.date") as reloj, patch(
                "interfaz_aplicacion_especificaciones.messagebox"
            ) as mensajes, patch("generador_hii.conectar"), patch("generador_hii.vincular"):
                reloj.today.return_value = date(2026, 10, 14)
                self.assertTrue(editor._actualizar_revision_automatica())
                mensajes.askyesno.return_value = True
                editor._guardar_bilingue()
                mensajes.showerror.assert_not_called()
                mensajes.showwarning.assert_not_called()
            for ruta in editor._rutas_idioma.values():
                datos = json.loads(ruta.read_text(encoding="utf-8"))
                self.assertEqual(datos["control_documental"]["fecha"], "14/10/26")
                self.assertEqual(datos["control_documental"]["revision"], "08")
                self.assertEqual(datos["documento"]["fecha_documento_normalizada"], "2026-10-14")

    def test_desarrollador_obligatorio_impide_cualquier_guardado(self):
        for modo in ("json", "bilingue", "progreso"):
            with self.subTest(modo=modo), tempfile.TemporaryDirectory() as carpeta:
                editor = self.editor(carpeta)
                editor.desarrollador.get.return_value = "  "
                editor.idioma_destino = None
                editor._cerrar_editor_celda = Mock()
                with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch(
                    "interfaz_aplicacion_especificaciones.filedialog.asksaveasfilename"
                ) as dialogo:
                    if modo == "json":
                        editor.guardar()
                    elif modo == "progreso":
                        editor.guardar_en_progreso()
                    else:
                        editor._guardar_bilingue()
                    mensajes.showerror.assert_called_once()
                    mensajes.askyesno.assert_not_called()
                    dialogo.assert_not_called()
                editor.entrada_desarrollador.focus_set.assert_called_once()
                self.assertFalse(list(Path(carpeta).iterdir()))

    def test_guardado_final_valida_tambien_celdas_del_otro_idioma(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            editor._bilingue.borradores["en"]["pruebas"][0]["valor_tipico"] = " "
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch(
                "interfaz_aplicacion_especificaciones.filedialog.asksaveasfilename"
            ) as dialogo:
                editor._guardar_bilingue()
                mensajes.showerror.assert_called_once()
                self.assertIn("ENG", mensajes.showerror.call_args.args[1])
                self.assertIn("Typical values", mensajes.showerror.call_args.args[1])
                dialogo.assert_not_called()
            self.assertFalse(list(Path(carpeta).iterdir()))

    def test_codigo_restringido_en_cualquier_borrador_impide_guardar_la_pareja(self):
        for idioma in ("es", "en"):
            for en_progreso in (True, False):
                with self.subTest(idioma=idioma, en_progreso=en_progreso), tempfile.TemporaryDirectory() as carpeta:
                    editor = self.editor(carpeta)
                    editor._bilingue.borradores[idioma]["control"]["codigo_unico_producto"] = "20000000"
                    with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch(
                        "interfaz_aplicacion_especificaciones.filedialog.asksaveasfilename"
                    ) as dialogo, patch("interfaz_aplicacion_especificaciones.guardar_pareja_json") as escribir:
                        editor._guardar_bilingue(en_progreso=en_progreso)
                        mensajes.showerror.assert_called_once()
                        self.assertEqual(mensajes.showerror.call_args.args[1],
                            "El código 20000000 corresponde al formato, está restringido para su uso")
                        dialogo.assert_not_called()
                        escribir.assert_not_called()
                    self.assertFalse(list(Path(carpeta).iterdir()))

    def test_revision_manual_cancelada_impide_escribir_ambos_json(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            editor._revision_calculada = ("P1", "08")
            editor.variables_control["revision"].get.return_value = "09"
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch(
                "interfaz_aplicacion_especificaciones.filedialog.asksaveasfilename"
            ) as dialogo:
                mensajes.askyesno.return_value = False
                editor._guardar_bilingue()
                mensajes.askyesno.assert_called_once()
                self.assertIn("sistema de gestión", mensajes.askyesno.call_args.args[1])
                dialogo.assert_not_called()
            self.assertFalse(list(Path(carpeta).iterdir()))

    def test_json_terminado_requiere_reanalisis_de_dos_digitos(self):
        for valor in ("", "6", "123", "ab", "12.0", "-1", "１２"):
            with self.subTest(valor=valor), tempfile.TemporaryDirectory() as carpeta:
                editor = self.editor(carpeta)
                editor.fecha_reanalisis.get.return_value = valor
                with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                    editor._guardar_bilingue()
                    mensajes.showerror.assert_called_once()
                    mensajes.askyesno.assert_not_called()
                editor.entrada_reanalisis.focus_set.assert_called_once()
                self.assertFalse(list(Path(carpeta).iterdir()))

    def test_reanalisis_se_edita_en_campo_existente_y_se_comparte_entre_idiomas(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor, _ = self.editor_con_vinculo(carpeta, revision="01")
            editor._iniciar_bilingue()
            contenido = copy.deepcopy(editor._bilingue.borradores["es"])
            contenido["fecha_reanalisis"] = "06"
            editor._bilingue.actualizar("es", contenido)
            editor.fecha_reanalisis.get.return_value = "06"
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch("generador_hii.conectar"), patch("generador_hii.vincular"):
                mensajes.askyesno.return_value = True
                editor._guardar_bilingue()
                mensajes.showerror.assert_not_called()
            for idioma, ruta in editor._rutas_idioma.items():
                datos = json.loads(ruta.read_text(encoding="utf-8"))
                self.assertEqual(datos["FECHA DE REANALISIS"], "06 months" if idioma == "en" else "06 meses")
                self.assertNotIn("fecha_reanalisis", datos["control_documental"])
                from generador_hii import preparar_datos
                self.assertEqual(preparar_datos(ruta).reanalisis, datos["FECHA DE REANALISIS"])

    def test_maestro_traduce_con_fuente_intacta_y_no_usa_fallback_como_traduccion(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            en = editor._bilingue.borradores["en"]
            self.assertEqual(en["pruebas"][0]["prueba"], "DENSITY")
            self.assertEqual(en["pruebas"][0]["metodo"], "ASTM D445")
            self.assertEqual(en["notas"], [NOTAS_PLANTILLA["en"]])
            self.assertFalse(editor._bilingue.pendientes["en"])
            editor.maestro[0]["NOMBRE IQM INGLES"] = ""
            _, seguro = editor._traducir_campo_bilingue("prueba", "DENSIDAD", "es", "en", {"prueba": "DENSIDAD"})
            self.assertFalse(seguro)

    def test_selector_de_idioma_traduce_la_pareja_guardada_automaticamente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            pareja = copy.deepcopy(editor._bilingue.borradores["en"])
            pareja["pruebas"][0]["especificacion"] = "149 MÍNIMO"
            editor._bilingue.cargar_pareja(pareja)
            editor._idioma = Mock(get=Mock(return_value="English"))
            editor._nombres_sugeridos_por_fila = {}
            editor._limpiar_sugerencia = Mock()
            editor._cerrar_editor_celda = Mock()
            editor._seleccion = Mock(return_value=None)
            editor._actualizar_tabla = Mock()
            editor._programar_resaltado = Mock()
            editor.variables_control = {c: Mock(get=Mock(return_value="Persona")) for c, _, _ in CAMPOS_CONTROL_GENERADOR}
            editor._etiquetas = {c: Mock() for c, _, _ in CAMPOS_CONTROL_GENERADOR}
            editor.desarrollador = Mock()
            editor.notas = Mock(get=Mock(return_value=NOTAS_PLANTILLA["en"]))
            editor.etiqueta_codigo_invalido = None
            for nombre in ("etiqueta_idioma", "selector_idioma", "etiqueta_producto", "etiqueta_desarrollador", "etiqueta_titulo_control", "marco_pruebas", "tabla_pruebas", "_ordenador_pruebas", "boton_eliminar", "boton_alinear", "etiqueta_notas", "boton_guardar", "boton_cancelar"):
                setattr(editor, nombre, Mock())
            editor._actualizar_idioma()
            self.assertEqual(editor.pruebas[0]["especificacion"], "149 MINIMUM")
            self.assertEqual(editor._idioma_anterior, "en")
            self.assertEqual(editor._bilingue.borradores["es"]["pruebas"][0]["especificacion"], "149 MÍNIMO")

    def test_guarda_dos_json_y_referencias_correctas_sin_ids_internos(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch("generador_hii.conectar"), patch("generador_hii.vincular"):
                mensajes.askyesno.return_value = True
                editor._guardar_bilingue()
                mensajes.showerror.assert_not_called()
            es = json.loads((Path(carpeta) / "es.json").read_text(encoding="utf-8"))
            en = json.loads((Path(carpeta) / "en.json").read_text(encoding="utf-8"))
            self.assertEqual(es["documento"]["idioma"], "es")
            self.assertEqual(en["documento"]["idioma"], "en")
            self.assertEqual(en["pruebas"][0]["especificacion"], "149 MINIMUM")
            self.assertEqual(es["versiones_idioma"]["INGLES"][0]["archivo_json"], "en.json")
            self.assertNotIn("_id_bilingue", en["pruebas"][0])
            self.assertEqual(editor.al_guardar.call_count, 2)
            editor.app.generar_hii.assert_called_once()
            self.assertEqual(set(editor.app.generar_hii.call_args.kwargs["rutas_json"]),
                             {Path(carpeta) / "es.json", Path(carpeta) / "en.json"})
            editor.ventana.destroy.assert_called_once()
            catalogo = cargar_catalogo(Path(carpeta))
            self.assertEqual(len(parejas_idioma(catalogo.especificaciones)), 2)

    def test_borrador_guardado_no_abre_hii_ni_cierra_editor(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            editor._confirmar_revision_manual = Mock(return_value=False)
            rutas = [Path(carpeta) / "es_trabajo.json", Path(carpeta) / "en_trabajo.json"]
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch(
                "interfaz_aplicacion_especificaciones.filedialog.asksaveasfilename",
                side_effect=[str(r) for r in rutas],
            ):
                mensajes.askyesno.return_value = True
                editor._guardar_bilingue(en_progreso=True)
                mensajes.showerror.assert_not_called()
            self.assertTrue(all(r.exists() for r in rutas))
            editor.app.generar_hii.assert_not_called()
            editor.ventana.destroy.assert_not_called()
            editor._confirmar_revision_manual.assert_not_called()

    def test_cancelar_y_rutas_iguales_no_guardan_archivos(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                mensajes.askyesno.return_value = False
                editor._guardar_bilingue()
                self.assertFalse(list(Path(carpeta).iterdir()))
                editor._rutas_idioma["en"] = editor._rutas_idioma["es"]
                editor._guardar_bilingue()
                mensajes.showerror.assert_called_once()
                self.assertFalse(list(Path(carpeta).iterdir()))

    def test_guardado_conserva_historial_hii_reciente_de_cada_idioma(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes, patch("generador_hii.conectar"), patch("generador_hii.vincular"):
                mensajes.askyesno.return_value = True
                editor._guardar_bilingue()
                # El generador registra una HII mientras el editor permanece abierto.
                for idioma, ruta in editor._rutas_idioma.items():
                    datos = json.loads(ruta.read_text(encoding="utf-8"))
                    datos["hii_generadas"] = [{"id": idioma, "archivo": f"{idioma}.xlsm"}]
                    ruta.write_text(json.dumps(datos), encoding="utf-8")
                editor._guardar_bilingue()
                mensajes.showerror.assert_not_called()
            for idioma, ruta in editor._rutas_idioma.items():
                datos = json.loads(ruta.read_text(encoding="utf-8"))
                self.assertEqual(datos["hii_generadas"], [{"id": idioma, "archivo": f"{idioma}.xlsm"}])

    def test_pendiente_rojo_conserva_fondo_fuente_y_asterisco_solo_visual(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            editor._idioma_anterior = "en"
            editor.pruebas = editor._bilingue.borradores["en"]["pruebas"]
            editor._bilingue.pendientes["en"].add((editor.pruebas[0]["_id_bilingue"], "especificacion"))
            editor._cambios_alineacion["en"] = {editor.pruebas[0]["_id_bilingue"]: {"especificacion"}}
            editor._marcas_pendientes = []
            editor._celda_editada = None
            editor._celda_activa = None
            editor.tabla_pruebas = Mock()
            editor.tabla_pruebas.bbox.return_value = (5, 25, 200, 25)
            editor.tabla_pruebas.winfo_width.return_value = 800
            editor.tabla_pruebas.cget.return_value = "Treeview"
            editor.tabla_pruebas.selection.return_value = ()
            for nombre in ("notas", "boton_confirmar_traduccion", "boton_confirmar_notas", "referencia_notas"):
                setattr(editor, nombre, Mock())
            with patch("interfaz_aplicacion_especificaciones.tk.Label") as etiqueta, patch("interfaz_aplicacion_especificaciones.ttk.Style") as estilo, patch("interfaz_aplicacion_especificaciones.ui.paleta", return_value={"panel": "#ffffff", "texto": "#102030", "fondo": "#eef3f9", "acento": "#1261a0"}):
                estilo.return_value.lookup.side_effect = lambda _nombre, campo, *args: {"fieldbackground": "#ffffff", "font": ("Segoe UI", 12)}.get(campo, "")
                editor._resaltar_pendientes()
                self.assertEqual(etiqueta.call_count, 1)
                self.assertEqual(etiqueta.call_args.kwargs["background"], "#ffffff")
                self.assertEqual(etiqueta.call_args.kwargs["foreground"], "#c62828")
                self.assertEqual(etiqueta.call_args.kwargs["font"], ("Segoe UI", 12))
                editor.pruebas[0]["especificacion"] = ""
                etiqueta.reset_mock()
                editor._resaltar_pendientes()
                self.assertEqual(etiqueta.call_args.kwargs["text"], "*")
                self.assertEqual(editor.pruebas[0]["especificacion"], "")
                editor._bilingue.pendientes["en"].clear()
                etiqueta.reset_mock()
                editor._resaltar_pendientes()
                self.assertEqual(etiqueta.call_args.kwargs["foreground"], "#3b82f6")

    def test_guardado_con_pendientes_enfoca_celda_y_no_escribe(self):
        with tempfile.TemporaryDirectory() as carpeta:
            editor = self.editor(carpeta)
            editor._idioma_anterior = "en"
            editor.pruebas = editor._bilingue.borradores["en"]["pruebas"]
            editor._bilingue.pendientes["en"].add((editor.pruebas[0]["_id_bilingue"], "especificacion"))
            editor.tabla_pruebas = Mock()
            editor._mostrar_referencia_bilingue = Mock()
            editor._programar_resaltado = Mock()
            with patch("interfaz_aplicacion_especificaciones.messagebox") as mensajes:
                editor._guardar_bilingue()
                mensajes.showwarning.assert_called_once()
                mensajes.askyesno.assert_not_called()
            editor.tabla_pruebas.see.assert_called_once_with("0")
            self.assertEqual(editor._celda_activa, ("0", "especificacion"))
            self.assertFalse(list(Path(carpeta).iterdir()))


if __name__ == "__main__":
    unittest.main()
