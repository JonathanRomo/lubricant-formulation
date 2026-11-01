import json
from pathlib import Path
import queue
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import interfaz_aplicacion_especificaciones as modulo
from editor_comportamientos import RevisionCalculada, siguiente_revision


def variable(valor):
    var = Mock()
    var.get.return_value = valor
    var.set.side_effect = lambda nuevo: setattr(var.get, "return_value", nuevo)
    return var


def documento_historial(codigo="P1", revision="00", idioma="es", valor="50 MÁXIMO", **extras):
    return {"producto": {"nombre": "Producto"}, "documento": {"idioma": idioma},
            "control_documental": {"codigo_unico_producto": codigo, "revision": revision},
            "pruebas": [{"prueba": "Viscosidad", "metodo": "ASTM D445", "unidades": "mm²/s", "especificacion": valor}], **extras}


class RevisionControlEditorTests(unittest.TestCase):
    def editor(self, historial="08"):
        g = modulo.GeneradorEspecificacion.__new__(modulo.GeneradorEspecificacion)
        g.variables_control = {c: variable("-") for c, _, _ in modulo.CAMPOS_CONTROL_GENERADOR}
        g.variables_control.update(codigo_unico_producto=variable("P1"), revision=variable("01"))
        g.datos_originales = {"control_documental": {"revision": "01"}}
        g._originales_idioma = {}
        g._contenido_inicial_revision = ("original",)
        g._contenido_para_revision = Mock(return_value=("editado",))
        g._revision_inicial = "01"
        g._revision_aplicada = False
        g._cambio_alineacion_revision = False
        g._resultados_revision = queue.Queue()
        g._consultas_revision = {"P1"}
        g._revisiones_historial = {"P1": historial}
        g.ventana = Mock()
        return g

    def test_producto_sin_alineados_inicia_en_cero_y_revision_vacia_da_error(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            self.assertEqual(siguiente_revision("P1", [base]), "00")
            (base / "producto.json").write_text(json.dumps({"control_documental": {"codigo_unico_producto": "P1"}}))
            with self.assertRaisesRegex(ValueError, "producto.json"):
                siguiente_revision("P1", [base])

    def test_usa_mayor_revision_ignora_duplicados_idiomas_y_borradores(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            for nombre, codigo, revision in (("viejo.json", "P1", "00"),
                                            ("intermedio.json", "P1", "02"),
                                            ("es.json", "P1", "05"),
                                            ("en.json", "P1", "05"),
                                            ("copia.json", "P1", "05"),
                                            ("otro.json", "P2", "90"),
                                            ("borrador.json", "P1", "99")):
                documento = documento_historial(codigo, revision, "en" if nombre == "en.json" else "es")
                if nombre == "es.json":
                    documento["versiones_idioma"] = {"INGLES": [{"archivo_json": "en.json"}]}
                (base / nombre).write_text(json.dumps(documento))
            (base / ".borrador.json.wip").write_text("{}")
            self.assertEqual(siguiente_revision("P1", [base, base]), "06")
            self.assertEqual(siguiente_revision(" p1 ", [base]), "06")
            calculada = siguiente_revision("P1", [base])
            self.assertIn("3 revisiones distintas", calculada.explicacion)
            self.assertIn("mayor revisión numérica es 05", calculada.explicacion)
            self.assertIn("consecutivo es 06", calculada.explicacion)

    def test_revisiones_invalidas_no_se_ocultan_entre_las_validas(self):
        for revision in ("", "A", "01/02/26", "100", None):
            with self.subTest(revision=revision), tempfile.TemporaryDirectory() as carpeta:
                base = Path(carpeta)
                for nombre, valor in (("correcto.json", "05"), ("invalido.json", revision)):
                    (base / nombre).write_text(json.dumps({"control_documental": {
                        "codigo_unico_producto": "P1", "revision": valor}}))
                with self.assertRaisesRegex(ValueError, "invalido.json"):
                    siguiente_revision("P1", [base])

    def test_revision_99_no_desborda_el_campo_de_dos_digitos(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            (base / "producto.json").write_text(json.dumps({"control_documental": {
                "codigo_unico_producto": "P1", "revision": "99"}}))
            with self.assertRaisesRegex(ValueError, "99"):
                siguiente_revision("P1", [base])

    def test_anio_de_origen_se_cuenta_sin_incrementar_el_anio_ni_los_idiomas(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            for nombre, idioma, ordinal in (("es.json", "es", "01"), ("en.json", "en", "00"), ("copia.json", "es", "02")):
                documento = documento_historial(revision="2023", idioma=idioma, revision_json=ordinal)
                documento["documento"]["fecha_documento_normalizada"] = "2023-07-24"
                if nombre == "es.json":
                    documento["versiones_idioma"] = {"INGLES": [{"archivo_json": "en.json"}]}
                (base / nombre).write_text(json.dumps(documento))
            antes = {p: p.read_bytes() for p in base.iterdir()}
            self.assertEqual(siguiente_revision("P1", [base]), "01")
            self.assertEqual({p: p.read_bytes() for p in base.iterdir()}, antes)

    def test_cuenta_versiones_de_origen_y_continua_tras_ediciones_numericas(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            for nombre, revision, fecha in (("viejo.json", "2023", "24/07/2023"),
                                           ("otra_fecha.json", "2023", "10/10/2023"),
                                           ("otro_anio.json", "2024", "10/10/2024")):
                (base / nombre).write_text(json.dumps({"control_documental": {
                    "codigo_unico_producto": "P1", "revision": revision, "fecha": fecha}}))
            self.assertEqual(siguiente_revision("P1", [base]), "03")
            (base / "editado.json").write_text(json.dumps({"control_documental": {
                "codigo_unico_producto": "P1", "revision": "03", "revision_fuente": "2023"}}))
            self.assertEqual(siguiente_revision("P1", [base]), "04")

    def test_editor_aplica_consecutivo_a_un_origen_con_anio_sin_avisos(self):
        g = self.editor("01")
        g.datos_originales = {"control_documental": {"revision": "2023"}}
        g.variables_control["revision"].set("2023")
        with patch.object(modulo.messagebox, "showerror") as aviso:
            self.assertTrue(g._actualizar_revision_automatica())
            self.assertEqual(g.variables_control["revision"].get(), "01")
            aviso.assert_not_called()

    def test_guardado_conserva_anio_de_origen_en_ediciones_sucesivas(self):
        original = {"control_documental": {"revision": "2023"}}
        primera = modulo._integrar_edicion_json(original, {"control_documental": {"revision": "01"}})
        segunda = modulo._integrar_edicion_json(primera, {"control_documental": {"revision": "02"}})
        self.assertEqual(segunda["control_documental"], {"revision": "02", "revision_fuente": "2023"})
        self.assertEqual(original, {"control_documental": {"revision": "2023"}})

    def test_resumen_de_guardado_explica_historial_y_revision_a_guardar(self):
        explicacion = "Se encontraron 3 revisiones distintas. La mayor revisión numérica es 05; su consecutivo es 06."
        g = self.editor(RevisionCalculada("06", explicacion))
        self.assertTrue(g._actualizar_revision_automatica())
        resumen = g._resumen_revision_guardado()
        self.assertIn("Part ID: P1", resumen)
        self.assertIn("Revisión a guardar: 06", resumen)
        self.assertIn(explicacion, resumen)
        g.variables_control["revision"].set("07")
        self.assertIn("Revisión a guardar: 07", g._resumen_revision_guardado())
        self.assertIn("Revisión calculada: 06", g._resumen_revision_guardado())

    def test_mismo_numero_con_tres_contenidos_distintos_cuenta_tres_versiones(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            for i, valor in enumerate(("40 MÁXIMO", "45 MÁXIMO", "50 MÁXIMO")):
                (base / f"version_{i}.json").write_text(json.dumps(documento_historial(revision="01", valor=valor)))
            resultado = siguiente_revision("P1", [base])
            self.assertEqual(resultado, "03")
            self.assertIn("3 revisiones distintas", resultado.explicacion)
            self.assertIn("mismo número con contenido diferente", resultado.explicacion)

    def test_mismo_anio_y_fecha_no_ocultan_diferencias_que_detecta_verificador(self):
        for campo, valor in (("especificacion", "45 MÁXIMO"), ("unidades", "cP"), ("metodo", "ASTM D7042")):
            with self.subTest(campo=campo), tempfile.TemporaryDirectory() as carpeta:
                base = Path(carpeta)
                original = documento_historial(revision="2023")
                original["documento"]["fecha_documento_normalizada"] = "2023-07-24"
                (base / "original.json").write_text(json.dumps(original))
                (base / "copia.json").write_text(json.dumps(dict(original, revision_json="09", archivo_origen="otro.doc")))
                original["pruebas"][0][campo] = valor
                (base / "distinto.json").write_text(json.dumps(original))
                self.assertEqual(siguiente_revision("P1", [base]), "02")

    def test_idiomas_sin_vinculo_no_se_fusionan_por_numero_y_fecha(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            for idioma in ("es", "en"):
                (base / f"{idioma}.json").write_text(json.dumps(documento_historial(revision="2023", idioma=idioma)))
            self.assertEqual(siguiente_revision("P1", [base]), "02")

    def test_vinculos_ambiguos_no_fusionan_dos_especificaciones_distintas(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            for nombre, valor in (("es1.json", "40 MÁXIMO"), ("es2.json", "50 MÁXIMO")):
                documento = documento_historial(revision="2023", valor=valor,
                    versiones_idioma={"INGLES": [{"archivo_json": "en.json"}]})
                (base / nombre).write_text(json.dumps(documento))
            (base / "en.json").write_text(json.dumps(documento_historial(revision="2023", idioma="en")))
            self.assertEqual(siguiente_revision("P1", [base]), "03")

    def test_revision_calculada_se_conserva_como_referencia_al_editar(self):
        g = self.editor()
        self.assertTrue(g._actualizar_revision_automatica())
        self.assertEqual(g._revision_calculada, ("P1", "08"))
        g.variables_control["revision"].set("09")
        self.assertTrue(g._actualizar_revision_automatica())
        self.assertEqual(g.variables_control["revision"].get(), "09")
        self.assertEqual(g._revision_calculada, ("P1", "08"))

    def test_sin_historial_aplica_00_una_sola_vez_sin_preguntar(self):
        g = self.editor(None)
        with patch.object(modulo.simpledialog, "askstring") as preguntar:
            self.assertTrue(g._actualizar_revision_automatica())
            self.assertTrue(g._actualizar_revision_automatica())
            preguntar.assert_not_called()
        self.assertEqual(g.variables_control["revision"].get(), "00")
        g.variables_control["revision"].set.assert_called_once_with("00")

    def test_error_de_historial_avisa_una_vez_y_no_asigna_cero(self):
        g = self.editor(ValueError("Revisión inválida en alineado.json"))
        with patch.object(modulo.messagebox, "showerror") as aviso:
            self.assertFalse(g._actualizar_revision_automatica())
            self.assertFalse(g._actualizar_revision_automatica())
            aviso.assert_called_once()
        g.variables_control["revision"].set.assert_not_called()

    def test_no_pregunta_mientras_busca_ni_oculta_errores_de_consulta(self):
        for historial in ({}, {"P1": ValueError("No se pudo consultar")}):
            with self.subTest(historial=historial):
                g = self.editor()
                g._revisiones_historial = historial
                with patch.object(modulo.simpledialog, "askstring") as preguntar, patch.object(modulo.messagebox, "showerror"):
                    self.assertFalse(g._actualizar_revision_automatica())
                    preguntar.assert_not_called()

    def test_editor_nuevo_calcula_al_editar_tabla_sin_json_original(self):
        for historial in ("00", "06"):
            with self.subTest(historial=historial):
                g = self.editor(historial)
                g.datos_originales = None
                self.assertTrue(g._actualizar_revision_automatica())
                self.assertEqual(g.variables_control["revision"].get(), historial)
                self.assertTrue(g._actualizar_revision_automatica())
                g.variables_control["revision"].set.assert_called_once_with(historial)

    def test_editor_nuevo_no_incrementa_por_notas_o_cabecera_sin_editar_tabla(self):
        g = self.editor("06")
        g.datos_originales = None
        g._contenido_inicial_revision = (("misma tabla",), "notas originales")
        g._contenido_para_revision.return_value = (("misma tabla",), "notas editadas")
        self.assertTrue(g._actualizar_revision_automatica())
        g.variables_control["revision"].set.assert_not_called()

    def test_editor_nuevo_bilingue_detecta_tabla_y_comparte_revision(self):
        g = self.editor("06")
        g.datos_originales = None
        g._bilingue = object()
        g._contenido_inicial_revision = (("es", (), "nota"), ("en", (), "note"))
        g._contenido_para_revision.return_value = (("es", (("prueba",),), "nota"), ("en", (), "note"))
        self.assertTrue(g._actualizar_revision_automatica())
        self.assertEqual(g.variables_control["revision"].get(), "06")

    def test_consulta_solo_alineados_y_recalcula_al_cambiar_part_id(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            g = self.editor("06")
            g.app = SimpleNamespace(rutas={"json_alineadas": base, "json_anteriores": base / "originales"})
            g._actualizar_revision_automatica()
            g.variables_control["codigo_unico_producto"].set("P2")
            # Un resultado anterior de P2 tampoco debe impedir una nueva consulta.
            g._consultas_revision.add("P2")
            g._revisiones_historial["P2"] = "20"
            with patch.object(modulo.threading, "Thread", side_effect=lambda target, daemon: SimpleNamespace(start=target)), \
                 patch("editor_comportamientos.siguiente_revision", return_value="03") as consultar:
                self.assertFalse(g._actualizar_revision_automatica())
                consultar.assert_called_once_with("P2", [base])
            self.assertTrue(g._actualizar_revision_automatica())
            self.assertEqual(g.variables_control["revision"].get(), "03")

    def test_revision_igual_no_advierte_y_revision_invalida_no_se_confirma(self):
        g = self.editor()
        g._revision_calculada = ("P1", "08")
        g.variables_control["revision"].set("08")
        with patch.object(modulo, "messagebox") as mensajes:
            self.assertTrue(g._confirmar_revision_manual())
            self.assertFalse(mensajes.mock_calls)
            g.variables_control["revision"].set("abc")
            self.assertFalse(g._confirmar_revision_manual())
            mensajes.showerror.assert_called_once()
            mensajes.askyesno.assert_not_called()

    def test_guardar_revision_distinta_advierte_y_respeta_cancelar_o_continuar(self):
        for aceptar in (False, True):
            with self.subTest(aceptar=aceptar), tempfile.TemporaryDirectory() as carpeta:
                g = self.editor()
                g._actualizar_revision_automatica()
                g.variables_control["revision"].set("09")
                g._idioma = variable("Español")
                g.nombre_producto = variable("Producto")
                g.desarrollador = variable("Autor")
                g.fecha_reanalisis = variable("12")
                g.pruebas = [{"prueba": "Color", "unidades": "N/A", "metodo": "Visual",
                              "especificacion": "ÁMBAR", "valor_tipico": "ÁMBAR", "periodicidad": "▶"}]
                g.idioma_destino = None
                g._cerrar_editor_celda = Mock()
                g._actualizar_validacion_codigo = Mock(return_value=True)
                g._bilingue = None
                g.notas = Mock(get=Mock(return_value="Nota"))
                g.datos_originales = None
                g.archivo_inicial = g.destino_inicial = None
                g.al_guardar = Mock()
                g.app = SimpleNamespace(rutas={"json_alineadas": Path(carpeta)}, generar_hii=Mock())
                destino = Path(carpeta) / "producto.json"
                with patch.object(modulo, "messagebox") as mensajes, patch.object(modulo, "filedialog") as dialogos:
                    mensajes.askyesno.return_value = aceptar
                    dialogos.asksaveasfilename.return_value = str(destino)
                    g.guardar()
                    mensajes.askyesno.assert_called_once()
                    self.assertIn("falta contra el sistema de gestión", mensajes.askyesno.call_args.args[1])
                    generados = list(Path(carpeta).glob("*.JSON"))
                    self.assertEqual(bool(generados), aceptar)
                    if aceptar:
                        self.assertRegex(
                            generados[0].name,
                            r"^P1 09_\d{2}_\d{2}_\d{2}_ESP\.JSON$",
                        )
                        self.assertEqual(json.loads(generados[0].read_text(encoding="utf-8"))["control_documental"]["revision"], "09")
                    else:
                        dialogos.asksaveasfilename.assert_not_called()
                        g.ventana.destroy.assert_not_called()
