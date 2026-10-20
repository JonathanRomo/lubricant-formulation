import copy
import unittest
from unittest.mock import Mock, patch

import interfaz_aplicacion_especificaciones as modulo


class GuardadoEditorCompletoTests(unittest.TestCase):
    def editor(self):
        g = modulo.GeneradorEspecificacion.__new__(modulo.GeneradorEspecificacion)
        g.variables_control = {
            campo: Mock(get=Mock(return_value="Dato"))
            for campo, _, _ in modulo.CAMPOS_CONTROL_GENERADOR
        }
        g.pruebas = [{"prueba": "Color", "unidades": "N/A", "metodo": "Visual",
                      "especificacion": "ÁMBAR", "valor_tipico": "ÁMBAR", "periodicidad": "▶"}]
        g._idioma = Mock(get=Mock(return_value="Español"))
        g.desarrollador = Mock(get=Mock(return_value="Autor"))
        g.fecha_reanalisis = Mock(get=Mock(return_value="12"))
        g.ventana = Mock()
        g.idioma_destino = None
        g._cerrar_editor_celda = Mock()
        return g

    def test_cada_campo_de_control_es_obligatorio_antes_de_guardar(self):
        for campo, etiqueta, _ in modulo.CAMPOS_CONTROL_GENERADOR:
            with self.subTest(campo=campo):
                g = self.editor()
                g.variables_control[campo].get.return_value = "  "
                with patch.object(modulo, "messagebox") as mensajes, patch.object(modulo, "filedialog") as dialogos:
                    g.guardar()
                    mensajes.showerror.assert_called_once()
                    self.assertIn(etiqueta, mensajes.showerror.call_args.args[1])
                    dialogos.asksaveasfilename.assert_not_called()

    def test_cada_columna_de_cada_fila_es_obligatoria(self):
        for campo, etiqueta in zip((*modulo.CAMPOS_TABLA, "periodicidad"), modulo.ETIQUETAS_TABLA_GENERADOR["es"]):
            for valor in (None, "", "   "):
                with self.subTest(campo=campo, valor=valor):
                    g = self.editor()
                    g.pruebas.append(copy.deepcopy(g.pruebas[0]))
                    g.pruebas[1][campo] = valor
                    with patch.object(modulo, "messagebox") as mensajes, patch.object(modulo, "filedialog") as dialogos:
                        g.guardar()
                        mensajes.showerror.assert_called_once()
                        self.assertIn("Fila 2", mensajes.showerror.call_args.args[1])
                        self.assertIn(etiqueta, mensajes.showerror.call_args.args[1])
                        dialogos.asksaveasfilename.assert_not_called()

    def test_no_se_ignoran_filas_vacias_y_se_necesita_al_menos_una(self):
        for filas in ([], [{}]):
            with self.subTest(filas=filas):
                g = self.editor()
                g.pruebas = filas
                with patch.object(modulo, "messagebox") as mensajes:
                    self.assertFalse(g._validar_campos_guardado())
                    mensajes.showerror.assert_called_once()

    def test_valores_cero_y_no_aplica_son_datos_completos(self):
        g = self.editor()
        g.pruebas[0].update(valor_tipico=0, especificacion="0", unidades="N/A")
        with patch.object(modulo, "messagebox") as mensajes:
            self.assertTrue(g._validar_campos_guardado())
            mensajes.showerror.assert_not_called()

    def test_borradores_siguen_permitiendo_tabla_y_control_incompletos(self):
        g = self.editor()
        g.pruebas = [{}]
        g.variables_control = {}
        with patch.object(modulo, "messagebox") as mensajes:
            self.assertTrue(g._validar_campos_guardado(en_progreso=True))
            mensajes.showerror.assert_not_called()
