"""Edición, inserción, eliminación y scroll con Treeview y actualización reales."""
import copy
import unittest
from types import MethodType, SimpleNamespace
from unittest.mock import Mock, patch

import interfaz_aplicacion_especificaciones as m
import test_autocompletado_celdas_tk as entorno_tk


class TablaEditorIntegracionTkTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        entorno_tk.AutocompletadoCeldasTkTests.setUpClass()

    @classmethod
    def tearDownClass(cls):
        entorno_tk.AutocompletadoCeldasTkTests.tearDownClass()

    def setUp(self):
        self.entorno = entorno_tk.AutocompletadoCeldasTkTests()
        self.entorno.setUp()
        self.raiz = self.entorno.raiz
        g = self.g = self.entorno.generador
        g._bilingue = None
        g._celda_activa = None
        g._cambios_alineacion = {}
        g._idioma_anterior = 'es'
        g._programar_resaltado = Mock()
        g.boton_eliminar = m.ttk.Button(g.ventana, command=g._eliminar_seleccionada)
        g._ordenador_pruebas = m.OrdenadorColumnasTreeview(g.tabla_pruebas, {'prueba': 'Prueba'})
        g._actualizar_tabla = MethodType(m.GeneradorEspecificacion._actualizar_tabla, g)
        g._sugerencias = MethodType(m.GeneradorEspecificacion._sugerencias, g)
        g.tabla_pruebas.bind('<Delete>', g._borrar_celda_activa)
        g.tabla_pruebas.bind('<<TreeviewSelect>>', g._actualizar_boton_eliminar)
        g.tabla_pruebas.bind('<Configure>', g._reposicionar_editor, add='+')
        g.tabla_pruebas.bind('<ButtonRelease-1>', g._programar_posicion_editor, add='+')
        g.tabla_pruebas.configure(yscrollcommand=lambda *_: g._programar_posicion_editor(),
                                 xscrollcommand=lambda *_: g._programar_posicion_editor())
        self.filas(3)

    def tearDown(self):
        try:
            self.entorno.tearDown()
        finally:
            self.entorno.doCleanups()

    def filas(self, cantidad):
        self.g.pruebas = [dict(prueba=f'PRUEBA_{i:02}', metodo=f'METODO_{i}', unidades=f'U{i}',
                              especificacion=f'{i + 1} MAXIMO', valor_tipico=str(i + 1), periodicidad='▶')
                         for i in range(cantidad)]
        self.g._actualizar_tabla()
        self.raiz.update()

    def abrir(self, iid='0', campo='unidades'):
        columna = (*m.CAMPOS_TABLA, 'periodicidad').index(campo) + 1
        self.g._editar_celda(SimpleNamespace(iid=str(iid), columna=f'#{columna}'))
        self.raiz.update()
        if self.g._editor_celda is not None:
            self.g._editor_celda.focus_force()
            self.raiz.update()
        return self.g._editor_celda

    def confirmar(self, entrada):
        entrada.event_generate('<Return>')
        self.raiz.update()

    def suprimir(self):
        self.g.tabla_pruebas.focus_force()
        self.raiz.update()
        self.g.tabla_pruebas.event_generate('<Delete>')
        self.raiz.update()

    def test_eliminar_intermedia_o_ultima_invalida_celda_y_suprimir_no_borra_otras(self):
        for indice in ('1', '2'):
            with self.subTest(indice=indice):
                self.filas(3)
                self.confirmar(self.abrir(indice))
                with patch.object(m.messagebox, 'askyesno', return_value=True):
                    self.g._eliminar_seleccionada()
                self.raiz.update()
                antes = copy.deepcopy(self.g.pruebas)
                self.assertEqual(len(antes), 2)
                self.assertIsNone(self.g._celda_activa)
                self.suprimir()
                self.assertEqual(self.g.pruebas, antes)

    def test_flechas_cambian_foco_y_suprimir_afecta_solo_la_fila_seleccionada(self):
        self.confirmar(self.abrir())
        tabla = self.g.tabla_pruebas
        self.assertEqual(tabla.focus(), '0')
        tabla.event_generate('<Down>')
        self.raiz.update()
        self.assertEqual(tabla.selection(), ('1',))
        self.assertEqual(self.g._celda_activa, ('1', 'unidades'))
        self.suprimir()
        self.assertEqual([f['unidades'] for f in self.g.pruebas], ['U0', '', 'U2'])

    def test_guardar_celda_reutiliza_las_filas_y_conserva_orden(self):
        tabla = self.g.tabla_pruebas
        self.g.pruebas[0]['prueba'] = 'ZZZ'
        self.g._actualizar_tabla()
        self.g._ordenador_pruebas.alternar('prueba')
        self.raiz.update()
        orden = tabla.get_children()
        entrada = self.abrir('1')
        entrada.insert('end', '-NUEVA')
        with patch.object(tabla, 'delete', wraps=tabla.delete) as borrar, \
                patch.object(tabla, 'insert', wraps=tabla.insert) as insertar:
            self.confirmar(entrada)
            borrar.assert_not_called()
            insertar.assert_not_called()
        self.assertEqual(tabla.get_children(), orden)
        self.assertEqual(tabla.focus(), '1')
        self.assertEqual(self.g.pruebas[1]['unidades'], 'U1-NUEVA')

    def test_desplazar_fuera_de_vista_confirma_en_origen_y_retira_el_editor(self):
        self.filas(30)
        entrada = self.abrir()
        entrada.delete(0, 'end')
        entrada.insert(0, 'EDITADA')
        self.g.tabla_pruebas.yview_scroll(3, 'units')
        self.raiz.update()
        self.assertIsNone(self.g._editor_celda)
        self.assertEqual(self.g.pruebas[0]['unidades'], 'EDITADA')
        self.assertEqual(self.g.pruebas[3]['unidades'], 'U3')

    def test_scroll_visible_reposiciona_el_editor_y_horizontal_lo_cierra(self):
        self.filas(30)
        entrada = self.abrir('4')
        tabla = self.g.tabla_pruebas
        tabla.yview_scroll(1, 'units')
        self.raiz.update()
        self.assertIs(self.g._editor_celda, entrada)
        self.assertEqual(int(entrada.place_info()['y']), tabla.bbox('4', '#2')[1])
        for campo in (*m.CAMPOS_TABLA, 'periodicidad'):
            tabla.column(campo, width=300, stretch=False)
        tabla.xview_moveto(1)
        self.raiz.update()
        self.assertIsNone(self.g._editor_celda)

    def test_escape_sin_sugerencias_cancela_y_con_sugerencias_requiere_dos(self):
        g = self.g
        entrada = self.abrir()
        entrada.insert('end', '-DESCARTAR')
        entrada.event_generate('<Escape>')
        self.raiz.update()
        self.assertIsNone(g._editor_celda)
        self.assertEqual(g.pruebas[0]['unidades'], 'U0')
        g._sugerencias = Mock(return_value=('mm²/s', 'mPa*s'))
        entrada = self.abrir()
        entrada.delete(0, 'end')
        entrada.insert(0, 'm')
        self.raiz.update()
        self.assertIsNotNone(g._popup_sugerencias)
        entrada.event_generate('<Escape>')
        self.raiz.update()
        self.assertIs(g._editor_celda, entrada)
        entrada.event_generate('<Escape>')
        self.raiz.update()
        self.assertIsNone(g._editor_celda)
        self.assertEqual(g.pruebas[0]['unidades'], 'U0')

    def test_fila_nueva_cancelada_o_confirmada_vacia_no_persiste(self):
        for confirmar in (False, True):
            with self.subTest(confirmar=confirmar):
                entrada = self.abrir('nueva')
                entrada.event_generate('<Return>' if confirmar else '<Escape>')
                self.raiz.update()
                self.assertEqual(len(self.g.pruebas), 3)
                self.assertEqual(self.g.tabla_pruebas.get_children(), ('0', '1', '2', 'nueva'))
                self.assertIsNone(self.g._fila_provisional)

    def test_fila_nueva_con_contenido_se_confirma_y_se_puede_completar(self):
        entrada = self.abrir('nueva', 'prueba')
        entrada.insert(0, 'PRUEBA NUEVA')
        self.confirmar(entrada)
        self.assertEqual(len(self.g.pruebas), 4)
        self.assertEqual(self.g.pruebas[3]['prueba'], 'PRUEBA NUEVA')
        self.assertEqual(self.g._celda_activa, ('3', 'prueba'))
        entrada = self.abrir('3')
        entrada.insert(0, 'mg/kg')
        self.confirmar(entrada)
        self.assertEqual(self.g.pruebas[3]['unidades'], 'mg/kg')

    def test_cancelar_metodo_revierte_nombre_y_unidad_autocompletados(self):
        g = self.g
        g.pruebas[0].update(prueba='', metodo='', unidades='')
        original = copy.deepcopy(g.pruebas[0])
        g.maestro = [{'METODO': 'ASTM D445', 'NOMBRE IQM': 'VISCOSIDAD', 'UNIDADES': 'mm²/s'}]
        entrada = self.abrir('0', 'metodo')
        g._elegir_sugerencia_fila('0', 'metodo', 'ASTM D445')
        self.assertEqual(g.pruebas[0]['unidades'], 'mm²/s')
        entrada.event_generate('<Escape>')
        self.raiz.update()
        self.assertEqual(g.pruebas[0], original)

    def test_vigilancia_bilingue_no_publica_ediciones_provisionales(self):
        g = self.g
        def contenido(idioma):
            return dict(idioma=idioma, nombre_producto='PRUEBA', control={},
                        pruebas=copy.deepcopy(g.pruebas), notas=[], desarrollador='Prueba')
        g._contenido_bilingue = contenido
        g._bilingue = m.BorradoresBilingues(contenido('es'), lambda campo, valor, *args: (valor, True))
        g._capturar_bilingue('es')
        g._actualizar_tabla()
        self.raiz.update()
        self.abrir('nueva')
        g._contenido_para_revision()
        self.assertEqual(len(g.pruebas), 4)
        for idioma in ('es', 'en'):
            self.assertEqual(len(g._bilingue.borradores[idioma]['pruebas']), 3)
        g._cerrar_editor_celda(guardar=False)
        self.assertEqual(len(g.pruebas), 3)
        antes = copy.deepcopy(g.pruebas[0])
        entrada = self.abrir()
        # Una sugerencia puede completar otras columnas antes de confirmar el Entry.
        g.pruebas[0]['prueba'] = 'SUGERENCIA SIN CONFIRMAR'
        g._contenido_para_revision()
        self.assertEqual(g._bilingue.borradores['es']['pruebas'][0], antes)
        g._cerrar_editor_celda(guardar=False)
        self.assertEqual(g.pruebas[0], antes)
        entrada = self.abrir('nueva', 'prueba')
        entrada.insert(0, 'NUEVA CONFIRMADA')
        self.confirmar(entrada)
        for idioma in ('es', 'en'):
            self.assertEqual(len(g._bilingue.borradores[idioma]['pruebas']), 4)

    def test_enter_y_escape_en_periodicidad_restauran_el_foco_de_tabla(self):
        g = self.g
        self.abrir('0', 'periodicidad')
        g._agregar_selector_periodicidad('♦')
        selector = g._periodicidad_editores[-1]
        selector.focus_force()
        self.raiz.update()
        selector.event_generate('<Return>')
        self.raiz.update()
        self.assertEqual(g.pruebas[0]['periodicidad'], '▶♦')
        self.assertIsNone(g._editor_celda)
        self.assertEqual(str(self.raiz.focus_get()), str(g.tabla_pruebas))
        self.abrir('0', 'periodicidad')
        g._agregar_selector_periodicidad('Q')
        selector = g._periodicidad_editores[-1]
        selector.focus_force()
        self.raiz.update()
        selector.event_generate('<Escape>')
        self.raiz.update()
        self.assertEqual(g.pruebas[0]['periodicidad'], '▶♦')
        self.assertIsNone(g._editor_celda)
        self.assertEqual(str(self.raiz.focus_get()), str(g.tabla_pruebas))


if __name__ == '__main__':
    unittest.main()
