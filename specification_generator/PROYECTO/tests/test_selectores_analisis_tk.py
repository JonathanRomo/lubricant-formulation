"""Regresiones de los selectores reales: catálogo grande, filtros y cierre."""
import unittest
from unittest.mock import Mock, patch

import analisis_productos as modulo


class SelectoresAnalisisTkTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = modulo.ctk.CTk()
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()

    def setUp(self):
        self.errores = []
        self.root.report_callback_exception = lambda *args: self.errores.append(args)
        # Ejecutar los widgets reales sin mostrar ventanas en el escritorio.
        ocultar = patch.object(modulo.ctk.CTkToplevel, 'deiconify')
        ocultar.start()
        self.addCleanup(ocultar.stop)

    def selector(self, multiple=False):
        callback = Mock()
        clase = modulo.SelectorMultipleProductos if multiple else modulo.SelectorProductos
        selector = clase(self.root, command=callback)
        selector.pack()
        valores = [f'Aceite sintético {i:04d} · P{i:04d} · 10 pruebas' for i in range(710)]
        selector.establecer_valores(valores)
        self.addCleanup(selector.destroy)
        self.addCleanup(selector.cerrar)
        return selector, callback, valores

    def esperar_eventos(self):
        fin = modulo.ctk.BooleanVar(master=self.root)
        self.root.after(180, lambda: fin.set(True))
        self.root.wait_variable(fin)
        self.assertFalse(self.errores, self.errores)

    def test_filtrar_reutiliza_710_filas_y_no_renormaliza_el_catalogo(self):
        selector, callback, valores = self.selector()
        selector.abrir()
        lista = selector.lista
        ids = lista.tabla.get_children()
        self.assertEqual(len(ids), 710)
        with patch.object(lista.tabla, 'insert', wraps=lista.tabla.insert) as insertar, \
                patch.object(modulo, 'normalizar', wraps=modulo.normalizar) as normalizar:
            lista.consulta.set('ACEITE SINTETICO 0012')
            self.assertEqual(lista.valores_visibles(), [valores[12]])
            lista.consulta.set('P0709')
            self.assertEqual(lista.valores_visibles(), [valores[709]])
            lista.consulta.set('inexistente')
            self.assertEqual(lista.valores_visibles(), [])
            lista.activar()
            callback.assert_not_called()
            lista.consulta.set('')
            self.assertEqual(lista.valores_visibles(), valores)
            self.assertEqual(lista.tabla.get_children(), ids)
            insertar.assert_not_called()
            self.assertLess(normalizar.call_count, 10)
        lista.consulta.set('P0709')
        lista.activar()  # Enter aplica también una búsqueda todavía pendiente.
        callback.assert_called_once_with(valores[709])
        self.assertEqual(selector.get(), valores[709])
        self.assertIsNone(selector.popup)

    def test_varias_teclas_programan_un_solo_filtrado(self):
        selector, _, valores = self.selector()
        selector.abrir()
        lista = selector.lista
        with patch.object(lista, 'aplicar_filtro', wraps=lista.aplicar_filtro) as filtrar:
            for consulta in ('P', 'P0', 'P07', 'P070', 'P0709'):
                lista.consulta.set(consulta)
            filtrar.assert_not_called()
            self.esperar_eventos()
            filtrar.assert_called_once()
            self.assertEqual([lista._filas[i] for i in lista.tabla.get_children()], [valores[709]])

    def test_multiseleccion_conserva_ocultos_y_notifica_una_vez_al_cerrar(self):
        selector, callback, valores = self.selector(multiple=True)
        selector.seleccionados = {valores[0]}
        selector.abrir()
        lista = selector.lista
        lista.consulta.set('P0709')
        selector._marcar_visibles(True)  # Usa la consulta actual sin esperar el timer.
        self.assertEqual(selector.get(), {valores[0], valores[709]})
        callback.assert_not_called()
        lista.consulta.set('P0000')
        selector._marcar_visibles(False)
        self.assertEqual(selector.get(), {valores[709]})
        lista.consulta.set('P0012')
        lista.activar()
        self.assertEqual(selector.get(), {valores[12], valores[709]})
        callback.assert_not_called()
        selector.cerrar()
        callback.assert_called_once_with({valores[12], valores[709]})
        selector.cerrar()
        self.assertEqual(callback.call_count, 1)
        selector.abrir()
        self.assertEqual(selector.lista.tabla.set('12', 'marca'), '☑')
        self.assertEqual(selector.lista.tabla.set('709', 'marca'), '☑')
        self.assertEqual(selector.lista.tabla.set('0', 'marca'), '☐')
        selector.cerrar()
        self.assertEqual(callback.call_count, 1)

    def test_cerrar_cancela_busqueda_y_foco_sin_afectar_la_reapertura(self):
        for multiple in (False, True):
            with self.subTest(multiple=multiple):
                selector, callback, _ = self.selector(multiple)
                selector.abrir()
                consulta = selector.lista.consulta
                traza = selector.lista._traza
                consulta.set('P00')
                filtro_id = selector.lista._pendiente
                selector._revisar_foco()
                foco_id = selector._foco_pendiente
                selector.cerrar()
                pendientes = self.root.tk.call('after', 'info')
                self.assertNotIn(filtro_id, pendientes)
                self.assertNotIn(foco_id, pendientes)
                self.assertNotIn(traza, [nombre for _, nombre in consulta.trace_info()])
                selector.abrir()
                self.esperar_eventos()
                self.assertIsNotNone(selector.popup)
                callback.assert_not_called()
                selector.cerrar()


if __name__ == '__main__':
    unittest.main()
