from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock, patch

from analisis_productos import VentanaAnalisis
from analisis_productos import comparar, comparar_multidimensional, leer_producto, filas_grupo
import json
import tempfile
from pathlib import Path


class RequisitosRepetidosTests(TestCase):
    def producto(self, limites, tipicos=None):
        temporal = tempfile.TemporaryDirectory()
        self.addCleanup(temporal.cleanup)
        ruta = Path(temporal.name) / "producto.json"
        pruebas = [{"numero": i + 1, "prueba": "COLOR", "metodo": "ASTM D1500", "unidades": "NA",
                    "especificacion": limite, "valor_tipico": (tipicos or [""] * len(limites))[i]}
                   for i, limite in enumerate(limites)]
        ruta.write_text(json.dumps({"producto": {"nombre": "PRODUCTO"}, "pruebas": pruebas,
            "alineacion_metodos_lab": {"detalle": [{"numero_prueba": p["numero"], "estado": "ALINEADO"} for p in pruebas]}}), encoding="utf-8")
        return leer_producto(ruta)

    def test_conserva_todos_los_requisitos_y_no_declara_cobertura_total(self):
        origen = self.producto(["6 MÁXIMO", "3 MÁXIMO"])
        candidato = self.producto(["4 MÁXIMO"])
        self.assertEqual(origen.numero_pruebas, 2)
        self.assertIn("2 pruebas", origen.etiqueta)
        resultado = comparar(origen, candidato)
        self.assertEqual((resultado.requisitos, resultado.cubiertos), (2, 1))
        self.assertEqual(resultado.cobertura, .5)
        self.assertEqual(len(resultado.filas), 2)

    def test_reportar_no_oculta_el_requisito_de_otra_fila(self):
        origen = self.producto(["REPORTAR", "20 - 22"])
        resultado = comparar(origen, self.producto(["21 - 21.5"]))
        self.assertEqual((resultado.requisitos, resultado.cubiertos), (1, 1))
        self.assertEqual(len(resultado.filas), 2)

    def test_candidatos_repetidos_son_ambiguos_independientemente_del_orden(self):
        origen = self.producto(["4 MÁXIMO"])
        for limites in (["6 MÁXIMO", "3 MÁXIMO"], ["3 MÁXIMO", "6 MÁXIMO"]):
            candidato = self.producto(limites, [6, 3])
            resultado = comparar(origen, candidato)
            self.assertEqual(resultado.cobertura, 0)
            self.assertEqual(resultado.filas[0].estado, "NO_COMPARABLE")
            multi = comparar_multidimensional(origen, candidato)
            self.assertEqual(multi.evaluables, 0)
            self.assertIsNone(multi.distancia)
            grupo = next(iter(candidato.pruebas.values()))
            self.assertEqual(len(filas_grupo(grupo)), 2)
            celda = VentanaAnalisis._valor_matriz(grupo)
            self.assertIn("Revisar", celda)
            self.assertIn("6", celda)
            self.assertIn("3", celda)


class ResultadosCoberturaTests(TestCase):
    def test_oculta_cero_y_conserva_cobertura_pequena(self):
        origen = SimpleNamespace(etiqueta='Origen')
        candidatos = [SimpleNamespace(etiqueta=str(i), nombre=str(i), codigo=str(i)) for i in range(3)]
        def resultado(candidato, cobertura):
            return SimpleNamespace(candidato=candidato, cobertura=cobertura, cubiertos=int(cobertura > 0),
                indicativos=0, fuera=0, faltantes=0, parciales=0, sin_dato=0, requisitos=1,
                dictamen='Cobertura parcial')
        resultados = [resultado(candidatos[0], 0), resultado(candidatos[1], .004), resultado(candidatos[2], .2)]
        tabla = Mock()
        tabla.get_children.side_effect = [(), ('0', '1')]
        ventana = SimpleNamespace(tabla=tabla, selector=SimpleNamespace(get=lambda: 'Origen'),
            productos=[origen, *candidatos], _escribir_detalle=Mock(), mostrar_detalle=Mock())
        with patch('analisis_productos.comparar', side_effect=resultados):
            VentanaAnalisis.actualizar_resultados(ventana)
        self.assertEqual([r.cobertura for r in ventana.comparaciones], [.2, .004])
        self.assertEqual(tabla.insert.call_count, 2)
        self.assertEqual(tabla.insert.call_args_list[1].kwargs['values'][2], '<1%')
