import copy
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from cambios_alineacion import comparar_valores_pruebas
from interfaz_aplicacion_especificaciones import Comparador
from traductor_especificaciones import DiccionarioTraductor


def prueba(valor, tipico="", nombre="PUNTO DE FLUIDEZ", metodo="ASTM D97", unidades="°C"):
    return {"prueba": nombre, "metodo": metodo, "unidades": unidades, "especificacion": valor, "valor_tipico": tipico}


def espec(*filas, idioma="es"):
    return {"pruebas": filas, "idioma": idioma}


class ComparadorResaltadoTests(unittest.TestCase):
    def test_d97_50_y_30_resalta_ambas_celdas_sin_mutar_documentos(self):
        docs = [espec(prueba("50", "10")), espec(prueba("30", "10"))]
        antes = copy.deepcopy(docs)
        cambios, resumen = comparar_valores_pruebas(docs)
        self.assertEqual(cambios, [{0: {"especificacion"}}, {0: {"especificacion"}}])
        self.assertEqual(resumen["celdas"], 2)
        self.assertEqual(docs, antes)

    def test_tres_especificaciones_y_prueba_solo_en_dos(self):
        docs = [espec(prueba("50", "10")), espec(prueba("50", "20")),
                espec(prueba("30", "10"), prueba("8", nombre="COLOR", metodo="ASTM D1500"))]
        cambios, resumen = comparar_valores_pruebas(docs)
        self.assertTrue(all(c[0] == {"especificacion", "valor_tipico"} for c in cambios))
        self.assertNotIn(1, cambios[2])
        docs[2] = espec(prueba("8", nombre="COLOR", metodo="ASTM D1500"))
        cambios, resumen = comparar_valores_pruebas(docs)
        self.assertEqual(cambios, [{0: {"valor_tipico"}}, {0: {"valor_tipico"}}, {}])

    def test_tres_pruebas_de_metodo_unico_con_nombres_alternativos(self):
        cambios, resumen = comparar_valores_pruebas([espec(prueba("50")), espec(prueba("50")),
                                                    espec(prueba("30", nombre="Pour point"))])
        self.assertEqual(cambios, [{0: {"especificacion"}}] * 3)
        self.assertEqual(resumen["comunes"], 1)

    def test_formatos_equivalentes_limites_y_miles_ambiguos(self):
        cambios, _ = comparar_valores_pruebas([espec(prueba("50", "50.0")), espec(prueba("50,0", "50"))])
        self.assertEqual(cambios, [{}, {}])
        cambios, _ = comparar_valores_pruebas([espec(prueba("50 MÍNIMO")), espec(prueba("50 MÁXIMO"))])
        self.assertEqual(cambios, [{0: {"especificacion"}}, {0: {"especificacion"}}])
        cambios, _ = comparar_valores_pruebas([espec(prueba("1,000")), espec(prueba("1"))])
        self.assertTrue(cambios[0])

    def test_unidades_distintas_no_se_comparan(self):
        cambios, resumen = comparar_valores_pruebas([espec(prueba("50", unidades="°C")), espec(prueba("30", unidades="°F"))])
        self.assertEqual(cambios, [{}, {}])
        self.assertEqual(resumen["unidades_distintas"], 1)

    def test_d445_distingue_temperaturas_y_no_confunde_filas(self):
        docs = [espec(prueba("50", nombre="VISCOSIDAD A 40 °C", metodo="ASTM D445"),
                      prueba("30", nombre="VISCOSIDAD A 100 °C", metodo="ASTM D445")),
                espec(prueba("20", nombre="VISCOSIDAD A 100 °C", metodo="ASTM D445"),
                      prueba("50", nombre="VISCOSIDAD A 40 °C", metodo="ASTM D445"))]
        cambios, _ = comparar_valores_pruebas(docs)
        self.assertEqual(cambios, [{1: {"especificacion"}}, {0: {"especificacion"}}])
        cambios, _ = comparar_valores_pruebas([espec(docs[0]["pruebas"][0]), espec(docs[1]["pruebas"][0])])
        self.assertEqual(cambios, [{}, {}])

    def test_metodo_unico_admite_nombres_distintos_pero_no_duplicados(self):
        cambios, _ = comparar_valores_pruebas([espec(prueba("50", nombre="Fluidez")), espec(prueba("30", nombre="Pour point"))])
        self.assertTrue(cambios[0])
        cambios, _ = comparar_valores_pruebas([espec(prueba("50"), prueba("40")), espec(prueba("30"))])
        self.assertEqual(cambios, [{}, {}])

    def test_pruebas_bilingues_del_maestro_y_equivalencias_de_valores(self):
        maestro = [{"NOMBRE IQM": "FLUIDEZ", "NOMBRE IQM INGLES": "POUR POINT"}]
        traductor = DiccionarioTraductor([{"es": "MÍNIMO", "en": "MINIMUM", "activo": True}])
        cambios, _ = comparar_valores_pruebas([
            espec(prueba("50 MÍNIMO", nombre="FLUIDEZ")),
            espec(prueba("50 MINIMUM", nombre="POUR POINT"), idioma="en")], maestro, traductor.traducir)
        self.assertEqual(cambios, [{}, {}])

    def comparador(self, paneles):
        c = Comparador.__new__(Comparador)
        c._cerrando = False
        c._maestro_comparacion = ()
        c.paneles = paneles
        c.resaltado_activo = False
        c.boton_resaltar = Mock()
        c.estado_resaltado = Mock()
        return c

    def panel(self, valor, alineada):
        vista = {**espec(prueba(valor)), "alineada": alineada, "tabla": Mock(), "filas": ("fila",), "vertical": Mock(), "horizontal": Mock()}
        return SimpleNamespace(vista=vista, resaltador=None)

    def test_boton_aparece_con_alineada_y_requiere_dos_paneles(self):
        c = self.comparador([self.panel("50", False)])
        c.actualizar_resaltado()
        c.boton_resaltar.pack_forget.assert_called_once()
        c.paneles[0].vista["alineada"] = True
        c.actualizar_resaltado()
        c.boton_resaltar.configure.assert_called_with(state="disabled", text="RESALTAR")
        c.paneles.append(self.panel("30", False))
        c.actualizar_resaltado()
        c.boton_resaltar.configure.assert_called_with(state="normal", text="RESALTAR")

    def test_alternar_resaltado_recalcula_y_limpia_las_marcas(self):
        c = self.comparador([self.panel("50", True), self.panel("30", True)])
        with patch("interfaz_aplicacion_especificaciones.ResaltadorCambios") as resaltador:
            c.alternar_resaltado()
            self.assertEqual(resaltador.call_count, 2)
            self.assertEqual(resaltador.call_args.args[1], {"fila": {"especificacion"}})
            c.paneles[1].vista["pruebas"] = (prueba("50"),)
            c.actualizar_resaltado()
            resaltador.return_value.establecer_cambios.assert_called_with({})
            c.alternar_resaltado()
            self.assertFalse(c.resaltado_activo)
            resaltador.return_value.establecer_cambios.assert_called_with({})


if __name__ == "__main__":
    unittest.main()
