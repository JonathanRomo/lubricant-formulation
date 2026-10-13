import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

import unificar_metodos_lab_v3 as unificador


class UnificadorMetodosTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        unificador.cargar_openpyxl()

    def test_lee_metodo_ingles_de_fuentes_que_lo_incluyen(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "fuente.xlsx"
            libro = Workbook()
            hoja = libro.active
            hoja.append(
                [
                    "METODO",
                    "MÉTODO INGLÉS",
                    "NOMBRE IQM",
                    "NOMBRE IQM INGLES",
                    "UNIDADES",
                    "UNITS",
                ]
            )
            hoja.append(
                [
                    "MÉTODO INTERNO",
                    "INTERNAL METHOD",
                    "Método interno",
                    "Internal test",
                    "mg/kg",
                    "mg/kg",
                ]
            )
            libro.save(ruta)
            libro.close()

            registros = unificador.leer(ruta, "NORMAL")

        self.assertEqual(len(registros), 1)
        self.assertEqual(registros[0]["METODO"], "MÉTODO INTERNO")
        self.assertEqual(registros[0]["MÉTODO INGLÉS"], "INTERNAL METHOD")
        unificador.conservar_metodos_ingles(
            registros,
            {},
            {"INTERNAL TEST": ["OLDER TRANSLATION"]},
        )
        self.assertEqual(registros[0]["MÉTODO INGLÉS"], "INTERNAL METHOD")

    def test_conserva_traducciones_al_reconstruir_la_lista(self):
        fila = {
            "METODO": "MÉTODO INTERNO",
            "MÉTODO INGLÉS": "",
            "NOMBRE IQM": "Método interno",
            "NOMBRE IQM INGLES": "Internal test",
            "UNIDADES": "mg/kg",
            "UNITS": "mg/kg",
            "ACREDITACION NOMBRE": "NO ACREDITADO",
            "_METODO_INGLES_FUENTE": False,
        }
        clave = unificador.claves_fila(fila)

        unificador.conservar_metodos_ingles([fila], {clave: ["INTERNAL METHOD"]})

        self.assertEqual(fila["MÉTODO INGLÉS"], "INTERNAL METHOD")
        self.assertNotIn("_METODO_INGLES_FUENTE", fila)

    def test_metodo_espanol_es_fallback_solo_si_no_hay_traduccion(self):
        fila = {
            "METODO": "MÉTODO INTERNO",
            "MÉTODO INGLÉS": "",
            "NOMBRE IQM": "Método interno",
            "NOMBRE IQM INGLES": "Internal test",
            "UNIDADES": "mg/kg",
            "UNITS": "mg/kg",
            "ACREDITACION NOMBRE": "NO ACREDITADO",
            "_METODO_INGLES_FUENTE": False,
        }

        unificador.conservar_metodos_ingles([fila], {})

        self.assertEqual(fila["MÉTODO INGLÉS"], "MÉTODO INTERNO")

    def test_recupera_traduccion_por_nombre_ingles_si_cambia_metodo_fuente(self):
        fila = {
            "METODO": "MÉTODO INTERNO / METHOD FOR INTERNAL CLIENT",
            "MÉTODO INGLÉS": "",
            "NOMBRE IQM": "Método interno",
            "NOMBRE IQM INGLES": "Internal test",
            "UNIDADES": "NA",
            "UNITS": "NA",
            "ACREDITACION NOMBRE": "NO ACREDITADO",
            "_METODO_INGLES_FUENTE": False,
        }

        unificador.conservar_metodos_ingles(
            [fila],
            {},
            {"INTERNAL TEST": ["INTERNAL METHOD"]},
        )

        self.assertEqual(fila["MÉTODO INGLÉS"], "INTERNAL METHOD")


if __name__ == "__main__":
    unittest.main()
