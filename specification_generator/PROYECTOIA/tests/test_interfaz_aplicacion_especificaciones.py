import unittest
import json
import tempfile
from pathlib import Path

from docx import Document

from interfaz_aplicacion_especificaciones import (
    crear_documento_alineado_manual,
    crear_json_temporal_comparacion,
)


class DocumentoAlineadoManualTests(unittest.TestCase):
    def test_crea_detalle_con_los_valores_editados_en_el_json(self):
        datos = {
            "producto": {"nombre": "Producto de prueba"},
            "pruebas": [
                {
                    "prueba": "COLOR NUEVO",
                    "unidades": "N/A",
                    "metodo": "USP",
                    "especificacion": "AMBAR",
                    "valor_tipico": "",
                }
            ],
        }

        resultado = crear_documento_alineado_manual(datos)

        self.assertEqual(datos.get("alineacion_metodos_lab"), None)
        self.assertEqual(
            resultado["alineacion_metodos_lab"]["detalle"][0]["valores_alineados"],
            {
                "prueba": "COLOR NUEVO",
                "unidades": "N/A",
                "metodo": "USP",
                "especificacion": "AMBAR",
                "valor_tipico": "",
            },
        )
        self.assertEqual(
            resultado["alineacion_metodos_lab"]["detalle"][0]["estado"],
            "EDICION_MANUAL",
        )

    def test_rechaza_lista_pruebas_con_elementos_invalidos(self):
        with self.assertRaisesRegex(ValueError, "'pruebas' debe ser una lista de objetos"):
            crear_documento_alineado_manual(
                {"producto": {"nombre": "Producto"}, "pruebas": ["no es una fila"]}
            )

    def test_conserva_valores_alineados_existentes_distintos_de_pruebas_raiz(self):
        datos = {
            "producto": {"nombre": "Producto de prueba"},
            "pruebas": [{"prueba": "Nombre extraído", "metodo": "VISUAL"}],
            "alineacion_metodos_lab": {
                "detalle": [
                    {
                        "valores_alineados": {
                            "prueba": "Nombre maestro",
                            "unidades": "N/A",
                            "metodo": "USP",
                            "especificacion": "AMBAR",
                            "valor_tipico": "",
                        }
                    }
                ]
            },
        }

        resultado = crear_documento_alineado_manual(datos)

        self.assertEqual(
            resultado["alineacion_metodos_lab"]["detalle"][0]["valores_alineados"]["prueba"],
            "Nombre maestro",
        )

    def test_no_fabrica_valores_alineados_desde_pruebas_raiz_si_hay_auditoria_antigua(self):
        with self.assertRaisesRegex(ValueError, "no tiene valores_alineados completos"):
            crear_documento_alineado_manual(
                {
                    "producto": {"nombre": "Producto"},
                    "pruebas": [{"prueba": "Nombre extraído"}],
                    "alineacion_metodos_lab": {"detalle": [{"nombre_original": "Nombre extraído"}]},
                }
            )

    def test_rechaza_campos_con_objetos_no_serializables_en_la_tabla(self):
        with self.assertRaisesRegex(ValueError, "pruebas\\[1\\]\\.metodo"):
            crear_documento_alineado_manual(
                {
                    "producto": {"nombre": "Producto"},
                    "pruebas": [{"prueba": "Color", "metodo": {"valor": "USP"}}],
                }
            )


class ComparacionTemporalWordTests(unittest.TestCase):
    def test_extrae_word_a_json_temporal_sin_escribir_registro(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal)
            origen = base / "comparacion.docx"
            documento = Document()
            documento.sections[0].header.paragraphs[0].text = "Hydraulic Oil"
            tabla = documento.add_table(rows=2, cols=5)
            for celda, encabezado in zip(
                tabla.rows[0].cells,
                ("TEST", "UNITS", "METHOD", "SPECIFICATION", "TYPICAL VALUE"),
            ):
                celda.text = encabezado
            for celda, valor in zip(
                tabla.rows[1].cells,
                ("Kinematic viscosity", "mm2/s", "ASTM D445", "10 MIN", "12.5"),
            ):
                celda.text = valor
            documento.save(origen)

            directorio_temporal = base / "temporal"
            directorio_temporal.mkdir()
            salida = crear_json_temporal_comparacion(origen, directorio_temporal)
            datos = json.loads(salida.read_text(encoding="utf-8"))

            self.assertEqual(datos["producto"]["nombre"], "Hydraulic Oil")
            self.assertEqual(datos["documento"]["idioma"], "en")
            self.assertEqual(datos["pruebas"][0]["prueba"], "Kinematic viscosity")
            self.assertEqual(datos["pruebas"][0]["metodo"], "ASTM D445")
            self.assertFalse((base / "especificaciones.db").exists())
            self.assertTrue(origen.is_file())


if __name__ == "__main__":
    unittest.main()
