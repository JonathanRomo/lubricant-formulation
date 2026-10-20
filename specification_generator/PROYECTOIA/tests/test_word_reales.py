"""Pruebas optativas: WORD_CORPUS_TEMP debe apuntar al corpus temporal preparado."""
import json
import os
import unittest
from pathlib import Path

import gestor_especificaciones as gestor


@unittest.skipUnless(os.environ.get("WORD_CORPUS_TEMP"), "Requiere corpus real temporal y conversión de Word")
class WordRealesTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temporal = Path(os.environ["WORD_CORPUS_TEMP"])
        cls.casos = json.loads((temporal / "casos.json").read_text(encoding="utf-8"))
        cls.catalogo = gestor.CatalogoPreparado(gestor.leer_catalogo(Path(gestor.__file__).parent / gestor.CATALOGO)[0])

    def extraer(self, nombre):
        caso = next(c for c in self.casos if Path(c["fuente"]).name == nombre)
        ruta = Path(caso["lectura"])
        with gestor.lectura_word_compartida():
            return gestor.extraer(ruta, "en" if gestor.contiene_marcadores_ingles(ruta) else "es", catalogo=self.catalogo)

    def test_transchutz_conserva_13_pruebas_y_controles(self):
        datos = self.extraer("TRANSCHUTZ SYN GL-1 SAE 140.doc")
        self.assertEqual(len(datos["pruebas"]), 13)
        self.assertEqual(datos["producto"]["nombre"], "TRANSCHUTZ SYN GL-1 SAE 140")
        self.assertEqual(datos["control_documental"]["codigo_unico_producto"], "20500194")
        self.assertEqual(datos["control_documental"]["fecha"], "08/01/2024")
        self.assertEqual(datos["control_documental"]["revision"], "2023")
        self.assertIn("ASTM D 7042", datos["pruebas"][0]["metodo"])

    def test_firestone_conserva_5_pruebas_nota_y_formato(self):
        datos = self.extraer("SCHUTZ PRO FR RB1 FIRESTONE.doc")
        self.assertEqual(len(datos["pruebas"]), 5)
        self.assertEqual(datos["documento"]["codigo_formato"], "R01-PRO-VAP-01")
        self.assertEqual(datos["documento"]["revision_formato"], "00")
        self.assertEqual(datos["control_documental"]["fecha"], "27/11/12")
        self.assertEqual(datos["control_documental"]["folio_apqp"], "VAP12K331")
        self.assertEqual(datos["control_documental"]["elaboro_y_aprobo"], "Validación de Producto")
        self.assertTrue(any("Característica" in nota and nota.startswith("◊") for nota in datos["notas"]))

    def test_ingles_conserva_12_pruebas_y_typical_literal(self):
        datos = self.extraer("SCHUTZ PRO II HVI 110-115 ENG 2021(1).doc")
        self.assertEqual(len(datos["pruebas"]), 12)
        self.assertEqual(datos["documento"]["idioma"], "en")
        self.assertEqual(datos["control_documental"]["fecha"], "08/06/2021")
        self.assertEqual(datos["control_documental"]["codigo_unico_producto"], "20300017")
        self.assertEqual(datos["control_documental"]["nombre_cliente"], "TOYODA GOSEI AUTOMOTIVE SEALING DE MEXICO")
        self.assertEqual(datos["pruebas"][1]["especificacion"], "15.86 TYPICAL")
        self.assertEqual(datos["pruebas"][1]["valor_tipico"], "")

    def test_originales_del_corpus_conservan_sus_bytes(self):
        for caso in self.casos:
            with self.subTest(archivo=caso["fuente"]):
                self.assertEqual(gestor.sha256(Path(caso["fuente"])), caso["sha256"])

    def test_documentos_de_entrada_conservan_los_campos_validos(self):
        excepciones = {
            "91600060 ESPECIFICACION REV 02.docx": "no contiene pruebas validas",
            "91600060 ESPECIFICACION REV. 03 - copia (6) - copia.docx": "no contiene pruebas validas",
            "ESPECIFICACION ENG (nombre producto).docx": "no contiene pruebas validas",
            "SCHUTZ FG + E 85 333215.doc": "duplicadas",
            "SCHUTZ FG +E 85 90 333215.doc": "duplicadas",
        }
        aceptados = 0
        for caso in self.casos:
            if caso["grupo"] != "buenos":
                continue
            nombre = Path(caso["fuente"]).name
            with self.subTest(archivo=nombre), gestor.lectura_word_compartida():
                if nombre in excepciones:
                    with self.assertRaisesRegex(ValueError, excepciones[nombre]):
                        gestor.extraer(Path(caso["lectura"]), catalogo=self.catalogo)
                    continue
                actual = gestor.extraer(Path(caso["lectura"]), catalogo=self.catalogo)
                anterior = caso["anterior"]["datos"]
                for campo, valor in anterior["control_documental"].items():
                    if valor:
                        self.assertEqual(actual["control_documental"].get(campo), valor)
                por_fila = {p["fila_origen"]: p for p in actual["pruebas"]}
                for prueba in anterior["pruebas"]:
                    etiqueta = gestor.clave(prueba["prueba"])
                    if (etiqueta.startswith(("NOTA", "NOTE")) or etiqueta in gestor.COLUMNAS
                            or etiqueta in gestor.CONTROL):
                        continue
                    fila = por_fila[prueba["fila_origen"]]
                    for campo in ("prueba", "unidades", "metodo", "especificacion", "valor_tipico"):
                        if not prueba[campo]:
                            continue  # Los nombres omitidos por el lector previo ahora se recuperan.
                        if nombre == "ESPECIFICACION SCHUTZ SOLFLUID.docx" and campo == "especificacion":
                            continue  # El Word contiene correcciones insertadas, antes ignoradas.
                        self.assertEqual(fila[campo], prueba[campo])
                aceptados += 1
        self.assertEqual(aceptados, 515)

    def test_revisar_separa_datos_extraibles_de_estructuras_ambiguas(self):
        excepciones = {
            "SCHUTZ FG 85 333215 (ESP).doc": "duplicadas",
            "SCHUTZ FG 85 90 REV 2020 ESP.doc": "duplicadas",
            "SCHUTZ PRO HVI 400-HUT01.doc": "nombre y especificacion",
            "SCHUTZ WS-LV VACCINE GRADE 55 - OCT17 BOHERINGER REV 01.doc": "sin nombre",
        }
        aceptados = 0
        regla_yoi = 0
        for caso in self.casos:
            if caso["grupo"] != "revisar":
                continue
            nombre = Path(caso["fuente"]).name
            ruta = Path(caso["lectura"])
            with self.subTest(archivo=nombre), gestor.lectura_word_compartida():
                if nombre in excepciones:
                    with self.assertRaisesRegex(ValueError, excepciones[nombre]):
                        gestor.extraer(ruta, catalogo=self.catalogo)
                    continue
                datos = gestor.extraer(ruta, catalogo=self.catalogo)
                self.assertTrue(datos["pruebas"])
                self.assertTrue(all(p["prueba"] for p in datos["pruebas"]))
                regla_yoi += bool(gestor.menciones_yves_yoi(ruta))
                aceptados += 1
        self.assertEqual(aceptados, 310)
        self.assertEqual(regla_yoi, 5)


if __name__ == "__main__":
    unittest.main()
