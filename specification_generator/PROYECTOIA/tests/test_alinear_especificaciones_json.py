import json
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from alinear_especificaciones_json import (
    cargar_maestro,
    limpiar_marca_acreditacion,
    metodo_maestro_canonico,
    mejor_coincidencia,
    normalizar_espacio_unidades,
    normalizar_formato_numerico,
    procesar_json,
    referencias_astm,
    referencias_astm_fuera_maestro,
    referencias_farmacopea,
    similitud_metodo,
)


MAESTRO = {
    "METODO": "ASTM D445",
    "NOMBRE IQM": "Viscosidad cinemática",
    "NOMBRE IQM INGLES": "Kinematic viscosity",
    "UNIDADES": "cSt",
    "UNITS": "mm2/s",
    "ACREDITACION NOMBRE": "ACREDITADO",
    "FILA_MAESTRO": "2",
}


class AlineacionEspecificacionesTests(unittest.TestCase):
    def test_normaliza_rangos_decimales_y_agrupacion_de_miles(self):
        casos = (
            ("50-70", "50 - 70"),
            ("-50--20", "-50 - -20"),
            ("40-40-0 (5)", "40-40-0 (5)"),
            ("5.0 mL", "5,0 mL"),
            (".25", "0,25"),
            ("0.25", "0,25"),
            ("7500.50", "7\u202f500,50"),
            ("7.500,50", "7\u202f500,50"),
            ("1.234.567,89", "1\u202f234\u202f567,89"),
            ("3 608", "3\u202f608"),
            ("0,095 7", "0,0957"),
            ("cm-1", "cm-1"),
        )
        for entrada, esperado in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(normalizar_formato_numerico(entrada), esperado)

    def test_preserva_decimales_ambiguos_como_decimales(self):
        self.assertEqual(normalizar_formato_numerico("3,351"), "3,351")
        self.assertEqual(normalizar_formato_numerico("7.500"), "7,500")

    def test_separa_cantidad_y_unidad_en_celdas_de_tabla(self):
        casos = (
            ("0,6ml", "0,6 mL"),
            ("0.6 ML", "0.6 mL"),
            ("40°C", "40 °C"),
            ("40º C", "40 °C"),
            ("4µm/6µm/14µm", "4 µm/6 µm/14 µm"),
            ("600rpm / 40kg", "600 rpm / 40 kg"),
            ("40°C / 0,6ml / ASTM D445", "40 °C / 0,6 mL / ASTM D445"),
            ("ISO 4406; ASTM D445", "ISO 4406; ASTM D445"),
        )
        for entrada, esperado in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(normalizar_espacio_unidades(entrada), esperado)

    def test_detecta_astm_que_no_existen_en_la_lista_maestra(self):
        self.assertEqual(
            referencias_astm_fuera_maestro(
                "ASTM D445 / ASTM D999", {"D445"}
            ),
            ("D999",),
        )
        self.assertEqual(
            referencias_astm_fuera_maestro("ASTM D445", {"D445"}),
            (),
        )

    def test_json_alineado_incluye_mensaje_de_astm_fuera_del_maestro(self):
        with tempfile.TemporaryDirectory() as temporal:
            origen = Path(temporal) / "entrada.json"
            destino = Path(temporal) / "salida.json"
            origen.write_text(
                json.dumps(
                    {
                        "documento": {"idioma": "es"},
                        "pruebas": [
                            {
                                "prueba": "Viscosidad cinemática",
                                "metodo": "ASTM D445 / ASTM D999",
                            },
                            {
                                "prueba": "Viscosidad cinemática",
                                "metodo": "ASTM D445",
                            },
                        ],
                    }
                ),
                encoding="utf-8",
            )

            conteo = procesar_json(origen, destino, [MAESTRO])
            alineado = json.loads(destino.read_text(encoding="utf-8"))
            alineacion = alineado["alineacion_metodos_lab"]
            primer_detalle, segundo_detalle = alineacion["detalle"]
            mensaje = (
                "Prueba 1 (Viscosidad cinemática): Referencias no encontradas "
                "en la lista maestra: ASTM D999."
            )

            self.assertEqual(conteo["alineados"], 2)
            self.assertEqual(alineado["pruebas"][0]["metodo"], "ASTM D445")
            self.assertEqual(
                primer_detalle["referencias_astm_no_encontradas"], ["D999"]
            )
            self.assertEqual(primer_detalle["mensaje_astm_no_encontradas"], mensaje)
            self.assertEqual(
                alineacion["mensajes_astm_no_encontradas"], [mensaje]
            )
            self.assertEqual(segundo_detalle["referencias_astm_no_encontradas"], [])
            self.assertEqual(segundo_detalle["mensaje_astm_no_encontradas"], "")

    def test_detalle_guarda_propuestas_como_valores_alineados_aunque_requiera_revision(self):
        with tempfile.TemporaryDirectory() as temporal:
            origen = Path(temporal) / "entrada.json"
            destino = Path(temporal) / "salida.json"
            origen.write_text(
                json.dumps(
                    {
                        "documento": {"idioma": "es"},
                        "pruebas": [
                            {
                                "prueba": "COLOR",
                                "unidades": "NA",
                                "metodo": "VISUAL",
                                "especificacion": "AMBAR",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )
            maestro_color = {
                "METODO": "USP / FEUM",
                "NOMBRE IQM": "COLOR",
                "NOMBRE IQM INGLES": "COLOR",
                "UNIDADES": "N/A",
                "UNITS": "N/A",
                "FILA_MAESTRO": "80",
            }

            procesar_json(origen, destino, [maestro_color])
            alineado = json.loads(destino.read_text(encoding="utf-8"))
            detalle = alineado["alineacion_metodos_lab"]["detalle"][0]

            self.assertEqual(detalle["estado"], "SIN_COINCIDENCIA")
            self.assertEqual(
                detalle["valores_alineados"],
                {
                    "prueba": "COLOR",
                    "unidades": "N/A",
                    "metodo": "USP / FEUM",
                    "especificacion": "AMBAR",
                    "valor_tipico": "",
                },
            )

    def test_procesar_json_normaliza_unidades_y_valor_tipico(self):
        with tempfile.TemporaryDirectory() as temporal:
            origen = Path(temporal) / "entrada.json"
            destino = Path(temporal) / "salida.json"
            origen.write_text(
                json.dumps(
                    {
                        "documento": {"idioma": "es"},
                        "notas": ["Mantener a 40°C con 0,6ml de muestra."],
                        "pruebas": [
                            {
                                "prueba": "Viscosidad cinemática",
                                "metodo": "ASTM D445",
                                "unidades": "mL NaOH 0.01 N",
                                "valor_tipico": "0.25-7500.50",
                                "especificacion": "40°C / 0,6ml",
                            },
                            {
                                "prueba": "Prueba sin coincidencia",
                                "metodo": "ASTM D999",
                                "unidades": "5.0 mL",
                                "valor_tipico": "50-70",
                                "nombre": "Temperatura 40°C",
                            },
                        ],
                    }
                ),
                encoding="utf-8",
            )
            maestro = {
                **MAESTRO,
                "UNIDADES": "mL NaOH 0.01 N",
            }

            conteo = procesar_json(origen, destino, [maestro])
            salida = json.loads(destino.read_text(encoding="utf-8"))
            alineada, sin_coincidencia = salida["pruebas"]
            detalle = salida["alineacion_metodos_lab"]["detalle"]

            self.assertEqual(conteo["alineados"], 1)
            self.assertEqual(conteo["sin_coincidencia"], 1)
            self.assertEqual(
                salida["notas"],
                ["Mantener a 40 °C con 0,6 mL de muestra."],
            )
            self.assertEqual(alineada["unidades"], "mL NaOH 0,01 N")
            self.assertEqual(alineada["valor_tipico"], "0,25 - 7\u202f500,50")
            self.assertEqual(alineada["especificacion"], "40 °C / 0,6 mL")
            self.assertEqual(sin_coincidencia["unidades"], "5,0 mL")
            self.assertEqual(sin_coincidencia["valor_tipico"], "50 - 70")
            self.assertEqual(sin_coincidencia["nombre"], "Temperatura 40 °C")
            self.assertEqual(detalle[0]["unidades_originales"], "mL NaOH 0.01 N")
            self.assertEqual(detalle[0]["valor_tipico_original"], "0.25-7500.50")
            self.assertEqual(
                detalle[0]["valores_alineados"],
                {
                    "prueba": "Viscosidad cinemática",
                    "unidades": "mL NaOH 0,01 N",
                    "metodo": "ASTM D445",
                    "especificacion": "40 °C / 0,6 mL",
                    "valor_tipico": "0,25 - 7\u202f500,50",
                },
            )

    def test_quita_asteriscos_de_acreditacion_al_inicio_del_nombre_maestro(self):
        self.assertEqual(
            limpiar_marca_acreditacion("  **  Viscosidad cinemática"),
            "Viscosidad cinemática",
        )
        self.assertEqual(
            limpiar_marca_acreditacion("Prueba con * en el nombre"),
            "Prueba con * en el nombre",
        )

    def test_salida_alineada_no_copia_asteriscos_de_nombres_maestros(self):
        casos = (
            ("es", "Viscosidad cinemática"),
            ("en", "Kinematic viscosity"),
        )
        for idioma, nombre_esperado in casos:
            with self.subTest(idioma=idioma), tempfile.TemporaryDirectory() as temporal:
                maestro_con_acreditacion = {
                    **MAESTRO,
                    "NOMBRE IQM": "** Viscosidad cinemática",
                    "NOMBRE IQM INGLES": "* Kinematic viscosity",
                }
                origen = Path(temporal) / "entrada.json"
                destino = Path(temporal) / "salida.json"
                origen.write_text(
                    json.dumps(
                        {
                            "documento": {"idioma": idioma},
                            "pruebas": [
                                {
                                    "prueba": nombre_esperado,
                                    "nombre": nombre_esperado,
                                    "metodo": "ASTM D445",
                                }
                            ],
                        }
                    ),
                    encoding="utf-8",
                )

                conteo = procesar_json(origen, destino, [maestro_con_acreditacion])
                salida = json.loads(destino.read_text(encoding="utf-8"))
                prueba = salida["pruebas"][0]
                auditoria = salida["alineacion_metodos_lab"]["detalle"][0]

                self.assertEqual(conteo["alineados"], 1)
                self.assertEqual(prueba["prueba"], nombre_esperado)
                self.assertEqual(prueba["nombre"], nombre_esperado)
                self.assertNotRegex(prueba["prueba"], r"^\s*\*")
                self.assertEqual(
                    auditoria["nombre_maestro_propuesto"],
                    "** Viscosidad cinemática",
                )
                self.assertEqual(
                    auditoria["nombre_maestro_ingles_propuesto"],
                    "* Kinematic viscosity",
                )

    def test_alinea_campos_en_ingles_y_conserva_datos_adicionales(self):
        with tempfile.TemporaryDirectory() as temporal:
            origen = Path(temporal) / "entrada.json"
            destino = Path(temporal) / "salida.json"
            origen.write_text(
                json.dumps(
                    {
                        "documento": {"idioma": "en"},
                        "pruebas": [
                            {
                                "prueba": "Kinematic viscosities",
                                "metodo": "ASTM D445",
                                "unidades": "old",
                                "especificacion": "10 min",
                                "campo_extra": "preservar",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )

            conteo = procesar_json(origen, destino, [MAESTRO])
            salida = json.loads(destino.read_text(encoding="utf-8"))
            prueba = salida["pruebas"][0]

            self.assertEqual(conteo["alineados"], 1)
            self.assertEqual(prueba["prueba"], "Kinematic viscosity")
            self.assertEqual(prueba["metodo"], "ASTM D445")
            self.assertEqual(prueba["unidades"], "mm2/s")
            self.assertEqual(prueba["campo_extra"], "preservar")
            self.assertEqual(
                salida["alineacion_metodos_lab"]["detalle"][0]["idioma"], "INGLES"
            )

    def test_alinea_nombre_en_espanol_y_sincroniza_alias(self):
        with tempfile.TemporaryDirectory() as temporal:
            origen = Path(temporal) / "entrada.json"
            destino = Path(temporal) / "salida.json"
            origen.write_text(
                json.dumps(
                    {
                        "documento": {"idioma": "es"},
                        "pruebas": [
                            {
                                "nombre": "Viscosidad cinematica",
                                "metodo": "ASTM D445",
                                "unidades": "cSt",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )

            conteo = procesar_json(origen, destino, [MAESTRO])
            prueba = json.loads(destino.read_text(encoding="utf-8"))["pruebas"][0]

            self.assertEqual(conteo["alineados"], 1)
            self.assertEqual(prueba["prueba"], "Viscosidad cinemática")
            self.assertEqual(prueba["nombre"], "Viscosidad cinemática")

    def test_metodo_estandar_distinto_no_se_autocorrige(self):
        resultado = mejor_coincidencia(
            "Viscosidad cinemática",
            "ASTM D999",
            [MAESTRO],
        )

        self.assertEqual(resultado["estado"], "SIN_COINCIDENCIA")
        self.assertEqual(resultado["puntaje_metodo"], 0.0)

    def test_reconoce_variaciones_de_referencia_astm(self):
        variaciones = (
            "D445",
            "D 445",
            "D-445",
            "ASTM D445",
            "ASTM D 445",
            "ASTM D-445",
            "ASTM D445-12",
            "ASTM D445 - 12 (2017)",
            "ASTM-445-12",
            "ASTM 445",
        )

        for metodo in variaciones:
            with self.subTest(metodo=metodo):
                self.assertEqual(referencias_astm(metodo), ("D445",))
                self.assertEqual(similitud_metodo(metodo, "ASTM D445 - 17"), 1.0)

    def test_reconoce_metodos_astm_conjugados_y_omisiones(self):
        variaciones = (
            "ASTM D445 / D7042",
            "ASTM D445 / ASTM D7042",
            "ASTM D445/D7042",
            "ASTM D-445-12 / ASTM D-7042-14",
            "ASTM D445 / 7042",
        )

        for metodo in variaciones:
            with self.subTest(metodo=metodo):
                self.assertEqual(
                    referencias_astm(metodo),
                    ("D445", "D7042"),
                )
                self.assertEqual(
                    metodo_maestro_canonico(metodo),
                    "ASTM D445 / ASTM D7042",
                )

    def test_astm_de_maestro_se_escribe_sin_edicion_y_con_servicios_adicionales(self):
        self.assertEqual(
            metodo_maestro_canonico("ASTM D445 - 17 / ASTM D7042 - 14"),
            "ASTM D445 / ASTM D7042",
        )
        self.assertEqual(
            metodo_maestro_canonico(
                "ASTM D445 - 17 / NMX-J-123-ANCE-2019 P 6,7"
            ),
            "ASTM D445",
        )
        self.assertEqual(metodo_maestro_canonico("ASTM E 1252"), "ASTM E1252")
        self.assertEqual(
            metodo_maestro_canonico(
                "ASTM 1275 / NMX-J-123-ANCE-2019 P 6,8"
            ),
            "ASTM D1275",
        )

    def test_reconoce_variaciones_de_farmacopeas_de_json(self):
        variaciones = {
            "USP 32": ("USP",),
            "CURRENT USP": ("USP",),
            "USP/NF": ("USP",),
            "USP 467": ("USP",),
            "FEUM 13/USP 32": ("FEUM", "USP"),
            "FEUM EDITION 13.0": ("FEUM",),
            "FEUM 11VA EDICION": ("FEUM",),
            "Ph. Eur 10th Edition 2.2.9.": ("Ph Eur",),
            "Ph Eur 13.0": ("Ph Eur",),
            "Farmacopea Europea13ª Edición 2.2.9.": ("Ph Eur",),
            "European Pharmacopoeia 10th Edition": ("Ph Eur",),
        }

        for metodo, referencias in variaciones.items():
            with self.subTest(metodo=metodo):
                self.assertEqual(referencias_farmacopea(metodo), referencias)

    def test_farmacopeas_se_escriben_sin_edicion_capitulo_ni_mga(self):
        self.assertEqual(metodo_maestro_canonico("USP 32"), "USP")
        self.assertEqual(
            metodo_maestro_canonico("FEUM (MGA. 0471)"),
            "FEUM",
        )
        self.assertEqual(metodo_maestro_canonico("PH. EUR"), "Ph Eur")
        self.assertEqual(
            metodo_maestro_canonico("USP (271) / FEUM (MGA. 0881)"),
            "USP / FEUM",
        )

    def test_versiones_farmacopeicas_equivalen_al_metodo_maestro(self):
        candidatos = [
            {
                **MAESTRO,
                "METODO": "FEUM (MGA 0571)",
                "NOMBRE IQM": "Límites microbianos",
                "NOMBRE IQM INGLES": "Microbial limits",
            }
        ]
        resultado = mejor_coincidencia(
            "Microbial limits",
            "FEUM EDITION 13.0",
            candidatos,
            "INGLES",
        )

        self.assertEqual(resultado["estado"], "ALINEADO")
        self.assertEqual(resultado["puntaje_metodo"], 1.0)

    def test_alineacion_canoniza_farmacopeas_y_registra_variantes(self):
        maestro_farmacopeas = {
            **MAESTRO,
            "METODO": "USP / FEUM",
            "NOMBRE IQM": "Acidez",
            "NOMBRE IQM INGLES": "Acidity",
        }
        with tempfile.TemporaryDirectory() as temporal:
            origen = Path(temporal) / "entrada.json"
            destino = Path(temporal) / "salida.json"
            origen.write_text(
                json.dumps(
                    {
                        "documento": {"idioma": "en"},
                        "pruebas": [
                            {
                                "prueba": "Acidity",
                                "metodo": "FEUM 13/USP 32",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )

            conteo = procesar_json(origen, destino, [maestro_farmacopeas])
            salida = json.loads(destino.read_text(encoding="utf-8"))
            prueba = salida["pruebas"][0]
            detalle = salida["alineacion_metodos_lab"]["detalle"][0]

            self.assertEqual(conteo["alineados"], 1)
            self.assertEqual(prueba["metodo"], "USP / FEUM")
            self.assertEqual(
                detalle["metodo_original_normalizado"],
                "FEUM / USP",
            )
            self.assertEqual(
                detalle["referencias_farmacopea_originales"],
                ["FEUM", "USP"],
            )
            self.assertEqual(
                detalle["referencias_farmacopea_maestras"],
                ["USP", "FEUM"],
            )

    def test_conserva_astm_y_farmacopea_con_formato_canonico(self):
        self.assertEqual(
            metodo_maestro_canonico(
                "Ph. Eur 13th Edition 2.2.9. / ASTM D-7042 - 14 / ASTM D445 - 17"
            ),
            "ASTM D7042 / ASTM D445 / Ph Eur",
        )

    def test_variante_letra_separada_no_cambia_designacion(self):
        self.assertEqual(referencias_astm("D 4172 B MODIFIED"), ("D4172",))
        self.assertEqual(referencias_astm("ASTM D877M"), ("D877M",))

    def test_metodo_astm_exacto_pesa_suficiente_con_nombre_abreviado(self):
        resultado = mejor_coincidencia(
            "FLASH POINT",
            "D 92",
            [
                {
                    **MAESTRO,
                    "METODO": "ASTM D92 - 18",
                    "NOMBRE IQM INGLES": (
                        "FLASH POINT AT 101,3 kPa, CLEVELAND OPEN CUP"
                    ),
                }
            ],
            "INGLES",
        )

        self.assertEqual(resultado["estado"], "ALINEADO")
        self.assertEqual(resultado["puntaje_metodo"], 1.0)

    def test_ignora_ediciones_al_distinguir_metodos_del_maestro(self):
        version_2017 = {**MAESTRO, "METODO": "ASTM D445 - 17"}
        version_2024 = {**MAESTRO, "METODO": "ASTM D445 - 24"}
        resultado = mejor_coincidencia(
            "Kinematic viscosity",
            "ASTM D445",
            [version_2017, version_2024],
            "INGLES",
        )

        self.assertEqual(resultado["estado"], "ALINEADO")
        self.assertFalse(resultado["ambiguo"])

    def test_alineacion_escribe_metodo_canonico_proveniente_del_maestro(self):
        maestro_conjugado = {
            **MAESTRO,
            "METODO": "ASTM D445 - 17 / ASTM D7042 - 14",
        }
        with tempfile.TemporaryDirectory() as temporal:
            origen = Path(temporal) / "entrada.json"
            destino = Path(temporal) / "salida.json"
            origen.write_text(
                json.dumps(
                    {
                        "documento": {"idioma": "en"},
                        "pruebas": [
                            {
                                "prueba": "Kinematic viscosity",
                                "metodo": "ASTM D-445-12 / D-7042-14",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )

            conteo = procesar_json(origen, destino, [maestro_conjugado])
            salida = json.loads(destino.read_text(encoding="utf-8"))

            self.assertEqual(conteo["alineados"], 1)
            self.assertEqual(
                salida["pruebas"][0]["metodo"],
                "ASTM D445 / ASTM D7042",
            )
            detalle = salida["alineacion_metodos_lab"]["detalle"][0]
            self.assertEqual(
                detalle["referencias_astm_originales"],
                ["D445", "D7042"],
            )
            self.assertEqual(
                detalle["referencias_astm_maestras"],
                ["D445", "D7042"],
            )

    def test_nombre_ambiguo_se_deja_para_revision(self):
        segundo_metodo = {**MAESTRO, "METODO": "ISO 3104"}
        resultado = mejor_coincidencia(
            "Viscosidad cinemática",
            "",
            [MAESTRO, segundo_metodo],
        )

        self.assertEqual(resultado["estado"], "REVISION_REQUERIDA")
        self.assertTrue(resultado["ambiguo"])

    def test_nombre_de_otra_secuencia_no_es_ambiguo(self):
        otra_secuencia = {
            **MAESTRO,
            "METODO": "ASTM D892",
            "NOMBRE IQM": "Formación de espuma (secuencia II)",
            "NOMBRE IQM INGLES": "Foaming characteristics (sequence II)",
        }
        resultado = mejor_coincidencia(
            "Foaming characteristics (sequence I)",
            "ASTM D892",
            [
                {
                    **MAESTRO,
                    "METODO": "ASTM D892",
                    "NOMBRE IQM": "Formación de espuma (secuencia I)",
                    "NOMBRE IQM INGLES": "Foaming characteristics (sequence I)",
                },
                otra_secuencia,
            ],
            "INGLES",
        )

        self.assertEqual(resultado["estado"], "ALINEADO")
        self.assertFalse(resultado["ambiguo"])

    def test_unidades_conflictivas_tambien_requieren_revision(self):
        segunda_unidad = {**MAESTRO, "UNIDADES": "mPa.s"}
        resultado = mejor_coincidencia(
            "Viscosidad cinemática",
            "",
            [MAESTRO, segunda_unidad],
        )

        self.assertEqual(resultado["estado"], "REVISION_REQUERIDA")
        self.assertTrue(resultado["ambiguo"])

    def test_no_vacia_nombre_si_falta_traduccion_maestra(self):
        sin_traduccion = {**MAESTRO, "NOMBRE IQM": ""}
        with tempfile.TemporaryDirectory() as temporal:
            origen = Path(temporal) / "entrada.json"
            destino = Path(temporal) / "salida.json"
            original = "Kinematic viscosity"
            origen.write_text(
                json.dumps(
                    {
                        "documento": {"idioma": "es"},
                        "pruebas": [{"prueba": original, "metodo": "ASTM D445"}],
                    }
                ),
                encoding="utf-8",
            )

            conteo = procesar_json(origen, destino, [sin_traduccion])
            prueba = json.loads(destino.read_text(encoding="utf-8"))["pruebas"][0]

            self.assertEqual(conteo["revision"], 1)
            self.assertEqual(prueba["prueba"], original)

    def test_nombre_vacio_no_se_alinea_por_metodo_solo(self):
        resultado = mejor_coincidencia("", "ASTM D445", [MAESTRO])

        self.assertEqual(resultado["estado"], "SIN_COINCIDENCIA")
        self.assertEqual(resultado["registro"], {})

    def test_carga_maestro_con_pandas_y_omite_filas_duplicadas(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "maestro.xlsx"
            pd.DataFrame(
                [
                    {key: value for key, value in MAESTRO.items() if key != "FILA_MAESTRO"},
                    {key: value for key, value in MAESTRO.items() if key != "FILA_MAESTRO"},
                ]
            ).to_excel(ruta, sheet_name="METODOS", index=False)

            registros = cargar_maestro(ruta)

            self.assertEqual(len(registros), 1)
            self.assertEqual(registros[0]["FILA_MAESTRO"], "2")


if __name__ == "__main__":
    unittest.main()
