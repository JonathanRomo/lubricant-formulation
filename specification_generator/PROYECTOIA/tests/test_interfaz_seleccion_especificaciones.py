import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from interfaz_seleccion_especificaciones import (
    Especificacion,
    buscar_especificaciones,
    buscar_contraparte,
    cargar_catalogo,
    encabezados_tabla,
    preparar_vista_previa,
    preparar_vista_previa_alineada,
)


def json_especificacion(
    idioma="es",
    codigo="20100001",
    nombre="Aceite Hidráulico",
    revision="00",
):
    return {
        "producto": {"nombre": nombre},
        "documento": {"idioma": idioma},
        "control_documental": {
            "codigo_unico_producto": codigo,
            "nombre_cliente": "Cliente de Prueba",
            "revision": revision,
            "fecha": "01-01-24",
        },
        "pruebas": [
            {
                "prueba": "Viscosidad cinemática",
                "unidades": "cSt",
                "metodo": "ASTM D445",
                "especificacion": "10 MIN",
                "valor_tipico": "",
            },
            {"prueba": "Notas: conservar en lugar fresco"},
        ],
        "notas": ["Usar envase cerrado"],
        "datos_institucionales": {"pie_pagina": "Dirección de prueba"},
    }


class SelectorEspecificacionesTests(unittest.TestCase):
    def test_busca_contraparte_por_nombre_de_archivo_y_metadatos(self):
        anterior = Especificacion(
            ruta=Path("anteriores/20100001 00_11_10_26.json"),
            nombre_producto="Aceite Hidráulico",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        alineada_mismo_archivo = Especificacion(
            ruta=Path("alineadas/20100001 00_11_10_26.json"),
            nombre_producto="Aceite Hidráulico alineado",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        alineada_otro_nombre = Especificacion(
            ruta=Path("alineadas/especificacion_limpia.json"),
            nombre_producto="Aceite Hidráulico",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )

        self.assertIs(
            buscar_contraparte(anterior, [alineada_mismo_archivo, alineada_otro_nombre]),
            alineada_mismo_archivo,
        )
        self.assertIs(
            buscar_contraparte(anterior, [alineada_otro_nombre]),
            alineada_otro_nombre,
        )

    def test_no_relaciona_revisiones_distintas_o_parejas_ambiguas(self):
        anterior = Especificacion(
            ruta=Path("anteriores/producto.json"),
            nombre_producto="Aceite Hidráulico",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        revision_distinta = Especificacion(
            ruta=Path("alineadas/producto.json"),
            nombre_producto="Aceite Hidráulico",
            codigo="20100001",
            cliente="Cliente",
            revision="01",
            busqueda="",
        )
        primera = Especificacion(
            ruta=Path("alineadas/primera.json"),
            nombre_producto="Aceite Hidráulico",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        segunda = Especificacion(
            ruta=Path("alineadas/segunda.json"),
            nombre_producto="Aceite Hidráulico",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )

        self.assertIsNone(buscar_contraparte(anterior, [revision_distinta]))
        self.assertIsNone(buscar_contraparte(anterior, [primera, segunda]))

    def test_busqueda_muestra_todos_los_documentos_sin_limite(self):
        especificaciones = [
            Especificacion(
                ruta=Path(f"{indice:03}.json"),
                nombre_producto=f"Aceite {indice}",
                codigo=str(indice),
                cliente="Cliente",
                revision="00",
                busqueda=f"ACEITE {indice} CLIENTE",
            )
            for indice in range(45)
        ]

        self.assertEqual(len(buscar_especificaciones(especificaciones, "")), 45)
        self.assertEqual(len(buscar_especificaciones(especificaciones, "cliente")), 45)
        self.assertEqual(
            len(buscar_especificaciones(especificaciones, "cliente", limite=10)),
            10,
        )

    def test_catalogo_busca_por_codigo_nombre_cliente_y_pruebas(self):
        with tempfile.TemporaryDirectory() as temporal:
            carpeta = Path(temporal)
            (carpeta / "producto.json").write_text(
                json.dumps(json_especificacion(), ensure_ascii=False),
                encoding="utf-8",
            )
            (carpeta / "english.json").write_text(
                json.dumps(json_especificacion(idioma="en")),
                encoding="utf-8",
            )

            catalogo = cargar_catalogo(carpeta)

            self.assertEqual(len(catalogo.especificaciones), 2)
            self.assertEqual(catalogo.especificaciones[0].codigo, "20100001")
            self.assertEqual(
                next(
                    especificacion.idioma
                    for especificacion in catalogo.especificaciones
                    if especificacion.ruta.name == "english.json"
                ),
                "INGLES",
            )
            for consulta in ("hidraulico", "viscosidad"):
                self.assertIn(
                    "producto.json",
                    {
                        especificacion.ruta.name
                        for especificacion in buscar_especificaciones(
                            catalogo.especificaciones, consulta
                        )
                    },
                )

    def test_vistas_previas_en_ingles_y_encabezados_de_tabla(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal)
            anterior = raiz / "producto_english.json"
            datos = json_especificacion(
                idioma="en",
                nombre="Hydraulic Oil",
            )
            datos["pruebas"][0].update(
                {
                    "prueba": "Kinematic viscosity",
                    "unidades": "mm2/s",
                    "metodo": "ASTM D445",
                    "especificacion": "10 MIN",
                }
            )
            anterior.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")

            previa = preparar_vista_previa(anterior)

            self.assertEqual(previa.idioma, "INGLES")
            self.assertEqual(previa.nombre_producto, "Hydraulic Oil")
            self.assertIn(("Revision", "01"), previa.control)
            self.assertEqual(previa.pruebas[0]["prueba"], "Kinematic viscosity")
            self.assertEqual(
                encabezados_tabla(previa.idioma),
                (
                    ("Test", "prueba"),
                    ("Units", "unidades"),
                    ("Method", "metodo"),
                    ("Specification", "especificacion"),
                    ("Typical value", "valor_tipico"),
                ),
            )

            alineado = raiz / "producto_english_aligned.json"
            datos["alineacion_metodos_lab"] = {
                "detalle": [
                    {
                        "valores_alineados": {
                            "prueba": "Kinematic viscosity (aligned)",
                            "unidades": "mm²/s",
                            "metodo": "ASTM D445",
                            "especificacion": "10 MIN",
                            "valor_tipico": "12.5",
                        }
                    }
                ]
            }
            alineado.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")

            previa_alineada = preparar_vista_previa_alineada(alineado)

            self.assertEqual(previa_alineada.idioma, "INGLES")
            self.assertEqual(
                previa_alineada.pruebas[0]["prueba"],
                "Kinematic viscosity (aligned)",
            )
            self.assertEqual(previa_alineada.pruebas[0]["unidades"], "mm²/s")

    def test_pestana_alineada_usa_los_valores_actualizados_de_su_json(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal)
            anteriores = raiz / "ESPECIFICACIONES_JSON"
            nuevas = raiz / "ESPECIFICACIONES_JSON_ALINEADAS"
            anteriores.mkdir()
            nuevas.mkdir()
            original = json_especificacion()
            alineado = json_especificacion()
            original["pruebas"][0].update(
                {
                    "prueba": "CORROSION AL COBRE",
                    "metodo": "D130",
                    "unidades": "A/D",
                    "especificacion": "50-70",
                    "valor_tipico": "5.0mL",
                }
            )
            alineado["pruebas"][0].update(
                {
                    "prueba": "CORROSIVIDAD AL COBRE (TIRA DE COBRE)",
                    "metodo": "ASTM D130",
                    "unidades": "N/A",
                    "especificacion": "50 - 70",
                    "valor_tipico": "5,0 mL",
                }
            )
            alineado["alineacion_metodos_lab"] = {
                "detalle": [
                    {
                        "numero_prueba": 1,
                        "valores_alineados": {
                            "prueba": "NOMBRE PROPUESTO DESDE DETALLE",
                            "unidades": "UNIDAD DESDE DETALLE",
                            "metodo": "METODO DESDE DETALLE",
                            "especificacion": "ESPECIFICACION DESDE DETALLE",
                            "valor_tipico": "VALOR TIPICO DESDE DETALLE",
                        },
                    }
                ]
            }
            (anteriores / "producto.json").write_text(
                json.dumps(original, ensure_ascii=False),
                encoding="utf-8",
            )
            (nuevas / "producto.json").write_text(
                json.dumps(alineado, ensure_ascii=False),
                encoding="utf-8",
            )

            catalogo_anterior = cargar_catalogo(anteriores)
            catalogo_nuevo = cargar_catalogo(nuevas)
            previa_anterior = preparar_vista_previa(catalogo_anterior.especificaciones[0].ruta)
            previa_nueva = preparar_vista_previa_alineada(
                catalogo_nuevo.especificaciones[0].ruta
            )

            self.assertEqual(previa_anterior.pruebas[0]["metodo"], "D130")
            self.assertEqual(previa_anterior.pruebas[0]["especificacion"], "50-70")
            self.assertEqual(previa_anterior.pruebas[0]["valor_tipico"], "5.0mL")
            self.assertIn(("Revisión", "01"), previa_anterior.control)
            self.assertEqual(
                previa_nueva.pruebas[0]["prueba"],
                "NOMBRE PROPUESTO DESDE DETALLE",
            )
            self.assertEqual(previa_nueva.pruebas[0]["metodo"], "METODO DESDE DETALLE")
            self.assertEqual(previa_nueva.pruebas[0]["unidades"], "UNIDAD DESDE DETALLE")
            self.assertEqual(previa_nueva.pruebas[0]["especificacion"], "ESPECIFICACION DESDE DETALLE")
            self.assertEqual(previa_nueva.pruebas[0]["valor_tipico"], "VALOR TIPICO DESDE DETALLE")
            self.assertIn(("Revisión", "00"), previa_nueva.control)
            self.assertIn(("Fecha", "01-01-24"), previa_nueva.control)

    def test_vista_previa_anterior_detecta_astm_ausente_del_maestro(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "producto.json"
            datos = json_especificacion()
            datos["pruebas"][0]["metodo"] = "ASTM D445 / ASTM D999"
            ruta.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")

            with patch(
                "interfaz_seleccion_especificaciones._referencias_maestras_astm",
                return_value=frozenset({"D445"}),
            ):
                previa = preparar_vista_previa(ruta)

            self.assertEqual(
                previa.mensajes_astm_no_encontradas,
                (
                    "Prueba 1 (Viscosidad cinemática): Referencias no encontradas "
                    "en la lista maestra: ASTM D999.",
                ),
            )

    def test_vista_previa_alineada_muestra_avisos_guardados_en_json(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "producto.json"
            datos = json_especificacion()
            mensaje = (
                "Prueba 1 (Viscosidad cinemática): Referencias no encontradas "
                "en la lista maestra: ASTM D999."
            )
            datos["alineacion_metodos_lab"] = {
                "mensajes_astm_no_encontradas": [mensaje],
                "detalle": [
                    {
                        "numero_prueba": 1,
                        "referencias_astm_no_encontradas": ["D999"],
                    }
                ],
            }
            ruta.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")

            with patch(
                "interfaz_seleccion_especificaciones._referencias_maestras_astm",
                side_effect=AssertionError("El aviso guardado debe tener prioridad."),
            ):
                previa = preparar_vista_previa(ruta)

            self.assertEqual(previa.mensajes_astm_no_encontradas, (mensaje,))

    def test_vista_previa_aplica_reglas_del_generador_sin_modificar_json(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "producto.json"
            original = json_especificacion()
            ruta.write_text(json.dumps(original, ensure_ascii=False), encoding="utf-8")

            previa = preparar_vista_previa(ruta)
            despues = json.loads(ruta.read_text(encoding="utf-8"))

            self.assertEqual(previa.nombre_producto, "Aceite Hidráulico")
            self.assertIn(("Revisión", "01"), previa.control)
            self.assertTrue(any(etiqueta == "Fecha" for etiqueta, _ in previa.control))
            self.assertEqual(len(previa.pruebas), 1)
            self.assertEqual(
                previa.notas,
                ("Usar envase cerrado", "conservar en lugar fresco"),
            )
            self.assertEqual(previa.pie_pagina, "Dirección de prueba")
            self.assertEqual(despues["control_documental"]["revision"], "00")
            self.assertEqual(len(despues["pruebas"]), 2)

    def test_vista_previa_muestra_especificacion_con_revision_no_numerica(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "producto.json"
            original = json_especificacion(revision="")
            original["pruebas"][0]["especificacion"] = "10 MIN"
            ruta.write_text(
                json.dumps(original, ensure_ascii=False),
                encoding="utf-8",
            )

            previa = preparar_vista_previa(ruta)

            self.assertEqual(previa.nombre_producto, "Aceite Hidráulico")
            self.assertEqual(len(previa.pruebas), 1)
            self.assertEqual(previa.pruebas[0]["especificacion"], "10 MIN")
            self.assertEqual(
                previa.mensajes_control,
                (
                    "revision no numerica en control_documental.revision: ''",
                ),
            )

    def test_vista_previa_admite_especificaciones_en_ingles(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "producto.json"
            ruta.write_text(
                json.dumps(json_especificacion(idioma="en")),
                encoding="utf-8",
            )

            previa = preparar_vista_previa(ruta)

            self.assertEqual(previa.idioma, "INGLES")


if __name__ == "__main__":
    unittest.main()
