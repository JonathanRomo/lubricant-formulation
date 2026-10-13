import json
import queue
import tempfile
import unittest
from types import SimpleNamespace
from pathlib import Path
from unittest.mock import Mock, patch

from interfaz_seleccion_especificaciones import (
    Especificacion,
    OrdenadorColumnasTreeview,
    SelectorEspecificaciones,
    buscar_especificaciones,
    buscar_contraparte,
    cargar_catalogo,
    especificaciones_pendientes_realineacion,
    encabezados_tabla,
    preparar_vista_previa,
    preparar_vista_previa_alineada,
    requiere_realineacion,
    VERSION_ALINEADOR,
)


class TreeviewFalso:
    def __init__(self, elementos, valores):
        self.elementos = list(elementos)
        self.valores = valores
        self.encabezados = {}

    def cget(self, opcion):
        return "Treeview" if opcion == "style" else ""

    def column(self, _columna, opcion):
        return 360 if opcion == "width" else ""

    def bind(self, *_args, **_kwargs):
        return "bind-id"

    def winfo_exists(self):
        return True

    def get_children(self, _padre=""):
        return tuple(self.elementos)

    def heading(self, columna, **opciones):
        self.encabezados[columna] = opciones

    def set(self, elemento, columna):
        return self.valores[elemento][columna]

    def move(self, elemento, _padre, indice):
        self.elementos.remove(elemento)
        self.elementos.insert(indice, elemento)

    def exists(self, elemento):
        return elemento in self.elementos


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
    def test_ordenar_una_columna_y_volver_al_orden_original(self):
        tabla = TreeviewFalso(
            ("0", "1", "2", "3"),
            {
                "0": {"prueba": "Zinc"},
                "1": {"prueba": "Ácido"},
                "2": {"prueba": "Boro"},
                "3": {"prueba": "Azul"},
            },
        )
        fuente = SimpleNamespace(measure=lambda texto: len(texto) * 8)
        with (
            patch(
                "interfaz_seleccion_especificaciones.ttk.Style",
                return_value=Mock(lookup=Mock(return_value="TkDefaultFont")),
            ),
            patch(
                "interfaz_seleccion_especificaciones.tkfont.Font",
                return_value=fuente,
            ),
        ):
            ordenador = OrdenadorColumnasTreeview(tabla, {"prueba": "Prueba"})

            ordenador.alternar("prueba")

            self.assertEqual(tabla.elementos, ["1", "3", "2", "0"])
            texto_encabezado = tabla.encabezados["prueba"]["text"]
            self.assertIn("▲", texto_encabezado)
            self.assertGreater(texto_encabezado.index("▲"), len("Prueba") + 1)
            self.assertEqual(tabla.encabezados["prueba"]["anchor"], "e")

            ordenador.alternar("prueba")

        self.assertEqual(tabla.elementos, ["0", "1", "2", "3"])
        self.assertIn("▼", tabla.encabezados["prueba"]["text"])

    def test_cambiar_columna_reemplaza_el_orden_activo_sin_perder_original(self):
        tabla = TreeviewFalso(
            ("0", "1", "2"),
            {
                "0": {"prueba": "Zinc", "metodo": "ASTM D1"},
                "1": {"prueba": "Ácido", "metodo": "ASTM D3"},
                "2": {"prueba": "Boro", "metodo": "ASTM D2"},
            },
        )
        fuente = SimpleNamespace(measure=lambda texto: len(texto) * 8)
        with (
            patch(
                "interfaz_seleccion_especificaciones.ttk.Style",
                return_value=Mock(lookup=Mock(return_value="TkDefaultFont")),
            ),
            patch(
                "interfaz_seleccion_especificaciones.tkfont.Font",
                return_value=fuente,
            ),
        ):
            ordenador = OrdenadorColumnasTreeview(
                tabla,
                {"prueba": "Prueba", "metodo": "Método"},
            )

            ordenador.alternar("prueba")
            ordenador.alternar("metodo")

            self.assertEqual(tabla.elementos, ["0", "2", "1"])
            ordenador.alternar("metodo")

        self.assertEqual(tabla.elementos, ["0", "1", "2"])

    def test_versiones_del_alineador_determinan_si_se_debe_realinear(self):
        self.assertTrue(requiere_realineacion("1.5.9", "1.6.0"))
        self.assertTrue(requiere_realineacion("", "1.6.0"))
        self.assertFalse(requiere_realineacion("1.6", "1.6.0"))
        self.assertFalse(requiere_realineacion("1.7.0", "1.6.0"))

    def test_catalogo_lee_la_version_del_alineador_guardada_en_el_json(self):
        with tempfile.TemporaryDirectory() as temporal:
            carpeta = Path(temporal)
            datos = json_especificacion()
            datos["alineacion_metodos_lab"] = {"version_programa": "1.5.2"}
            (carpeta / "producto.json").write_text(
                json.dumps(datos, ensure_ascii=False),
                encoding="utf-8",
            )

            catalogo = cargar_catalogo(carpeta)

        self.assertEqual(
            catalogo.especificaciones[0].version_alineador,
            "1.5.2",
        )

    def test_seleccion_mixta_solo_realinea_archivos_con_version_anterior(self):
        origen_viejo = Especificacion(
            ruta=Path("origen/viejo.json"),
            nombre_producto="Producto viejo",
            codigo="P-1",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        origen_actual = Especificacion(
            ruta=Path("origen/actual.json"),
            nombre_producto="Producto actual",
            codigo="P-2",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        alineada_vieja = Especificacion(
            ruta=Path("alineadas/viejo.json"),
            nombre_producto="Producto viejo",
            codigo="P-1",
            cliente="Cliente",
            revision="00",
            busqueda="",
            version_alineador="1.5.0",
        )
        alineada_actual = Especificacion(
            ruta=Path("alineadas/actual.json"),
            nombre_producto="Producto actual",
            codigo="P-2",
            cliente="Cliente",
            revision="00",
            busqueda="",
            version_alineador=VERSION_ALINEADOR,
        )
        selector = SimpleNamespace(
            paneles={
                "nuevas": SimpleNamespace(
                    catalogo=SimpleNamespace(
                        especificaciones=[alineada_vieja, alineada_actual]
                    ),
                    documentos_marcados={
                        alineada_vieja.ruta,
                        alineada_actual.ruta,
                    },
                ),
                "anteriores": SimpleNamespace(
                    catalogo=SimpleNamespace(
                        especificaciones=[origen_viejo, origen_actual]
                    )
                ),
            },
            raiz=Mock(),
            _alinear_especificaciones=Mock(),
        )

        SelectorEspecificaciones._realinear_seleccionadas(selector, "nuevas")

        selector._alinear_especificaciones.assert_called_once_with(
            [origen_viejo.ruta],
            destinos_seleccionados=[alineada_vieja.ruta],
        )

    def test_boton_realinear_se_deshabilita_si_no_hay_versiones_pendientes(self):
        boton = Mock()
        especificacion = Especificacion(
            ruta=Path("alineadas/actual.json"),
            nombre_producto="Producto",
            codigo="P-1",
            cliente="Cliente",
            revision="00",
            busqueda="",
            version_alineador=VERSION_ALINEADOR,
        )
        selector = SimpleNamespace(
            paneles={
                "nuevas": SimpleNamespace(
                    catalogo=SimpleNamespace(especificaciones=[especificacion]),
                    documentos_marcados={especificacion.ruta},
                    boton_realinear=boton,
                )
            }
        )

        SelectorEspecificaciones._actualizar_estado_accion(selector, "nuevas")

        boton.configure.assert_called_once_with(state="disabled")

    def test_worker_entrega_los_destinos_seleccionados_al_alineador(self):
        selector = SimpleNamespace(
            _etiqueta_progreso_alineacion=None,
            _barra_progreso_alineacion=None,
            _cola_alineacion=queue.Queue(),
            carpetas_json={
                "anteriores": Path("origen"),
                "nuevas": Path("alineadas"),
            },
        )
        origen = Path("origen/producto.json")
        destino = Path("alineadas/producto_actualizado.json")

        with (
            patch(
                "interfaz_seleccion_especificaciones.ejecutar_alineacion",
                return_value={"archivos": 1},
            ) as alinear,
            patch("interfaz_seleccion_especificaciones.threading.Thread") as thread,
        ):
            SelectorEspecificaciones._ejecutar_tarea_alineacion(
                selector,
                [origen],
                [destino],
            )
            thread.call_args.kwargs["target"]()

        alinear.assert_called_once()
        self.assertEqual(
            alinear.call_args.kwargs["destinos_seleccionados"],
            [destino],
        )

    def test_boton_bulk_alinea_solo_las_especificaciones_marcadas(self):
        marcada = Path("anteriores/marcada.json")
        especificacion_marcada = Especificacion(
            ruta=marcada,
            nombre_producto="Especificación marcada",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        actual_no_marcada = Especificacion(
            ruta=Path("anteriores/actual.json"),
            nombre_producto="Otra especificación",
            codigo="20100002",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        selector = SimpleNamespace(
            paneles={
                "anteriores": SimpleNamespace(
                    catalogo=SimpleNamespace(
                        especificaciones=[especificacion_marcada]
                    ),
                    documentos_marcados={marcada},
                    especificacion_actual=actual_no_marcada,
                )
            },
            _alinear_especificaciones=Mock(),
        )

        SelectorEspecificaciones._alinear_seleccionadas(selector, "anteriores")

        selector._alinear_especificaciones.assert_called_once_with([marcada])

    def test_marcar_una_especificacion_no_inicia_alineacion(self):
        ruta = Path("anteriores/producto.json")
        especificacion = Especificacion(
            ruta=ruta,
            nombre_producto="Aceite Hidráulico",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        lista = SimpleNamespace(
            nearest=Mock(return_value=0),
            yview=Mock(return_value=(0.0, 1.0)),
            yview_moveto=Mock(),
            selection_set=Mock(),
            activate=Mock(),
        )
        selector = SimpleNamespace(
            paneles={
                "anteriores": SimpleNamespace(
                    documentos_marcados=set(),
                    lista=lista,
                    sugerencias=[especificacion],
                )
            },
            _actualizar_sugerencias=Mock(),
            _actualizar_estado_accion=Mock(),
            _mostrar_previa=Mock(),
            _alinear_especificaciones=Mock(),
        )

        SelectorEspecificaciones._clic_en_lista(
            selector, "anteriores", SimpleNamespace(x=0, y=1)
        )

        selector._alinear_especificaciones.assert_not_called()

    def test_desmarcar_una_especificacion_no_vuelve_a_alinearla(self):
        ruta = Path("anteriores/producto.json")
        especificacion = Especificacion(
            ruta=ruta,
            nombre_producto="Aceite Hidráulico",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        lista = SimpleNamespace(
            nearest=Mock(return_value=0),
            yview=Mock(return_value=(0.0, 1.0)),
            yview_moveto=Mock(),
            selection_set=Mock(),
            activate=Mock(),
        )
        selector = SimpleNamespace(
            paneles={
                "anteriores": SimpleNamespace(
                    documentos_marcados={ruta},
                    lista=lista,
                    sugerencias=[especificacion],
                )
            },
            _actualizar_sugerencias=Mock(),
            _actualizar_estado_accion=Mock(),
            _mostrar_previa=Mock(),
            _alinear_especificaciones=Mock(),
        )

        SelectorEspecificaciones._clic_en_lista(
            selector, "anteriores", SimpleNamespace(x=0, y=1)
        )

        self.assertNotIn(ruta, selector.paneles["anteriores"].documentos_marcados)
        selector._alinear_especificaciones.assert_not_called()

    def test_editar_abre_la_especificacion_actual_en_el_callback(self):
        especificacion = Especificacion(
            ruta=Path("alineadas/producto.json"),
            nombre_producto="Aceite Hidráulico",
            codigo="20100001",
            cliente="Cliente",
            revision="00",
            busqueda="",
        )
        abrir_editor = Mock()
        selector = SimpleNamespace(
            paneles={
                "nuevas": SimpleNamespace(especificacion_actual=especificacion)
            },
            al_editar=abrir_editor,
        )

        SelectorEspecificaciones._editar_especificacion(selector, "nuevas")

        abrir_editor.assert_called_once_with(especificacion.ruta)

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
