import json
import queue
import tempfile
import unittest
from dataclasses import replace
from types import SimpleNamespace
from pathlib import Path
from unittest.mock import Mock, patch

import interfaz_seleccion_especificaciones as modulo
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
    parejas_idioma,
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
    def test_guardar_hii_se_habilita_por_checkbox_incluso_para_mas_de_diez(self):
        especificaciones = tuple(Especificacion(Path(f"producto_{i}.json"), f"PRODUCTO {i}",
                                               str(i), "Cliente", "01", "") for i in range(12))
        panel = SimpleNamespace(catalogo=modulo.Catalogo(especificaciones, ()), documentos_marcados=set(),
                                boton_comparar=Mock(), boton_realinear=None, boton_guardar_hii=Mock())
        selector = SelectorEspecificaciones.__new__(SelectorEspecificaciones)
        selector.paneles = {"nuevas": panel}
        for cantidad in (0, 1, 12):
            with self.subTest(cantidad=cantidad):
                panel.documentos_marcados = {e.ruta for e in especificaciones[:cantidad]}
                selector._actualizar_estado_accion("nuevas")
                panel.boton_guardar_hii.configure.assert_called_with(state="normal" if cantidad else "disabled")
        panel.boton_comparar.configure.assert_called_with(state="disabled")
        panel.documentos_marcados = {Path("ya_no_existe.json")}
        selector._actualizar_estado_accion("nuevas")
        panel.boton_guardar_hii.configure.assert_called_with(state="disabled")

    def test_guardar_hii_entrega_todas_las_marcas_y_no_la_fila_sin_marcar(self):
        especificaciones = tuple(Especificacion(Path(f"producto_{i}.json"), f"PRODUCTO {i}",
                                               str(i), "Cliente", "01", "") for i in range(3))
        marcadas = {especificaciones[0].ruta, especificaciones[2].ruta}
        selector = SelectorEspecificaciones.__new__(SelectorEspecificaciones)
        selector.al_generar_hii = Mock()
        selector.paneles = {"nuevas": SimpleNamespace(
            catalogo=modulo.Catalogo(especificaciones, ()), documentos_marcados=marcadas.copy(),
            sugerencias=[especificaciones[1]], especificacion_actual=especificaciones[1],
        )}
        selector._guardar_hii_seleccionadas("nuevas")
        selector.al_generar_hii.assert_called_once_with((especificaciones[0].ruta, especificaciones[2].ruta))
        self.assertEqual(selector.paneles["nuevas"].documentos_marcados, marcadas)
        selector.al_generar_hii.reset_mock()
        selector.paneles["nuevas"].documentos_marcados.clear()
        selector._guardar_hii_seleccionadas("nuevas")
        selector._guardar_hii_seleccionadas("anteriores")
        selector.al_generar_hii.assert_not_called()

    def test_selector_independiente_abre_hii_y_reutiliza_la_ventana(self):
        especificacion = Especificacion(Path("producto.json"), "PRODUCTO", "20100001", "Cliente", "01", "")
        selector = SelectorEspecificaciones.__new__(SelectorEspecificaciones)
        selector.raiz = Mock()
        selector.carpetas_json = {"nuevas": Path("mis_alineadas")}
        selector.al_generar_hii = None
        selector.paneles = {"nuevas": SimpleNamespace(catalogo=modulo.Catalogo((especificacion,), ()),
                                                      documentos_marcados={especificacion.ruta})}
        generador = Mock()
        generador._cerrada = False
        generador.ventana.state.return_value = "iconic"
        with patch("generador_hii.GeneradorHII", return_value=generador) as crear:
            selector._guardar_hii_seleccionadas("nuevas")
            selector._guardar_hii_seleccionadas("nuevas")
            crear.assert_called_once_with(selector.raiz, carpeta_json=Path("mis_alineadas"))
        self.assertEqual(generador.seleccionar_json_guardados.call_count, 2)
        generador.seleccionar_json_guardados.assert_called_with((especificacion.ruta,))
        generador.ventana.deiconify.assert_called()

    def test_delta_solo_aparece_con_original_y_a_la_derecha_del_boton(self):
        actual = Especificacion(Path("alineada.json"), "PRODUCTO", "20100001", "Cliente", "01", "")
        original = replace(actual, ruta=Path("original.json"))
        panel = SimpleNamespace(especificacion_actual=actual, boton_contraparte=Mock(), boton_diferencias=Mock())
        anteriores = SimpleNamespace(catalogo=SimpleNamespace(especificaciones=(original,)))
        selector = SelectorEspecificaciones.__new__(SelectorEspecificaciones)
        selector.paneles = {"nuevas": panel, "anteriores": anteriores}

        selector._actualizar_boton_contraparte("nuevas")
        panel.boton_contraparte.configure.assert_called_with(state="normal")
        panel.boton_diferencias.pack.assert_called_once_with(
            side="left", after=panel.boton_contraparte, padx=(6, 0))

        for sin_original in (True, False):
            anteriores.catalogo.especificaciones = () if sin_original else (original,)
            panel.especificacion_actual = actual if sin_original else None
            selector._actualizar_boton_contraparte("nuevas")
            panel.boton_contraparte.configure.assert_called_with(state="disabled")
        self.assertEqual(panel.boton_diferencias.pack_forget.call_count, 2)

    def test_delta_compara_original_con_instantanea_alineada_y_copia_resumen(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            antes = json_especificacion()
            despues = json_especificacion()
            despues["alineacion_metodos_lab"] = {"detalle": [{"valores_alineados": {
                **despues["pruebas"][0], "especificacion": "20 MIN"}},
                {"valores_alineados": despues["pruebas"][1]}]}
            origen, destino = base / "original.json", base / "alineada.json"
            for ruta, datos in ((origen, antes), (destino, despues)):
                ruta.write_text(json.dumps(datos), encoding="utf-8")
            originales = {ruta: ruta.read_bytes() for ruta in (origen, destino)}
            actual = Especificacion(destino, "Aceite Hidráulico", "20100001", "Cliente de Prueba", "00", "")
            original = replace(actual, ruta=origen)
            selector = SelectorEspecificaciones.__new__(SelectorEspecificaciones)
            selector.raiz = Mock()
            selector.paneles = {
                "nuevas": SimpleNamespace(especificacion_actual=actual),
                "anteriores": SimpleNamespace(catalogo=SimpleNamespace(especificaciones=(original,))),
            }
            with patch.object(modulo.ui, "Ventana") as ventana, patch.multiple(
                modulo.ttk, Frame=Mock(), Label=Mock(), Button=Mock()
            ), patch.object(modulo, "ScrolledText") as texto, patch.object(
                modulo, "cargar_maestro", return_value=()
            ), patch.object(modulo, "messagebox") as mensajes:
                selector._mostrar_diferencias_original("nuevas")
                mensajes.showerror.assert_not_called()
                resumen = texto.return_value.insert.call_args.args[1]
                self.assertIn("«10 MIN» a «20 MIN»", resumen)
                self.assertIn("Origen: Original · original.json", resumen)
                self.assertIn("Destino: Alineada · alineada.json", resumen)
                self.assertIn("Hash JSON origen (SHA-256, contenido):", resumen)
                self.assertIn("Hash JSON destino (SHA-256, contenido):", resumen)
                texto.return_value.configure.assert_called_once_with(state="disabled")
                modulo.ttk.Button.call_args.kwargs["command"]()
                ventana.return_value.clipboard_append.assert_called_once_with(resumen)
            self.assertEqual(originales, {ruta: ruta.read_bytes() for ruta in originales})

    def test_parejas_idioma_exige_relacion_vigente_y_revision_igual(self):
        es = Especificacion(Path("es.json"), "Producto", "P1", "Cliente", "01", "",
                            version_alineador=VERSION_ALINEADOR,
                            versiones_idioma=(Path("en.json").resolve(),))
        en = replace(es, ruta=Path("en.json"), idioma="INGLES", versiones_idioma=())
        self.assertEqual(parejas_idioma((es, en)), {es.ruta: en, en.ruta: es})
        self.assertEqual(parejas_idioma((es, replace(en, revision="02"))), {})
        self.assertEqual(parejas_idioma((es, replace(en, version_alineador=""))), {})
        self.assertEqual(parejas_idioma((es,)), {})
        self.assertEqual(parejas_idioma((replace(es, versiones_idioma=()), en)), {})

    def test_parejas_idioma_no_agrupa_relaciones_ambiguas(self):
        es = Especificacion(Path("es.json"), "Producto", "P1", "Cliente", "01", "",
                            version_alineador=VERSION_ALINEADOR,
                            versiones_idioma=(Path("en.json").resolve(), Path("en2.json").resolve()))
        en = replace(es, ruta=Path("en.json"), idioma="INGLES", versiones_idioma=())
        en2 = replace(en, ruta=Path("en2.json"))
        self.assertEqual(parejas_idioma((es, en, en2)), {})

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
                    boton_comparar=None,
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
            get=Mock(return_value="\u2610  Producto"),
            itemcget=Mock(return_value="blue"),
            curselection=Mock(return_value=(0,)),
            delete=Mock(),
            insert=Mock(),
            itemconfigure=Mock(),
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
            get=Mock(return_value="\u2610  Producto"),
            itemcget=Mock(return_value="blue"),
            curselection=Mock(return_value=(0,)),
            delete=Mock(),
            insert=Mock(),
            itemconfigure=Mock(),
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
            self.assertIn(("Revision", "00"), previa.control)
            self.assertIn(("Date", "01-01-24"), previa.control)
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
            self.assertIn(("Revisión", "00"), previa_anterior.control)
            self.assertIn(("Fecha", "01-01-24"), previa_anterior.control)
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

    def test_vista_previa_conserva_control_original_y_separa_notas_sin_modificar_json(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "producto.json"
            original = json_especificacion()
            ruta.write_text(json.dumps(original, ensure_ascii=False), encoding="utf-8")

            previa = preparar_vista_previa(ruta)
            despues = json.loads(ruta.read_text(encoding="utf-8"))

            self.assertEqual(previa.nombre_producto, "Aceite Hidráulico")
            self.assertIn(("Revisión", "00"), previa.control)
            self.assertIn(("Fecha", "01-01-24"), previa.control)
            self.assertEqual(len(previa.pruebas), 1)
            self.assertEqual(
                previa.notas,
                ("Usar envase cerrado", "conservar en lugar fresco"),
            )
            self.assertEqual(previa.pie_pagina, "Dirección de prueba")
            self.assertEqual(despues["control_documental"]["revision"], "00")
            self.assertEqual(len(despues["pruebas"]), 2)

    def test_vistas_sin_fecha_original_no_usan_fecha_de_generacion_ni_metadatos(self):
        ausente = object()
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "producto_19_10_26.json"
            for idioma in ("es", "en"):
                for fecha in (ausente, None, "", "   "):
                    for preparar in (preparar_vista_previa, preparar_vista_previa_alineada):
                        with self.subTest(idioma=idioma, fecha=fecha, vista=preparar.__name__):
                            datos = json_especificacion(idioma=idioma)
                            if fecha is ausente:
                                del datos["control_documental"]["fecha"]
                            else:
                                datos["control_documental"]["fecha"] = fecha
                            datos["fecha_creacion_json"] = "2026-10-19T10:00:00"
                            datos["alineacion_metodos_lab"] = {
                                "fecha": "2026-10-19T11:00:00",
                                "detalle": [{"valores_alineados": p} for p in datos["pruebas"]],
                            }
                            ruta.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")
                            original = ruta.read_bytes()
                            with patch.object(modulo, "_referencias_maestras_astm", return_value=frozenset()):
                                previa = preparar(ruta)
                            self.assertFalse(any(etiqueta in ("Fecha", "Date") for etiqueta, _ in previa.control))
                            self.assertIn(("Revision" if idioma == "en" else "Revisión", "00"), previa.control)
                            self.assertEqual(ruta.read_bytes(), original)

    def test_vista_anterior_no_modifica_los_datos_recibidos_en_memoria(self):
        datos = json_especificacion()
        datos["control_documental"].pop("fecha")
        original = json.dumps(datos, ensure_ascii=False, sort_keys=True)
        with patch.object(modulo, "_referencias_maestras_astm", return_value=frozenset()):
            previa = preparar_vista_previa(Path("producto.json"), datos=datos)
        self.assertFalse(any(etiqueta == "Fecha" for etiqueta, _ in previa.control))
        self.assertEqual(json.dumps(datos, ensure_ascii=False, sort_keys=True), original)

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
