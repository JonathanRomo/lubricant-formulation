import json
import os
import tempfile
import threading
import unittest
from concurrent.futures import CancelledError
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch

import generador_hii as cache_hii
import interfaz_seleccion_especificaciones as selector
from especificaciones_wip import ruta_estado


def escribir(carpeta, nombre="producto.json", idioma="es", **extras):
    datos = {"producto": {"nombre": "Aceite Hidráulico"}, "documento": {"idioma": idioma},
             "control_documental": {"codigo_unico_producto": "P1", "revision": "01", "nombre_cliente": "Cliente"},
             "pruebas": [{"prueba": "Viscosidad", "metodo": "ASTM D445"}], **extras}
    ruta = carpeta / nombre
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")
    return ruta


class CacheCatalogoTests(unittest.TestCase):
    def setUp(self):
        temporal = tempfile.TemporaryDirectory()
        self.addCleanup(temporal.cleanup)
        self.base = Path(temporal.name).resolve()
        self.carpeta = self.base / "alineadas"
        self.es = escribir(self.carpeta, versiones_idioma={"INGLES": [{"archivo_json": "sub/en.JSON"}]})
        self.en = escribir(self.carpeta, "sub/en.JSON", "en")

    def test_cache_equivale_al_catalogo_y_no_relee_json_sin_cambios(self):
        esperado = selector.cargar_catalogo(self.carpeta)
        antes = {p: p.read_bytes() for p in (self.es, self.en)}
        self.assertEqual(cache_hii.cargar_catalogo_hii(self.carpeta), esperado)
        with patch.object(selector, "leer_datos_json", side_effect=AssertionError("No releer JSON")), \
             patch("generador_hii.conectar", side_effect=AssertionError("No abrir SQLite")):
            self.assertEqual(cache_hii.cargar_catalogo_hii(self.carpeta), esperado)
        self.assertEqual({p: p.read_bytes() for p in antes}, antes)

    def test_relee_solo_modificados_y_detecta_altas_bajas_y_renombres(self):
        cache_hii.cargar_catalogo_hii(self.carpeta)
        datos = json.loads(self.es.read_text(encoding="utf-8"))
        datos["producto"]["nombre"] = "Producto actualizado"
        self.es.write_text(json.dumps(datos), encoding="utf-8")
        nuevo = escribir(self.carpeta, "nuevo.json")
        with patch.object(selector, "leer_datos_json", wraps=selector.leer_datos_json) as leer:
            actual = cache_hii.cargar_catalogo_hii(self.carpeta)
            self.assertEqual({c.args[0] for c in leer.call_args_list}, {self.es, nuevo})
        self.assertEqual(actual, selector.cargar_catalogo(self.carpeta))
        nuevo.unlink()
        renombrada = self.en.with_name("renombrada.JSON")
        self.en.rename(renombrada)
        with patch.object(selector, "leer_datos_json", wraps=selector.leer_datos_json) as leer:
            actual = cache_hii.cargar_catalogo_hii(self.carpeta)
            self.assertEqual([c.args[0] for c in leer.call_args_list], [renombrada])
        self.assertEqual({e.ruta for e in actual.especificaciones}, {self.es, renombrada})

    def test_mismo_tamano_con_mtime_distinto_invalida_cache_y_wip_tambien(self):
        cache_hii.cargar_catalogo_hii(self.carpeta)
        info = self.es.stat()
        self.es.write_bytes(self.es.read_bytes().replace(b"P1", b"P2"))
        os.utime(self.es, ns=(info.st_atime_ns, info.st_mtime_ns + 1_000_000))
        actual = cache_hii.cargar_catalogo_hii(self.carpeta)
        self.assertEqual(next(e.codigo for e in actual.especificaciones if e.ruta == self.es), "P2")
        ruta_estado(self.es).write_text("borrador", encoding="utf-8")
        actual = cache_hii.cargar_catalogo_hii(self.carpeta)
        self.assertTrue(next(e.en_progreso for e in actual.especificaciones if e.ruta == self.es))

    def test_cache_danada_o_incompatible_se_reconstruye(self):
        cache = self.carpeta / cache_hii.NOMBRE_CACHE
        esperado = selector.cargar_catalogo(self.carpeta)
        for contenido in ("{", "[]", '{"version": 999, "archivos": {}}'):
            with self.subTest(contenido=contenido):
                cache.write_text(contenido, encoding="utf-8")
                self.assertEqual(cache_hii.cargar_catalogo_hii(self.carpeta), esperado)
        datos = json.loads(cache.read_text(encoding="utf-8"))
        datos["archivos"]["producto.json"]["datos"]["versiones_idioma"] = 123
        cache.write_text(json.dumps(datos), encoding="utf-8")
        self.assertEqual(cache_hii.cargar_catalogo_hii(self.carpeta), esperado)

    def test_json_invalido_se_reporta_y_se_recupera_cuando_se_corrige(self):
        self.es.write_text("{", encoding="utf-8")
        catalogo = cache_hii.cargar_catalogo_hii(self.carpeta)
        self.assertEqual(len(catalogo.errores), 1)
        self.assertEqual(len(catalogo.especificaciones), 1)
        escribir(self.carpeta)
        self.assertEqual(cache_hii.cargar_catalogo_hii(self.carpeta).errores, ())

    def test_cache_no_escribible_y_cancelacion_no_modifican_documentos(self):
        with patch.object(cache_hii, "escribir_atomico", side_effect=PermissionError("solo lectura")):
            self.assertEqual(len(cache_hii.cargar_catalogo_hii(self.carpeta).especificaciones), 2)
        cancelar = threading.Event()
        cancelar.set()
        with patch.object(selector, "leer_datos_json", side_effect=AssertionError("Carga cancelada")):
            with self.assertRaises(CancelledError):
                cache_hii.cargar_catalogo_hii(self.carpeta, cancelar=cancelar)
        self.assertFalse((self.carpeta / cache_hii.NOMBRE_CACHE).exists())

    def test_cache_portable_reubica_especificaciones_y_relaciones(self):
        cache_hii.cargar_catalogo_hii(self.carpeta)
        movida = self.base / "otra carpeta"
        self.carpeta.rename(movida)
        with patch.object(selector, "leer_datos_json", side_effect=AssertionError("No releer")):
            catalogo = cache_hii.cargar_catalogo_hii(movida)
        es = next(e for e in catalogo.especificaciones if e.idioma == "ESPANOL")
        self.assertEqual(es.ruta, movida / self.es.name)
        self.assertEqual(es.versiones_idioma, (movida / "sub/en.JSON",))

    def test_cancelar_durante_lectura_no_publica_cache_parcial(self):
        cache_hii.cargar_catalogo_hii(self.carpeta)
        cache = self.carpeta / cache_hii.NOMBRE_CACHE
        antes = cache.read_bytes()
        escribir(self.carpeta, producto={"nombre": "Producto cambiado"})
        cancelar = threading.Event()
        lector = selector.leer_datos_json
        def leer(ruta):
            datos = lector(ruta)
            cancelar.set()
            return datos
        with patch.object(selector, "leer_datos_json", side_effect=leer):
            with self.assertRaises(CancelledError):
                cache_hii.cargar_catalogo_hii(self.carpeta, cancelar=cancelar)
        self.assertEqual(cache.read_bytes(), antes)

    def test_carpeta_vacia_tambien_reconstruye_cache_danada(self):
        vacia = self.base / "vacia"
        vacia.mkdir()
        cache = vacia / cache_hii.NOMBRE_CACHE
        cache.write_text("{", encoding="utf-8")
        self.assertEqual(cache_hii.cargar_catalogo_hii(vacia), selector.Catalogo((), ()))
        self.assertEqual(json.loads(cache.read_text(encoding="utf-8"))["archivos"], {})

    def test_candidatas_respetan_revision_cliente_enlaces_y_ambiguedad(self):
        catalogo = cache_hii.cargar_catalogo_hii(self.carpeta)
        es = next(e for e in catalogo.especificaciones if e.idioma == "ESPANOL")
        en = next(e for e in catalogo.especificaciones if e.idioma == "INGLES")
        otra = replace(en, ruta=self.carpeta / "otra.json", revision="02")
        # Un enlace explícito puede cruzar revisiones y prevalece sobre coincidencias automáticas.
        self.assertEqual(cache_hii.candidatas_catalogo([otra, en], replace(es, versiones_idioma=(otra.ruta,)), "INGLES"), [otra.ruta])
        sin_enlaces = replace(es, versiones_idioma=())
        self.assertEqual(cache_hii.candidatas_catalogo([otra, en], sin_enlaces, "INGLES"), [en.ruta])
        self.assertEqual(cache_hii.candidatas_catalogo([replace(en, cliente="Otro")], es, "INGLES"), [])
        self.assertEqual(cache_hii.candidatas_catalogo([replace(en, codigo="P2")], es, "INGLES"), [])
        self.assertEqual(cache_hii.candidatas_catalogo([replace(otra, versiones_idioma=(es.ruta,))], sin_enlaces, "INGLES"), [otra.ruta])
        otra = replace(otra, revision=en.revision)
        self.assertEqual(len(cache_hii.candidatas_catalogo([en, otra], sin_enlaces, "INGLES")), 2)
        self.assertEqual(cache_hii.candidatas_catalogo([en], replace(sin_enlaces, revision=""), "INGLES"), [])


if __name__ == "__main__":
    unittest.main()
