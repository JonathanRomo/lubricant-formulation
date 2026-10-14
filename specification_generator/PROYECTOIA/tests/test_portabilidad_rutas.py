import json
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unittest
from contextlib import closing, contextmanager
from pathlib import Path
from unittest.mock import patch

import alinear_especificaciones_json as alineador
import generador_especs_hii as generador_word
import gestor_especificaciones as gestor
import registro_hii
from interfaz_aplicacion_especificaciones import _guardar_configuracion, _cargar_configuracion
from rutas_proyecto import BASE, CARPETAS, resolver_ruta, ruta_portatil, rutas_configuradas


@contextmanager
def otro_directorio(carpeta):
    anterior = Path.cwd()
    os.chdir(carpeta)
    try:
        yield
    finally:
        os.chdir(anterior)


def escribir_especificacion(ruta, idioma):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(json.dumps({
        "producto": {"nombre": "Producto"}, "documento": {"idioma": idioma},
        "control_documental": {"codigo_unico_producto": "P1", "revision": "01", "nombre_cliente": "Cliente"},
        "pruebas": [],
    }), encoding="utf-8")


class PortabilidadRutasTests(unittest.TestCase):
    def test_configuracion_relativa_sigue_al_proyecto_movido(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal).resolve()
            proyecto = raiz / "Proyecto original"
            proyecto.mkdir()
            rutas = {clave: proyecto / nombre for clave, nombre in CARPETAS.items()}
            for ruta in rutas.values():
                ruta.mkdir()
            for nombre in ("METODOS_LAB_UNIFICADO.xlsx", "TRADUCCIONES_ESPECIFICACIONES.xlsx"):
                (rutas["metodos_lab"] / nombre).write_bytes(b"lista")
            with patch("interfaz_aplicacion_especificaciones.ARCHIVO_CONFIGURACION", proyecto / "configuracion_interfaz.json"):
                _guardar_configuracion(rutas)
                self.assertEqual(_cargar_configuracion(), rutas)
            config = json.loads((proyecto / "configuracion_interfaz.json").read_text(encoding="utf-8"))
            self.assertEqual(config, CARPETAS)
            nueva = raiz / "Proyecto trasladado"
            proyecto.rename(nueva)
            with otro_directorio(raiz):
                self.assertEqual(rutas_configuradas(nueva), {clave: nueva / nombre for clave, nombre in CARPETAS.items()})

    def test_rutas_externas_se_conservan_y_relativas_no_dependen_de_cwd(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal).resolve()
            base = raiz / "proyecto"
            base.mkdir()
            externo = raiz / "externo"
            externo.mkdir()
            with otro_directorio(externo):
                self.assertEqual(resolver_ruta("FORMATOS", base), base / "FORMATOS")
                self.assertEqual(resolver_ruta("FORMATOS"), BASE / "FORMATOS")
                self.assertEqual(ruta_portatil(base / "FORMATOS", base), "FORMATOS")
                self.assertEqual(ruta_portatil(externo, base), str(externo))

    def test_configuracion_antigua_con_rutas_absolutas_recupera_carpetas_movidas(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal).resolve()
            proyecto = raiz / "anterior"
            proyecto.mkdir()
            (proyecto / "METODOS_LAB").mkdir()
            (proyecto / "configuracion_interfaz.json").write_text(json.dumps({"metodos_lab": str(proyecto / "METODOS_LAB")}), encoding="utf-8")
            nueva = raiz / "actual"
            proyecto.rename(nueva)
            self.assertEqual(rutas_configuradas(nueva)["metodos_lab"], nueva / "METODOS_LAB")

    def test_todos_los_modulos_copiados_parten_del_programa_desde_otra_carpeta(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal).resolve()
            proyecto = raiz / "Programa portable"
            proyecto.mkdir()
            modulos = []
            for fuente in BASE.glob("*.py"):
                shutil.copy2(fuente, proyecto / fuente.name)
                modulos.append(fuente.stem)
            codigo = """
import importlib, json, sys
from pathlib import Path
base = Path(sys.argv[1]).resolve()
sys.path.insert(0, str(base))
for nombre in json.loads(sys.argv[2]):
    importlib.import_module(nombre)
from rutas_proyecto import BASE, resolver_ruta, rutas_configuradas
import estilo_interfaz, interfaz_aplicacion_especificaciones, interfaz_seleccion_especificaciones, generador_hii
from gestor_especificaciones import base_app
assert BASE == base == base_app()
assert all(modulo.BASE == base for modulo in (estilo_interfaz, interfaz_aplicacion_especificaciones, interfaz_seleccion_especificaciones, generador_hii))
assert resolver_ruta('especificaciones.db') == base / 'especificaciones.db'
assert rutas_configuradas()['metodos_lab'] == base / 'METODOS_LAB'
"""
            resultado = subprocess.run([sys.executable, "-c", codigo, str(proyecto), json.dumps(modulos)],
                                       cwd=raiz, capture_output=True, text=True, timeout=30)
            self.assertEqual(resultado.returncode, 0, resultado.stderr)
            self.assertFalse((raiz / "especificaciones.db").exists())
            self.assertFalse((proyecto / "especificaciones.db").exists())

    def test_generador_word_resuelve_argumentos_relativos_al_programa(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal).resolve()
            with patch.object(generador_word, "BASE", base), patch.object(generador_word, "detectar_json", return_value=[]) as detectar, patch("builtins.print"):
                self.assertEqual(generador_word.main(["--json-dir", "origen"]), 0)
                detectar.assert_called_once_with(base / "origen")
            (base / "configuracion_interfaz.json").write_text(json.dumps({"json_anteriores": "otros_json"}), encoding="utf-8")
            with patch.object(generador_word, "BASE", base), patch.object(generador_word, "detectar_json", return_value=[]) as detectar, patch("builtins.print"):
                self.assertEqual(generador_word.main([]), 0)
                detectar.assert_called_once_with(base / "otros_json")

    def test_alineador_resuelve_base_relativa_fuera_del_proyecto(self):
        with tempfile.TemporaryDirectory() as temporal, otro_directorio(temporal):
            with patch.object(sys, "argv", ["alineador", "--base", "."]), patch.object(alineador, "ejecutar", return_value={"errores": 0}) as ejecutar:
                self.assertEqual(alineador.main(), 0)
                ejecutar.assert_called_once_with(BASE, forzar=False)

    def test_cache_alineacion_evitar_realinear_al_mover_proyecto(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal).resolve()
            base = raiz / "original"
            (base / "METODOS_LAB").mkdir(parents=True)
            (base / "METODOS_LAB" / "METODOS_LAB_UNIFICADO.xlsx").write_bytes(b"maestro")
            origen = base / "ESPECIFICACIONES_JSON" / "producto.json"
            escribir_especificacion(origen, "es")
            with patch.object(alineador, "cargar_maestro", return_value=[]), patch("builtins.print"):
                self.assertEqual(alineador.ejecutar(base)["errores"], 0)
                cache = json.loads((base / "ESPECIFICACIONES_JSON_ALINEADAS" / ".cache_alineacion").read_text(encoding="utf-8"))
                self.assertIn("ESPECIFICACIONES_JSON_ALINEADAS/producto.json", cache)
                self.assertEqual(next(iter(cache.values()))["origen"], "ESPECIFICACIONES_JSON/producto.json")
                nueva = raiz / "trasladado"
                base.rename(nueva)
                with otro_directorio(raiz):
                    resultado = alineador.ejecutar(nueva)
                self.assertEqual(resultado["omitidos"], 1)
                self.assertEqual(resultado["archivos"], 0)

    def test_relaciones_y_hii_siguen_funcionando_al_mover_proyecto(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal).resolve()
            base = raiz / "original"
            es = base / "ESPECIFICACIONES_JSON_ALINEADAS" / "es.json"
            en = es.with_name("en.json")
            escribir_especificacion(es, "es")
            escribir_especificacion(en, "en")
            excel = base / "HISTORICO_ESPECIFICACIONES" / "producto.xlsm"
            excel.parent.mkdir()
            excel.write_bytes(b"HII")
            with registro_hii.conectar(base / "especificaciones.db") as con:
                registro_hii.vincular(con, es, en)
                registro_hii.registrar(con, excel, es, "ESPANOL", "2026-10-14")
                self.assertTrue(all(not Path(fila[0]).is_absolute() for fila in con.execute("SELECT ruta FROM hii_especificaciones")))
            nueva = raiz / "trasladado"
            base.rename(nueva)
            es, en = nueva / es.relative_to(base), nueva / en.relative_to(base)
            with otro_directorio(raiz), registro_hii.conectar(nueva / "especificaciones.db") as con:
                self.assertEqual(registro_hii.candidatas(con, es, "INGLES", [en]), [en])
                fila = con.execute("SELECT * FROM hii_generadas").fetchone()
                self.assertTrue((nueva / fila["ruta"]).is_file())
                self.assertTrue((nueva / fila["especificacion"]).is_file())
                self.assertEqual(registro_hii.exportar_relaciones(con, [es, en]), set())

    def test_migracion_de_registros_antiguos_conserva_parejas_y_hii(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal).resolve()
            base = raiz / "original"
            es = base / "ESPECIFICACIONES_JSON_ALINEADAS" / "es.json"
            en = es.with_name("en.json")
            escribir_especificacion(es, "es")
            escribir_especificacion(en, "en")
            excel = base / "HISTORICO_ESPECIFICACIONES" / "producto.xlsm"
            excel.parent.mkdir()
            excel.write_bytes(b"HII")
            with registro_hii.conectar(base / "especificaciones.db", recuperar=False) as con:
                # Simular el formato anterior sin tocar ninguna base del usuario.
                con.execute("DELETE FROM hii_estado")
                con.executemany("INSERT INTO hii_especificaciones VALUES (?, ?, ?, ?, ?)",
                                [(str(es), "P1", "01", "CLIENTE", "ESPANOL"), (str(en), "P1", "01", "CLIENTE", "INGLES")])
                con.execute("INSERT INTO hii_traducciones VALUES (?, ?)", (str(es), str(en)))
                con.execute("INSERT INTO hii_enlaces_automaticos VALUES (?, ?)", (str(es), str(en)))
                con.execute("INSERT INTO hii_generadas VALUES (?, ?, ?, ?)", (str(excel), str(es), "ESPANOL", "2026-10-14"))
            nueva = raiz / "trasladado"
            base.rename(nueva)
            es, en = nueva / es.relative_to(base), nueva / en.relative_to(base)
            contenido = es.read_bytes()
            with patch.object(registro_hii, "reubicar_ruta_heredada", side_effect=[es, OSError("fallo simulado")]):
                with self.assertRaises(OSError), registro_hii.conectar(nueva / "especificaciones.db"):
                    pass
            # Una migración interrumpida no deja claves nuevas ni rompe las anteriores.
            with closing(sqlite3.connect(nueva / "especificaciones.db")) as comprobar:
                self.assertEqual(comprobar.execute("SELECT COUNT(*) FROM hii_especificaciones").fetchone()[0], 2)
                self.assertTrue(all(Path(fila[0]).is_absolute() for fila in comprobar.execute("SELECT ruta FROM hii_especificaciones")))
                self.assertEqual(list(comprobar.execute("PRAGMA foreign_key_check")), [])
            with registro_hii.conectar(nueva / "especificaciones.db") as con:
                self.assertEqual(registro_hii.candidatas(con, es, "INGLES", [en]), [en])
                for tabla in ("hii_traducciones", "hii_enlaces_automaticos", "hii_generadas"):
                    self.assertEqual(con.execute(f"SELECT COUNT(*) FROM {tabla}").fetchone()[0], 1)
                self.assertEqual(list(con.execute("PRAGMA foreign_key_check")), [])
                self.assertFalse(Path(con.execute("SELECT ruta FROM hii_generadas").fetchone()[0]).is_absolute())
            self.assertEqual(es.read_bytes(), contenido)
            with patch.object(registro_hii, "reubicar_ruta_heredada", side_effect=AssertionError("No repetir migración")):
                with registro_hii.conectar(nueva / "especificaciones.db"):
                    pass

    def test_catalogo_registrado_se_reubica_sin_cambiar_el_historial(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal).resolve()
            base = raiz / "original"
            base.mkdir()
            catalogo = base / "CODIGOS_200.xlsx"
            catalogo.write_bytes(b"catalogo")
            con = gestor.conectar(base / "especificaciones.db")
            try:
                con.execute("DELETE FROM gestor_estado WHERE clave='rutas_catalogo_portatiles'")
                con.execute("INSERT INTO catalogo_fuentes (nombre_archivo,ruta_archivo,hash_catalogo,tamano_bytes,fecha_modificacion_ns,fecha_lectura,total_registros) VALUES (?, ?, 'hash', 8, 1, '2026-10-14', 2)",
                            (catalogo.name, str(catalogo)))
                con.commit()
            finally:
                con.close()
            nueva = raiz / "trasladado"
            base.rename(nueva)
            con = gestor.conectar(nueva / "especificaciones.db")
            try:
                fila = con.execute("SELECT * FROM catalogo_fuentes").fetchone()
                self.assertEqual(fila["ruta_archivo"], "CODIGOS_200.xlsx")
                self.assertEqual(fila["total_registros"], 2)
                self.assertEqual(fila["hash_catalogo"], "hash")
            finally:
                con.close()


if __name__ == "__main__":
    unittest.main()
