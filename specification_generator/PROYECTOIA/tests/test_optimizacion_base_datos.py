import json
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

import generador_hii as hii
import gestor_especificaciones as gestor
import optimizar_base_datos as optimizacion
import reconstruir_base_datos as recuperacion


class OptimizacionTests(unittest.TestCase):
    def setUp(self):
        temporal = tempfile.TemporaryDirectory()
        self.addCleanup(temporal.cleanup)
        self.base = Path(temporal.name).resolve()
        self.db = self.base / "especificaciones.db"

    def documento(self, nombre, idioma="es", revision="01", **extras):
        ruta = self.base / "ESPECIFICACIONES_JSON" / nombre
        ruta.parent.mkdir(exist_ok=True)
        datos = {"producto": {"nombre": "Producto"}, "documento": {"idioma": idioma},
                 "control_documental": {"codigo_unico_producto": "P1", "revision": revision, "nombre_cliente": "Cliente"}, **extras}
        ruta.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")
        return ruta

    def heredada(self, es, en):
        with closing(gestor.conectar(self.db, recuperar=False)) as con:
            con.executescript("""
                CREATE TABLE hii_especificaciones (ruta TEXT PRIMARY KEY, codigo TEXT, revision TEXT, cliente TEXT, idioma TEXT);
                CREATE TABLE hii_traducciones (origen TEXT REFERENCES hii_especificaciones(ruta), destino TEXT REFERENCES hii_especificaciones(ruta));
                CREATE TABLE hii_enlaces_automaticos (origen TEXT, destino TEXT);
                CREATE TABLE hii_generadas (ruta TEXT PRIMARY KEY, especificacion TEXT REFERENCES hii_especificaciones(ruta), idioma TEXT, fecha TEXT);
                CREATE TABLE hii_estado (clave TEXT PRIMARY KEY, valor TEXT);
                CREATE TABLE recuperacion_documentos (ruta TEXT PRIMARY KEY);
                CREATE VIEW hii_versiones_idiomas AS SELECT * FROM hii_generadas;
                CREATE VIEW recuperacion_actuales AS SELECT * FROM recuperacion_documentos;
            """)
            for ruta, idioma in ((es, "ESPANOL"), (en, "INGLES")):
                con.execute("INSERT INTO hii_especificaciones VALUES (?, 'P1', '01', 'CLIENTE', ?)", (ruta.relative_to(self.base).as_posix(), idioma))
            con.execute("INSERT INTO hii_traducciones VALUES (?, ?)", (es.relative_to(self.base).as_posix(), en.relative_to(self.base).as_posix()))
            con.commit()

    def test_retira_auxiliares_y_respalda_sin_reescribir_json_ya_exportados(self):
        es = self.documento("es.json", versiones_idioma={"INGLES": [{"archivo_json": "en.json"}]})
        en = self.documento("en.json", "en")
        self.heredada(es, en)
        antes = {p: p.read_bytes() for p in (es, en)}
        resultado = optimizacion.optimizar(self.base)
        self.assertEqual(set(resultado["tablas_retiradas"]), set(optimizacion.TABLAS_AUXILIARES))
        self.assertEqual(resultado["vinculos_rescatados"], 0)
        self.assertEqual({p: p.read_bytes() for p in antes}, antes)
        with closing(sqlite3.connect(resultado["respaldo"])) as copia:
            self.assertEqual(copia.execute("SELECT COUNT(*) FROM hii_traducciones").fetchone()[0], 1)
        with closing(sqlite3.connect(self.db)) as con:
            self.assertEqual(con.execute("PRAGMA foreign_key_check").fetchall(), [])
            self.assertFalse(set(optimizacion.TABLAS_AUXILIARES) & {r[0] for r in con.execute("SELECT name FROM sqlite_master")})
        self.assertIsNone(optimizacion.optimizar(self.base)["respaldo"])
        self.assertEqual(len(list((self.base / "RESPALDOS_BASE_DATOS").iterdir())), 1)

    def test_error_json_revierte_archivos_y_tablas_de_la_migracion(self):
        es, en = self.documento("es.json"), self.documento("en.json", "en")
        self.heredada(es, en)
        antes = {p: p.read_bytes() for p in (es, en)}
        escribir = optimizacion._escribir_atomico
        llamadas = 0
        def fallar_segundo(ruta, contenido):
            nonlocal llamadas
            llamadas += 1
            if llamadas == 2:
                raise PermissionError("JSON bloqueado")
            escribir(ruta, contenido)
        with patch.object(optimizacion, "_escribir_atomico", side_effect=fallar_segundo):
            with self.assertRaises(PermissionError):
                optimizacion.optimizar(self.base)
        self.assertEqual({p: p.read_bytes() for p in antes}, antes)
        with closing(sqlite3.connect(self.db)) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM hii_traducciones").fetchone()[0], 1)
            self.assertEqual(con.execute("PRAGMA foreign_key_check").fetchall(), [])
        self.assertFalse(list(self.base.rglob(".migracion_*.tmp")))

    def test_gestor_migra_vinculos_heredados_una_sola_vez(self):
        es, en = self.documento("es.json"), self.documento("en.json", "en", revision="02")
        self.heredada(es, en)
        with closing(gestor.conectar(self.db)) as con:
            self.assertIsNone(con.execute("SELECT name FROM sqlite_master WHERE name='hii_traducciones'").fetchone())
        with hii.conectar(self.db) as con:
            self.assertEqual(hii.candidatas(con, es, "INGLES", [en]), [en])
        with patch.object(optimizacion, "_rescatar_hii", side_effect=AssertionError("No repetir")):
            with closing(gestor.conectar(self.db)):
                pass

    def test_hii_no_crea_archivos_ni_persiste_coincidencias_automaticas(self):
        es, en = self.documento("es.json"), self.documento("en.json", "en")
        antes = {p: p.read_bytes() for p in (es, en)}
        for _ in range(2):
            with hii.conectar(self.db) as con:
                self.assertEqual(hii.sincronizar_idiomas(con, es.parent), set())
                self.assertEqual(hii.candidatas(con, es, "INGLES", [en]), [en])
        self.assertEqual({p: p.read_bytes() for p in antes}, antes)
        self.assertEqual(set(self.base.iterdir()), {es.parent})

    def test_nuevo_vinculo_conserva_otras_traducciones_en_nueva_sesion(self):
        es = self.documento("es.json")
        en1 = self.documento("en1.json", "en")
        en2 = self.documento("en2.json", "en", revision="02")
        with hii.conectar(self.db) as con:
            hii.vincular(con, es, en1)
        with hii.conectar(self.db) as con:
            hii.vincular(con, es, en2, permitir_revision_distinta=True)
        with hii.conectar(self.db) as con:
            self.assertEqual(set(hii.candidatas(con, es, "INGLES", [en1, en2])), {en1, en2})
        self.assertEqual(len(json.loads(es.read_text(encoding="utf-8"))["versiones_idioma"]["INGLES"]), 2)

    def test_realineacion_y_traslado_conservan_registros_heredados(self):
        origen = self.documento("origen.json")
        destino = self.base / "alineadas" / "destino.json"
        destino.parent.mkdir()
        registro = {"archivo": "../HII/archivo.xlsm", "idioma": "ESPANOL", "fecha": "01/01/26"}
        destino.write_text(json.dumps({"hii_registros_heredados": [registro]}), encoding="utf-8")
        datos = json.loads(origen.read_text(encoding="utf-8"))
        hii.conservar_historial(datos, origen, destino)
        self.assertEqual(datos["hii_registros_heredados"], [registro])
        movida = self.base / "nuevo.json"
        hii.conservar_historial(datos, destino, movida)
        self.assertEqual(datos["hii_registros_heredados"][0]["archivo"], "HII/archivo.xlsm")

    def test_reconstruccion_no_publica_indices_auxiliares(self):
        self.documento("es.json")
        self.documento("en.json", "en")
        resultado = recuperacion.reconstruir(self.base)
        self.assertEqual(resultado["relaciones"], 1)
        with closing(sqlite3.connect(self.db)) as con:
            tablas = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            self.assertFalse(tablas & set(optimizacion.TABLAS_AUXILIARES))
            self.assertEqual(con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0], 2)

    def test_catalogo_conserva_solo_ultima_copia_y_repara_incompletas(self):
        archivo = self.base / "CODIGOS_200.xlsx"
        archivo.write_bytes(b"catalogo")
        registro = {"hoja": "Hoja1", "fila_excel": 1, "part_id": "P1", "descripcion": "Producto",
                    "part_id_clave": "P1", "descripcion_clave": "PRODUCTO"}
        with closing(gestor.conectar(self.db, recuperar=False)) as con:
            for digest in ("anterior", "vigente", "vigente"):
                gestor.registrar_catalogo(con, [registro], digest, "2026-10-19", archivo)
            self.assertEqual([r[0] for r in con.execute("SELECT hash_catalogo FROM catalogo_fuentes")], ["vigente"])
            self.assertEqual([r[0] for r in con.execute("SELECT hash_catalogo FROM catalogo_codigos")], ["vigente"])
            con.execute("DELETE FROM catalogo_codigos")
            gestor.registrar_catalogo(con, [registro], "vigente", "2026-10-19", archivo)
            self.assertEqual(con.execute("SELECT COUNT(*) FROM catalogo_codigos").fetchone()[0], 1)


if __name__ == "__main__":
    unittest.main()
