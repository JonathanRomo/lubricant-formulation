import copy
import hashlib
import json
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

import reconstruir_base_datos as duplicados
import reconstruir_base_datos as recuperacion
from rutas_proyecto import rutas_configuradas


def datos(**extras):
    return dict({
        "producto": {"nombre": "PRODUCTO"},
        "documento": {"tipo": "ESPECIFICACION", "idioma": "es"},
        "control_documental": {"codigo_unico_producto": "20100001", "revision": "00", "nombre_cliente": "Cliente"},
        "pruebas": [{"numero": 1, "fila_origen": 2, "prueba": "VISCOSIDAD", "metodo": "ASTM D445",
                     "unidades": "mm²/s", "especificacion": "50 MÁXIMO", "valor_tipico": "40"}],
        "notas": ["Nota uno"], "datos_institucionales": {"pie_pagina": "Empresa"},
        "fecha_creacion_json": "2026-01-01T10:00:00",
    }, **extras)


class DuplicadosTests(unittest.TestCase):
    def setUp(self):
        temporal = tempfile.TemporaryDirectory()
        self.addCleanup(temporal.cleanup)
        self.base = Path(temporal.name).resolve()
        self.rutas = rutas_configuradas(self.base)
        self.db = self.base / "especificaciones.db"
        self.destino = self.base / "REVISAR" / "DUPLICADOS"

    def json(self, nombre, contenido=None, tipo="json_anteriores"):
        ruta = self.rutas[tipo] / nombre
        ruta.parent.mkdir(parents=True, exist_ok=True)
        ruta.write_text(json.dumps(contenido or datos(), ensure_ascii=False), encoding="utf-8")
        return ruta

    def word(self, nombre, contenido=b"word"):
        ruta = self.rutas["especificaciones_viejas"] / nombre
        ruta.parent.mkdir(parents=True, exist_ok=True)
        ruta.write_bytes(contenido)
        return ruta

    def extraer(self, nombre, contenido=b"word", **extras):
        word = self.word(nombre + ".docx", contenido)
        documento = datos(archivo_origen=word.relative_to(self.base).as_posix(),
                          hash_sha256_origen=hashlib.sha256(contenido).hexdigest(), **extras)
        return word, self.json(nombre + ".json", documento)

    def revisar(self):
        return recuperacion.revisar_duplicados(self.base)

    def test_compara_datos_sin_nombre_archivo_metadatos_ni_espacios(self):
        original = self.json("a.json")
        copia = datos(fecha_creacion_json="2026-02-01", archivo_json="otro.json", revision_json="10")
        copia["producto"]["nombre"] = "  PRODUCTO\n  "
        copia["pruebas"][0].update(numero=12, fila_origen=90)
        repetida = self.json("b.json", copia)
        resumen = self.revisar()
        self.assertEqual(resumen["json_duplicados"], 1)
        self.assertTrue(original.is_file())
        self.assertFalse(repetida.exists())
        self.assertTrue((self.destino / repetida.name).is_file())
        self.assertFalse(self.db.exists())
        self.assertFalse(any(p.is_dir() for p in self.destino.iterdir()))
        self.assertEqual(self.revisar()["json_duplicados"], 0)

    def test_diferencias_significativas_se_conservan(self):
        base = datos()
        variantes = []
        for campo, valor in (("especificacion", "50 MÍNIMO"), ("unidades", "cP"), ("metodo", "ASTM D97"), ("valor_tipico", "41")):
            otra = copy.deepcopy(base)
            otra["pruebas"][0][campo] = valor
            variantes.append(otra)
        for campo, valor in (("revision", "01"), ("nombre_cliente", "Otro"), ("codigo_unico_producto", "20100002")):
            otra = copy.deepcopy(base)
            otra["control_documental"][campo] = valor
            variantes.append(otra)
        variantes.extend([datos(notas=["Otra nota"]), datos(documento={"tipo": "ESPECIFICACION", "idioma": "en"}),
                          datos(datos_institucionales={"pie_pagina": "Otra empresa"})])
        rutas = [self.json("base.json", base)]
        rutas += [self.json(f"variante_{i}.json", d) for i, d in enumerate(variantes)]
        self.assertEqual(self.revisar()["json_duplicados"], 0)
        self.assertTrue(all(r.is_file() for r in rutas))

    def test_mueve_word_y_json_y_deja_el_ya_registrado(self):
        conservado, original = self.extraer("z_anterior", b"primero")
        recuperacion.reconstruir(self.base)
        word, repetido = self.extraer("a_nuevo", b"otro archivo binario", fecha_creacion_json="2025-01-01")
        resumen = self.revisar()
        self.assertEqual((resumen["json_duplicados"], resumen["word_duplicados"]), (1, 1))
        self.assertTrue(conservado.is_file())
        self.assertTrue(original.is_file())
        self.assertFalse(word.exists())
        self.assertFalse(repetido.exists())
        self.assertEqual((self.destino / word.name).read_bytes(), b"otro archivo binario")
        with closing(sqlite3.connect(self.db)) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0], 1)
            self.assertEqual(con.execute("PRAGMA foreign_key_check").fetchall(), [])

    def test_reconstruccion_aparta_duplicados_y_solo_registra_conservado(self):
        a, original = self.extraer("a", b"uno")
        b, repetido = self.extraer("b", b"dos")
        resultado = recuperacion.reconstruir(self.base)
        self.assertEqual((resultado["documentos"], resultado["versiones"]), (1, 1))
        self.assertEqual(resultado["duplicados"]["word_duplicados"], 1)
        self.assertTrue(a.is_file())
        self.assertTrue(original.is_file())
        self.assertFalse(b.exists())
        self.assertFalse(repetido.exists())
        with closing(sqlite3.connect(self.db)) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM archivos").fetchone()[0], 1)
            self.assertEqual(con.execute("SELECT archivo_json FROM versiones").fetchone()[0], "a.json")

    def test_word_exacto_sin_json_y_nombres_distintos(self):
        a = self.word("a.doc", b"identico")
        b = self.word("b.doc", b"identico")
        c = self.word("c.doc", b"distinto")
        resultado = self.revisar()
        self.assertEqual(resultado["word_duplicados"], 1)
        self.assertTrue(a.exists())
        self.assertTrue(c.exists())
        self.assertFalse(b.exists())

    def test_mismo_word_compartido_no_se_mueve(self):
        word, _ = self.extraer("a")
        copia = json.loads((self.rutas["json_anteriores"] / "a.json").read_text(encoding="utf-8"))
        b = self.json("b.json", copia)
        self.assertEqual(self.revisar()["word_duplicados"], 0)
        self.assertTrue(word.exists())
        self.assertFalse(b.exists())

    def test_word_modificado_tras_extraccion_no_se_mueve(self):
        self.extraer("a")
        word, _ = self.extraer("b", b"viejo")
        word.write_bytes(b"modificado")
        resumen = self.revisar()
        self.assertTrue(word.exists())
        self.assertEqual(resumen["word_duplicados"], 0)
        self.assertTrue(resumen["advertencias"])

    def test_incompletos_y_wip_no_se_mueven(self):
        original = self.json("a.json")
        wip = self.json("b.json")
        wip.with_name("." + wip.name + ".wip").touch()
        incompleto = self.json("c.json", datos(pruebas=[]))
        otra = self.json("d.json", datos(pruebas=[]))
        self.assertEqual(self.revisar()["json_duplicados"], 0)
        self.assertTrue(all(p.exists() for p in (original, wip, incompleto, otra)))

    def test_original_y_alineado_no_se_comparan_entre_si(self):
        original = self.json("a.json")
        alineada = self.json("a.json", tipo="json_alineadas")
        self.assertEqual(self.revisar()["json_duplicados"], 0)
        self.assertTrue(original.exists())
        self.assertTrue(alineada.exists())

    def test_ediciones_de_instantanea_alineada_se_conservan(self):
        a = datos()
        a["alineacion_metodos_lab"] = {"detalle": [{"valores_alineados": copy.deepcopy(a["pruebas"][0])}]}
        b = copy.deepcopy(a)
        b["alineacion_metodos_lab"]["detalle"][0]["valores_alineados"]["especificacion"] = "OTRO VALOR"
        self.json("a.json", a, "json_alineadas")
        self.json("b.json", b, "json_alineadas")
        self.assertEqual(self.revisar()["json_duplicados"], 0)

    def test_colision_no_sobrescribe_y_no_crea_subcarpetas(self):
        self.json("a.json")
        b = self.json("b.json")
        self.destino.mkdir(parents=True)
        existente = self.destino / b.name
        existente.write_text("existente", encoding="utf-8")
        self.revisar()
        self.assertEqual(existente.read_text(encoding="utf-8"), "existente")
        self.assertTrue((self.destino / "b_01.json").is_file())

    def test_error_al_publicar_revierte_movimientos_y_conserva_base(self):
        self.extraer("a")
        recuperacion.reconstruir(self.base)
        word, repetido = self.extraer("b", b"otra version del zip")
        with patch.object(recuperacion, "_publicar", side_effect=OSError("no publicar")):
            with self.assertRaises(OSError):
                recuperacion.reconstruir(self.base, reemplazar=True)
        self.assertTrue(word.exists())
        self.assertTrue(repetido.exists())
        self.assertEqual(list(self.destino.iterdir()), [])
        with closing(sqlite3.connect(self.db)) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0], 1)

    def test_error_tras_primer_movimiento_revierte_todos(self):
        self.extraer("a")
        word, repetido = self.extraer("b", b"otro binario")
        renombrar = Path.rename
        llamadas = []
        def fallo(ruta, destino):
            if destino.parent == self.destino:
                llamadas.append(ruta)
                if len(llamadas) == 2:
                    raise PermissionError("archivo en uso")
            return renombrar(ruta, destino)
        with patch.object(Path, "rename", fallo), self.assertRaises(PermissionError):
            self.revisar()
        self.assertTrue(word.is_file())
        self.assertTrue(repetido.is_file())
        self.assertFalse(any(self.destino.iterdir()))

    def test_archivo_cambiado_durante_revision_aborta(self):
        self.json("a.json")
        b = self.json("b.json")
        plan = duplicados.planificar_duplicados(self.base, self.rutas, self.db)
        b.write_text("modificado", encoding="utf-8")
        with self.assertRaisesRegex(RuntimeError, "cambió"):
            with duplicados.aplicar_plan(plan):
                pass
        self.assertTrue(b.exists())

    def test_elimina_solo_indices_de_copias_y_sus_relaciones(self):
        # Construir una base histórica anterior a esta función sin apartar las copias.
        self.extraer("a", b"word a")
        word, b = self.extraer("b", b"word b")
        with patch.object(duplicados, "planificar_duplicados", return_value=duplicados.PlanDuplicados(self.destino)):
            recuperacion.reconstruir(self.base)
        with closing(sqlite3.connect(self.db)) as con:
            # Simular las tablas persistentes de una versión anterior.
            con.executescript("""
                CREATE TABLE hii_especificaciones (ruta TEXT PRIMARY KEY);
                CREATE TABLE hii_generadas (ruta TEXT, especificacion TEXT, idioma TEXT, fecha TEXT);
                CREATE TABLE hii_traducciones (origen TEXT, destino TEXT);
                CREATE TABLE recuperacion_documentos (ruta TEXT PRIMARY KEY);
            """)
            for ruta in (b.parent / "a.json", b):
                clave = ruta.relative_to(self.base).as_posix()
                con.execute("INSERT INTO hii_especificaciones VALUES (?)", (clave,))
                con.execute("INSERT INTO recuperacion_documentos VALUES (?)", (clave,))
            con.execute("INSERT INTO hii_generadas VALUES ('generada.xlsm', ?, 'ESPANOL', '01/01/26')", (b.relative_to(self.base).as_posix(),))
            con.execute("INSERT INTO hii_traducciones VALUES (?, ?)", ((b.parent / "a.json").relative_to(self.base).as_posix(), b.relative_to(self.base).as_posix()))
            con.commit()
        self.revisar()
        with closing(sqlite3.connect(self.db)) as con:
            for tabla in ("archivos", "versiones", "recuperacion_documentos", "hii_especificaciones"):
                self.assertEqual(con.execute(f"SELECT COUNT(*) FROM {tabla}").fetchone()[0], 1)
            self.assertEqual(con.execute("SELECT COUNT(*) FROM hii_generadas").fetchone()[0], 0)
            self.assertEqual(con.execute("SELECT COUNT(*) FROM hii_traducciones").fetchone()[0], 0)
            self.assertEqual(con.execute("PRAGMA foreign_key_check").fetchall(), [])
        self.assertFalse(word.exists())

    def test_repara_enlaces_en_json_conservados(self):
        self.json("a.json")
        self.json("b.json")
        en = self.json("en.json", datos(documento={"tipo": "ESPECIFICACION", "idioma": "en"},
            versiones_idioma={"es": [{"archivo_json": "b.json"}]}))
        self.revisar()
        d = json.loads(en.read_text(encoding="utf-8"))
        self.assertEqual(d["versiones_idioma"]["es"][0]["archivo_json"], "a.json")

    def test_depuracion_de_cache_no_conserva_entradas_de_copias(self):
        self.json("a.json")
        self.json("b.json")
        cache = self.rutas["json_alineadas"] / ".cache_alineacion"
        cache.parent.mkdir()
        cache.write_text(json.dumps({
            "ESPECIFICACIONES_JSON_ALINEADAS/a.json": {"origen": "ESPECIFICACIONES_JSON/a.json"},
            "ESPECIFICACIONES_JSON_ALINEADAS/b.json": {"origen": "ESPECIFICACIONES_JSON/b.json"},
        }), encoding="utf-8")
        self.revisar()
        self.assertEqual(list(json.loads(cache.read_bytes())), ["ESPECIFICACIONES_JSON_ALINEADAS/a.json"])

    def test_alineados_conservan_el_mismo_origen_que_el_original(self):
        word_a, a = self.extraer("a", b"word a")
        word_b, b = self.extraer("b", b"word b", fecha_creacion_json="2026-02-01")
        for ruta, fecha in ((a, "2026-03-01"), (b, "2025-01-01")):
            contenido = json.loads(ruta.read_bytes())
            contenido["fecha_creacion_json"] = fecha
            self.json(ruta.name, contenido, "json_alineadas")
        resultado = self.revisar()
        self.assertEqual((resultado["json_duplicados"], resultado["word_duplicados"]), (2, 1))
        self.assertTrue(word_a.is_file())
        self.assertFalse(word_b.exists())
        self.assertTrue((self.rutas["json_alineadas"] / "a.json").exists())
        self.assertFalse((self.rutas["json_alineadas"] / "b.json").exists())

    def test_error_sqlite_revierte_archivos_e_indices(self):
        self.extraer("a", b"word a")
        word, repetido = self.extraer("b", b"word b")
        with patch.object(duplicados, "planificar_duplicados", return_value=duplicados.PlanDuplicados(self.destino)):
            recuperacion.reconstruir(self.base)
        conectar = sqlite3.connect
        class Conexion:
            def __init__(self, *args, **kwargs):
                self.real = conectar(*args, **kwargs)
            def __getattr__(self, nombre):
                return getattr(self.real, nombre)
            def commit(self):
                raise sqlite3.OperationalError("fallo al confirmar")
        with patch.object(recuperacion.sqlite3, "connect", Conexion):
            with self.assertRaises(sqlite3.OperationalError):
                self.revisar()
        self.assertTrue(word.is_file())
        self.assertTrue(repetido.is_file())
        with closing(conectar(self.db)) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0], 2)

    def test_soporta_destino_en_otra_unidad(self):
        import errno
        a = self.word("a.docx", b"misma copia")
        b = self.word("b.docx", b"misma copia")
        with patch.object(Path, "rename", side_effect=OSError(errno.EXDEV, "otra unidad")):
            self.assertEqual(self.revisar()["word_duplicados"], 1)
        self.assertTrue(a.is_file())
        self.assertFalse(b.exists())
        self.assertEqual((self.destino / b.name).read_bytes(), a.read_bytes())

    def test_extractor_no_reprocesa_archivos_movidos_por_recuperacion_automatica(self):
        import gestor_especificaciones as gestor
        from openpyxl import Workbook
        libro = Workbook()
        libro.active.append(["ID", "Description"])
        libro.active.append(["20100001", "PRODUCTO"])
        catalogo = self.base / "CODIGOS_200.xlsx"
        libro.save(catalogo)
        libro.close()
        huella_catalogo = hashlib.sha256(catalogo.read_bytes()).hexdigest()
        self.extraer("a", b"word a", hash_sha256_catalogo=huella_catalogo, version_lector=gestor.VERSION_LECTOR)
        word, repetido = self.extraer("b", b"word b", hash_sha256_catalogo=huella_catalogo, version_lector=gestor.VERSION_LECTOR)
        with patch.object(gestor, "preparar", side_effect=AssertionError("No extraer nuevamente")), patch("builtins.print"):
            resultado = gestor.procesar(self.base)
        self.assertEqual((resultado["ignorados"], resultado["errores"], resultado["nuevos"]), (1, 0, 0))
        self.assertFalse(word.exists())
        self.assertFalse(repetido.exists())


if __name__ == "__main__":
    unittest.main()
