import hashlib
import json
import sqlite3
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import gestor_especificaciones as gestor
import reconstruir_base_datos as recuperacion
import registro_hii


def documento(idioma="es", revision="01", fecha="2026-01-01", **extras):
    return dict({"producto": {"nombre": "Producto"}, "documento": {"idioma": idioma, "fecha_documento_normalizada": fecha},
        "control_documental": {"codigo_unico_producto": "20100001", "nombre_cliente": "Cliente", "revision": revision},
        "pruebas": [], "fecha_creacion_json": "2026-02-01T10:00:00"}, **extras)


def escribir(base, nombre="original.json", datos=None, tipo="ESPECIFICACIONES_JSON"):
    ruta = base / tipo / nombre
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(json.dumps(datos or documento(), ensure_ascii=False), encoding="utf-8")
    return ruta


class ReconstruccionTests(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporal.cleanup)
        self.base = Path(self.temporal.name).resolve()
        self.db = self.base / "especificaciones.db"

    def test_auto_gestor_registra_json_y_no_reconstruye_si_ya_existe(self):
        ruta = escribir(self.base)
        contenido = ruta.read_bytes()
        with closing(gestor.conectar(self.db)) as con:
            self.assertEqual(con.execute("SELECT archivo_json FROM versiones").fetchone()[0], ruta.name)
            self.assertEqual(con.execute("PRAGMA integrity_check").fetchone()[0], "ok")
        with patch.object(recuperacion, "_leer_documentos", side_effect=AssertionError("No repetir")):
            with closing(gestor.conectar(self.db)):
                pass
        self.assertEqual(ruta.read_bytes(), contenido)

    def test_auto_relaciones_crea_tambien_tablas_del_gestor(self):
        escribir(self.base)
        with registro_hii.conectar(self.db) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0], 1)

    def test_recupera_archivo_de_cero_bytes_con_respaldo(self):
        ruta = escribir(self.base)
        antes = ruta.read_bytes()
        self.db.touch()
        self.assertEqual(recuperacion.motivo_recuperacion(self.base), "vacia")
        resultado = recuperacion.asegurar_base(self.base)
        self.assertEqual(resultado["motivo"], "vacia")
        self.assertIsNotNone(resultado["respaldo"])
        self.assertEqual(resultado["documentos"], 1)
        self.assertEqual(ruta.read_bytes(), antes)

    def test_gestor_recupera_base_con_tablas_pero_sin_registros(self):
        with closing(gestor.conectar(self.db, recuperar=False)):
            pass
        escribir(self.base)
        with closing(gestor.conectar(self.db)) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0], 1)

    def test_registro_hii_recupera_sqlite_vacio(self):
        with closing(sqlite3.connect(self.db)):
            pass
        escribir(self.base)
        with registro_hii.conectar(self.db) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM hii_especificaciones").fetchone()[0], 1)

    def test_catalogo_solo_no_equivale_a_especificaciones_registradas(self):
        with closing(gestor.conectar(self.db, recuperar=False)) as con:
            con.execute("INSERT INTO gestor_estado VALUES ('dato_de_control', '1')")
            con.execute("INSERT INTO catalogo_fuentes (nombre_archivo,ruta_archivo,hash_catalogo,tamano_bytes,fecha_modificacion_ns,fecha_lectura,total_registros) VALUES ('catalogo.xlsx','catalogo.xlsx','hash',1,1,'2026-01-01',1)")
            con.commit()
        escribir(self.base)
        self.assertTrue(recuperacion.necesita_recuperacion(self.base))
        self.assertEqual(recuperacion.asegurar_base(self.base)["documentos"], 1)

    def test_base_vacia_sin_json_permanece_intacta(self):
        self.db.touch()
        self.assertFalse(recuperacion.necesita_recuperacion(self.base))
        self.assertIsNone(recuperacion.asegurar_base(self.base))
        self.assertEqual(self.db.read_bytes(), b"")
        self.assertFalse((self.base / "RESPALDOS_BASE_DATOS").exists())

    def test_no_interpreta_corrupcion_como_vacio(self):
        escribir(self.base)
        self.db.write_bytes(b"sqlite corrupto")
        self.assertIsNone(recuperacion.motivo_recuperacion(self.base))
        self.assertIsNone(recuperacion.asegurar_base(self.base))
        self.assertEqual(self.db.read_bytes(), b"sqlite corrupto")

    def test_dos_solicitudes_para_base_vacia_no_reemplazan_la_ya_recuperada(self):
        escribir(self.base)
        self.db.touch()
        with ThreadPoolExecutor(max_workers=2) as hilos:
            resultados = list(hilos.map(lambda _: recuperacion.asegurar_base(self.base), range(2)))
        self.assertEqual(sum(r is not None for r in resultados), 1)
        self.assertEqual(len(list((self.base / "RESPALDOS_BASE_DATOS").iterdir())), 1)

    def test_base_con_registros_no_se_reconstruye_aunque_haya_mas_json(self):
        escribir(self.base)
        recuperacion.asegurar_base(self.base)
        escribir(self.base, "otro.json")
        with patch.object(recuperacion, "_leer_documentos", side_effect=AssertionError("No sustituir una base con datos")):
            self.assertIsNone(recuperacion.asegurar_base(self.base))
        with closing(sqlite3.connect(self.db)) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0], 1)

    def test_original_y_alineado_se_registran_sin_duplicar_extraccion(self):
        original = escribir(self.base, "mismo.json")
        alineada = escribir(self.base, "mismo.json", tipo="ESPECIFICACIONES_JSON_ALINEADAS")
        antes = {p: p.read_bytes() for p in (original, alineada)}
        resultado = recuperacion.reconstruir(self.base)
        self.assertEqual(resultado["documentos"], 2)
        self.assertEqual(resultado["versiones"], 1)
        self.assertEqual({p: p.read_bytes() for p in antes}, antes)

    def test_alineada_sin_original_tiene_ruta_recuperable(self):
        ruta = escribir(self.base, tipo="ESPECIFICACIONES_JSON_ALINEADAS")
        recuperacion.reconstruir(self.base)
        with closing(sqlite3.connect(self.db)) as con:
            referencia = con.execute("SELECT archivo_json FROM versiones").fetchone()[0]
        self.assertEqual((self.base / "ESPECIFICACIONES_JSON" / referencia).resolve(), ruta)

    def test_actuales_usan_fecha_del_documento_y_no_mtime(self):
        escribir(self.base, "antiguo.json", documento(fecha="2020-01-01", fecha_creacion_json="2030-01-01"))
        escribir(self.base, "reciente.json", documento(fecha="2026-01-01", fecha_creacion_json="2026-01-01"))
        recuperacion.reconstruir(self.base)
        with closing(sqlite3.connect(self.db)) as con:
            self.assertEqual(con.execute("SELECT archivo_json FROM recuperacion_actuales").fetchone()[0], "reciente.json")
            self.assertEqual(con.execute("SELECT COUNT(*) FROM versiones").fetchone()[0], 2)

    def test_restaura_referencias_explicitamente_sin_escribir_json(self):
        es = escribir(self.base, "es.json", documento(versiones_idioma={"en": [{"archivo_json": "en.json"}]}))
        en = escribir(self.base, "en.json", documento("en", revision="02"))
        antes = {p: p.read_bytes() for p in (es, en)}
        resultado = recuperacion.reconstruir(self.base)
        self.assertEqual(resultado["relaciones"], 1)
        with registro_hii.conectar(self.db) as con:
            self.assertEqual(registro_hii.candidatas(con, es, "INGLES", [en]), [en])
        self.assertEqual({p: p.read_bytes() for p in antes}, antes)

    def test_no_inventa_parejas_ambiguas(self):
        escribir(self.base, "es.json", documento())
        escribir(self.base, "en1.json", documento("en"))
        escribir(self.base, "en2.json", documento("en"))
        self.assertEqual(recuperacion.reconstruir(self.base)["relaciones"], 0)

    def test_wip_conserva_auxiliar_y_no_se_cuenta_como_extraccion_terminada(self):
        ruta = escribir(self.base)
        auxiliar = ruta.with_name("." + ruta.name + ".wip")
        auxiliar.write_text("estado de borrador", encoding="utf-8")
        resultado = recuperacion.reconstruir(self.base)
        self.assertEqual((resultado["wip"], resultado["versiones"]), (1, 0))
        self.assertEqual(auxiliar.read_text(encoding="utf-8"), "estado de borrador")

    def test_omite_invalido_y_rechaza_reemplazo_si_ninguno_es_valido(self):
        ruta = escribir(self.base)
        invalido = escribir(self.base, "invalido.json")
        invalido.write_text("{", encoding="utf-8")
        resultado = recuperacion.reconstruir(self.base)
        self.assertEqual(len(resultado["omitidos"]), 1)
        antes = self.db.read_bytes()
        ruta.write_text("[]", encoding="utf-8")
        with self.assertRaises(ValueError):
            recuperacion.reconstruir(self.base, reemplazar=True)
        self.assertEqual(self.db.read_bytes(), antes)

    def test_manual_respalda_base_sana_y_corrupta(self):
        escribir(self.base)
        recuperacion.reconstruir(self.base)
        resultado = recuperacion.reconstruir(self.base, reemplazar=True)
        self.assertTrue((Path(resultado["respaldo"]) / self.db.name).is_file())
        self.db.write_bytes(b"base corrupta")
        resultado = recuperacion.reconstruir(self.base, reemplazar=True)
        self.assertEqual((Path(resultado["respaldo"]) / (self.db.name + ".corrupta")).read_bytes(), b"base corrupta")
        with closing(sqlite3.connect(self.db)) as con:
            self.assertEqual(con.execute("PRAGMA integrity_check").fetchone()[0], "ok")

    def test_auxiliares_huerfanos_no_se_reutilizan(self):
        escribir(self.base)
        auxiliar = Path(str(self.db) + "-wal")
        auxiliar.write_bytes(b"diario antiguo")
        with self.assertRaisesRegex(RuntimeError, "auxiliares"):
            recuperacion.reconstruir(self.base)
        self.assertFalse(self.db.exists())
        self.assertEqual(auxiliar.read_bytes(), b"diario antiguo")
        self.assertFalse(list(self.base.glob(".reconstruccion_*")))

    def test_error_al_publicar_preserva_base_anterior(self):
        escribir(self.base)
        recuperacion.reconstruir(self.base)
        antes = self.db.read_bytes()
        with patch.object(recuperacion, "_publicar", side_effect=OSError("fallo")), self.assertRaises(OSError):
            recuperacion.reconstruir(self.base, reemplazar=True)
        self.assertEqual(self.db.read_bytes(), antes)
        self.assertFalse(list(self.base.glob(".reconstruccion_*")))

    def test_no_vacia_base_existente_si_faltan_los_json(self):
        ruta = escribir(self.base)
        recuperacion.reconstruir(self.base)
        antes = self.db.read_bytes()
        ruta.unlink()
        with self.assertRaisesRegex(ValueError, "conserva la base anterior"):
            recuperacion.reconstruir(self.base, reemplazar=True)
        self.assertEqual(self.db.read_bytes(), antes)

    def test_dos_solicitudes_reconstruyen_una_sola_vez(self):
        escribir(self.base)
        with ThreadPoolExecutor(max_workers=2) as hilos:
            resultados = list(hilos.map(lambda _: recuperacion.asegurar_base(self.base), range(2)))
        self.assertEqual(sum(r is not None for r in resultados), 1)

    def test_gestor_omite_word_recuperado_sin_crear_otro_json(self):
        libro = Workbook()
        libro.active.append(["ID", "Description"])
        libro.active.append(["20100001", "Producto"])
        catalogo = self.base / "CODIGOS_200.xlsx"
        libro.save(catalogo)
        libro.close()
        word = self.base / "ESPECIFICACIONES_VIEJAS" / "documento.docx"
        word.parent.mkdir()
        word.write_bytes(b"documento previamente procesado")
        ruta = escribir(self.base, datos=documento(archivo_origen="ESPECIFICACIONES_VIEJAS/documento.docx",
            hash_sha256_origen=hashlib.sha256(word.read_bytes()).hexdigest(),
            hash_sha256_catalogo=hashlib.sha256(catalogo.read_bytes()).hexdigest()))
        contenido = ruta.read_bytes()
        with patch.object(gestor, "preparar", side_effect=AssertionError("No extraer nuevamente")), patch("builtins.print"):
            resultado = gestor.procesar(self.base)
        self.assertEqual(resultado["ignorados"], 1)
        self.assertEqual(resultado["nuevos"], 0)
        self.assertEqual(list(ruta.parent.glob("*.json")), [ruta])
        self.assertEqual(ruta.read_bytes(), contenido)


if __name__ == "__main__":
    unittest.main()
