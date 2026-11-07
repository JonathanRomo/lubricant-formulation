import json
import sqlite3
from contextlib import closing
import tempfile
from pathlib import Path
from unittest import TestCase

from revisiones_cronologicas import planificar_revisiones, aplicar_plan, asignar_revision_nueva, revisar_ingresos_json, IndiceRevisiones
from unittest.mock import patch


def crear(ruta, revision, fecha, idioma='es', enlaces=None):
    datos = {'producto': {'nombre': 'Producto'}, 'documento': {'idioma': idioma},
             'control_documental': {'codigo_unico_producto': 'P1', 'revision': revision, 'fecha': fecha},
             'pruebas': [{'prueba': 'Color', 'metodo': 'ASTM D1500', 'unidades': 'NA', 'especificacion': '1'}]}
    if enlaces:
        datos['versiones_idioma'] = enlaces
    ruta.write_text(json.dumps(datos, ensure_ascii=False), encoding='utf-8')
    return datos


class RevisionesCronologicasTests(TestCase):
    def test_indice_temporal_reutiliza_el_lote_y_detecta_cambios_externos(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta).resolve(); ruta = base / 'uno.json'
            crear(ruta, '00', '01/01/2023')
            indice = IndiceRevisiones([base])
            with patch.object(indice, '_reconstruir', wraps=indice._reconstruir) as reconstruir:
                datos = {'producto': {'nombre': 'Producto'}, 'control_documental': {'codigo_unico_producto': 'P1', 'revision': '2024'}}
                self.assertEqual(asignar_revision_nueva(datos, [base], indice=indice), '01')
                segunda = base / 'dos.json'; crear(segunda, '01', '02/01/2024')
                indice.registrar(segunda, json.loads(segunda.read_text(encoding='utf-8')))
                nuevo = {'producto': {'nombre': 'Producto'}, 'control_documental': {'codigo_unico_producto': 'P1', 'revision': '2025'}}
                self.assertEqual(asignar_revision_nueva(nuevo, [base], indice=indice), '02')
                reconstruir.assert_not_called()
                crear(base / 'externo.json', '07', '01/01/2025')
                otro = {'producto': {'nombre': 'Producto'}, 'control_documental': {'codigo_unico_producto': 'P1', 'revision': '2026'}}
                self.assertEqual(asignar_revision_nueva(otro, [base], indice=indice), '08')
                self.assertTrue(reconstruir.called)
    def test_ingreso_con_conflicto_y_duplicado_preserva_local_y_procedencia(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); alineadas = base / 'alineadas'; alineadas.mkdir(); db = base / 'base.db'
            local = alineadas / 'local.json'; crear(local, '03', '01/01/2023')
            with closing(sqlite3.connect(db)) as con:
                con.execute('CREATE TABLE versiones (archivo_json TEXT)')
                con.execute('INSERT INTO versiones VALUES (?)', ('local.json',)); con.commit()
            inicial = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertTrue(inicial['inicializado'])
            antes = local.read_bytes()
            entrante = alineadas / 'recibido.json'
            recibido = crear(entrante, '03', '02/02/2024')
            recibido['control_documental']['revision_fuente'] = '2024'
            entrante.write_text(json.dumps(recibido), encoding='utf-8')
            resultado = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            control = json.loads(entrante.read_text(encoding='utf-8'))['control_documental']
            self.assertEqual((control['revision'], control['revision_recibida'], control['revision_fuente']), ('04', '03', '2024'))
            self.assertEqual(resultado['ajustados'], 1)
            self.assertEqual(local.read_bytes(), antes)
            self.assertTrue(Path(resultado['respaldo']).is_file())
            reenviado = alineadas / 'reenviado.json'
            reenviado.write_text(json.dumps(recibido), encoding='utf-8')
            resultado = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertIn('reenviado.json', resultado['archivos_duplicados'])
            self.assertEqual(resultado['ajustados'], 0)
            self.assertIn('reenviado.json', revisar_ingresos_json(alineadas, db, base / 'respaldos')['archivos_duplicados'])
            (alineadas / 'copia.json').write_bytes(local.read_bytes())
            duplicados = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertGreaterEqual(duplicados['duplicados'], 1)
            self.assertEqual(json.loads((alineadas / 'copia.json').read_text(encoding='utf-8'))['control_documental']['revision'], '03')

    def test_primera_ejecucion_no_renumera_archivo_no_registrado(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); alineadas = base / 'alineadas'; alineadas.mkdir(); db = base / 'base.db'
            ruta = alineadas / 'desconocido.json'; crear(ruta, '2023', '01/01/2023'); antes = ruta.read_bytes()
            with closing(sqlite3.connect(db)) as con:
                con.execute('CREATE TABLE versiones (archivo_json TEXT)'); con.commit()
            resultado = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertEqual(resultado['pendientes'], 1)
            self.assertEqual(ruta.read_bytes(), antes)
            datos = json.loads(antes.decode('utf-8'))
            datos['control_documental']['revision'] = '00'
            ruta.write_text(json.dumps(datos), encoding='utf-8')
            resuelto = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertEqual(resuelto['pendientes'], 0)

    def test_archivo_conocido_modificado_no_se_renumera(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); alineadas = base / 'alineadas'; alineadas.mkdir(); db = base / 'base.db'
            ruta = alineadas / 'local.json'; crear(ruta, '03', '01/01/2023')
            with closing(sqlite3.connect(db)) as con:
                con.execute('CREATE TABLE versiones (archivo_json TEXT)')
                con.execute('INSERT INTO versiones VALUES (?)', ('local.json',)); con.commit()
            revisar_ingresos_json(alineadas, db, base / 'respaldos')
            crear(ruta, '03', '01/01/2024'); antes = ruta.read_bytes()
            resultado = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertEqual(resultado['modificados'], 1)
            self.assertEqual(ruta.read_bytes(), antes)

    def test_conflicto_historico_se_avisa_sin_renumerar_locales(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); alineadas = base / 'alineadas'; alineadas.mkdir(); db = base / 'base.db'
            a = alineadas / 'uno.json'; b = alineadas / 'dos.json'
            crear(a, '03', '01/01/2023'); crear(b, '03', '01/01/2024')
            with closing(sqlite3.connect(db)) as con:
                con.execute('CREATE TABLE versiones (archivo_json TEXT)')
                con.executemany('INSERT INTO versiones VALUES (?)', [('uno.json',), ('dos.json',)]); con.commit()
            antes = (a.read_bytes(), b.read_bytes())
            resultado = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertEqual(resultado['conflictos_existentes'], 1)
            self.assertEqual((a.read_bytes(), b.read_bytes()), antes)

    def test_ingreso_incompleto_queda_pendiente_en_vez_de_suponer_conflicto(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); alineadas = base / 'alineadas'; alineadas.mkdir(); db = base / 'base.db'
            crear(alineadas / 'local.json', '03', '01/01/2023')
            with closing(sqlite3.connect(db)) as con:
                con.execute('CREATE TABLE versiones (archivo_json TEXT)')
                con.execute('INSERT INTO versiones VALUES (?)', ('local.json',)); con.commit()
            revisar_ingresos_json(alineadas, db, base / 'respaldos')
            ruta = alineadas / 'incompleto.json'
            ruta.write_text(json.dumps({'producto': {'nombre': 'Producto'}, 'control_documental': {'codigo_unico_producto': 'P1', 'revision': '03'}}), encoding='utf-8')
            antes = ruta.read_bytes()
            resultado = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertEqual(resultado['pendientes'], 1)
            self.assertEqual(ruta.read_bytes(), antes)

    def test_registro_de_json_se_conserva_al_reconstruir_base(self):
        from reconstruir_base_datos import _conservar_registro_json_revision
        with tempfile.TemporaryDirectory() as carpeta:
            anterior = Path(carpeta) / 'anterior.db'; nueva = Path(carpeta) / 'nueva.db'
            with closing(sqlite3.connect(anterior)) as con:
                con.execute('CREATE TABLE registro_json_revision (ruta TEXT PRIMARY KEY, huella TEXT, estado TEXT)')
                con.execute("INSERT INTO registro_json_revision VALUES ('x.json', 'abc', 'conocido')"); con.commit()
                con.execute('CREATE TABLE registro_json_revision_v2 (ambito TEXT, ruta TEXT, huella TEXT, estado TEXT, PRIMARY KEY(ambito, ruta))')
                con.execute('CREATE TABLE registro_json_revision_ambitos (ambito TEXT PRIMARY KEY)')
                con.execute("INSERT INTO registro_json_revision_v2 VALUES ('carpeta', 'x.json', 'abc', 'conocido')")
                con.execute("INSERT INTO registro_json_revision_ambitos VALUES ('carpeta')"); con.commit()
            with closing(sqlite3.connect(nueva)) as con:
                avisos = []
                _conservar_registro_json_revision(con, anterior, avisos)
                self.assertEqual(con.execute('SELECT ruta, huella, estado FROM registro_json_revision').fetchone(), ('x.json', 'abc', 'conocido'))
                self.assertEqual(con.execute('SELECT ambito, ruta FROM registro_json_revision_v2').fetchone(), ('carpeta', 'x.json'))
                self.assertEqual(avisos, [])

    def test_nueva_carpeta_se_inicializa_con_precaucion_sin_perder_registro_anterior(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); primera = base / 'uno'; segunda = base / 'dos'; primera.mkdir(); segunda.mkdir(); db = base / 'base.db'
            crear(primera / 'local.json', '03', '01/01/2023')
            with closing(sqlite3.connect(db)) as con:
                con.execute('CREATE TABLE versiones (archivo_json TEXT)')
                con.execute("INSERT INTO versiones VALUES ('local.json')"); con.commit()
            revisar_ingresos_json(primera, db, base / 'respaldos')
            ruta = segunda / 'local.json'; crear(ruta, '03', '01/01/2024'); antes = ruta.read_bytes()
            resultado = revisar_ingresos_json(segunda, db, base / 'respaldos')
            self.assertEqual(resultado['pendientes'], 1)
            self.assertEqual(ruta.read_bytes(), antes)
            regreso = revisar_ingresos_json(primera, db, base / 'respaldos')
            self.assertEqual(regreso['pendientes'], 0)

    def test_pareja_confirmada_recibida_comparte_nueva_revision(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); alineadas = base / 'alineadas'; alineadas.mkdir(); db = base / 'base.db'
            crear(alineadas / 'local.json', '03', '01/01/2023')
            with closing(sqlite3.connect(db)) as con:
                con.execute('CREATE TABLE versiones (archivo_json TEXT)')
                con.execute('INSERT INTO versiones VALUES (?)', ('local.json',)); con.commit()
            revisar_ingresos_json(alineadas, db, base / 'respaldos')
            crear(alineadas / 'es.json', '03', '02/02/2024', enlaces={'INGLES': [{'archivo_json': 'en.json'}]})
            crear(alineadas / 'en.json', '03', '02/02/2024', idioma='en')
            resultado = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertEqual(resultado['ajustados'], 2)
            for nombre in ('es.json', 'en.json'):
                self.assertEqual(json.loads((alineadas / nombre).read_text(encoding='utf-8'))['control_documental']['revision'], '04')

    def test_enlace_recibido_no_resuelto_conserva_revision_y_queda_pendiente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); alineadas = base / 'alineadas'; alineadas.mkdir(); db = base / 'base.db'
            crear(alineadas / 'local.json', '03', '01/01/2023')
            with closing(sqlite3.connect(db)) as con:
                con.execute('CREATE TABLE versiones (archivo_json TEXT)')
                con.execute('INSERT INTO versiones VALUES (?)', ('local.json',)); con.commit()
            revisar_ingresos_json(alineadas, db, base / 'respaldos')
            ruta = alineadas / 'recibido.json'
            crear(ruta, '03', '02/02/2024', enlaces={'INGLES': [{'archivo_json': 'ausente.json'}]})
            anterior = ruta.read_bytes()
            resultado = revisar_ingresos_json(alineadas, db, base / 'respaldos')
            self.assertEqual(resultado['pendientes'], 1)
            self.assertEqual(ruta.read_bytes(), anterior)
    def test_no_aplica_plan_si_archivo_cambio_desde_su_preparacion(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); alineadas = base / 'alineadas'; anteriores = base / 'anteriores'
            alineadas.mkdir(); anteriores.mkdir(); ruta = alineadas / 'uno.json'
            crear(ruta, '2023', '01/01/2023')
            plan = planificar_revisiones(alineadas)
            crear(ruta, '2024', '01/01/2024')
            db = base / 'base.db'; db.touch()
            with self.assertRaisesRegex(RuntimeError, 'cambió después'):
                aplicar_plan(plan, alineadas, anteriores, db, base / 'respaldos')
            self.assertEqual(json.loads(ruta.read_text(encoding='utf-8'))['control_documental']['revision'], '2024')

    def test_plan_por_fecha_con_enlace_confirmado_y_sin_fecha_al_final(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            crear(base / 'uno.json', '2022', '01/01/2022', enlaces={'INGLES': [{'archivo_json': 'en.json'}]})
            crear(base / 'en.json', '2022', '01/01/2022', idioma='en')
            crear(base / 'sin_fecha.json', '', '')
            plan = {x['ruta'].name: x for x in planificar_revisiones(base)}
            self.assertEqual(plan['uno.json']['revision_nueva'], '00')
            self.assertEqual(plan['en.json']['revision_nueva'], '00')
            self.assertEqual(plan['sin_fecha.json']['revision_nueva'], '01')
            self.assertEqual(plan['uno.json']['datos']['control_documental']['revision_fuente'], '2022')

    def test_aplicacion_respalda_actualiza_json_y_base_y_siguiente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta); alineadas = base / 'alineadas'; anteriores = base / 'anteriores'
            alineadas.mkdir(); anteriores.mkdir(); archivo = 'uno.json'
            original = crear(alineadas / archivo, '2023', '01/01/2023')
            crear(anteriores / archivo, '2023', '01/01/2023')
            db = base / 'base.db'
            with closing(sqlite3.connect(db)) as con:
                con.executescript('CREATE TABLE versiones (archivo_id INTEGER, archivo_json TEXT, revision_especificacion TEXT);'
                                  'CREATE TABLE archivos (id INTEGER, ultimo_json TEXT, revision_especificacion TEXT);')
                con.execute('INSERT INTO versiones VALUES (1, ?, ?)', (archivo, '2023'))
                con.execute('INSERT INTO archivos VALUES (1, ?, ?)', (archivo, '2023'))
                con.commit()
            respaldo, _ = aplicar_plan(planificar_revisiones(alineadas), alineadas, anteriores, db, base / 'respaldos')
            self.assertTrue(respaldo.is_file())
            for ruta in (alineadas / archivo, anteriores / archivo):
                control = json.loads(ruta.read_text(encoding='utf-8'))['control_documental']
                self.assertEqual((control['revision'], control['revision_fuente']), ('00', '2023'))
            with closing(sqlite3.connect(db)) as con:
                self.assertEqual(con.execute('SELECT revision_especificacion FROM versiones').fetchone()[0], '00')
            self.assertEqual(asignar_revision_nueva(original, [alineadas, anteriores]), '01')
            self.assertEqual(original['control_documental']['revision_fuente'], '2023')
