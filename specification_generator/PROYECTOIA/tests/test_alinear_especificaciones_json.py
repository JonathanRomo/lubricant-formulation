import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from alinear_especificaciones_json import (
    VERSION,
    cargar_maestro,
    ejecutar,
    es_prueba_brookfield,
    limpiar_marca_acreditacion,
    metodo_maestro_canonico,
    metodo_maestro_para_idioma,
    metodos_maestros,
    mejor_coincidencia,
    normalizar_especificacion,
    normalizar_espacio_unidades,
    normalizar_formato_numerico,
    normalizar_terminos_especificacion,
    normalizar_valor_medicion,
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


def simular_json_para_cache(origen, destino, maestro):
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text("{}", encoding="utf-8")
    return {"alineados": 1, "revision": 0, "sin_coincidencia": 0}


class ReglasConfirmadasLaboratorioTests(unittest.TestCase):
    def test_unidades_declaradas_se_normalizan_con_y_sin_coincidencia(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        from generador_hii import preparar_datos
        maestro = [self.registro('ESPECTRO INFRARROJO', 'INFRARED SPECTRUM', 'ASTM E1252', 'NA')]
        for idioma in ('es', 'en'):
            for unidad in ('N/A', 'N / A', 'NA', 'A/D', 'AD', 'ADIMENSIONAL', 'DIMENSIONLESS'):
                esperado = 'NA' if unidad.replace(' ', '') in {'N/A', 'NA'} else 'DIMENSIONLESS' if idioma == 'en' else 'ADIMENSIONAL'
                esperado_original = esperado
                for metodo in ('ASTM E1252', 'NOM-999-SSA1'):
                    esperado = 'NA' if metodo == 'ASTM E1252' else esperado_original
                    with self.subTest(idioma=idioma, unidad=unidad, metodo=metodo), tempfile.TemporaryDirectory() as carpeta:
                        fila = {'prueba': 'NOMBRE ANTIGUO', 'metodo': metodo, 'unidades': unidad, 'especificacion': '<1000'}
                        errores = {}
                        editadas, conteo = alinear_filas_generador([fila], maestro, idioma, errores_metodo=errores)
                        self.assertEqual(editadas[0]['unidades'], esperado)
                        if metodo.startswith('NOM'):
                            self.assertIn(0, errores)
                            self.assertEqual(editadas[0]['prueba'], fila['prueba'])
                            self.assertEqual(editadas[0]['metodo'], metodo)
                            self.assertEqual(editadas[0]['especificacion'], '<1000')
                        ruta = Path(carpeta) / 'prueba.json'
                        ruta.write_text(json.dumps({'producto': {'nombre': 'Producto'}, 'documento': {'idioma': idioma},
                            'control_documental': {'codigo_unico_producto': 'P1', 'revision': '00'}, 'pruebas': [fila]}), encoding='utf-8')
                        for _ in range(2):
                            procesar_json(ruta, ruta, maestro)
                            datos = json.loads(ruta.read_text(encoding='utf-8'))
                            self.assertEqual(datos['pruebas'][0]['unidades'], esperado)
                            self.assertEqual(datos['alineacion_metodos_lab']['detalle'][0]['valores_alineados']['unidades'], esperado)
                            self.assertEqual(datos['alineacion_metodos_lab']['detalle'][0]['unidades_originales'], unidad)
                            self.assertEqual(preparar_datos(ruta).pruebas[0]['unidades'], esperado)
        # Tampoco depende de que el candidato pase las restricciones de temperatura.
        fija = self.registro('VISCOSIDAD CINEMÁTICA A 40 °C', 'KINEMATIC VISCOSITY AT 40 °C', 'ASTM D445', 'mm²/s')
        filas, conteo = alinear_filas_generador([{'prueba': 'VISCOSIDAD A 100 °F', 'metodo': 'ASTM D445', 'unidades': 'A/D'}], [fija], 'en')
        self.assertEqual(conteo['alineados'], 1)
        self.assertEqual(filas[0]['unidades'], 'mm²/s')
        # AD y A/D también representan la misma unidad al validar grupos ASTM.
        from alinear_especificaciones_json import MaestroPreparado
        base = self.registro('PRUEBA DE LABORATORIO', 'LABORATORY TEST', 'ASTM D123', 'A/D')
        otra = {**base, 'METODO': 'ASTM D456', 'UNIDADES': 'ADIMENSIONAL', 'UNITS': 'DIMENSIONLESS'}
        r = mejor_coincidencia('PRUEBA DE LABORATORIO', 'ASTM D123 / ASTM D456', MaestroPreparado([base, otra]), 'ESPANOL', 'AD')
        self.assertEqual(r['estado'], 'ALINEADO')
        r = mejor_coincidencia('PRUEBA DE LABORATORIO', 'ASTM D123 / ASTM D456', [base, {**otra, 'UNIDADES': 'NA', 'UNITS': 'NA'}], 'ESPANOL', 'AD')
        self.assertNotEqual(r['estado'], 'ALINEADO')

    def test_saybolt_admite_solo_equivalencias_nominales_confirmadas(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        maestro = [self.registro('VISCOSIDAD SAYBOLT A '+str(f)+' °F', 'SAYBOLT VISCOSITY AT '+str(f)+' °F', 'ASTM D2161', 'SUS') for f in (100, 210)]
        for celsius, fahrenheit in (('37.8', '100'), ('98.9', '210')):
            for original in ('SAYBOLT VISCOSITY @ '+celsius+' °C', 'VISCOSIDAD SAYBOLT A '+celsius+' °C ('+fahrenheit+' °F)'):
                for idioma in ('es', 'en'):
                    with self.subTest(original=original, idioma=idioma), tempfile.TemporaryDirectory() as carpeta:
                        fila = {'prueba': original, 'metodo': 'ASTM D-2161 -25 (2025)', 'unidades': 'SUS'}
                        editadas, conteo = alinear_filas_generador([fila], maestro, idioma)
                        self.assertEqual(conteo['alineados'], 1)
                        self.assertIn(fahrenheit+' °F', editadas[0]['prueba'])
                        ruta = Path(carpeta) / 'prueba.json'
                        ruta.write_text(json.dumps({'documento': {'idioma': idioma}, 'pruebas': [fila]}), encoding='utf-8')
                        for _ in range(2):
                            procesar_json(ruta, ruta, maestro)
                            datos = json.loads(ruta.read_text(encoding='utf-8'))
                            self.assertIn(fahrenheit+' °F', datos['pruebas'][0]['prueba'])
                            self.assertEqual(datos['alineacion_metodos_lab']['detalle'][0]['nombre_original'], original)
        self.assertEqual(mejor_coincidencia('SAYBOLT VISCOSITY @ 38 °C', 'ASTM D2161', maestro, 'INGLES')['estado'], 'ALINEADO')
        self.assertEqual(mejor_coincidencia('SAYBOLT VISCOSITY @ 37.8 °C', 'ASTM D2161', maestro[1:], 'INGLES')['estado'], 'ALINEADO')

    def test_nuevos_identificadores_de_crepitacion_y_appearance(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        crepitacion = self.registro('HUMEDAD POR CREPITACIÓN', 'MOISTURE BY CRACKLE TEST', 'MÉTODO INTERNO (91300029)')
        crepitacion['MÉTODO INGLÉS'] = 'INTERNAL METHOD (91300029)'
        for codigo in ('INS-CCA-07', 'INS-CCA-07-00', 'CLIENT METHOD (91300029)'):
            for idioma in ('es', 'en'):
                filas, conteo = alinear_filas_generador([{'prueba': 'NOMBRE ANTIGUO', 'metodo': codigo}], [crepitacion], idioma)
                self.assertEqual(conteo['alineados'], 1)
                self.assertEqual(filas[0]['prueba'], crepitacion['NOMBRE IQM INGLES' if idioma == 'en' else 'NOMBRE IQM'])
        self.assertTrue(mejor_coincidencia('NOMBRE ANTIGUO', 'CLIENT METHOD (913000290)', [crepitacion])['metodo_no_encontrado'])
        apariencia = self.registro('APARIENCIA VISUAL', 'VISUAL APPEAREANCE', 'ORGANOLÉPTICO')
        apariencia['MÉTODO INGLÉS'] = 'ORGANOLEPTIC'
        filas, conteo = alinear_filas_generador([{'prueba': 'APPEARANCE', 'metodo': 'VISUAL', 'unidades': 'N/A'}], [apariencia], 'en')
        self.assertEqual(conteo['alineados'], 1)
        self.assertEqual(filas[0]['prueba'], 'VISUAL APPEARANCE')
        self.assertEqual(filas[0]['unidades'], 'NA')
        self.assertEqual(filas[0]['metodo'], 'ORGANOLEPTIC')

    def test_identificadores_confirmados_reconocen_nombres_distintos_y_vacios(self):
        from alinear_especificaciones_json import MaestroPreparado
        casos = [
            ('ASTM D2272 - 16 (2025)', 'ESTABILIDAD A LA OXIDACIÓN, MÉTODO B', 'OXIDATION STABILITY, METHOD B', 'ASTM D2272', 'min'),
            ('ASTM E1252 -25 (2025)', 'ESPECTRO INFRARROJO', 'INFRARED SPECTRUM', 'ASTM E1252', 'NA'),
            ('91300092', 'HUMEDAD POR CREPITACIÓN', 'MOISTURE BY CRACKLE TEST', 'MÉTODO INTERNO (91300029)', 'NA'),
            ('MÉTODO INTERNO (91300092)', 'HUMEDAD POR CREPITACIÓN', 'MOISTURE BY CRACKLE TEST', 'MÉTODO INTERNO (91300029)', 'NA'),
        ]
        for fuente, es, en, metodo, unidad in casos:
            registro = self.registro(es, en, metodo, unidad)
            if '91300029' in metodo: registro['MÉTODO INGLÉS'] = 'INTERNAL METHOD (91300029)'
            for maestro in ([registro], MaestroPreparado([registro])):
                for idioma, esperado, campo in (('ESPANOL', es, 'NOMBRE IQM'), ('INGLES', en, 'NOMBRE IQM INGLES')):
                    for nombre in ('', 'NOMBRE ANTIGUO SIN PARECIDO', 'DENSIDAD RELATIVA MAL TITULADA', 'CARBONOS AROMATICOS'):
                        with self.subTest(fuente=fuente, nombre=nombre, idioma=idioma):
                            match = mejor_coincidencia(nombre, fuente, maestro, idioma, unidad)
                            self.assertEqual(match['estado'], 'ALINEADO')
                            self.assertEqual(match['registro'][campo], esperado)
                            self.assertTrue(match['reglas_equivalencia'])

    def test_d2272_predetermina_b_y_respeta_metodo_a_explicito(self):
        a = self.registro('ESTABILIDAD A LA OXIDACIÓN, MÉTODO A', 'OXIDATION STABILITY, METHOD A', 'ASTM D2272', 'min')
        b = self.registro('ESTABILIDAD A LA OXIDACIÓN, MÉTODO B', 'OXIDATION STABILITY, METHOD B', 'ASTM D2272', 'min')
        for nombre, metodo, esperado in (('PRUEBA ANTIGUA RPVOT', 'ASTM D2272 -16', b),
                                         ('OXIDATION STABILITY, METHOD A', 'ASTM D2272', a),
                                         ('PRUEBA ANTIGUA', 'ASTM D2272 (PROCEDIMIENTO A)', a)):
            match = mejor_coincidencia(nombre, metodo, [a, b], 'ESPANOL', 'min')
            self.assertEqual(match['estado'], 'ALINEADO')
            self.assertEqual(match['registro']['NOMBRE IQM'], esperado['NOMBRE IQM'])
        self.assertNotEqual(mejor_coincidencia('PRUEBA ANTIGUA', 'ASTM D2272', [a], 'ESPANOL')['estado'], 'ALINEADO')
        self.assertNotEqual(mejor_coincidencia('MÉTODO A', 'ASTM D2272', [b], 'ESPANOL')['estado'], 'ALINEADO')

    def test_reglas_por_metodo_se_aplican_en_editor_json_y_conservan_origen(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        registro = self.registro('HUMEDAD POR CREPITACIÓN', 'MOISTURE BY CRACKLE TEST', 'MÉTODO INTERNO (91300029)')
        registro['MÉTODO INGLÉS'] = 'INTERNAL METHOD (91300029)'
        for idioma, nombre in (('es', registro['NOMBRE IQM']), ('en', registro['NOMBRE IQM INGLES'])):
            fila = {'prueba': 'NOMBRE ANTIGUO', 'metodo': '91300092', 'unidades': 'NA'}
            salida, conteo = alinear_filas_generador([fila], [registro], idioma)
            self.assertEqual(conteo['alineados'], 1)
            self.assertEqual(salida[0]['prueba'], nombre)
            with tempfile.TemporaryDirectory() as temporal:
                ruta = Path(temporal) / 'datos.json'
                ruta.write_text(json.dumps({'documento': {'idioma': idioma}, 'pruebas': [fila]}), encoding='utf-8')
                for _ in range(2):
                    procesar_json(ruta, ruta, [registro])
                    datos = json.loads(ruta.read_text(encoding='utf-8'))
                    self.assertEqual(datos['pruebas'][0]['prueba'], nombre)
                    self.assertEqual(datos['alineacion_metodos_lab']['detalle'][0]['metodo_original'], '91300092')
        self.assertTrue(mejor_coincidencia('NOMBRE ANTIGUO', '913000920', [registro], 'ESPANOL')['metodo_no_encontrado'])
        ir = self.registro('ESPECTRO INFRARROJO', 'INFRARED SPECTRUM', 'ASTM E1252')
        self.assertTrue(mejor_coincidencia('PRUEBA IR', 'ASTM E1252 / ASTM D9999', [ir], 'ESPANOL')['metodo_no_encontrado'])

    def test_recupera_encabezado_ingles_sobrescrito_solo_en_diseno_conocido(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / 'maestro.xlsx'
            w = Workbook(); s = w.active; s.title = 'METODOS'
            s.append(['METODO', 'MÉTODO INGLÉS', 'NOMBRE IQM', 'LOWER POUR POINT', 'UNIDADES', 'UNITS', 'ACREDITACION NOMBRE'])
            s.append(['ASTM D2140', 'ASTM D2140', 'CARBONOS AROMÁTICOS', 'AROMATIC CARBONS', '%', '%', 'NO ACREDITADO'])
            w.save(ruta)
            self.assertEqual(cargar_maestro(ruta)[0]['NOMBRE IQM INGLES'], 'AROMATIC CARBONS')
            s['C1'] = 'OTRO CAMPO'; w.save(ruta)
            with self.assertRaises(ValueError): cargar_maestro(ruta)

    def test_ediciones_y_anos_no_cambian_la_identidad_astm(self):
        from alinear_especificaciones_json import MaestroPreparado
        for norma in ('D123', 'D1234', 'D2140'):
            registro = self.registro('PRUEBA DE LABORATORIO', 'LABORATORY TEST', 'ASTM '+norma+' - 15 (2015)')
            for variante in ('ASTM '+norma+' - 25', 'ASTM '+norma+' -25 (AAAA)', 'ASTM '+norma+'-25 (2025)', 'ASTM '+norma+' - 19a (2023)'):
                for maestro in ([registro], MaestroPreparado([registro])):
                    with self.subTest(norma=norma, variante=variante):
                        r = mejor_coincidencia('PRUEBA DE LABORATORIO', variante, maestro, 'ESPANOL')
                        self.assertEqual(r['estado'], 'ALINEADO')
                        self.assertEqual(referencias_astm(variante), (norma,))
                        self.assertEqual(metodo_maestro_canonico(variante), 'ASTM '+norma)
            distinto = mejor_coincidencia('PRUEBA DE LABORATORIO', 'ASTM D9999 - 25', [registro], 'ESPANOL')
            self.assertTrue(distinto['metodo_no_encontrado'])

    def test_carbonos_se_relacionan_con_sus_tres_filas_sin_intercambiarlos(self):
        from alinear_especificaciones_json import MaestroPreparado
        datos = [('CARBONOS AROMATICOS', 'PORCENTAJE DE CARBONOS EN ANILLOS AROMÁTICOS (%CA)', 'AROMATIC RING CARBONS'),
                 ('CARBONOS NAFTENICOS', 'PORCENTAJE DE CARBONOS EN ANILLOS NAFTÉNICOS (%CN)', 'NAPHTHENIC RING CARBONS'),
                 ('CARBONOS PARAFINICOS', 'PORCENTAJE DE CARBONOS EN CADENAS PARAFÍNICAS (%CP)', 'PARAFFINIC CHAIN CARBONS')]
        registros = [self.registro(es, en, 'ASTM D2140', '%') for original, es, en in datos]
        for maestro in (registros, MaestroPreparado(registros)):
            for original, es, en in datos:
                r = mejor_coincidencia(original, 'ASTM D2140 -25 (2025)', maestro, 'ESPANOL', '%')
                self.assertEqual(r['estado'], 'ALINEADO')
                self.assertEqual(r['registro']['NOMBRE IQM'], es)
        r = mejor_coincidencia('CARBONOS AROMATICOS', 'ASTM D2140', registros[1:], 'ESPANOL')
        self.assertNotEqual(r['estado'], 'ALINEADO')

        r = mejor_coincidencia('COMPOSICIÓN POR TIPO DE CARBONO (%CA, %CN, %CP)', 'ASTM D2140 -25', registros, 'ESPANOL')
        self.assertEqual(r['estado'], 'REVISION_REQUERIDA')
        self.assertEqual(r['tipo_pendiente'], 'COMPONENTE_CARBONO_POR_DEFINIR')

    def test_alkalinity_con_nota_rosa_y_marca_b(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        registro = self.registro('ALCALINIDAD', 'ALKALINITY', 'USP / FEUM')
        nombres = ('ALKALINITY   (THE SOLUTION DOES NOT AQUIRE A PINK COLOR)',
                   'ALKALINITY   (THE SOLUTION DOES NOT AQUIRE A PINK COLOR)B',
                   'ALKALINITY(THE   SOLUTION DOES NOT AQUIRE A PINK COLOR)',
                   'ALKALINITY (THE SOLUTION DOES NOT ACQUIRE A PINK COLOR)')
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre in nombres:
                for idioma, esperado in (('es', 'ALCALINIDAD'), ('en', 'ALKALINITY')):
                    with self.subTest(nombre=nombre, idioma=idioma):
                        for _ in range(2):
                            filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': 'USP 32', 'unidades': 'NA'}], maestro, idioma)
                            self.assertEqual(conteo['alineados'], 1)
                            self.assertEqual(filas[0]['prueba'], esperado)
                            self.assertEqual(filas[0]['metodo'], 'USP')
        nombre = nombres[0]
        self.assertEqual(equivalencias_confirmadas(nombre, 'ASTM D9999')[0], nombre)

    def test_polinucleares_limite_marcas_y_feum_confirmados(self):
        from alinear_especificaciones_json import MaestroPreparado
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'L\u00cdMITE DE HIDROCARBUROS POLIC\u00cdCLICOS AROM\u00c1TICOS'
        en = 'LIMIT OF POLYCYCLIC AROMATIC HYDROCARBONS'
        registro = self.registro(es, en, 'USP')
        nombres = ('COMPUESTOS POLINUCLEARES', 'COMPUESTOS POLINUCLEARES \u25ba \u25ca',
                   'COMPUESTOS POLINUCLEARES \u25ba\u25ca', 'COMPUESTOS POLINUCLEARES \u25ca',
                   'COMPUESTOS POLINUCLEARES \u25ca \u25ba', 'COMPUESTOS POLINUCLEARES\u25ba',
                   'COMPUESTOS POLINUCLEARES\u25ba*', 'LIMITE DE COMPUESTOS POLINUCLEARES',
                   'L\u00cdMITE DE COMPUESTOS POLINUCLEARES')
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre in nombres:
                for metodo in ('USP (VIGENTE)', 'USP 24', 'USP 32', 'FEUM 13', 'FEUM 13/USP 32'):
                    for idioma, esperado in (('es', es), ('en', en)):
                        with self.subTest(nombre=nombre, metodo=metodo, idioma=idioma):
                            for _ in range(2):
                                filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo, 'unidades': 'NA'}], maestro, idioma)
                                self.assertEqual(conteo['alineados'], 1)
                                self.assertEqual(filas[0]['prueba'], esperado)
                                self.assertEqual(filas[0]['metodo'], 'USP')
        self.assertTrue(mejor_coincidencia('COMPUESTOS POLINUCLEARES', 'ASTM D9999', [registro], 'ESPANOL')['metodo_no_encontrado'])

    def test_iso4406_variantes_limpieza_del_laboratorio(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        registro = self.registro('C\u00d3DIGO DE LIMPIEZA', 'CLEANLINESS CODE', 'ISO 4406')
        nombres = ('AUTOMATIC   CLEANING CODE', 'CODIGO DE LIMPIEZA (CONTROL INTERNO)',
                   'CODIGO DE LIMPIEZA, PARTES /1 ml, MAXIMO',
                   'CODIGO DE LIMPIEZA, PARTICULAS /1 ml',
                   'CODIGO DE LIMPIEZA, PARTICULAS /1 ml \u25ba',
                   'CODING THE LEVEL OF CONTAMINATION', 'NIVEL DE LIMPIEZA',
                   'Nivel de Limpieza', 'PUREZA DE PARTICULA')
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre in nombres:
                for idioma, esperado in (('es', registro['NOMBRE IQM']), ('en', 'CLEANLINESS CODE')):
                    with self.subTest(nombre=nombre, idioma=idioma):
                        for _ in range(2):
                            filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': 'ISO 4406:09', 'unidades': 'A/D'}], maestro, idioma)
                            self.assertEqual(conteo['alineados'], 1)
                            self.assertEqual(filas[0]['prueba'], esperado)
                            self.assertEqual(filas[0]['metodo'], 'ISO 4406')
                            self.assertEqual(filas[0]['unidades'], 'NA')
        for nombre, metodo in (('NIVEL DE LIMPIEZA', 'ISO 4407'),
                               ('NIVEL DE LIMPIEZA', 'ISO 4406 / ISO 4407'),
                               ('', 'ISO 4406'), ('COLOR', 'ISO 4406')):
            self.assertEqual(equivalencias_confirmadas(nombre, metodo)[0], nombre)

    def test_tendencia_espuma_con_guiones_es_separador(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        from cambios_alineacion import cambios_documento
        from alinear_especificaciones_json import es_separador_pruebas
        fila = {'prueba': 'TENDENCIA   A LA ESPUMA', 'metodo': '---', 'unidades': ''}
        errores = {}
        filas, conteo = alinear_filas_generador([fila], [MAESTRO], 'es', errores_metodo=errores)
        self.assertEqual(conteo['alineados'], 0)
        self.assertEqual(filas[0]['metodo'], '---')
        self.assertIn('no es una prueba', errores[0])
        self.assertIn('no es una prueba', cambios_documento({'pruebas': [fila]}, [fila])[0].mensaje_metodo)
        self.assertTrue(mejor_coincidencia(fila['prueba'], '---', [MAESTRO], 'ESPANOL')['fila_separadora'])
        self.assertFalse(es_separador_pruebas(fila['prueba'], 'ASTM D892'))

    def test_separador_cada_lote_no_se_alinea_y_se_marca_en_rojo(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        from cambios_alineacion import cambios_documento
        fila = {'prueba': '\u25baPRUEBA   CADA LOTE', 'metodo': '', 'unidades': ''}
        r = mejor_coincidencia(fila['prueba'], '', [MAESTRO], 'ESPANOL')
        self.assertTrue(r['fila_separadora'])
        self.assertFalse(r.get('metodo_no_encontrado'))
        errores = {}
        filas, conteo = alinear_filas_generador([fila], [MAESTRO], 'es', errores_metodo=errores)
        self.assertEqual(conteo['alineados'], 0)
        self.assertIn('no es una prueba', errores[0])
        cambios = cambios_documento({'pruebas': [fila]}, [fila])
        self.assertIn('no es una prueba', cambios[0].mensaje_metodo)

    def test_densidad_relativa_15_6_conjunto_confirmado(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'DENSIDAD RELATIVA 15,6 \u00b0C/15,6 \u00b0C'
        en = 'RELATIVE DENSITY AT 15,6 \u00b0C/15,6 \u00b0C'
        registro = self.registro(es, en, 'ASTM D1298', 'ADIMENSIONAL')
        registro['UNITS'] = 'DIMENSIONLESS'
        nombre = 'DENSIDAD   RELATIVA A 15.6/15.6 \u00b0C'
        metodo = 'ASTM D1298 - 12b (2017) / ASTM   D7042 - 14'
        for maestro in ([registro], MaestroPreparado([registro])):
            for idioma, esperado in (('es', es), ('en', en)):
                with self.subTest(idioma=idioma):
                    filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo}], maestro, idioma)
                    self.assertEqual(conteo['alineados'], 1)
                    self.assertEqual(filas[0]['prueba'], esperado)
                    self.assertEqual(filas[0]['metodo'], 'ASTM D1298 / ASTM D7042')
        for n, m in ((nombre.replace('15.6', '20'), metodo),
                     (nombre, metodo + ' / ASTM D9999')):
            self.assertEqual(equivalencias_confirmadas(n, m)[1], m)

    def test_separabildad_fluidos_sinteticos_54_confirmada(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'SEPARABILIDAD AGUA-ACEITE A 54 \u00b0C'
        en = 'WATER-OIL SEPARABILITY AT 54 \u00b0C'
        registros = [self.registro('SEPARABILIDAD AGUA-ACEITE A 82 \u00b0C', 'WATER-OIL SEPARABILITY AT 82 \u00b0C', 'ASTM D1401'),
                     self.registro(es, en, 'ASTM D1401')]
        nombre = 'SEPARABILDAD   DEL AGUA Y DEL ACEITE Y FLUIDOS SINTETICOS A 54\u00b0C'
        for maestro in (registros, MaestroPreparado(registros)):
            for metodo in ('', 'NA', 'ASTM D1401 - 19'):
                for idioma, esperado in (('es', es), ('en', en)):
                    with self.subTest(metodo=metodo, idioma=idioma):
                        filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo}], maestro, idioma)
                        self.assertEqual(conteo['alineados'], 1)
                        self.assertEqual(filas[0]['prueba'], esperado)
                        self.assertEqual(filas[0]['metodo'], 'ASTM D1401')
        self.assertEqual(equivalencias_confirmadas(nombre, 'ASTM D2711')[0], nombre)
        otro = nombre.replace('54', '82')
        self.assertEqual(equivalencias_confirmadas(otro, '')[0], otro)

    def test_d972_variantes_evaporacion_confirmadas(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'P\u00c9RDIDA POR EVAPORACI\u00d3N DE GRASAS LUBRICANTES Y ACEITES'
        en = 'EVAPORATION LOSS OF LUBRICATING GREASES AND OILS'
        registro = self.registro(es, en, 'ASTM D972', '%')
        nombres = ('VOLATILIDAD NOACK, 22 hrs, 225 \u00b0F', 'PERDIDA POR EVAPORACION',
                   'VOLATILITY, 22 HRS, 225 \u00b0F', 'P\u00c9RDIDA POR EVAPORACI\u00d3N',
                   'VOLATILIDAD, 22 HRS A 225 \u00b0F \u2666', 'VOLATILIDAD, 22 HORAS A 225 \u00b0F',
                   'VOLATILIDAD NOACK', 'VOLATILITY, 22 hr/225\u00ba F',
                   'PERDIDA POR EVAPORACI\u00d3N (99\u00b0C, 22 HR)',
                   'PERDIDAD POR EVAPORACION DE GRASAS Y ACEITES LUBRICANTES',
                   'P\u00c9RDIDA POR EVAPORACI\u00d3N, 22 HRS, 225 \u00b0F',
                   'VOLATILITY, 22 hr @ 225\u00b0F', 'Volatility, 22 hr @ 225 \u00b0F, mass%',
                   'Volatility, 22 hr @ 2250F, mass%', 'Volatility, 22 hr @ 225\u00b0F, mass%',
                   'Volatility, 22 hr @ 225\u00baF, mass%', 'PERDIDA POR EVAPORACI\u00d3N A 120\u00b0C')
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre in nombres:
                for idioma, esperado in (('es', es), ('en', en)):
                    with self.subTest(nombre=nombre, idioma=idioma):
                        filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': 'ASTM D972 - 02 (2008)'}], maestro, idioma)
                        self.assertEqual(conteo['alineados'], 1)
                        self.assertEqual(filas[0]['prueba'], esperado)
                        self.assertEqual(filas[0]['metodo'], 'ASTM D972')
        nombre = 'VOLATILIDAD NOACK'
        self.assertEqual(equivalencias_confirmadas(nombre, 'ASTM D5800')[0], nombre)

    def test_espuma_secuencias_confirmadas_cortas_y_tendencia(self):
        from alinear_especificaciones_json import MaestroPreparado, es_separador_pruebas, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        registros = [self.registro('FORMACI\u00d3N DE ESPUMA (SECUENCIA '+sec+')',
                     'FOAMING CHARACTERISTICS (SEQUENCE '+sec+')', 'ASTM D892', 'ml/ml') for sec in ('I', 'II', 'III')]
        for maestro in (registros, MaestroPreparado(registros)):
            for sec in ('I', 'II', 'III'):
                for nombre in ('TENDENCIA   A LA ESPUMA, SECUENCIA '+sec, 'SECUENCIA   '+sec, 'SECUENCIA '+sec+'\u25ba'):
                    for metodo in ('', '---', 'ASTM D892 - 18'):
                        for idioma in ('es', 'en'):
                            with self.subTest(sec=sec, nombre=nombre, metodo=metodo, idioma=idioma):
                                filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo}], maestro, idioma)
                                self.assertEqual(conteo['alineados'], 1)
                                self.assertEqual(filas[0]['prueba'], registros[('I','II','III').index(sec)]['NOMBRE IQM' if idioma == 'es' else 'NOMBRE IQM INGLES'])
                                self.assertEqual(filas[0]['metodo'], 'ASTM D892')
        self.assertTrue(es_separador_pruebas('TENDENCIA A LA ESPUMA', '---'))
        self.assertEqual(equivalencias_confirmadas('SECUENCIA I', 'ASTM D9999')[0], 'SECUENCIA I')

    def test_separacion_agua_aceite_82_selecciona_fila_correcta(self):
        from alinear_especificaciones_json import MaestroPreparado
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'SEPARABILIDAD AGUA-ACEITE A 82 \u00b0C'
        en = 'WATER-OIL SEPARABILITY AT 82 \u00b0C'
        registros = [self.registro('SEPARABILIDAD AGUA-ACEITE A 54 \u00b0C', 'WATER-OIL SEPARABILITY AT 54 \u00b0C', 'ASTM D1401', 'mL/mL/mL (min)'),
                     self.registro(es, en, 'ASTM D1401', 'mL/mL/mL (min)')]
        for maestro in (registros, MaestroPreparado(registros)):
            for nombre in ('SEPARACI\u00d3N   AGUA ACEITE A 82\u00b0C', 'SEPARACION AGUA ACEITE A 82 \u00b0C'):
                for idioma, esperado in (('es', es), ('en', en)):
                    with self.subTest(nombre=nombre, idioma=idioma):
                        filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': 'ASTM D1401 - 19'}], maestro, idioma)
                        self.assertEqual(conteo['alineados'], 1)
                        self.assertEqual(filas[0]['prueba'], esperado)
                        self.assertEqual(filas[0]['metodo'], 'ASTM D1401')

    def test_separacion_agua_aceite_54_es_separabilidad(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'SEPARABILIDAD AGUA-ACEITE A 54 \u00b0C'
        en = 'WATER-OIL SEPARABILITY AT 54 \u00b0C'
        registros = [self.registro('SEPARABILIDAD AGUA-ACEITE A 82 \u00b0C', 'WATER-OIL SEPARABILITY AT 82 \u00b0C', 'ASTM D1401', 'mL/mL/mL (min)'),
                     self.registro(es, en, 'ASTM D1401', 'mL/mL/mL (min)')]
        for maestro in (registros, MaestroPreparado(registros)):
            for nombre in ('SEPARACION   AGUA ACEITE A 54\u00b0C', 'SEPARACI\u00d3N AGUA-ACEITE A 54 \u00b0C', 'DEMULSIBILIDAD'):
                for idioma, esperado in (('es', es), ('en', en)):
                    with self.subTest(nombre=nombre, idioma=idioma):
                        for _ in range(2):
                            filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': 'ASTM D1401 - 19', 'unidades': ''}], maestro, idioma)
                            self.assertEqual(conteo['alineados'], 1)
                            self.assertEqual(filas[0]['prueba'], esperado)
                            self.assertEqual(filas[0]['metodo'], 'ASTM D1401')
        for nombre, metodo in (('SEPARACION AGUA ACEITE A 90\u00b0C', 'ASTM D1401'),
                               ('SEPARACION AGUA ACEITE A 54\u00b0F', 'ASTM D1401'),
                               ('SEPARACION AGUA ACEITE A 54\u00b0C', 'ASTM D2711')):
            self.assertEqual(equivalencias_confirmadas(nombre, metodo)[0], nombre)

    def test_viscosidad_100_d445_confirmada_es_cinematica(self):
        from alinear_especificaciones_json import MaestroPreparado
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'VISCOSIDAD CINEM\u00c1TICA A 100 \u00b0C'
        en = 'KINEMATIC VISCOSITY AT 100 \u00b0C'
        registros = [self.registro('VISCOSIDAD DIN\u00c1MICA A 100 \u00b0C', 'DYNAMIC VISCOSITY AT 100 \u00b0C', 'ASTM D445', 'mPa*s'),
                     self.registro('VISCOSIDAD CINEM\u00c1TICA A 40 \u00b0C', 'KINEMATIC VISCOSITY AT 40 \u00b0C', 'ASTM D445', 'mm\u00b2/s'),
                     self.registro(es, en, 'ASTM D445', 'mm\u00b2/s')]
        for maestro in (registros, MaestroPreparado(registros)):
            for idioma, esperado in (('es', es), ('en', en)):
                with self.subTest(idioma=idioma):
                    for _ in range(2):
                        filas, conteo = alinear_filas_generador([{'prueba': 'VISCOSIDAD   A 100\u00b0C', 'metodo': 'ASTM   D445 - 19a', 'unidades': ''}], maestro, idioma)
                        self.assertEqual(conteo['alineados'], 1)
                        self.assertEqual(filas[0]['prueba'], esperado)
                        self.assertEqual(filas[0]['unidades'], 'mm\u00b2/s')
        conflicto = mejor_coincidencia('VISCOSIDAD A 100\u00b0C', 'ASTM D445', registros, 'ESPANOL', 'cP')
        self.assertEqual(conflicto['registro']['UNIDADES'], 'mPa*s')

    def test_viscosidad_40_d445_confirmada_es_cinematica(self):
        from alinear_especificaciones_json import MaestroPreparado
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'VISCOSIDAD CINEM\u00c1TICA A 40 \u00b0C'
        en = 'KINEMATIC VISCOSITY AT 40 \u00b0C'
        registros = [self.registro('VISCOSIDAD DIN\u00c1MICA A 40 \u00b0C', 'DYNAMIC VISCOSITY AT 40 \u00b0C', 'ASTM D445', 'mPa*s'),
                     self.registro(es, en, 'ASTM D445', 'mm\u00b2/s')]
        for maestro in (registros, MaestroPreparado(registros)):
            for idioma, esperado in (('es', es), ('en', en)):
                with self.subTest(idioma=idioma):
                    filas, conteo = alinear_filas_generador([{'prueba': 'VISCOSIDAD   A 40\u00b0C', 'metodo': 'ASTM   D445 - 19a', 'unidades': ''}], maestro, idioma)
                    self.assertEqual(conteo['alineados'], 1)
                    self.assertEqual(filas[0]['prueba'], esperado)
                    self.assertEqual(filas[0]['unidades'], 'mm\u00b2/s')
        conflicto = mejor_coincidencia('VISCOSIDAD A 40\u00b0C', 'ASTM D445', registros, 'ESPANOL', 'cP')
        self.assertEqual(conflicto['registro']['UNIDADES'], 'mPa*s')

    def test_escurrimiento_d97_es_fluidez_general(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        registros = [self.registro('PUNTO DE FLUIDEZ SUPERIOR', 'HIGHER POUR POINT', 'ASTM D97', '\u00b0C'),
                     self.registro('PUNTO DE FLUIDEZ INFERIOR', 'LOWER POUR POINT', 'ASTM D97', '\u00b0C'),
                     self.registro('PUNTO DE FLUIDEZ', 'POUR POINT', 'ASTM D97', '\u00b0C')]
        for maestro in (registros, MaestroPreparado(registros)):
            for idioma, esperado in (('es', 'PUNTO DE FLUIDEZ'), ('en', 'POUR POINT')):
                with self.subTest(idioma=idioma):
                    for _ in range(2):
                        filas, conteo = alinear_filas_generador([{'prueba': 'PUNTO DE   ESCURRIMIENTO', 'metodo': 'ASTM D97   - 17b', 'unidades': '\u00b0C'}], maestro, idioma)
                        self.assertEqual(conteo['alineados'], 1)
                        self.assertEqual(filas[0]['prueba'], esperado)
                        self.assertEqual(filas[0]['metodo'], 'ASTM D97')
        self.assertEqual(equivalencias_confirmadas('PUNTO DE ESCURRIMIENTO', 'ASTM D5950')[0], 'PUNTO DE ESCURRIMIENTO')

    def test_puntode_anilina_d611_metodo_e(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'PUNTO DE ANILINA, M\u00c9TODO E'
        en = 'ANILINE POINT, METHOD E'
        registro = self.registro(es, en, 'ASTM D611', '\u00b0C, \u00b0F')
        for maestro in ([registro], MaestroPreparado([registro])):
            for idioma, esperado in (('es', es), ('en', en)):
                with self.subTest(idioma=idioma):
                    for _ in range(2):
                        filas, conteo = alinear_filas_generador([{'prueba': 'PUNTODE   ANILINA', 'metodo': 'ASTM D611 - 12 (M\u00e9todo E) (2016)', 'unidades': '\u00b0C'}], maestro, idioma)
                        self.assertEqual(conteo['alineados'], 1)
                        self.assertEqual(filas[0]['prueba'], esperado)
                        self.assertEqual(filas[0]['metodo'], 'ASTM D611')
        for metodo in ('ASTM D611 (METHOD A)', 'ASTM D611', 'ASTM D612 (METHOD E)'):
            self.assertEqual(equivalencias_confirmadas('PUNTODE ANILINA', metodo)[0], 'PUNTODE ANILINA')

    def test_punto_inflamacion_d92_usa_cleveland_101_3(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'PUNTO DE INFLAMACI\u00d3N A 101,3 kPa, COPA ABIERTA CLEVELAND'
        en = 'FLASH POINT AT 101,3 kPa, CLEVELAND OPEN CUP'
        registros = [self.registro('PUNTO DE INFLAMACI\u00d3N A 760 mmHg, COPA ABIERTA CLEVELAND',
                                  'FLASH POINT AT 760 mmHg, CLEVELAND OPEN CUP', 'ASTM D92', '\u00b0C, \u00b0F'),
                     self.registro(es, en, 'ASTM D92', '\u00b0C')]
        for maestro in (registros, MaestroPreparado(registros)):
            for nombre in ('PUNTO DE   INFLAMACI\u00d3N', 'PUNTO DE   INFLAMACION', 'PUNTO DE   INFLAMACION, \u00b0C'):
                for idioma, esperado in (('es', es), ('en', en)):
                    with self.subTest(nombre=nombre, idioma=idioma):
                        filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': 'ASTM D92 - 18', 'unidades': '\u00b0C'}], maestro, idioma)
                        self.assertEqual(conteo['alineados'], 1)
                        self.assertEqual(filas[0]['prueba'], esperado)
                        self.assertEqual(filas[0]['metodo'], 'ASTM D92')
        nombre = 'PUNTO DE INFLAMACION'
        self.assertEqual(equivalencias_confirmadas(nombre, 'ASTM D93')[0], nombre)
        especifico = 'PUNTO DE INFLAMACION A 760 mmHg'
        self.assertEqual(equivalencias_confirmadas(especifico, 'ASTM D92')[0], especifico)

    def test_azufre_farmacopeas_confirmado_sin_confundir_otras_normas(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'L\u00cdMITE DE COMPUESTOS DE AZUFRE'
        en = 'LIMIT OF SULFUR COMPOUNDS'
        registro = self.registro(es, en, 'USP')
        casos = (('AZUFRE', 'USP 41'), ('AZUFRE Y   SULFUROS', 'USP 24'),
                 ('L\u00cdMITE   DE COMPUESTOS DE AZUFRE', 'FEUM 13/USP 32'))
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre, metodo in casos:
                for idioma, esperado in (('es', es), ('en', en)):
                    with self.subTest(nombre=nombre, idioma=idioma):
                        for _ in range(2):
                            filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo, 'unidades': 'A/D'}], maestro, idioma)
                            self.assertEqual(conteo['alineados'], 1)
                            self.assertEqual(filas[0]['prueba'], esperado)
                            self.assertEqual(filas[0]['metodo'], 'USP')
                            self.assertEqual(filas[0]['unidades'], 'NA')
        for nombre, metodo in (('AZUFRE', 'ASTM D2622'), ('AZUFRE', ''),
                               ('AZUFRE', 'USP / ASTM D2622')):
            self.assertEqual(equivalencias_confirmadas(nombre, metodo)[0], nombre)

    def test_acidez_confirmada_usa_fila_usp_feum(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        registros = [self.registro('ACIDEZ', 'ACIDITY', m) for m in ('USP', 'USP / FEUM')]
        casos = (('ACIDEZ\u2666', 'USP 32'), ('ACIDITY', 'CURRENT USP'),
                 ('ACIDITY', 'USP (VIGENTE)'), ('ACIDITY', 'USP 38'), ('ACIDITY', 'USP 41'),
                 ('ACIDITY (NO RED OR PINK COLOR IS PRODUCED)', 'USP 32'),
                 ('ACIDITY AND ALKALINITY', 'FEUM 13/USP 32'),
                 ('ACIDITY AND ALKALINITY', 'USP 32'), ('Acidity', ''), ('Acidity', 'USP/NF'), ('ACIDEZ \u2666', 'USP 37'),
                 ('ACIDEZ', 'Ph. Eur 10th Edition'), ('ACIDEZ', 'USP VIGENTE'),
                 ('ACIDEZ %', ''))
        for nombre, metodo in casos:
            for idioma, esperado in (('es', 'ACIDEZ'), ('en', 'ACIDITY')):
                with self.subTest(nombre=nombre, metodo=metodo, idioma=idioma):
                    filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo, 'unidades': 'NA'}], registros, idioma)
                    self.assertEqual(conteo['alineados'], 1)
                    self.assertEqual(filas[0]['prueba'], esperado)
                    self.assertEqual(filas[0]['metodo'], 'USP / FEUM')

    def test_agua_d1533_liquidos_aislantes_conserva_apartado_nmx(self):
        from alinear_especificaciones_json import MaestroPreparado
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'CONTENIDO DE AGUA POR KARL FISCHER (LIQUIDOS AISLANTES)'
        en = 'WATER CONTENT BY KARL FISCHER (INSULATING LIQUIDS)'
        nmx = 'NMX-J-123-ANCE-2019 P 6,11'
        registros = [self.registro(es, en, m, 'mg/kg , ppm') for m in ('ASTM D1533', 'ASTM D1533 / '+nmx)]
        for maestro in (registros, MaestroPreparado(registros)):
            for metodo, esperado_metodo in (('ASTM D1533 - 12', 'ASTM D1533'),
                    ('ASTM D1533 - 12 /   '+nmx, 'ASTM D1533 / '+nmx)):
                for idioma, esperado in (('es', es), ('en', en)):
                    with self.subTest(metodo=metodo, idioma=idioma):
                        for _ in range(2):
                            filas, conteo = alinear_filas_generador([{'prueba': 'CONTENIDO   DE AGUA', 'metodo': metodo, 'unidades': 'ppm'}], maestro, idioma)
                            self.assertEqual(conteo['alineados'], 1)
                            self.assertEqual(filas[0]['prueba'], esperado)
                            self.assertEqual(filas[0]['metodo'], esperado_metodo)
                            self.assertEqual(filas[0]['unidades'], 'ppm')
        r = mejor_coincidencia('CONTENIDO DE AGUA', 'ASTM D1533 / '+nmx.replace('6,11', '6,12'), registros, 'ESPANOL')
        self.assertTrue(r['metodo_no_encontrado'])

    def test_agua_procedimiento_a_sin_metodo_confirmado(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'CONTENIDO DE AGUA POR KARL FISCHER (COULOM\u00c9TRICO)'
        en = 'WATER CONTENT BY KARL FISCHER (COULOMETRIC)'
        registro = self.registro(es, en, 'ASTM D6304', 'mg/kg')
        nombre = 'CONTENIDO   DE AGUA, PROCEDIMIENTO A'
        for maestro in ([registro], MaestroPreparado([registro])):
            for metodo in ('', 'NA', 'ASTM D6304 - 20 (PROCEDIMIENTO A)'):
                for idioma, esperado in (('es', es), ('en', en)):
                    with self.subTest(metodo=metodo, idioma=idioma):
                        for _ in range(2):
                            filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo, 'unidades': 'A/D'}], maestro, idioma)
                            self.assertEqual(conteo['alineados'], 1)
                            self.assertEqual(filas[0]['prueba'], esperado)
                            self.assertEqual(filas[0]['metodo'], 'ASTM D6304')
        for n, m in (('CONTENIDO DE AGUA', ''), ('CONTENIDO DE AGUA, PROCEDIMIENTO B', ''),
                     (nombre, 'ASTM D95')):
            self.assertEqual(equivalencias_confirmadas(n, m)[0], n)

    def test_d6304_procedimiento_a_variantes_de_agua(self):
        from alinear_especificaciones_json import MaestroPreparado, equivalencias_confirmadas
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        es = 'CONTENIDO DE AGUA POR KARL FISCHER (COULOM\u00c9TRICO)'
        en = 'WATER CONTENT BY KARL FISCHER (COULOMETRIC)'
        registro = self.registro(es, en, 'ASTM D6304', 'mg/kg')
        nombres = ('% DE   HUMEDAD', 'AGUA, %   M', 'AGUA, KF', 'CONTENIDO   DE AGUA',
                   'CONTENIDO   DE AGUA POR KARL FISHER', 'CONTENIDO   DE AGUA, PROCEDIMIENTO A',
                   'WATER   CONTENT, PROCEDURE A', 'Water')
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre in nombres:
                for metodo in ('ASTM D6304 - 20 (PROCEDIMIENTO   A)', 'ASTM D6304-25 (PROCEDURE A)'):
                    for idioma, esperado in (('es', es), ('en', en)):
                        with self.subTest(nombre=nombre, metodo=metodo, idioma=idioma):
                            filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo, 'unidades': 'A/D'}], maestro, idioma)
                            self.assertEqual(conteo['alineados'], 1)
                            self.assertEqual(filas[0]['prueba'], esperado)
                            self.assertEqual(filas[0]['metodo'], 'ASTM D6304')
                            self.assertEqual(filas[0]['unidades'], 'mg/kg')
        for nombre, metodo in (('Water', 'ASTM D6304 (PROCEDURE B)'),
                               ('Water', 'ASTM D6304 (PROCEDURE C)'),
                               ('Water', 'ASTM D6304 / ASTM D95 (PROCEDURE A)'),
                               ('COLOR', 'ASTM D6304 (PROCEDURE A)')):
            self.assertEqual(equivalencias_confirmadas(nombre, metodo)[0], nombre)

    def test_lovibond_variantes_de_celda_color_y_temperatura(self):
        from alinear_especificaciones_json import MaestroPreparado
        registro = self.registro('COLOR LOVIBOND RYBN', 'LOVIBOND COLOR RYBN', 'IP17 A')
        casos = (('LOVIBOND', 'NA'),
                 ('COLOR LOVIBOND, CELDA DE 2 PULGADAS', 'USP 32'),
                 ('COLOR ROJO LOVIBOND', ''), ('COLOR AMARILLO LOVIBOND', ''),
                 ('In 2-Inch Cell Lovibond @ 90 \u00b0C', ''),
                 ('COLOR LOVIBOND EN CELDA DE 2 "', 'LOVIBOND'),
                 ('COLOR IN 2-INCH CELL LOVIBOND', 'LOVIBOND'),
                 ('Color Lovibond en celda de 2 \u201c', 'LOVIBOND'),
                 ('DETERMINACION DEL COLOR LOVIBOND CELDA 2"', 'LOVIBOND'),
                 ('LOVIBOND COLOR', 'LOVIBOND'), ('LOVIBOND COLOR IN 2" CELL', 'LOVIBOND'),
                 ('In 2-Inch Cell Lovibond @ 90 \u00b0C', 'LOVIBOND'))
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre, metodo in casos:
                for idioma in ('ESPANOL', 'INGLES'):
                    with self.subTest(nombre=nombre, metodo=metodo, idioma=idioma):
                        r = mejor_coincidencia(nombre, metodo, maestro, idioma, 'NA')
                        self.assertEqual(r['estado'], 'ALINEADO')
                        self.assertEqual(r['registro']['NOMBRE IQM'], 'COLOR LOVIBOND RYBN')
                        self.assertEqual(r['registro']['METODO'], 'IP17 A')
                        from interfaz_aplicacion_especificaciones import alinear_filas_generador
                        filas, _ = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo, 'unidades': 'NA'}], maestro, 'es' if idioma == 'ESPANOL' else 'en')
                        self.assertEqual(filas[0]['metodo'], 'IP17 A')
        self.assertTrue(mejor_coincidencia('COLOR LOVIBOND', 'ASTM D1500', [registro], 'ESPANOL')['metodo_no_encontrado'])

    def test_variantes_color_visual_respetan_metodo_y_maestro(self):
        from alinear_especificaciones_json import MaestroPreparado
        registro = self.registro('COLOR', 'COLOR', 'VISUAL')
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre in ('COLOR VISUAL', 'VISUAL COLOR', 'DETERMINACI\u00d3N DE COLOR'):
                for idioma in ('ESPANOL', 'INGLES'):
                    with self.subTest(nombre=nombre, idioma=idioma):
                        r = mejor_coincidencia(nombre, 'VISUAL', maestro, idioma, 'NA')
                        self.assertEqual(r['estado'], 'ALINEADO')
                        self.assertEqual(r['registro']['NOMBRE IQM'], 'COLOR')
                        self.assertEqual(r['registro']['METODO'], 'VISUAL')
                self.assertTrue(mejor_coincidencia(nombre, 'ASTM D1500', maestro, 'ESPANOL')['metodo_no_encontrado'])

    def test_humidity_y_humedad_sin_metodo_o_cliente_son_crepitacion(self):
        from alinear_especificaciones_json import MaestroPreparado
        registro = self.registro('HUMEDAD POR CREPITACI\u00d3N', 'MOISTURE BY CRACKLE', 'M\u00c9TODO INTERNO (91300029)')
        nombres = ('HUMEDAD', 'HUMIDITY', 'HUMIDITY (SPLATTER TEST)',
                   'HUMIDITY CREPITATION', 'CREPITATION HUMIDITY',
                   'HUMEDAD POR CREPITACION', 'Humidity \u25ca', 'HUMEDAD \u25ca \u25ba')
        metodos = ('', 'NA', 'N/A', 'M\u00c9TODO DEL CLIENTE', 'Client Method', 'INS-CCA-07-00', '91300092')
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre in nombres:
                for metodo in metodos:
                    with self.subTest(nombre=nombre, metodo=metodo):
                        r = mejor_coincidencia(nombre, metodo, maestro, 'ESPANOL', 'NA')
                        self.assertEqual(r['estado'], 'ALINEADO')
                        self.assertEqual(r['registro']['NOMBRE IQM'], registro['NOMBRE IQM'])
                self.assertTrue(mejor_coincidencia(nombre, 'ASTM D6304', maestro, 'ESPANOL')['metodo_no_encontrado'])
        self.assertNotEqual(mejor_coincidencia('', 'CLIENT METHOD', [registro], 'ESPANOL')['estado'], 'ALINEADO')

    def test_aceites_fijos_resin_rosin_con_texto_adicional(self):
        from alinear_especificaciones_json import MaestroPreparado
        registro = self.registro('ACEITES FIJOS, GRASAS Y RESINA', 'FIXED OILS, FATS, AND RESINS', 'USP / FEUM')
        nombres = ('FIXED OILS, FATS AND RESIN (NO OILY OR SOLID MATTER SAPARATES)',
                   'FIXED OILS, FATS AND RESIN \n (NO OILY OR SOLID MATTER SAPARATES)',
                   'Fixed Oils, Fats And Rosin',
                   'FIXED OILS, FATS AND ROSIN (NO OILY OR SOLID MATTER SEPARATES) B',
                   'FIXED OILS, FATS AND RESIN (NO OILY OR SOLID MATTER SEPARATES) B',
                   'FIXED OILS, FATS AND ROSIN (NO OILY OR SOLID MATTER SEPARATES)')
        for maestro in ([registro], MaestroPreparado([registro])):
            for nombre in nombres:
                for idioma, campo in (('ESPANOL', 'NOMBRE IQM'), ('INGLES', 'NOMBRE IQM INGLES')):
                    with self.subTest(nombre=nombre, idioma=idioma):
                        r = mejor_coincidencia(nombre, 'USP 32', maestro, idioma, 'N/A')
                        self.assertEqual(r['estado'], 'ALINEADO')
                        self.assertEqual(r['registro'][campo], registro[campo])
                self.assertTrue(mejor_coincidencia(nombre, 'NOM-092-SSA1-1994', maestro, 'ESPANOL')['metodo_no_encontrado'])

    def test_carbonos_abreviaturas_y_familias_sin_listas_de_nombres(self):
        from alinear_especificaciones_json import _tipo_carbono
        casos = {'CA': ('%Ca', 'Ca', 'AROMATICS CONTENT', 'COMPUESTOS AROMATICOS (Ca)'),
                 'CN': ('%Cn', 'CN', '% CARBONES NAFTEN ICOS', 'Naphtenic Carbons', 'NAPHTHENICS COMPOUNDS (Cn)'),
                 'CP': ('%Cp', 'CP', 'PARAFINICS COMPOUNDS (Cp)', 'Parafinic Carbons')}
        for tipo, nombres in casos.items():
            for nombre in nombres:
                with self.subTest(nombre=nombre):
                    self.assertEqual(_tipo_carbono(nombre), tipo)
        for nombre in ('Refractivity intercept', 'Viscosity Gravity Constant (VGC)',
                       'COMPOSITION (%CA, %CN, %CP)', 'AROMATIC (CN)', 'CADENAS'):
            self.assertEqual(_tipo_carbono(nombre), '')

    def test_espuma_ingles_y_genericos_reconocen_metodo_sin_inventar_secuencia(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        maestro = [self.registro('FORMACIÓN DE ESPUMA (SECUENCIA '+s+')', 'FOAMING CHARACTERISTICS (SEQUENCE '+s+')', 'ASTM D892', 'ml/ml') for s in ('I', 'II', 'III')]
        for sec in ('I', 'II', 'III'):
            r = mejor_coincidencia('SEQUENCE '+sec+', TENDENCY/STABILITY', '', maestro, 'INGLES')
            self.assertEqual(r['estado'], 'ALINEADO')
            self.assertIn('(SEQUENCE '+sec+')', r['registro']['NOMBRE IQM INGLES'])
        for nombre in ('TENDENCIA A LA ESPUMA', 'ESPUMA', 'ESTABILIDAD A LA ESPUMACION', 'Prueba a la Espuma'):
            for metodo in ('', '---', 'ASTM D892 -25 (2025)'):
                filas, conteo = alinear_filas_generador([{'prueba': nombre, 'metodo': metodo}], maestro, 'es')
                if nombre == 'TENDENCIA A LA ESPUMA' and metodo == '---':
                    self.assertEqual(filas[0]['metodo'], '---')
                    self.assertEqual(filas[0]['prueba'], nombre)
                    self.assertEqual(conteo['revision'], 1)
                    continue
                self.assertEqual(filas[0]['metodo'], 'ASTM D892')
                self.assertEqual(filas[0]['prueba'], 'FORMACIÓN DE ESPUMA')
                self.assertEqual(conteo['revision'], 1)
        r = mejor_coincidencia('INDICE DE VISCOSIDAD', 'ASTM D2270 -25', [self.registro('ÍNDICE DE VISCOSIDAD', 'VISCOSITY INDEX', 'ASTM D2270'), *maestro], 'ESPANOL')
        self.assertEqual(r['registro']['METODO'], 'ASTM D2270')

    def test_visual_appearance_es_alias_confirmado(self):
        registro = self.registro('APARIENCIA VISUAL', 'VISUAL APPEAREANCE', 'ORGANOLÉPTICO')
        for metodo in ('VISUAL', 'ORGANOLEPTIC', ''):
            r = mejor_coincidencia('VISUAL APPEARANCE', metodo, [registro], 'INGLES')
            self.assertEqual(r['estado'], 'ALINEADO')
            self.assertEqual(r['registro']['NOMBRE IQM'], 'APARIENCIA VISUAL')
        registro['MÉTODO INGLÉS'] = 'ORGANOLEPTIC'
        r = mejor_coincidencia('VISUAL APPEARANCE', 'VISUAL', [registro], 'INGLES')
        self.assertEqual(r['estado'], 'ALINEADO')

    @staticmethod
    def registro(nombre, ingles, metodo, unidad='NA'):
        return {**MAESTRO, 'NOMBRE IQM': nombre, 'NOMBRE IQM INGLES': ingles, 'METODO': metodo,
                'UNIDADES': unidad, 'UNITS': unidad}

    def test_unidades_seleccionan_la_magnitud_sin_depender_del_orden(self):
        from alinear_especificaciones_json import MaestroPreparado
        cin = self.registro('VISCOSIDAD CINEMÁTICA A 40 °C', 'KINEMATIC VISCOSITY AT 40 °C', 'ASTM D445', 'mm²/s')
        din = self.registro('VISCOSIDAD DINÁMICA A 40 °C', 'DYNAMIC VISCOSITY AT 40 °C', 'ASTM D445', 'mPa*s')
        for unidades, esperado in (('cSt', cin), ('mm2/seg (cSt)', cin), ('cP', din), ('mPa.s', din)):
            for registros in ([cin, din], [din, cin]):
                for maestro in (registros, MaestroPreparado(registros)):
                    with self.subTest(unidades=unidades, preparado=isinstance(maestro, MaestroPreparado)):
                        match = mejor_coincidencia('VISCOSIDAD A 40 °C', 'ASTM D445', maestro, 'ESPANOL', unidades)
                        self.assertEqual(match['estado'], 'ALINEADO')
                        self.assertEqual(match['registro']['NOMBRE IQM'], esperado['NOMBRE IQM'])
        match = mejor_coincidencia('VISCOSIDAD A 40 °C cSt', 'ASTM D445', [din, cin], 'ESPANOL')
        self.assertEqual(match['registro']['UNIDADES'], 'mm²/s')
        self.assertNotEqual(mejor_coincidencia(cin['NOMBRE IQM'], 'ASTM D445', [cin, din], 'ESPANOL', 'cP')['estado'], 'ALINEADO')

    def test_regla_temperaturas_retirada_en_maestro_simple_y_preparado(self):
        from alinear_especificaciones_json import MaestroPreparado
        fija = self.registro('VISCOSIDAD CINEMÁTICA A 40 °C', 'KINEMATIC VISCOSITY AT 40 °C', 'ASTM D445', 'cSt')
        plantilla = self.registro('VISCOSIDAD CINEMÁTICA A X °C', 'KINEMATIC VISCOSITY AT X °C', 'ASTM D445', 'cSt')
        generica = self.registro('VISCOSIDAD CINEMÁTICA', 'KINEMATIC VISCOSITY', 'ASTM D445', 'cSt')
        for registro in (fija, plantilla, generica):
            for maestro in ([registro], MaestroPreparado([registro])):
                for nombre in ('VISCOSIDAD CINEMÁTICA A 15 °C', 'VISCOSIDAD CINEMÁTICA A 100 °F'):
                    with self.subTest(destino=registro['NOMBRE IQM'], nombre=nombre, preparado=isinstance(maestro, MaestroPreparado)):
                        match = mejor_coincidencia(nombre, 'ASTM D445', maestro, 'ESPANOL', 'cSt')
                        self.assertEqual(match['estado'], 'ALINEADO')
                        self.assertFalse(match.get('condiciones_incompatibles'))
                        self.assertEqual(match['registro']['NOMBRE IQM'], registro['NOMBRE IQM'])
                        self.assertEqual(maestro[0], registro)

    def test_alias_confirmados_se_aplican_en_json_editor_y_hii_y_se_auditan(self):
        from generador_hii import preparar_datos
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        from alinear_especificaciones_json import MaestroPreparado
        casos = [
            ('IV', 'ASTM D2270 - 10 (2016)', 'ÍNDICE DE VISCOSIDAD', 'VISCOSITY INDEX', 'ASTM D2270'),
            ('COMPUESTOS CARBONIZABLES', 'USP 32', 'SUSTANCIAS FÁCILMENTE CARBONIZABLES', 'EASILY CARBONIZABLE SUBSTANCES', 'USP / FEUM'),
            ('COMPUESTOS POLINUCLEARES', 'USP 32', 'LÍMITE DE HIDROCARBUROS POLICÍCLICOS AROMÁTICOS', 'LIMIT OF POLYCYCLIC AROMATIC HYDROCARBONS', 'USP'),
            ('ASPECTO', 'VISUAL', 'APARIENCIA VISUAL', 'VISUAL APPEARANCE', 'ORGANOLÉPTICO'),
            ('ASPCTO', 'VISUAL', 'APARIENCIA VISUAL', 'VISUAL APPEARANCE', 'ORGANOLÉPTICO'),
            ('APARIENCIA', 'VISUAL', 'APARIENCIA VISUAL', 'VISUAL APPEARANCE', 'ORGANOLÉPTICO'),
            ('COLOR LOVIBOND', 'LOVIBOND', 'COLOR LOVIBOND RYBN', 'LOVIBOND COLOR RYBN', 'IP17 A'),
            ('HUMEDAD', 'INS-CCA-07-00', 'HUMEDAD POR CREPITACIÓN', 'MOISTURE BY CRACKLE', 'MÉTODO INTERNO (91300029)'),
            ('HUMEDAD', '91300092', 'HUMEDAD POR CREPITACIÓN', 'MOISTURE BY CRACKLE', 'MÉTODO INTERNO (91300029)'),
        ]
        for original, metodo, es, en, canonico in casos:
            for idioma, esperado in (('es', es), ('en', en)):
                with self.subTest(original=original, idioma=idioma), tempfile.TemporaryDirectory() as temporal:
                    maestro = MaestroPreparado([self.registro(es, en, canonico)])
                    fila = {'prueba': original, 'metodo': metodo, 'unidades': 'NA', 'especificacion': 'PASA', 'valor_tipico': ''}
                    origen, salida = Path(temporal) / 'original.json', Path(temporal) / 'alineado.json'
                    origen.write_text(json.dumps({'producto': {'nombre': 'Producto'}, 'documento': {'idioma': idioma},
                        'control_documental': {'codigo_unico_producto': 'P1', 'revision': '00'}, 'pruebas': [fila]}), encoding='utf-8')
                    for entrada in (origen, salida):
                        conteo = procesar_json(entrada, salida, maestro)
                        self.assertEqual(conteo['alineados'], 1)
                        datos = json.loads(salida.read_text(encoding='utf-8'))
                        self.assertEqual(datos['pruebas'][0]['prueba'], esperado)
                        self.assertEqual(preparar_datos(salida).pruebas[0]['prueba'], esperado)
                        detalle = datos['alineacion_metodos_lab']['detalle'][0]
                        self.assertEqual(detalle['nombre_original'], original)
                        self.assertEqual(detalle['metodo_original'], metodo)
                        self.assertTrue(detalle['reglas_equivalencia'])
                        if metodo == 'USP 32': self.assertEqual(datos['pruebas'][0]['metodo'], 'USP')
                    errores = {}
                    editadas, conteo = alinear_filas_generador([{**fila, 'metodo': metodo + ' ¹'}], maestro, idioma, errores_metodo=errores)
                    self.assertEqual(conteo['alineados'], 1)
                    self.assertEqual(editadas[0]['prueba'], esperado)
                    self.assertTrue(editadas[0]['metodo'].endswith(' ¹'))
                    self.assertEqual(errores, {})

    def test_alias_no_se_extienden_a_normas_distintas_o_codigos_internos_desconocidos(self):
        for nombre, metodo, canonico, correcto in (
                ('IV', 'ASTM D9999', 'ÍNDICE DE VISCOSIDAD', 'ASTM D2270'),
                ('HUMEDAD', 'INS-CCA-08-00', 'HUMEDAD POR CREPITACIÓN', 'MÉTODO INTERNO (91300029)'),
                ('COMPUESTOS CARBONIZABLES', 'NOM-092-SSA1', 'SUSTANCIAS FÁCILMENTE CARBONIZABLES', 'USP')):
            with self.subTest(nombre=nombre, metodo=metodo):
                match = mejor_coincidencia(nombre, metodo, [self.registro(canonico, canonico, correcto)], 'ESPANOL')
                self.assertTrue(match['metodo_no_encontrado'])
                self.assertNotEqual(match['estado'], 'ALINEADO')

    def test_polinucleares_normaliza_nombre_confirmado_sin_inventar_unidades(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        nombre = 'LÍMITE DE HIDROCARBUROS POLICÍCLICOS AROMÁTICOS'
        maestro = [self.registro(nombre, 'LIMIT OF POLYCYCLIC AROMATIC HYDROCARBONS', 'USP', unidad) for unidad in ('A', 'NA')]
        fila = {'prueba': 'COMPUESTOS POLINUCLEARES', 'metodo': 'USP 32', 'unidades': 'nm -3'}
        editadas, conteo = alinear_filas_generador([fila], maestro, 'es')
        self.assertEqual(editadas[0]['prueba'], nombre)
        self.assertEqual(editadas[0]['metodo'], 'USP')
        self.assertEqual(editadas[0]['unidades'], 'nm -3')
        self.assertEqual(conteo['revision'], 1)
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / 'prueba.json'
            ruta.write_text(json.dumps({'documento': {'idioma': 'es'}, 'pruebas': [fila]}), encoding='utf-8')
            procesar_json(ruta, ruta, maestro)
            datos = json.loads(ruta.read_text(encoding='utf-8'))
            self.assertEqual(datos['pruebas'][0]['prueba'], nombre)
            self.assertEqual(datos['pruebas'][0]['unidades'], 'nm -3')
            self.assertEqual(datos['alineacion_metodos_lab']['detalle'][0]['tipo_pendiente'], 'UNIDADES_POR_REVISAR')

    def test_unidades_y_x_llegan_al_editor_json_e_hii_en_ambos_idiomas(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        from generador_hii import preparar_datos
        maestro = [self.registro('VISCOSIDAD CINEMÁTICA A X °C', 'KINEMATIC VISCOSITY AT X °C', 'ASTM D445', 'mm²/s'),
                   self.registro('VISCOSIDAD DINÁMICA A X °C', 'DYNAMIC VISCOSITY AT X °C', 'ASTM D445', 'mPa*s')]
        for idioma, original, nombre in (('es', 'VISCOSIDAD A 25 °C cSt', 'VISCOSIDAD CINEMÁTICA A X °C'),
                                         ('en', 'VISCOSITY AT 25 °C cSt', 'KINEMATIC VISCOSITY AT X °C')):
            with self.subTest(idioma=idioma), tempfile.TemporaryDirectory() as temporal:
                fila = {'prueba': original, 'metodo': 'ASTM D445', 'unidades': 'cSt'}
                filas, conteo = alinear_filas_generador([fila], maestro, idioma)
                self.assertEqual(conteo['alineados'], 1)
                self.assertEqual(filas[0]['prueba'], nombre)
                self.assertEqual(filas[0]['unidades'], 'mm²/s')
                ruta = Path(temporal) / 'prueba.json'
                ruta.write_text(json.dumps({'producto': {'nombre': 'Producto'}, 'documento': {'idioma': idioma},
                    'control_documental': {'codigo_unico_producto': 'P1', 'revision': '00'}, 'pruebas': [fila]}), encoding='utf-8')
                for _ in range(2):
                    procesar_json(ruta, ruta, maestro)
                    datos = json.loads(ruta.read_text(encoding='utf-8'))
                    self.assertEqual(datos['pruebas'][0]['prueba'], nombre)
                    self.assertEqual(datos['pruebas'][0]['unidades'], 'mm²/s')
                    self.assertEqual(preparar_datos(ruta).pruebas[0]['prueba'], nombre)
                    self.assertEqual(datos['alineacion_metodos_lab']['detalle'][0]['nombre_original'], original)

    def test_espuma_recupera_el_metodo_solo_para_la_secuencia_confirmada(self):
        maestro = [self.registro('FORMACIÓN DE ESPUMA (SECUENCIA '+s+')', 'FOAMING CHARACTERISTICS (SEQUENCE '+s+')', 'ASTM D892', 'ml/ml') for s in ('I', 'II', 'III')]
        for sec in ('I', 'II', 'III'):
            for metodo in ('', '---', 'ASTM D892 - 18'):
                with self.subTest(sec=sec, metodo=metodo):
                    r = mejor_coincidencia('SEC '+sec+', TENDENCIA/ESTABILIDAD', metodo, maestro, 'ESPANOL')
                    self.assertEqual(r['estado'], 'ALINEADO')
                    self.assertEqual(r['registro']['NOMBRE IQM'], 'FORMACIÓN DE ESPUMA (SECUENCIA '+sec+')')
        self.assertNotEqual(mejor_coincidencia('SEC III, TENDENCIA/ESTABILIDAD', '', maestro[:1], 'ESPANOL')['estado'], 'ALINEADO')

    def test_compuesto_astm_nmx_exige_misma_prueba_unidades_y_apartado(self):
        from alinear_especificaciones_json import MaestroPreparado, metodo_alineado_canonico
        base = self.registro('VISCOSIDAD CINEMÁTICA A X °C', 'KINEMATIC VISCOSITY AT X °C', 'ASTM D445 / NMX-J-123-ANCE-2019 P 6,7', 'mm²/s')
        otro = {**base, 'METODO': 'ASTM D7042'}
        original = 'ASTM D445 - 19a / D7042 - 14 / NMX-J-123-ANCE-2019 P 6,7'
        for maestro in ([base, otro], MaestroPreparado([base, otro])):
            r = mejor_coincidencia('VISCOSIDAD CINEMÁTICA A 25 °C', original, maestro, 'ESPANOL', 'cSt')
            self.assertEqual(r['estado'], 'ALINEADO')
            self.assertEqual(r['registro']['NOMBRE IQM'], base['NOMBRE IQM'])
            self.assertEqual(metodo_alineado_canonico(original, r['registro']['METODO']), 'ASTM D445 / ASTM D7042 / NMX-J-123-ANCE-2019 P 6,7')
            for invalido in (original.replace('6,7', '6,5'), original.replace('7042', '9999')):
                self.assertTrue(mejor_coincidencia('VISCOSIDAD CINEMÁTICA A 25 °C', invalido, maestro, 'ESPANOL')['metodo_no_encontrado'])
        for diferencia in ({'NOMBRE IQM': 'OTRA PRUEBA'}, {'UNIDADES': 'OTRA UNIDAD'}):
            r = mejor_coincidencia('VISCOSIDAD CINEMÁTICA A 25 °C', original, [base, {**otro, **diferencia}], 'ESPANOL')
            self.assertTrue(r['metodo_no_encontrado'])


class NormasAlineadorV2Tests(unittest.TestCase):
    def test_calificadores_de_vigencia_no_impiden_alinear_farmacopeas(self):
        from alinear_especificaciones_json import MaestroPreparado, metodo_alineado_canonico
        for base, canonico in (("Ph. Eur", "Ph Eur"), ("USP 32", "USP"), ("FEUM 13", "FEUM"),
                               ("USP 32 / FEUM 13", "USP / FEUM")):
            registro = {**MAESTRO, "METODO": canonico}
            for palabra in ("VIGENTE", "CURRENT", "CURRENCY"):
                for metodo in (base + " " + palabra, palabra.lower() + " " + base):
                    for idioma, nombre in (("ESPANOL", MAESTRO["NOMBRE IQM"]), ("INGLES", MAESTRO["NOMBRE IQM INGLES"])):
                        for maestro in ([registro], MaestroPreparado([registro])):
                            with self.subTest(metodo=metodo, idioma=idioma, preparado=isinstance(maestro, MaestroPreparado)):
                                match = mejor_coincidencia(nombre, metodo, maestro, idioma)
                                self.assertEqual(match["estado"], "ALINEADO")
                                self.assertEqual(match["puntaje_metodo"], 1.0)
                                self.assertEqual(metodo_alineado_canonico(metodo, registro["METODO"]), canonico)

    def test_limpia_vigencia_sin_sustituir_una_prueba_sin_coincidencia(self):
        from alinear_especificaciones_json import MaestroPreparado
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        from interfaz_seleccion_especificaciones import preparar_vista_previa_alineada
        from generador_hii import preparar_datos
        registros = [{**MAESTRO, "METODO": "Ph Eur", "NOMBRE IQM": "Solubilidad", "NOMBRE IQM INGLES": "Solubility"},
                     {**MAESTRO, "METODO": "USP / FEUM", "NOMBRE IQM": "DENSIDAD RELATIVA A 60 °C", "NOMBRE IQM INGLES": "RELATIVE DENSITY AT 60 °C"}]
        for idioma, nombre in (("es", "DENSIDAD RELATIVA A 60 °C"), ("en", "RELATIVE DENSITY AT 60 °C")):
            for palabra in ("VIGENTE", "CURRENT", "CURRENCY"):
                for maestro in (registros, MaestroPreparado(registros)):
                    with self.subTest(idioma=idioma, palabra=palabra, preparado=isinstance(maestro, MaestroPreparado)), tempfile.TemporaryDirectory() as carpeta:
                        metodo = "Ph. Eur " + palabra
                        fila = {"prueba": nombre, "metodo": metodo, "unidades": "A/D", "especificacion": "0,8400 - 0,8800", "valor_tipico": "0,843"}
                        origen, destino = Path(carpeta) / "original.json", Path(carpeta) / "alineado.json"
                        origen.write_text(json.dumps({"producto": {"nombre": "Producto"}, "documento": {"idioma": idioma},
                            "control_documental": {"codigo_unico_producto": "P1", "revision": "00"}, "pruebas": [fila]}), encoding="utf-8")
                        for entrada in (origen, destino):
                            conteo = procesar_json(entrada, destino, maestro)
                            self.assertEqual(conteo["alineados"], 0)
                            datos = json.loads(destino.read_text(encoding="utf-8"))
                            detalle = datos["alineacion_metodos_lab"]["detalle"][0]
                            self.assertEqual(detalle["metodo_original"], metodo)
                            self.assertEqual(detalle["valores_alineados"]["metodo"], "Ph Eur")
                            self.assertEqual(datos["pruebas"][0]["prueba"], nombre)
                            self.assertEqual(datos["pruebas"][0]["unidades"], "DIMENSIONLESS" if idioma == "en" else "ADIMENSIONAL")
                            self.assertEqual(preparar_vista_previa_alineada(destino).pruebas[0]["metodo"], "Ph Eur")
                            self.assertEqual(preparar_datos(destino).pruebas[0]["metodo"], "Ph Eur")
                        errores = {}
                        filas, conteo = alinear_filas_generador([{**fila, "metodo": metodo + " ¹"}], maestro, idioma, errores_metodo=errores)
                        self.assertEqual(conteo["alineados"], 0)
                        self.assertEqual(filas[0]["metodo"], "Ph Eur ¹")
                        self.assertEqual(filas[0]["prueba"], nombre)
                        self.assertEqual(errores, {})

    def test_vigencia_no_autoriza_farmacopeas_ausentes_ni_normas_distintas(self):
        from alinear_especificaciones_json import MaestroPreparado
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        registro = {**MAESTRO, "METODO": "USP / FEUM"}
        for metodo in ("Ph Eur VIGENTE", "USP CURRENT / NOM-092-SSA1-1994", "CURRENCY"):
            for maestro in ([registro], MaestroPreparado([registro])):
                with self.subTest(metodo=metodo, preparado=isinstance(maestro, MaestroPreparado)):
                    match = mejor_coincidencia(MAESTRO["NOMBRE IQM"], metodo, maestro, "ESPANOL")
                    self.assertNotEqual(match["estado"], "ALINEADO")
                    self.assertNotIn("metodo_normalizado", match)
                    fila = {"prueba": MAESTRO["NOMBRE IQM"], "metodo": metodo}
                    filas, _ = alinear_filas_generador([fila], maestro, "es")
                    self.assertEqual(filas[0]["metodo"], metodo)

    def test_astm_alternativos_de_la_misma_prueba_se_alinean_sin_perder_el_compuesto(self):
        from alinear_especificaciones_json import MaestroPreparado
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        metodo = "ASTM D445/ ASTM D7042"
        registros = [dict(MAESTRO), {**MAESTRO, "METODO": "ASTM D7042"},
                     {**MAESTRO, "METODO": "ASTM D445 / NMX-J-123-ANCE-2019 P 6,7"}]
        for idioma, nombre in (("ESPANOL", MAESTRO["NOMBRE IQM"]), ("INGLES", MAESTRO["NOMBRE IQM INGLES"])):
            for maestro in (registros, MaestroPreparado(registros)):
                with self.subTest(idioma=idioma, preparado=isinstance(maestro, MaestroPreparado)), tempfile.TemporaryDirectory() as carpeta:
                    match = mejor_coincidencia(nombre, metodo, maestro, idioma)
                    self.assertEqual(match["estado"], "ALINEADO")
                    self.assertFalse(match["ambiguo"])
                    fila = {"prueba": nombre, "metodo": metodo, "unidades": "cSt", "especificacion": "11 - 18", "valor_tipico": "15"}
                    origen, destino = Path(carpeta) / "original.json", Path(carpeta) / "alineado.json"
                    origen.write_text(json.dumps({"documento": {"idioma": idioma}, "pruebas": [fila]}), encoding="utf-8")
                    conteo = procesar_json(origen, destino, maestro)
                    self.assertEqual(conteo["alineados"], 1)
                    salida = json.loads(destino.read_text(encoding="utf-8"))
                    self.assertEqual(salida["pruebas"][0]["metodo"], "ASTM D445 / ASTM D7042")
                    self.assertFalse(salida["alineacion_metodos_lab"]["detalle"][0].get("metodo_no_encontrado"))
                    errores = {}
                    filas, conteo_editor = alinear_filas_generador([fila], maestro, "en" if idioma == "INGLES" else "es", errores_metodo=errores)
                    self.assertEqual(filas[0]["metodo"], "ASTM D445 / ASTM D7042")
                    self.assertEqual(conteo_editor["alineados"], 1)
                    self.assertEqual(errores, {})

    def test_no_combina_astm_de_pruebas_o_unidades_diferentes(self):
        from alinear_especificaciones_json import MaestroPreparado
        for diferencia in ({"NOMBRE IQM": "Otra prueba"}, {"UNIDADES": "Otra unidad"}):
            registros = [dict(MAESTRO), {**MAESTRO, "METODO": "ASTM D7042", **diferencia}]
            for maestro in (registros, MaestroPreparado(registros)):
                with self.subTest(diferencia=diferencia, preparado=isinstance(maestro, MaestroPreparado)):
                    match = mejor_coincidencia(MAESTRO["NOMBRE IQM"], "ASTM D445 / ASTM D7042", maestro, "ESPANOL")
                    self.assertEqual(match["estado"], "REVISION_REQUERIDA")
                    self.assertTrue(match["metodo_no_encontrado"])

    def test_identifica_normas_sin_confundir_familias_numeros_o_ediciones(self):
        from alinear_especificaciones_json import referencias_normativas
        self.assertEqual(referencias_normativas("NOM-092-SSA1 1994"), referencias_normativas("NOM 092 SSA1-2020"))
        self.assertNotEqual(referencias_normativas("NOM-092-SSA1"), referencias_normativas("NOM-093-SSA1"))
        self.assertEqual(referencias_normativas("NOM-115-SSA1 1994 / NOM-111-SSA1-1994"),
                         {"NOM-115-SSA1", "NOM-111-SSA1"})

    def test_no_propone_farmacopeas_para_nom_ausentes_y_conserva_fila_en_todas_las_salidas(self):
        from alinear_especificaciones_json import MaestroPreparado
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        from interfaz_seleccion_especificaciones import preparar_vista_previa_alineada
        from cambios_alineacion import cambios_documento, CAMPOS
        from generador_hii import preparar_datos
        maestro = {**MAESTRO, "NOMBRE IQM": "NO PATÓGENOS B", "NOMBRE IQM INGLES": "NO PATHOGENS B", "METODO": "USP / FEUM"}
        for idioma, nombre in (("es", "NO PATÓGENOS B"), ("en", "NO PATHOGENS B")):
            for metodo in ("NOM-092-SSA1 1994", "NOM-115-SSA1 1994 / NOM-111-SSA1 1994"):
                with self.subTest(idioma=idioma, metodo=metodo), tempfile.TemporaryDirectory() as carpeta:
                    fila = {"prueba": nombre, "metodo": metodo, "unidades": "UFC", "especificacion": "< 1000", "valor_tipico": ""}
                    for lista in ([maestro], MaestroPreparado([maestro])):
                        match = mejor_coincidencia(nombre, metodo, lista, "INGLES" if idioma == "en" else "ESPANOL")
                        self.assertEqual(match["estado"], "SIN_COINCIDENCIA")
                        self.assertTrue(match["metodo_no_encontrado"])
                        self.assertEqual(match["registro"], {})
                    origen, destino = Path(carpeta) / "original.json", Path(carpeta) / "alineado.json"
                    origen.write_text(json.dumps({"producto": {"nombre": "Producto"}, "documento": {"idioma": idioma},
                        "control_documental": {"codigo_unico_producto": "P1", "revision": "00"}, "pruebas": [fila]}), encoding="utf-8")
                    for entrada in (origen, destino):
                        procesar_json(entrada, destino, [maestro])
                        datos = json.loads(destino.read_text(encoding="utf-8"))
                        detalle = datos["alineacion_metodos_lab"]["detalle"][0]
                        self.assertEqual(datos["alineacion_metodos_lab"]["version_programa"], VERSION)
                        self.assertEqual(detalle["valores_alineados"], fila)
                        previa = preparar_vista_previa_alineada(destino)
                        self.assertEqual(previa.pruebas[0], fila)
                        self.assertEqual(preparar_datos(destino).pruebas[0]["metodo"], metodo)
                        marcas = cambios_documento(datos, previa.pruebas)
                        self.assertEqual(marcas[0], set(CAMPOS))
                        self.assertIn("Método no encontrado en lista maestra", marcas[0].mensaje_metodo)
                    errores = {}
                    resultado, conteo = alinear_filas_generador([fila], [maestro], idioma, errores_metodo=errores)
                    self.assertEqual(resultado, [fila])
                    self.assertEqual(conteo["sin_coincidencia"], 1)
                    self.assertIn(0, errores)

    def test_candidato_debe_contener_todas_las_normas_en_un_solo_registro(self):
        maestro = {**MAESTRO, "NOMBRE IQM": "NO PATOGENOS", "NOMBRE IQM INGLES": "NO PATHOGENS", "METODO": "NOM-092-SSA1-1994"}
        otro = {**maestro, "METODO": "NOM-111-SSA1-1994"}
        parcial = mejor_coincidencia("NO PATOGENOS", "NOM-092-SSA1 / NOM-111-SSA1", [maestro, otro], "ESPANOL")
        self.assertEqual(parcial["estado"], "REVISION_REQUERIDA")
        self.assertTrue(parcial["metodo_no_encontrado"])
        completo = {**maestro, "METODO": "NOM-092-SSA1 / NOM-111-SSA1"}
        resultado = mejor_coincidencia("NO PATOGENOS", "NOM 092 SSA1 1994 / NOM 111 SSA1 1994", [completo], "ESPANOL")
        self.assertEqual(resultado["estado"], "ALINEADO")

    def test_realinear_json_antiguo_recupera_norma_y_nombre_de_la_auditoria(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "viejo.json"
            sustituidos = {"prueba": "ACIDOS ORGANICOS", "metodo": "USP / FEUM", "unidades": "NA", "especificacion": "< 100", "valor_tipico": ""}
            ruta.write_text(json.dumps({"documento": {"idioma": "es"}, "pruebas": [sustituidos],
                "alineacion_metodos_lab": {"version_programa": "1.13.6", "detalle": [{
                    "nombre_original": "NO PATOGENOS B", "metodo_original": "NOM-115-SSA1 1994", "unidades_originales": "UFC",
                    "especificacion_original": "< 100", "valores_alineados": sustituidos}]}}), encoding="utf-8")
            procesar_json(ruta, ruta, [{**MAESTRO, "METODO": "USP / FEUM"}])
            fila = json.loads(ruta.read_text(encoding="utf-8"))["pruebas"][0]
            self.assertEqual(fila["prueba"], "NO PATOGENOS B")
            self.assertEqual(fila["metodo"], "NOM-115-SSA1 1994")
            self.assertEqual(fila["unidades"], "UFC")


class FormatoBrookfieldTests(unittest.TestCase):
    def test_detecta_nombre_aproximado_o_d2983_sin_confundir_otros_metodos(self):
        for nombre in ("BROOKFIELD", "Viscosidad Brookfild", "BROOKFEILD", "BROKFIELD", "Brook field", "BROOK-FIELD"):
            with self.subTest(nombre=nombre):
                self.assertTrue(es_prueba_brookfield(nombre, ""))
        for metodo in ("ASTM D2983", "ASTM D 2983", "ASTM D-2983", "ASTM 2983", "ATSM D2983", "ASMTD2983", "ASTM D 2 983"):
            with self.subTest(metodo=metodo):
                self.assertTrue(es_prueba_brookfield("VISCOSIDAD", metodo))
        self.assertTrue(es_prueba_brookfield("ASTM D2983", ""))
        for nombre, metodo in (("COLOR", "ASTM D1500"), ("VISCOSIDAD", "ASTM D2988"),
                               ("VISCOSIDAD", "ASTM D29830"), ("VISCOSIDAD CINEMÁTICA", "ASTM D445")):
            with self.subTest(nombre=nombre, metodo=metodo):
                self.assertFalse(es_prueba_brookfield(nombre, metodo))

    def test_coma_solo_para_miles_y_grupos_incorrectos_requieren_revision(self):
        casos = (
            ("3,610", "3 610", []), ("3,610.25", "3 610,25", []),
            ("1,234,567.890", "1 234 567,890", []), ("3.610", "3,610", []),
            ("3610", "3 610", []), ("0,130", "130", []),
            ("3,610 - 4,500", "3 610 - 4 500", []), ("-3,610", "− 3 610", []),
            ("<3,610.25 mPa.s", "< 3 610,25 mPa·s", []),
            ("3,610e3", "3 610 × 10³", []), ("3.610e3", "3,610 × 10³", []),
            ("ASTM D2983", "ASTM D2983", []),
            ("3,61", "3,61", ["3,61"]), ("3.610,00", "3.610,00", ["3.610,00"]),
            ("36,10.25", "36,10.25", ["36,10.25"]),
        )
        for entrada, esperado, ambiguos in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(normalizar_valor_medicion(entrada, coma_de_miles=True), (esperado, ambiguos))

    def test_excepcion_en_ambos_campos_y_realineacion_conserva_su_significado(self):
        maestro = {**MAESTRO, "METODO": "ASTM D2983", "NOMBRE IQM": "VISCOSIDAD BROOKFIELD",
                   "NOMBRE IQM INGLES": "BROOKFIELD VISCOSITY", "UNIDADES": "mPa*s", "UNITS": "mPa*s"}
        for valor, esperado, revision in (("3,610", "3 610", False), ("3,610.25", "3 610,25", False),
                                         ("3.610", "3,610", False), ("3,61", "3,61", True)):
            with self.subTest(valor=valor), tempfile.TemporaryDirectory() as temporal:
                origen, destino = Path(temporal) / "entrada.json", Path(temporal) / "alineada.json"
                origen.write_text(json.dumps({"documento": {"idioma": "es"}, "pruebas": [{
                    "prueba": "VISCOSIDAD BROOKFIELD", "metodo": "ASTM D2983",
                    "especificacion": valor, "valor_tipico": valor,
                }]}), encoding="utf-8")
                original = origen.read_bytes()
                for entrada in (origen, destino, destino):
                    conteo = procesar_json(entrada, destino, [maestro])
                    resultado = json.loads(destino.read_text(encoding="utf-8"))
                    prueba = resultado["pruebas"][0]
                    detalle = resultado["alineacion_metodos_lab"]["detalle"][0]
                    self.assertEqual(prueba["especificacion"], esperado)
                    self.assertEqual(prueba["valor_tipico"], esperado)
                    self.assertEqual(detalle["especificacion_original"], valor)
                    self.assertEqual(detalle["valor_tipico_original"], valor)
                    self.assertEqual(detalle["regla_formato_numerico"], "COMA_DE_MILES_BROOKFIELD_D2983")
                    self.assertEqual(bool(detalle["observaciones_limpieza"]), revision)
                    self.assertEqual(conteo["revision"], int(revision))
                self.assertEqual(origen.read_bytes(), original)

    def test_otra_prueba_conserva_la_regla_general_en_el_mismo_documento(self):
        with tempfile.TemporaryDirectory() as temporal:
            origen, destino = Path(temporal) / "entrada.json", Path(temporal) / "alineada.json"
            origen.write_text(json.dumps({"documento": {"idioma": "es"}, "pruebas": [
                {"prueba": "Viscosidad Brookfeild", "metodo": "METODO INTERNO", "especificacion": "3,610", "valor_tipico": "3,610"},
                {"prueba": "Viscosidad cinemática", "metodo": "ASTM D445", "especificacion": "3,610", "valor_tipico": "0,130"},
            ]}), encoding="utf-8")
            procesar_json(origen, destino, [MAESTRO, {**MAESTRO, "METODO": "METODO INTERNO",
                         "NOMBRE IQM": "Viscosidad Brookfield", "NOMBRE IQM INGLES": "Brookfield viscosity",
                         "UNIDADES": "mPa.s", "UNITS": "mPa.s"}])
            detalle = json.loads(destino.read_text(encoding="utf-8"))["alineacion_metodos_lab"]["detalle"]
            self.assertEqual(detalle[0]["valores_alineados"]["especificacion"], "3 610")
            self.assertEqual(detalle[0]["valores_alineados"]["valor_tipico"], "3 610")
            self.assertEqual(detalle[1]["valores_alineados"]["especificacion"], "3,610")
            self.assertEqual(detalle[1]["valores_alineados"]["valor_tipico"], "0,130")
            self.assertEqual(detalle[1]["regla_formato_numerico"], "GENERAL")
            self.assertEqual(detalle[1]["observaciones_limpieza"], ["Separador numérico ambiguo en especificacion: 3,610"])


class FormatoNumericoMexicoTests(unittest.TestCase):
    def test_agrupa_decimales_y_completa_precision_de_rangos(self):
        casos = (
            ("0,8750", "0,875\u202f0"),
            ("0,1234567", "0,123\u202f456\u202f7"),
            ("0,818 - 0,8750", "0,818\u202f0 - 0,875\u202f0"),
            ("7,00 - 8,000", "7,000 - 8,000"),
            ("1,2 - 3,45", "1,20 - 3,45"),
            ("10,000 - 12,5", "10,000 - 12,500"),
            ("7 - 8,0000", "7,000\u202f0 - 8,000\u202f0"),
            ("50 - 70", "50 - 70"),
            ("0.818–0.8750", "0,818\u202f0 - 0,875\u202f0"),
            ("0,000001", "0,000\u202f001"),
            ("0,123 456 7", "0,123\u202f456\u202f7"),
        )
        for original, esperado in casos:
            with self.subTest(original=original):
                self.assertEqual(normalizar_valor_medicion(original), (esperado, []))
                self.assertEqual(normalizar_formato_numerico(original), esperado)
                self.assertEqual(normalizar_formato_numerico(esperado), esperado)
                self.assertEqual(normalizar_valor_medicion(esperado)[0], esperado)
        self.assertEqual(normalizar_valor_medicion("12345,1234"), ("12 345,123\u202f4", []))
        self.assertEqual(normalizar_valor_medicion("7,000 - 8,000"),
                         ("7,000 - 8,000", ["7,000", "8,000"]))
        self.assertEqual(normalizar_valor_medicion("1.23456789e-4"),
                         ("1,234\u202f567\u202f89 × 10⁻⁴", []))

    def test_precision_decimal_se_conserva_en_json_editor_y_realineacion(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        for idioma, nombre in (("es", "Viscosidad cinemática"), ("en", "Kinematic viscosity")):
            for entrada, esperado in (("0,818 - 0,8750", "0,818\u202f0 - 0,875\u202f0"),
                                      ("7,00 - 8,000", "7,000 - 8,000")):
                with self.subTest(idioma=idioma, entrada=entrada), tempfile.TemporaryDirectory() as carpeta:
                    origen, destino = Path(carpeta) / "origen.json", Path(carpeta) / "alineado.json"
                    fila = {"prueba": nombre, "metodo": "ASTM D445", "especificacion": entrada,
                            "valor_tipico": "0,8350"}
                    origen.write_text(json.dumps({"documento": {"idioma": idioma}, "pruebas": [fila]}), encoding="utf-8")
                    for ruta in (origen, destino, destino):
                        resultado = procesar_json(ruta, destino, [MAESTRO])
                        datos = json.loads(destino.read_text(encoding="utf-8"))
                        self.assertEqual(resultado["alineados"], 1)
                        self.assertEqual(datos["pruebas"][0]["especificacion"], esperado)
                        self.assertEqual(datos["pruebas"][0]["valor_tipico"], "0,835\u202f0")
                        self.assertEqual(datos["alineacion_metodos_lab"]["detalle"][0]["especificacion_original"], entrada)
                    filas, resumen = alinear_filas_generador([dict(fila)], [MAESTRO], idioma)
                    self.assertEqual(resumen["alineados"], 1)
                    self.assertEqual(filas[0]["especificacion"], esperado)
                    self.assertEqual(filas[0]["valor_tipico"], "0,835\u202f0")

    def test_punto_decimal_miles_y_coma_ambigua_conservan_precision(self):
        casos = (
            ("7.500", "7,500", []),
            ("7,500.00", "7 500,00", []),
            ("7.500,00", "7 500,00", []),
            ("1,234,567.890", "1 234 567,890", []),
            ("1.234.567,890", "1 234 567,890", []),
            ("0.130 - 0.200", "0,130 - 0,200", []),
            ("0,130 - 0,200", "0,130 - 0,200", []),
            ("-0.130", "− 0,130", []),
            ("-0,130", "− 0,130", []),
            ("0.000", "0,000", []),
            ("0,000", "0,000", []),
            ("7,50", "7,50", []),
            ("7,500", "7,500", ["7,500"]),
            ("1,234", "1,234", ["1,234"]),
            ("7 500,123", "7 500,123", ["7 500,123"]),
            ("<0.130 mg/L", "< 0,130 mg/L", []),
        )
        for entrada, esperado, ambiguos in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(normalizar_valor_medicion(entrada), (esperado, ambiguos))

    def test_normalizacion_respetada_al_realinear_y_ediciones_se_vuelven_a_validar(self):
        for idioma in ("es", "en"):
            for valor, esperado, requiere_revision in (("7.500", "7,500", False),
                                                      ("7,500.00", "7 500,00", False),
                                                      ("0.130 - 0.200", "0,130 - 0,200", False),
                                                      ("7,500", "7,500", True)):
                with self.subTest(idioma=idioma, valor=valor), tempfile.TemporaryDirectory() as temporal:
                    origen, destino = Path(temporal) / "original.json", Path(temporal) / "alineada.json"
                    nombre = "Kinematic viscosity" if idioma == "en" else "Viscosidad cinemática"
                    origen.write_text(json.dumps({"documento": {"idioma": idioma}, "pruebas": [{
                        "prueba": nombre, "metodo": "ASTM D445", "especificacion": valor, "valor_tipico": valor,
                    }]}), encoding="utf-8")
                    original = origen.read_bytes()
                    for entrada in (origen, destino, destino):
                        conteo = procesar_json(entrada, destino, [MAESTRO])
                        resultado = json.loads(destino.read_text(encoding="utf-8"))
                        detalle = resultado["alineacion_metodos_lab"]["detalle"][0]
                        self.assertEqual(resultado["pruebas"][0]["especificacion"], esperado)
                        self.assertEqual(resultado["pruebas"][0]["valor_tipico"], esperado)
                        self.assertEqual(detalle["especificacion_original"], valor)
                        self.assertEqual(detalle["valor_tipico_original"], valor)
                        self.assertEqual(bool(detalle["observaciones_limpieza"]), requiere_revision)
                        self.assertEqual(conteo["revision"], int(requiere_revision))
                    self.assertEqual(origen.read_bytes(), original)
                    resultado["pruebas"][0]["especificacion"] = "9,500"
                    destino.write_text(json.dumps(resultado), encoding="utf-8")
                    procesar_json(destino, destino, [MAESTRO])
                    editada = json.loads(destino.read_text(encoding="utf-8"))["alineacion_metodos_lab"]["detalle"][0]
                    self.assertEqual(editada["especificacion_original"], "9,500")
                    self.assertIn("Separador numérico ambiguo en especificacion: 9,500", editada["observaciones_limpieza"])


class EncabezadosAlineacionTests(unittest.TestCase):
    def test_omite_encabezados_antes_de_buscar_coincidencias_y_guarda_auditoria(self):
        encabezados = (
            {"prueba": "PRUEBA", "metodo": "METODO", "especificacion": "ESPECIFICACION"},
            {"prueba": "TEST", "metodo": "TEST METHOD", "especificacion": "SPECIFICATION"},
            {"prueba": "PRUEBA / MÉTODO", "especificacion": "ESPECIFICACIÓN"},
            {"prueba": "  prueba  ", "metodo": "método", "especificacion": "especificación"},
            {"nombre": "PRUEBA", "metodo": "METODO", "especificacion": "ESPECIFICACION"},
        )
        for encabezado in encabezados:
            with self.subTest(encabezado=encabezado), tempfile.TemporaryDirectory() as temporal:
                origen, destino = Path(temporal) / "original.json", Path(temporal) / "alineada.json"
                fila_encabezado = {"numero": 1, "fila_origen": 2, **encabezado}
                datos = {"documento": {"idioma": "es"}, "pruebas": [fila_encabezado, {
                    "numero": 2, "fila_origen": 3, "prueba": "Viscosidad cinemática",
                    "metodo": "ASTM D445", "especificacion": "10 MIN", "unidades": "cSt",
                }]}
                origen.write_text(json.dumps(datos), encoding="utf-8")
                original = origen.read_bytes()
                with patch("alinear_especificaciones_json.mejor_coincidencia", wraps=mejor_coincidencia) as buscar:
                    conteo = procesar_json(origen, destino, [MAESTRO])
                    buscar.assert_called_once()
                    self.assertEqual(buscar.call_args.args[0], "Viscosidad cinemática")
                resultado = json.loads(destino.read_text(encoding="utf-8"))
                auditoria = resultado["alineacion_metodos_lab"]
                self.assertEqual(sum(conteo.values()), 1)
                self.assertEqual(len(resultado["pruebas"]), 1)
                self.assertEqual(resultado["pruebas"][0]["numero"], 1)
                self.assertEqual(resultado["pruebas"][0]["fila_origen"], 3)
                self.assertEqual(auditoria["detalle"][0]["numero_prueba"], 1)
                self.assertEqual(auditoria["encabezados_omitidos"], [{
                    "indice_original": 1, "motivo": "Encabezado de tabla, no es una prueba.",
                    "fila_original": fila_encabezado,
                }])
                self.assertEqual(origen.read_bytes(), original)

    def test_encabezados_repetidos_se_omiten_y_la_auditoria_sobrevive_realineacion(self):
        with tempfile.TemporaryDirectory() as temporal:
            origen, destino = Path(temporal) / "original.json", Path(temporal) / "alineada.json"
            cabecera = {"prueba": "PRUEBA", "metodo": "METODO", "especificacion": "ESPECIFICACION"}
            prueba = {"prueba": "Viscosidad cinemática", "metodo": "ASTM D445", "especificacion": "10 MIN"}
            filas = [{"numero": i, "fila_origen": i + 1, **fila}
                     for i, fila in enumerate((cabecera, prueba, cabecera, prueba, cabecera), 1)]
            origen.write_text(json.dumps({"documento": {"idioma": "es"}, "pruebas": filas}), encoding="utf-8")
            procesar_json(origen, destino, [MAESTRO])
            primera = json.loads(destino.read_text(encoding="utf-8"))
            self.assertEqual([p["numero"] for p in primera["pruebas"]], [1, 2])
            self.assertEqual([p["fila_origen"] for p in primera["pruebas"]], [3, 5])
            omitidas = primera["alineacion_metodos_lab"]["encabezados_omitidos"]
            self.assertEqual([fila["indice_original"] for fila in omitidas], [1, 3, 5])
            procesar_json(destino, destino, [MAESTRO])
            segunda = json.loads(destino.read_text(encoding="utf-8"))
            self.assertEqual(segunda["alineacion_metodos_lab"]["encabezados_omitidos"], omitidas)
            self.assertEqual(len(segunda["pruebas"]), 2)
            self.assertEqual(len(segunda["alineacion_metodos_lab"]["detalle"]), 2)

    def test_no_omite_pruebas_reales_que_contienen_palabras_de_encabezado(self):
        with tempfile.TemporaryDirectory() as temporal:
            origen, destino = Path(temporal) / "original.json", Path(temporal) / "alineada.json"
            filas = [
                {"prueba": "PRUEBA", "metodo": "ASTM D445", "especificacion": "10 MIN"},
                {"prueba": "PRUEBA", "metodo": "METODO", "especificacion": "CUMPLE"},
                {"prueba": "PRUEBA DE LODOS Y BARNICES", "metodo": "MÉTODO INTERNO", "especificacion": "CUMPLE"},
            ]
            origen.write_text(json.dumps({"pruebas": filas}), encoding="utf-8")
            conteo = procesar_json(origen, destino, [MAESTRO])
            resultado = json.loads(destino.read_text(encoding="utf-8"))
            self.assertEqual(len(resultado["pruebas"]), 3)
            self.assertEqual(sum(conteo.values()), 3)
            self.assertEqual(resultado["alineacion_metodos_lab"]["encabezados_omitidos"], [])

    def test_archivo_con_solo_encabezado_no_inventa_pruebas(self):
        with tempfile.TemporaryDirectory() as temporal:
            origen, destino = Path(temporal) / "original.json", Path(temporal) / "alineada.json"
            origen.write_text(json.dumps({"pruebas": [{
                "prueba": "TEST", "metodo": "METHOD", "especificacion": "SPECIFICATIONS",
            }]}), encoding="utf-8")
            conteo = procesar_json(origen, destino, [MAESTRO])
            resultado = json.loads(destino.read_text(encoding="utf-8"))
            self.assertEqual(sum(conteo.values()), 0)
            self.assertEqual(resultado["pruebas"], [])
            self.assertEqual(resultado["alineacion_metodos_lab"]["detalle"], [])
            self.assertEqual(len(resultado["alineacion_metodos_lab"]["encabezados_omitidos"]), 1)


class AlineacionEspecificacionesTests(unittest.TestCase):
    def test_periodicidad_vacia_usa_triangulo_y_conserva_valores_capturados(self):
        for campos, esperado in (({}, "▶"), ({"PERIODIDICAD DE ANALISIS": " "}, "▶"),
                                 ({"PERIODIDICAD DE ANALISIS": None}, "▶"),
                                 ({"PERIODIDICAD DE ANALISIS": "QIP"}, "QIP"),
                                 ({"periodicidad": "I"}, "I")):
            with self.subTest(campos=campos), tempfile.TemporaryDirectory() as temporal:
                origen = Path(temporal) / "entrada.json"
                destino = Path(temporal) / "salida.json"
                origen.write_text(json.dumps({"pruebas": [{"prueba": "Viscosidad cinemática", "metodo": "ASTM D445", **campos}]}), encoding="utf-8")
                procesar_json(origen, destino, [MAESTRO])
                self.assertEqual(json.loads(destino.read_text(encoding="utf-8"))["pruebas"][0]["PERIODIDICAD DE ANALISIS"], esperado)

    def test_ejecutar_acepta_carpetas_configuradas_de_entrada_y_salida(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal).resolve()
            entrada = base / "origen"
            salida = base / "alineadas"
            archivo = entrada / "producto.json"
            entrada.mkdir()
            archivo.write_text("{}", encoding="utf-8")
            (base / "METODOS_LAB").mkdir()
            (base / "METODOS_LAB" / "METODOS_LAB_UNIFICADO.xlsx").write_bytes(b"maestro simulado")

            with (
                patch(
                    "alinear_especificaciones_json.cargar_maestro",
                    return_value=[MAESTRO],
                ),
                patch(
                    "alinear_especificaciones_json.procesar_json",
                    return_value={
                        "alineados": 1,
                        "revision": 0,
                        "sin_coincidencia": 0,
                    },
                ) as procesar_json_mock,
            ):
                procesar_json_mock.side_effect = simular_json_para_cache
                resultado = ejecutar(
                    base,
                    archivos_seleccionados=[archivo],
                    carpeta_entrada=entrada,
                    carpeta_salida=salida,
                )

            procesar_json_mock.assert_called_once_with(
                archivo, salida / "producto.json", [MAESTRO]
            )
            self.assertEqual(resultado["archivos"], 1)
            self.assertTrue((salida / "REPORTE_ALINEACION.txt").is_file())

    def test_ejecutar_realinea_sobre_los_destinos_seleccionados(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal).resolve()
            entrada = base / "origen"
            salida = base / "alineadas"
            origen = entrada / "producto_original.json"
            destino = salida / "producto_alineado.json"
            entrada.mkdir()
            salida.mkdir()
            origen.write_text("{}", encoding="utf-8")
            (base / "METODOS_LAB").mkdir()
            (base / "METODOS_LAB" / "METODOS_LAB_UNIFICADO.xlsx").write_bytes(b"maestro simulado")

            with (
                patch(
                    "alinear_especificaciones_json.cargar_maestro",
                    return_value=[MAESTRO],
                ),
                patch(
                    "alinear_especificaciones_json.procesar_json",
                    return_value={
                        "alineados": 1,
                        "revision": 0,
                        "sin_coincidencia": 0,
                    },
                ) as procesar_json_mock,
            ):
                procesar_json_mock.side_effect = simular_json_para_cache
                ejecutar(
                    base,
                    archivos_seleccionados=[origen],
                    carpeta_entrada=entrada,
                    carpeta_salida=salida,
                    destinos_seleccionados=[destino],
                )

            procesar_json_mock.assert_called_once_with(origen, destino, [MAESTRO])

    def test_usa_el_metodo_del_maestro_en_el_idioma_del_json(self):
        maestro_bilingue = {
            **MAESTRO,
            "METODO": "MÉTODO INTERNO",
            "MÉTODO INGLÉS": "INTERNAL METHOD",
            "NOMBRE IQM": "Método interno",
            "NOMBRE IQM INGLES": "Internal test",
        }

        with tempfile.TemporaryDirectory() as temporal:
            for idioma, nombre, metodo in (
                ("en", "Internal test", "INTERNAL METHOD"),
                ("es", "Método interno", "MÉTODO INTERNO"),
            ):
                with self.subTest(idioma=idioma):
                    origen = Path(temporal) / f"entrada_{idioma}.json"
                    destino = Path(temporal) / f"salida_{idioma}.json"
                    origen.write_text(
                        json.dumps(
                            {
                                "documento": {"idioma": idioma},
                                "pruebas": [
                                    {"prueba": nombre, "metodo": metodo}
                                ],
                            }
                        ),
                        encoding="utf-8",
                    )

                    procesar_json(origen, destino, [maestro_bilingue])
                    alineado = json.loads(destino.read_text(encoding="utf-8"))

                    self.assertEqual(alineado["pruebas"][0]["metodo"], metodo)
                    self.assertEqual(
                        alineado["alineacion_metodos_lab"]["detalle"][0][
                            "metodo_maestro_fuente"
                        ],
                        metodo,
                    )

        self.assertEqual(
            metodo_maestro_para_idioma(maestro_bilingue, "ENGLISH"),
            "INTERNAL METHOD",
        )
        self.assertEqual(
            metodos_maestros(maestro_bilingue),
            ("MÉTODO INTERNO", "INTERNAL METHOD"),
        )

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
            ("0,095 7", "0,095\u202f7"),
            ("cm-1", "cm-1"),
        )
        for entrada, esperado in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(normalizar_formato_numerico(entrada), esperado)

    def test_preserva_decimales_ambiguos_como_decimales(self):
        self.assertEqual(normalizar_formato_numerico("3,351"), "3,351")
        self.assertEqual(normalizar_formato_numerico("7.500"), "7,500")

    def test_normaliza_variaciones_de_minimo_y_maximo_en_especificaciones(self):
        casos = (
            ("30 mínimo", "30 MÍNIMO"),
            ("30 minimo", "30 MÍNIMO"),
            ("30 MINIMO", "30 MÍNIMO"),
            ("30 MÍNIMO", "30 MÍNIMO"),
            ("30 MIN", "30 MÍNIMO"),
            ("30 MÍN.", "30 MÍNIMO"),
            ("30 min", "30 min"),
            ("30 min de agitación", "30 min de agitación"),
            ("30 max", "30 MÁXIMO"),
            ("30 MAXIMO", "30 MÁXIMO"),
            ("30 máximo", "30 MÁXIMO"),
            ("30 Máx.", "30 MÁXIMO"),
            ("30 MAX.", "30 MÁXIMO"),
        )
        for entrada, esperado in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(normalizar_especificacion(entrada), esperado)

    def test_normaliza_minimo_maximo_sin_confundirlos_con_metros(self):
        casos = (
            ("149 mínimo", "149 MÍNIMO", []),
            ("− 45 máximo", "− 45 MÁXIMO", []),
            ("10 MIN", "10 MÍNIMO", []),
            ("10 min", "10 min", []),
            ("1.000 mínimo", "1,000 MÍNIMO", []),
        )
        for entrada, esperado, ambiguos_esperados in casos:
            with self.subTest(entrada=entrada):
                especificacion = normalizar_especificacion(entrada)
                valor, ambiguos = normalizar_valor_medicion(especificacion)
                self.assertEqual(valor, esperado)
                self.assertEqual(ambiguos, ambiguos_esperados)

    def test_normaliza_terminos_frecuentes_en_especificacion_y_valor_tipico(self):
        casos = (
            ("Típico", "TÍPICO"),
            ("tipica", "TÍPICA"),
            ("typical", "TYPICAL"),
            ("Pasa", "PASA"),
            ("pass", "PASS"),
            ("reportar", "REPORTAR"),
            ("Report", "REPORT"),
            ("reported", "REPORTED"),
            ("negativo", "NEGATIVA"),
            ("NEGATIVA", "NEGATIVA"),
            ("negativos", "NEGATIVA"),
            ("NEGATIVAS", "NEGATIVA"),
            ("negative", "NEGATIVE"),
            ("NEGATIVE", "NEGATIVE"),
            ("negatives", "NEGATIVE"),
        )
        for entrada, esperado in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(normalizar_terminos_especificacion(entrada), esperado)

    def test_normaliza_valores_y_unidades_segun_reglas_de_limpieza(self):
        casos = (
            ("10000", "10 000", []),
            ("1234.56", "1 234,56", []),
            (".25", "0,25", []),
            ("25°C", "25 °C", []),
            ("100mg/L", "100 mg/L", []),
            ("10 kgs", "10 kg", []),
            ("kg/m3", "kg/m³", []),
            ("3.2e5", "3,2 × 10⁵", []),
            ("3.2e-5", "3,2 × 10⁻⁵", []),
            ("<0.05mg/L", "< 0,05 mg/L", []),
            ("10±2", "10 ± 2", []),
            ("10<20", "10 < 20", []),
            ("-0.25", "− 0,25", []),
            ("N.m", "N·m", []),
            ("Pa.s", "Pa·s", []),
            ("mPa.s", "mPa·s", []),
            ("MPa", "MPa", []),
            ("10 MPA", "10 MPa", []),
            ("10 M", "10 m", []),
            ("100 l, 10 secs", "100 L, 10 s", []),
            ("cm-1", "cm⁻¹", []),
            ("m2, m^3, s-1", "m², m³, s⁻¹", []),
            ("1.234.567,89", "1 234 567,89", []),
            ("1,234,567", "1 234 567", []),
            ("10 min", "10 min", []),
            ("10 kg.", "10 kg", []),
            ("N/A", "N/A", []),
            ("ND", "ND", []),
            ("2026-10-13", "2026-10-13", []),
            ("ASTM D445", "ASTM D445", []),
            ("ID: 10000", "ID: 10000", []),
            ("SN-1234", "SN-1234", []),
            ("1.2.3", "1.2.3", []),
            ("192.168.1.10", "192.168.1.10", []),
            ("7.500", "7,500", []),
        )
        for entrada, esperado, ambiguos in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(
                    normalizar_valor_medicion(entrada),
                    (esperado, ambiguos),
                )

    def test_coma_ambigua_se_conserva_y_marca_la_prueba_para_revision(self):
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
                                "metodo": "ASTM D445",
                                "especificacion": "7.500 mg/L",
                                "valor_tipico": "1,234",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )

            conteo = procesar_json(origen, destino, [MAESTRO])
            salida = json.loads(destino.read_text(encoding="utf-8"))
            prueba = salida["pruebas"][0]
            detalle = salida["alineacion_metodos_lab"]["detalle"][0]

            self.assertEqual(conteo["revision"], 1)
            self.assertEqual(conteo["alineados"], 0)
            self.assertEqual(prueba["especificacion"], "7,500 mg/L")
            self.assertEqual(prueba["valor_tipico"], "1,234")
            self.assertEqual(detalle["estado"], "REVISION_REQUERIDA")
            self.assertEqual(
                detalle["observaciones_limpieza"],
                [
                    "Separador numérico ambiguo en valor_tipico: 1,234",
                ],
            )

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

            self.assertEqual(conteo["alineados"], 1)
            self.assertEqual(conteo["revision"], 1)
            self.assertEqual(alineado["pruebas"][0]["metodo"], "ASTM D445 / ASTM D999")
            self.assertTrue(primer_detalle["metodo_no_encontrado"])
            self.assertEqual(
                primer_detalle["referencias_astm_no_encontradas"], ["D999"]
            )
            self.assertEqual(primer_detalle["mensaje_astm_no_encontradas"], mensaje)
            self.assertEqual(
                alineacion["mensajes_astm_no_encontradas"], [mensaje]
            )
            self.assertEqual(segundo_detalle["referencias_astm_no_encontradas"], [])
            self.assertEqual(segundo_detalle["mensaje_astm_no_encontradas"], "")

    def test_detalle_conserva_originales_si_no_hay_coincidencia(self):
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
                    "unidades": "NA",
                    "metodo": "VISUAL",
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
            self.assertEqual(alineada["valor_tipico"], "0,25 - 7 500,50")
            self.assertEqual(alineada["especificacion"], "40 °C / 0,6 mL")
            self.assertEqual(sin_coincidencia["unidades"], "5.0 mL")
            self.assertEqual(sin_coincidencia["valor_tipico"], "50-70")
            self.assertEqual(sin_coincidencia["nombre"], "Temperatura 40°C")
            self.assertEqual(detalle[0]["unidades_originales"], "mL NaOH 0.01 N")
            self.assertEqual(detalle[0]["valor_tipico_original"], "0.25-7500.50")
            self.assertEqual(
                detalle[0]["valores_alineados"],
                {
                    "prueba": "Viscosidad cinemática",
                    "unidades": "mL NaOH 0,01 N",
                    "metodo": "ASTM D445",
                    "especificacion": "40 °C / 0,6 mL",
                    "valor_tipico": "0,25 - 7 500,50",
                },
            )

    def test_unidades_maestras_separadas_por_comas_son_alternativas(self):
        maestro_unidades = {
            **MAESTRO,
            "UNIDADES": "mm²/s, g/cm³, kg/L",
            "UNITS": "mm²/s, g/ml, kg/L",
        }
        casos = (
            (
                "es",
                "Viscosidad cinemática",
                "g/cm³",
                "Pa",
                "g/cm³",
                "mm²/s",
            ),
            (
                "en",
                "Kinematic viscosity",
                "g/ml",
                "Pa",
                "g/ml",
                "mm²/s",
            ),
        )
        with tempfile.TemporaryDirectory() as temporal:
            for (
                idioma,
                nombre,
                unidad_valida,
                unidad_invalida,
                esperada_valida,
                esperada_invalida,
            ) in casos:
                with self.subTest(idioma=idioma):
                    origen = Path(temporal) / f"entrada_unidades_{idioma}.json"
                    destino = Path(temporal) / f"salida_unidades_{idioma}.json"
                    origen.write_text(
                        json.dumps(
                            {
                                "documento": {"idioma": idioma},
                                "pruebas": [
                                    {
                                        "prueba": nombre,
                                        "metodo": "ASTM D445",
                                        "unidades": unidad_valida,
                                    },
                                    {
                                        "prueba": nombre,
                                        "metodo": "ASTM D445",
                                        "unidades": unidad_invalida,
                                    },
                                ],
                            }
                        ),
                        encoding="utf-8",
                    )

                    conteo = procesar_json(origen, destino, [maestro_unidades])
                    salida = json.loads(destino.read_text(encoding="utf-8"))
                    detalle = salida["alineacion_metodos_lab"]["detalle"]

                    self.assertEqual(conteo["alineados"], 2)
                    self.assertEqual(
                        [prueba["unidades"] for prueba in salida["pruebas"]],
                        [esperada_valida, esperada_invalida],
                    )
                    self.assertEqual(
                        [
                            valores["valores_alineados"]["unidades"]
                            for valores in detalle
                        ],
                        [esperada_valida, esperada_invalida],
                    )

    def test_separa_notas_estandarizadas_de_las_pruebas_en_espanol_e_ingles(self):
        casos = (
            (
                "es",
                "Viscosidad cinemática",
                "ASTM D445",
                "Notas: La versión vigente del método aplicable "
                "corresponderá a la versión emitida proporcionada por el laboratorio.",
                True,
                "La versión vigente del método aplicable corresponderá a la "
                "versión emitida proporcionada por el laboratorio.",
            ),
            (
                "en",
                "Kinematic viscosity Notes: The active version of the applicable "
                "method shall be the version issued and provided by the laboratory.",
                "ASTM D445",
                "",
                False,
                "The active version of the applicable method shall be the "
                "version issued and provided by the laboratory.",
            ),
        )
        with tempfile.TemporaryDirectory() as temporal:
            for idioma, nombre, metodo, nota, fila_nota, nota_canonica in casos:
                with self.subTest(idioma=idioma):
                    origen = Path(temporal) / f"entrada_{idioma}.json"
                    destino = Path(temporal) / f"salida_{idioma}.json"
                    pruebas = [
                        {
                            "prueba": nombre,
                            "metodo": metodo,
                            "unidades": "cSt",
                        }
                    ]
                    notas = []
                    if fila_nota:
                        pruebas.append(
                            {
                                campo: nota
                                for campo in (
                                    "prueba",
                                    "unidades",
                                    "metodo",
                                    "especificacion",
                                    "valor_tipico",
                                )
                            }
                        )
                    origen.write_text(
                        json.dumps(
                            {
                                "documento": {"idioma": idioma},
                                "notas": notas,
                                "pruebas": pruebas,
                            }
                        ),
                        encoding="utf-8",
                    )

                    conteo = procesar_json(origen, destino, [MAESTRO])
                    salida = json.loads(destino.read_text(encoding="utf-8"))

                    self.assertEqual(conteo["alineados"], 1)
                    self.assertEqual(len(salida["pruebas"]), 1)
                    self.assertEqual(
                        salida["pruebas"][0]["prueba"],
                        "Viscosidad cinemática"
                        if idioma == "es"
                        else "Kinematic viscosity",
                    )
                    self.assertEqual(salida["notas"], [nota_canonica])
                    self.assertEqual(
                        salida["alineacion_metodos_lab"]["version_programa"],
                        VERSION,
                    )

    def test_normaliza_terminos_en_especificacion_y_valor_tipico_alineados(self):
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
                                "metodo": "ASTM D445",
                                "unidades": "cSt",
                                "especificacion": "típico / pasa / reportar / negativo",
                                "valor_tipico": "typical / pass / report / negative",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )

            procesar_json(origen, destino, [MAESTRO])
            prueba = json.loads(destino.read_text(encoding="utf-8"))["pruebas"][0]

            self.assertEqual(
                prueba["especificacion"],
                "TÍPICO / PASA / REPORTAR / NEGATIVA",
            )
            self.assertEqual(
                prueba["valor_tipico"],
                "TYPICAL / PASS / REPORT / NEGATIVE",
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
            self.assertEqual(prueba["especificacion"], "10 min")
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

    def test_no_agrega_farmacopeas_del_maestro_al_original(self):
        from interfaz_aplicacion_especificaciones import alinear_filas_generador
        maestro = {**MAESTRO, "METODO": "USP / FEUM", "MÉTODO INGLÉS": "USP / FEUM",
                   "NOMBRE IQM": "ALCALINIDAD", "NOMBRE IQM INGLES": "ALKALINITY"}
        casos = {"USP 32": "USP", "FEUM 13": "FEUM", "USP 32 / FEUM 13": "USP / FEUM",
                 "FEUM 13 / USP 32": "USP / FEUM", "USP-NF 2024": "USP"}
        for idioma, nombre in (("es", "ALCALINIDAD"), ("en", "ALKALINITY")):
            for original, esperado in casos.items():
                with self.subTest(idioma=idioma, original=original), tempfile.TemporaryDirectory() as carpeta:
                    origen, destino = Path(carpeta) / "entrada.json", Path(carpeta) / "alineado.json"
                    fila = {"prueba": nombre, "metodo": original, "especificacion": "PASA"}
                    origen.write_text(json.dumps({"documento": {"idioma": idioma}, "pruebas": [fila]}), encoding="utf-8")
                    conteo = procesar_json(origen, destino, [maestro])
                    datos = json.loads(destino.read_text(encoding="utf-8"))
                    detalle = datos["alineacion_metodos_lab"]["detalle"][0]
                    self.assertEqual(conteo["alineados"], 1)
                    self.assertEqual(datos["pruebas"][0]["metodo"], esperado)
                    self.assertEqual(detalle["valores_alineados"]["metodo"], esperado)
                    self.assertEqual(detalle["metodo_original"], original)
                    # Realinear el JSON generado conserva la regla y la auditoría.
                    procesar_json(destino, destino, [maestro])
                    self.assertEqual(json.loads(destino.read_text(encoding="utf-8"))["pruebas"][0]["metodo"], esperado)
                    filas, resumen = alinear_filas_generador([dict(fila)], [maestro], idioma)
                    self.assertEqual(resumen["alineados"], 1)
                    self.assertEqual(filas[0]["metodo"], esperado)

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

    def test_alineacion_conserva_metodos_astm_conjugados_ausentes_del_maestro(self):
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
                                "metodo": "ASTM D445 / ASTM D7042",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )

            conteo = procesar_json(origen, destino, [MAESTRO])
            salida = json.loads(destino.read_text(encoding="utf-8"))

            self.assertEqual(conteo["alineados"], 0)
            self.assertEqual(conteo["revision"], 1)
            self.assertEqual(
                salida["pruebas"][0]["metodo"],
                "ASTM D445 / ASTM D7042",
            )
            self.assertEqual(
                salida["alineacion_metodos_lab"]["detalle"][0]["valores_alineados"][
                    "metodo"
                ],
                "ASTM D445 / ASTM D7042",
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

    def test_descarta_registro_que_mezcla_unidades_cinematicas_y_dinamicas(self):
        segunda_unidad = {**MAESTRO, "UNIDADES": "mPa.s"}
        resultado = mejor_coincidencia(
            "Viscosidad cinemática",
            "",
            [MAESTRO, segunda_unidad],
        )

        self.assertEqual(resultado["estado"], "ALINEADO")
        self.assertFalse(resultado["ambiguo"])
        self.assertEqual(resultado["registro"]["UNIDADES"], "cSt")

    def test_orden_de_alternativas_no_provoca_ambiguedad_de_unidades(self):
        unidades_primera_opcion = {
            **MAESTRO,
            "UNIDADES": "cSt, g/cm³",
        }
        unidad_en_orden_distinto = {
            **MAESTRO,
            "UNIDADES": "g/cm³, cSt",
        }
        resultado = mejor_coincidencia(
            "Viscosidad cinemática",
            "ASTM D445",
            [unidades_primera_opcion, unidad_en_orden_distinto],
        )

        self.assertEqual(resultado["estado"], "ALINEADO")
        self.assertFalse(resultado["ambiguo"])

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
            self.assertEqual(registros[0]["MÉTODO INGLÉS"], "ASTM D445")


if __name__ == "__main__":
    unittest.main()
