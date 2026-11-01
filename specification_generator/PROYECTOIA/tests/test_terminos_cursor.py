import unittest

from interfaz_aplicacion_especificaciones import _sugerencias_terminos_cursor
from traductor_especificaciones import DiccionarioTraductor


class TerminosCursorTests(unittest.TestCase):
    def test_frase_mas_larga_y_variantes_sin_acentos(self):
        texto = "30 AMBAR C; 10 mm²/s ¹"
        inicio, fin, opciones = _sugerencias_terminos_cursor(
            texto, len("30 AMBAR C"), ("CLARO", "ÁMBAR CLARO", "AMBAR CLARO", "ÁMBAR CLARA"),
        )
        self.assertEqual(opciones, ("ÁMBAR CLARO", "ÁMBAR CLARA"))
        self.assertEqual(texto[:inicio] + opciones[0] + texto[fin:], "30 ÁMBAR CLARO; 10 mm²/s ¹")

    def test_cursor_en_medio_reemplaza_solo_la_palabra(self):
        texto = "30 MAXI; 40 mm²/s ¹"
        inicio, fin, opciones = _sugerencias_terminos_cursor(texto, 6, ("MÁXIMO",))
        self.assertEqual(texto[:inicio] + opciones[0] + texto[fin:], "30 MÁXIMO; 40 mm²/s ¹")

    def test_no_sugiere_para_numeros_espacios_o_una_letra(self):
        for texto in ("30", "30 ", "M", "30 ¹"):
            self.assertEqual(_sugerencias_terminos_cursor(texto, len(texto), ("MÁXIMO",))[2], ())

    def test_idioma_activos_y_limite(self):
        diccionario = DiccionarioTraductor([
            {"es": "MÁXIMO", "en": "MAXIMUM", "activo": True},
            {"es": "MÍNIMO", "en": "MINIMUM", "activo": False},
        ])
        self.assertEqual(_sugerencias_terminos_cursor("MA", 2, diccionario.mapas["en"])[2], ("MAXIMUM",))
        self.assertEqual(_sugerencias_terminos_cursor("MI", 2, diccionario.mapas["es"])[2], ())
        self.assertEqual(len(_sugerencias_terminos_cursor("AM", 2, (f"AMBAR {n}" for n in range(20)))[2]), 5)
