import copy
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from edicion_bilingue import BorradoresBilingues, guardar_pareja_json
from traductor_especificaciones import traducir as traducir_del_diccionario
from traductor_especificaciones import DiccionarioTraductor


def traducir(campo, texto, origen, destino, fila):
    pares = {"DENSIDAD": "DENSITY", "COLOR": "COLOR", "g/ml": "g/ml", "Nota": "Note"}
    if origen == "en":
        pares = {v: k for k, v in pares.items()}
    if texto in pares:
        return pares[texto], True
    return traducir_del_diccionario(texto, origen, destino)


def contenido():
    return {"idioma": "es", "nombre_producto": "Producto", "desarrollador": "Nombre",
            "control": {"codigo_unico_producto": "P1", "revision": "01", "nombre_cliente": "Cliente", "elaboro_y_aprobo": ""},
            "notas": ["Nota"], "pruebas": [
                {"prueba": "DENSIDAD", "unidades": "g/ml", "metodo": "ASTM D445",
                 "especificacion": "149 MÍNIMO", "valor_tipico": "150", "periodicidad": "A"}]}


class BilingueTests(unittest.TestCase):
    def test_patrones_conservan_cifras_y_detectan_texto_libre(self):
        self.assertEqual(traducir_del_diccionario("149 MÍNIMO", "es", "en"), ("149 MINIMUM", True))
        self.assertEqual(traducir_del_diccionario("0.130 – 0.200", "es", "en"), ("0.130 – 0.200", True))
        self.assertEqual(traducir_del_diccionario("ASTM D445", "es", "en"), ("ASTM D445", True))
        self.assertEqual(traducir_del_diccionario("INS-CCA-07", "es", "en"), ("INS-CCA-07", True))
        self.assertEqual(traducir_del_diccionario("GRIS", "es", "en"), ("GRAY", True))
        self.assertEqual(traducir_del_diccionario("MAXIMUM 30", "en", "es"), ("MÁXIMO 30", True))
        self.assertEqual(traducir_del_diccionario("GRIS CLARO", "es", "en"), ("LIGHT GRAY", True))
        self.assertFalse(traducir_del_diccionario("GRIS DESCONOCIDO", "es", "en")[1])
        self.assertEqual(traducir_del_diccionario("MAXIMOMETRO", "es", "en")[0], "MAXIMOMETRO")

    def test_colores_y_apariencia_en_ambos_idiomas(self):
        ejemplos = (("VERDE CLARO", "LIGHT GREEN"), ("AZUL OSCURO", "DARK BLUE"),
                    ("AMARILLO PÁLIDO", "PALE YELLOW"), ("LIGERAMENTE TURBIO", "SLIGHTLY HAZY"),
                    ("CLARO Y BRILLANTE", "CLEAR AND BRIGHT"), ("TRANSPARENTE", "TRANSPARENT"),
                    ("OPACO", "OPAQUE"), ("ÁMBAR", "AMBER"))
        # El Excel del usuario puede cambiar; estas equivalencias son datos de prueba.
        diccionario = DiccionarioTraductor([{"es": es, "en": en, "activo": True}
                                           for es, en in (*ejemplos, ("BRILANTE", "BRIGHT"), ("AMBAR", "AMBER"))])
        for es, en in ejemplos:
            with self.subTest(es=es):
                self.assertEqual(diccionario.traducir(es, "es", "en"), (en, True))
                self.assertEqual(diccionario.traducir(en, "en", "es"), (es, True))
        self.assertEqual(diccionario.traducir("BRILANTE", "es", "en"), ("BRIGHT", True))
        self.assertEqual(diccionario.traducir("AMBAR", "es", "en"), ("AMBER", True))

    def test_cambiar_idioma_traduce_pareja_guardada_sin_editar_origen(self):
        diccionario = DiccionarioTraductor([{"es": "CLARO Y BRILLANTE", "en": "BRIGHT AND CLEAR", "activo": True}])
        def convertir(campo, valor, origen, destino, fila):
            return diccionario.traducir(valor, origen, destino)
        fuente = contenido()
        fuente["pruebas"][0]["especificacion"] = "CLARO Y BRILLANTE"
        fuente["pruebas"][0]["valor_tipico"] = "CLARO Y BRILLANTE"
        modelo = BorradoresBilingues(fuente, convertir)
        pareja = copy.deepcopy(modelo.borradores["en"])
        pareja["pruebas"][0]["especificacion"] = "CLARO Y BRILLANTE"
        pareja["pruebas"][0]["valor_tipico"] = "CLARO Y BRILLANTE"
        modelo.cargar_pareja(pareja)
        modelo.actualizar_traducciones_automaticas("en", completar_sin_traducir=True)
        self.assertEqual(modelo.borradores["en"]["pruebas"][0]["especificacion"], "BRIGHT AND CLEAR")
        self.assertEqual(modelo.borradores["en"]["pruebas"][0]["valor_tipico"], "BRIGHT AND CLEAR")
        modelo.actualizar_traducciones_automaticas("es", completar_sin_traducir=True)
        self.assertEqual(modelo.borradores["es"]["pruebas"][0]["especificacion"], "CLARO Y BRILLANTE")

    def test_cambiar_idioma_respeta_correccion_y_confirmacion_manual(self):
        modelo = BorradoresBilingues(contenido(), traducir)
        en = copy.deepcopy(modelo.borradores["en"])
        en["pruebas"][0]["especificacion"] = "At least 149"
        modelo.actualizar("en", en)
        clave = (en["pruebas"][0]["_id_bilingue"], "especificacion")
        modelo.confirmar("en", clave)
        modelo.actualizar_traducciones_automaticas("en", completar_sin_traducir=True)
        self.assertEqual(modelo.borradores["en"]["pruebas"][0]["especificacion"], "At least 149")

    def test_confirmacion_de_texto_igual_al_origen_se_conserva(self):
        modelo = BorradoresBilingues(contenido(), traducir)
        fila = modelo.borradores["en"]["pruebas"][0]
        fila["especificacion"] = modelo.borradores["es"]["pruebas"][0]["especificacion"]
        modelo.confirmar("en", (fila["_id_bilingue"], "especificacion"))
        modelo.actualizar_traducciones_automaticas("en", completar_sin_traducir=True)
        self.assertEqual(fila["especificacion"], "149 MÍNIMO")

    def test_conserva_borradores_y_actualiza_traducciones_automaticas(self):
        modelo = BorradoresBilingues(contenido(), traducir)
        es = copy.deepcopy(modelo.borradores["es"])
        es["pruebas"][0]["especificacion"] = "160 MÍNIMO"
        modelo.actualizar("es", es)
        self.assertEqual(modelo.borradores["en"]["pruebas"][0]["especificacion"], "160 MINIMUM")
        self.assertEqual(modelo.borradores["es"]["pruebas"][0]["especificacion"], "160 MÍNIMO")
        self.assertFalse(modelo.pendientes["en"])

    def test_no_sobrescribe_traduccion_manual_y_permite_confirmarla(self):
        modelo = BorradoresBilingues(contenido(), traducir)
        en = copy.deepcopy(modelo.borradores["en"])
        en["pruebas"][0]["especificacion"] = "At least 149"
        modelo.actualizar("en", en)
        es = copy.deepcopy(modelo.borradores["es"])
        es["pruebas"][0]["especificacion"] = "160 MÍNIMO"
        modelo.actualizar("es", es)
        clave = (es["pruebas"][0]["_id_bilingue"], "especificacion")
        self.assertEqual(modelo.borradores["en"]["pruebas"][0]["especificacion"], "At least 149")
        self.assertIn(clave, modelo.pendientes["en"])
        modelo.confirmar("en", clave)
        self.assertNotIn(clave, modelo.pendientes["en"])

    def test_sincroniza_numeros_control_y_periodicidad(self):
        modelo = BorradoresBilingues(contenido(), traducir)
        pareja = copy.deepcopy(modelo.borradores["en"])
        modelo.cargar_pareja(pareja)
        es = copy.deepcopy(modelo.borradores["es"])
        es["control"]["revision"] = "02"
        es["pruebas"][0]["valor_tipico"] = "180"
        es["pruebas"][0]["periodicidad"] = "B"
        modelo.actualizar("es", es)
        self.assertEqual(modelo.borradores["en"]["control"]["revision"], "02")
        self.assertEqual(modelo.borradores["en"]["pruebas"][0]["valor_tipico"], "180")
        self.assertEqual(modelo.borradores["en"]["pruebas"][0]["periodicidad"], "B")

    def test_agregar_eliminar_y_reordenar_no_mezcla_filas(self):
        modelo = BorradoresBilingues(contenido(), traducir)
        es = copy.deepcopy(modelo.borradores["es"])
        es["pruebas"].append({"prueba": "COLOR", "unidades": "N/A", "metodo": "ASTM D1500", "especificacion": "GRIS", "valor_tipico": "", "periodicidad": "B"})
        es = modelo.actualizar("es", es)
        es["pruebas"].reverse()
        modelo.actualizar("es", es)
        self.assertEqual([f["prueba"] for f in modelo.borradores["en"]["pruebas"]], ["COLOR", "DENSITY"])
        es["pruebas"].pop()
        modelo.actualizar("es", es)
        self.assertEqual(len(modelo.borradores["en"]["pruebas"]), 1)

    def test_pareja_se_relaciona_por_prueba_no_por_posicion(self):
        inicial = contenido()
        inicial["pruebas"].append({"prueba": "COLOR", "metodo": "ASTM D1500", "unidades": "N/A", "especificacion": "GRIS", "valor_tipico": "", "periodicidad": "A"})
        modelo = BorradoresBilingues(inicial, traducir)
        pareja = copy.deepcopy(modelo.borradores["en"])
        pareja["pruebas"].reverse()
        modelo.cargar_pareja(pareja)
        self.assertEqual(modelo.borradores["en"]["pruebas"][0]["prueba"], "DENSITY")
        pareja["pruebas"].append(copy.deepcopy(pareja["pruebas"][0]))
        with self.assertRaises(ValueError):
            modelo.cargar_pareja(pareja)

    def test_texto_libre_y_notas_pendientes(self):
        inicial = contenido()
        inicial["notas"] = ["Texto sin equivalencia"]
        inicial["pruebas"][0]["especificacion"] = "Amarillo con matiz desconocido"
        modelo = BorradoresBilingues(inicial, traducir)
        self.assertIn(("cabecera", "notas"), modelo.pendientes["en"])
        self.assertEqual(len(modelo.pendientes["en"]), 2)
        en = copy.deepcopy(modelo.borradores["en"])
        en["notas"] = ["Light yellow product"]
        modelo.actualizar("en", en)
        self.assertNotIn(("cabecera", "notas"), modelo.pendientes["en"])
        self.assertEqual(modelo.borradores["es"]["notas"], ["Texto sin equivalencia"])
        self.assertNotIn(("cabecera", "notas"), modelo.pendientes["es"])

    def test_borrar_traduccion_no_borra_origen_y_origen_puede_completarla(self):
        modelo = BorradoresBilingues(contenido(), traducir)
        en = copy.deepcopy(modelo.borradores["en"])
        en["pruebas"][0]["especificacion"] = ""
        modelo.actualizar("en", en)
        clave = (en["pruebas"][0]["_id_bilingue"], "especificacion")
        self.assertIn(clave, modelo.pendientes["en"])
        self.assertEqual(modelo.borradores["es"]["pruebas"][0]["especificacion"], "149 MÍNIMO")
        es = copy.deepcopy(modelo.borradores["es"])
        es["pruebas"][0]["especificacion"] = "160 MÍNIMO"
        modelo.actualizar("es", es)
        self.assertEqual(modelo.borradores["en"]["pruebas"][0]["especificacion"], "160 MINIMUM")
        self.assertNotIn(clave, modelo.pendientes["en"])

    def test_campo_opcional_vacio_en_ambos_no_queda_pendiente(self):
        modelo = BorradoresBilingues(contenido(), traducir)
        en = copy.deepcopy(modelo.borradores["en"])
        en["pruebas"][0]["valor_tipico"] = ""
        modelo.actualizar("en", en)
        es = copy.deepcopy(modelo.borradores["es"])
        es["pruebas"][0]["valor_tipico"] = ""
        modelo.actualizar("es", es)
        clave = (es["pruebas"][0]["_id_bilingue"], "valor_tipico")
        self.assertNotIn(clave, modelo.pendientes["en"])
        self.assertNotIn(clave, modelo.pendientes["es"])

    def test_guardado_restaurado_si_falla_el_segundo_reemplazo(self):
        import os
        with tempfile.TemporaryDirectory() as carpeta:
            es, en = Path(carpeta) / "es.json", Path(carpeta) / "en.json"
            es.write_text("original es", encoding="utf-8")
            en.write_text("original en", encoding="utf-8")
            reemplazar = os.replace
            def fallo(origen, destino):
                if destino == en:
                    raise OSError("fallo simulado")
                reemplazar(origen, destino)
            with patch("os.replace", side_effect=fallo), self.assertRaises(OSError):
                guardar_pareja_json({es: {"idioma": "es"}, en: {"idioma": "en"}})
            self.assertEqual(es.read_text(encoding="utf-8"), "original es")
            self.assertEqual(en.read_text(encoding="utf-8"), "original en")
            self.assertEqual(len(list(Path(carpeta).iterdir())), 2)
            guardar_pareja_json({es: {"idioma": "es"}, en: {"idioma": "en"}})
            self.assertEqual(json.loads(en.read_text(encoding="utf-8"))["idioma"], "en")


if __name__ == "__main__":
    unittest.main()
