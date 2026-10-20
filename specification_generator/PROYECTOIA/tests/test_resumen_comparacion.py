import copy
import json
from pathlib import Path
import tempfile
import tkinter as tk
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import interfaz_aplicacion_especificaciones as interfaz
from resumen_comparacion import resumir_diferencias_alineadas, formatear_resumen_diferencias


def fila(nombre="PUNTO DE FLUIDEZ", metodo="ASTM D97", valor="−45 MÁXIMO", **campos):
    return {"prueba": nombre, "metodo": metodo, "unidades": "°C",
            "especificacion": valor, "valor_tipico": "", **campos}


def documento(*filas):
    return {
        "producto": {"nombre": "PRODUCTO"}, "documento": {"idioma": "es"},
        "control_documental": {"revision": "00"}, "notas": [],
        "pruebas": copy.deepcopy(list(filas)),
        "alineacion_metodos_lab": {"detalle": [{"valores_alineados": copy.deepcopy(f)} for f in filas]},
    }


class ResumenComparacionTests(unittest.TestCase):
    def test_compara_original_sin_alineacion_sin_mutar_y_con_hashes_reales(self):
        a = documento(fila(valor="10", periodicidad="▶"), fila("Notas: Conservar cerrado", metodo="", valor="", unidades=""))
        a.pop("alineacion_metodos_lab")
        b = documento(fila(valor="20", periodicidad="▶I"))
        b["notas"] = ["Conservar cerrado"]
        originales = copy.deepcopy((a, b))
        frases = resumir_diferencias_alineadas(a, b)
        resumen = formatear_resumen_diferencias(a, b, "Original", "Alineada", frases)
        self.assertIn("«10» a «20»", resumen)
        self.assertIn("«▶» a «▶I»", resumen)
        self.assertNotIn("Se retiró la nota", resumen)
        self.assertNotIn("Se retiró la prueba", resumen)
        self.assertIn(f"Hash JSON origen (SHA-256, contenido): {interfaz.hash_especificacion(a)}", resumen)
        self.assertIn(f"Hash JSON destino (SHA-256, contenido): {interfaz.hash_especificacion(b)}", resumen)
        self.assertEqual((a, b), originales)

    def test_redacta_cambios_de_celdas_control_y_notas_sin_mutar(self):
        a = documento(fila(valor_tipico="−50"))
        b = documento(fila(valor="−30 MÁXIMO", metodo="ASTM D97 / ASTM D5950", unidades="°F"))
        b["control_documental"]["revision"] = "01"
        b["control_documental"]["nombre_cliente"] = "Cliente"
        a["notas"] = ["Nota anterior"]
        b["notas"] = ["Nota nueva"]
        originales = copy.deepcopy((a, b))
        frases = resumir_diferencias_alineadas(a, b)
        resumen = "\n".join(frases)
        for parte in ("Revisión: cambió de «00» a «01»", "Nombre del cliente: se agregó «Cliente»",
                      "la especificación cambió de «−45 MÁXIMO» a «−30 MÁXIMO»",
                      "el método cambió", "el campo de unidades cambió", "se retiró el valor típico: «−50»",
                      "Se retiró la nota: «Nota anterior»", "Se agregó la nota: «Nota nueva»"):
            self.assertIn(parte, resumen)
        self.assertEqual((a, b), originales)

    def test_usa_instantanea_alineada_e_ignora_auditoria_y_original(self):
        a = documento(fila())
        b = copy.deepcopy(a)
        b["pruebas"][0]["especificacion"] = "1000"
        b["alineacion_metodos_lab"]["version_programa"] = "OTRA VERSION"
        b["alineacion_metodos_lab"]["detalle"][0]["nombre_original"] = "Nombre viejo"
        self.assertEqual(resumir_diferencias_alineadas(a, b), ["No se detectaron diferencias en el contenido comparado."])
        b["alineacion_metodos_lab"]["detalle"][0]["valores_alineados"]["especificacion"] = "−30 MÁXIMO"
        self.assertIn("«−45 MÁXIMO» a «−30 MÁXIMO»", resumir_diferencias_alineadas(a, b)[0])

    def test_reordenamiento_y_cambio_en_una_temperatura(self):
        f40 = fila("VISCOSIDAD A 40 °C", "ASTM D445", "50")
        f100 = fila("VISCOSIDAD A 100 °C", "ASTM D445", "10")
        a = documento(f40, f100)
        self.assertEqual(resumir_diferencias_alineadas(a, documento(f100, f40)),
                         ["No se detectaron diferencias en el contenido comparado."])
        b = documento({**f100, "especificacion": "20"}, f40)
        frases = resumir_diferencias_alineadas(a, b)
        self.assertEqual(len(frases), 1)
        self.assertIn("VISCOSIDAD A 100 °C", frases[0])
        self.assertIn("«10» a «20»", frases[0])

    def test_pruebas_agregadas_retiradas_y_temperaturas_no_se_confunden(self):
        a = documento(fila("VISCOSIDAD A 40 °C", "ASTM D445", "50"))
        b = documento(fila("VISCOSIDAD A 100 °C", "ASTM D445", "10"))
        frases = resumir_diferencias_alineadas(a, b)
        self.assertEqual(len(frases), 2)
        self.assertIn("Se retiró la prueba «VISCOSIDAD A 40 °C»", frases[0])
        self.assertIn("Se agregó la prueba «VISCOSIDAD A 100 °C»", frases[1])

    def test_cambio_de_nombre_por_metodo_y_unidades_vacias(self):
        frases = resumir_diferencias_alineadas(documento(fila()), documento(fila("FLUIDEZ", unidades="")))
        self.assertEqual(len(frases), 2)
        self.assertIn("el nombre de la prueba cambió", frases[0])
        self.assertIn("se retiró el campo de unidades", frases[1])

    def test_metodo_compartido_no_basta_para_deducir_cambio_de_nombre(self):
        apariencia = fila("APARIENCIA", "ORGANOLÉPTICO", "CLARO")
        a = documento(apariencia, fila("OLOR", "ORGANOLÉPTICO", "CARACTERÍSTICO"))
        b = documento(apariencia, fila("COLOR", "ORGANOLÉPTICO", "VERDE"))
        frases = resumir_diferencias_alineadas(a, b)
        self.assertEqual(len(frases), 2)
        self.assertIn("Se retiró la prueba «OLOR»", frases[0])
        self.assertIn("Se agregó la prueba «COLOR»", frases[1])

    def test_duplicados_no_se_emparejan_por_su_posicion(self):
        a = documento(fila(valor="1"), fila(valor="2"))
        b = documento(fila(valor="4"), fila(valor="3"))
        frases = resumir_diferencias_alineadas(a, b)
        self.assertIn("no tienen una correspondencia única", frases[0])
        self.assertEqual(sum("Se retiró la prueba" in f for f in frases), 2)
        self.assertEqual(sum("Se agregó la prueba" in f for f in frases), 2)
        # Una repetición idéntica se reconoce antes de emparejar la restante.
        frases = resumir_diferencias_alineadas(a, documento(fila(valor="2"), fila(valor="3")))
        self.assertEqual(len(frases), 1)
        self.assertIn("«1» a «3»", frases[0])

    def test_periodicidad_notas_integradas_y_formatos(self):
        a = documento(fila(valor="50.0"), fila("Notas: Nota anterior", metodo="", valor="", unidades=""))
        b = documento(fila(valor="50,0", periodicidad="▶I"))
        a["pruebas"][0]["PERIODIDICAD DE ANALISIS"] = "▶"
        frases = "\n".join(resumir_diferencias_alineadas(a, b))
        self.assertIn("«50.0» a «50,0»", frases)
        self.assertIn("la periodicidad de análisis cambió de «▶» a «▶I»", frases)
        self.assertIn("Se retiró la nota: «Nota anterior»", frases)
        self.assertNotIn("Se retiró la prueba", frases)

    def test_alias_maestro_y_sentido_inverso(self):
        maestro = [{"NOMBRE IQM": "FLUIDEZ", "NOMBRE IQM INGLES": "POUR POINT"}]
        a, b = documento(fila("FLUIDEZ", valor="10")), documento(fila("POUR POINT", valor="20", metodo="ASTM D5950"))
        frases = "\n".join(resumir_diferencias_alineadas(a, b, maestro))
        self.assertIn("«10» a «20»", frases)
        self.assertNotIn("Se retiró la prueba", frases)
        self.assertIn("«20» a «10»", "\n".join(resumir_diferencias_alineadas(b, a, maestro)))


class VentanaDiferenciasTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.raiz = tk.Tk()
        cls.raiz.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.raiz.destroy()

    def setUp(self):
        self.padre = tk.Toplevel(self.raiz)
        self.padre.withdraw()
        self.comparador = SimpleNamespace(
            ventana=self.padre, _cargar_maestro_comparacion=lambda: (),
            paneles=[SimpleNamespace(indice=i, ruta=Path(f"producto_{i}.json"),
                                     vista={"alineada": True, "datos": documento(fila(valor=str(i)))})
                     for i in (1, 2, 3)],
        )
        constructor = interfaz.ui.Ventana

        def invisible(*args, **kwargs):
            ventana = constructor(*args, **kwargs)
            ventana.attributes("-alpha", 0)
            ventana.withdraw()
            return ventana

        with patch.object(interfaz.ui, "Ventana", invisible):
            self.resumen = interfaz.VentanaDiferenciasComparacion(self.comparador)

    def tearDown(self):
        self.padre.destroy()

    def test_copia_exactamente_el_resumen_y_permite_cambiar_sentido(self):
        resumen = self.resumen
        for etiqueta, panel in zip(("origen", "destino"), self.comparador.paneles):
            huella = interfaz.hash_especificacion(panel.vista["datos"])
            self.assertIn(f"Hash JSON {etiqueta} (SHA-256, contenido): {huella}",
                          resumen.texto_resumen.split("\n\n", 1)[0])
        self.assertIn("«1» a «2»", resumen.texto.get("1.0", "end-1c"))
        self.assertEqual(str(resumen.texto.cget("state")), "disabled")
        resumen.boton_copiar.invoke()
        self.assertEqual(resumen.ventana.clipboard_get(), resumen.texto.get("1.0", "end-1c"))
        a, b = resumen.origen.get(), resumen.destino.get()
        resumen.origen.set(b)
        resumen.destino.set(a)
        resumen._redactar()
        self.assertIn("«2» a «1»", resumen.texto_resumen)
        resumen.destino.set("Especificación 3 · producto_3.json")
        resumen.selectores[1].event_generate("<<ComboboxSelected>>")
        self.assertIn("«2» a «3»", resumen.texto_resumen)

    def test_seleccion_repetida_o_panel_limpiado_invalida_resumen_anterior(self):
        resumen = self.resumen
        resumen.destino.set(resumen.origen.get())
        resumen._redactar()
        self.assertEqual(resumen.texto_resumen, "")
        self.assertEqual(resumen.boton_copiar.cget("state"), "disabled")
        self.comparador.paneles[1].vista = None
        self.comparador.paneles[2].vista["alineada"] = False
        resumen.actualizar()
        self.assertEqual(resumen.textos_cambios, [])
        self.assertNotIn("«1» a «2»", resumen.texto.get("1.0", "end-1c"))
        self.assertEqual(resumen.boton_copiar.cget("state"), "disabled")

    def test_carga_real_de_json_muestra_delta_reutiliza_ventana_y_cierra(self):
        constructor = interfaz.ui.Ventana

        def invisible(*args, **kwargs):
            ventana = constructor(*args, **kwargs)
            ventana.attributes("-alpha", 0)
            ventana.withdraw()
            return ventana

        with tempfile.TemporaryDirectory() as carpeta, patch.object(interfaz.ui, "Ventana", invisible):
            c = interfaz.Comparador(SimpleNamespace(raiz=self.padre))
            c._maestro_comparacion = ()
            self.assertEqual(c.boton_diferencias.winfo_manager(), "")
            for indice, panel in enumerate(c.paneles):
                ruta = Path(carpeta) / f"alineada_{indice}.json"
                ruta.write_text(json.dumps(documento(fila(valor=str(indice)))), encoding="utf-8")
                panel.cargar_archivo(ruta)
            self.assertEqual(c.boton_diferencias.winfo_manager(), "pack")
            self.assertEqual(c.boton_diferencias.cget("width"), c.boton_diferencias.cget("height"))
            c.boton_diferencias.invoke()
            ventana = c.ventana_diferencias
            self.assertIn("«0» a «1»", ventana.texto_resumen)
            c.mostrar_diferencias()
            self.assertIs(c.ventana_diferencias, ventana)
            c.paneles[0].limpiar()
            self.assertEqual(c.boton_diferencias.winfo_manager(), "")
            self.assertEqual(ventana.texto_resumen, "")
            c.cerrar()
            self.assertFalse(ventana.ventana.winfo_exists())


if __name__ == "__main__":
    unittest.main()
