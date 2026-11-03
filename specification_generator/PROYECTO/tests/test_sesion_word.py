import multiprocessing
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import gestor_especificaciones as gestor
from conversion_word import ErrorCOM, SesionWord, TiempoCOM


def _worker_simulado(ordenes, resultados):
    resultados.put(("ready", None))
    while True:
        orden = ordenes.get()
        if orden is None:
            return
        origen, destino = orden
        if Path(origen).name == "fallar.doc":
            resultados.put(("error", (False, "Documento no convertible")))
            return
        if Path(origen).name == "bloquear.doc":
            time.sleep(10)
            return
        Path(destino).write_text(str(os.getpid()), encoding="utf-8")
        resultados.put(("ok", None))


def _worker_inicio_fallido(ordenes, resultados):
    resultados.put(("error", (True, "No inicia Word")))


class SesionWordTests(unittest.TestCase):
    def setUp(self):
        temporal = tempfile.TemporaryDirectory()
        self.addCleanup(temporal.cleanup)
        self.base = Path(temporal.name)
        self.inspeccionar = Mock(return_value=set())
        self.terminar = Mock()
        self.sesion = SesionWord(self.inspeccionar, self.terminar, worker=_worker_simulado)
        self.addCleanup(self.sesion.cerrar)

    def convertir(self, nombre="a.doc", limite=5):
        destino = self.base / (Path(nombre).stem + ".docx")
        self.sesion.convertir(self.base / nombre, destino, tiempo_limite=limite)
        return destino

    def test_reutiliza_el_proceso_y_cierra_al_final(self):
        a = self.convertir("a.doc")
        b = self.convertir("b.docm")
        proceso = self.sesion.proceso
        self.assertEqual(a.read_text(), b.read_text())
        self.inspeccionar.assert_called_once()
        self.sesion.cerrar()
        self.assertFalse(proceso.is_alive())
        self.terminar.assert_not_called()

    def test_recicla_despues_del_limite_del_lote(self):
        self.sesion.max_documentos = 2
        self.convertir("a.doc")
        proceso = self.sesion.proceso
        self.convertir("b.doc")
        self.assertIs(self.sesion.proceso, proceso)
        self.convertir("c.doc")
        self.assertIsNot(self.sesion.proceso, proceso)
        self.assertFalse(proceso.is_alive())
        self.assertEqual(self.inspeccionar.call_count, 2)

    def test_fallo_de_documento_reinicia_para_el_siguiente(self):
        self.convertir()
        with self.assertRaises(ErrorCOM) as error:
            self.convertir("fallar.doc")
        self.assertFalse(error.exception.inicio)
        self.assertIsNone(self.sesion.proceso)
        self.assertTrue(self.convertir("siguiente.doc").exists())

    def test_timeout_termina_trabajador_y_permite_continuar(self):
        self.convertir()
        proceso = self.sesion.proceso
        with self.assertRaises(TiempoCOM) as error:
            self.convertir("bloquear.doc", limite=.2)
        self.assertFalse(error.exception.inicio)
        self.assertFalse(proceso.is_alive())
        self.assertTrue(self.convertir("siguiente.doc").exists())

    def test_identifica_fallo_global_de_inicio(self):
        self.sesion.worker = _worker_inicio_fallido
        with self.assertRaises(ErrorCOM) as error:
            self.convertir()
        self.assertTrue(error.exception.inicio)
        self.assertIsNone(self.sesion.proceso)

    def test_no_inicia_si_no_puede_proteger_instancias_preexistentes(self):
        self.inspeccionar.side_effect = RuntimeError("Inspección no disponible")
        with self.assertRaises(ErrorCOM) as error:
            self.convertir()
        self.assertTrue(error.exception.inicio)
        self.assertIsNone(self.sesion.proceso)

    def test_no_inicia_word_para_un_lote_solo_docx(self):
        with patch.object(gestor, "_procesos_word_automatizacion") as inspeccionar:
            with gestor.sesion_word_compartida() as sesion:
                self.assertIsNone(sesion.proceso)
            inspeccionar.assert_not_called()

    def test_no_termina_instancias_ambiguas_ni_preexistentes(self):
        self.sesion.proceso = Mock()
        self.sesion.proceso.is_alive.return_value = False
        self.sesion.preexistentes = {100}
        self.inspeccionar.return_value = {100, 321, 654}
        self.sesion._abortar()
        self.terminar.assert_not_called()

    def test_termina_solo_el_pid_conocido_de_su_word(self):
        self.sesion.proceso = Mock()
        self.sesion.proceso.is_alive.return_value = False
        self.sesion.preexistentes = {100}
        self.sesion.pid_word = 321
        self.sesion._abortar()
        self.terminar.assert_called_once_with(321)
        self.inspeccionar.assert_not_called()

    def test_inspeccion_fallida_al_abortar_no_oculta_error_ni_impide_reintento(self):
        self.convertir()
        self.inspeccionar.side_effect = RuntimeError("Inspección no disponible")
        with self.assertRaisesRegex(ErrorCOM, "Documento no convertible"):
            self.convertir("fallar.doc")
        self.assertIsNone(self.sesion.proceso)
        self.terminar.assert_not_called()
        self.inspeccionar.side_effect = None
        self.assertTrue(self.convertir("siguiente.doc").exists())


@unittest.skipUnless(os.environ.get("WORD_INTEGRATION"), "Prueba optativa con Microsoft Word real")
class SesionWordRealTests(unittest.TestCase):
    def test_documentos_doc_y_docm_comparten_instancia_y_conservan_datos(self):
        base = Path(gestor.__file__).parent
        casos = ((base / "REVISAR" / "TRANSCHUTZ SYN GL-1 SAE 140.doc", 13),
                 (base / "REVISAR" / "SCHUTZ PRO FR RB1 FIRESTONE.doc", 5),
                 (base / "REVISAR" / "SCHUTZ PRO II HVI 110-115 ENG 2021(1).doc", 12),
                 (base / gestor.ENTRADA / "ESPECIFICACION- SCHUTZ ALPHA LITE 380.docm", 12))
        pids = []
        tiempos = []
        with tempfile.TemporaryDirectory() as temporal, gestor.sesion_word_compartida() as sesion:
            gestor._restablecer_error_inicio_word()
            for i, (fuente, cantidad) in enumerate(casos):
                original = gestor.sha256(fuente)
                inicio = time.monotonic()
                destino = Path(temporal) / f"{i}.docx"
                gestor._convertir_word_com(fuente, destino)
                tiempos.append(round(time.monotonic() - inicio, 3))
                pids.append(sesion.pid_word)
                self.assertEqual(len(gestor.extraer(destino)["pruebas"]), cantidad)
                self.assertEqual(gestor.sha256(fuente), original)
            self.assertEqual(len(set(pids)), 1)
            self.assertIsNotNone(pids[0])
        print("Conversión real por documento (s):", tiempos, "; una instancia de Word.")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    unittest.main()
