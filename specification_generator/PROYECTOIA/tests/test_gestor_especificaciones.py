import queue
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import gestor_especificaciones


def _worker_bloqueado(path, destino, resultados):
    time.sleep(10)


class _ColaSimulada:
    def __init__(self):
        self.mensajes = [
            ("stage", "iniciar Microsoft Word"),
            ("word_pid", 321),
        ]

    def get(self, timeout=None):
        if self.mensajes:
            return self.mensajes.pop(0)
        raise queue.Empty

    def close(self):
        pass

    def cancel_join_thread(self):
        pass


class _ProcesoSimulado:
    def __init__(self, target, args):
        self.target = target
        self.args = args
        self.vivo = False
        self.exitcode = None

    def start(self):
        self.vivo = True

    def is_alive(self):
        return self.vivo

    def terminate(self):
        self.vivo = False

    def kill(self):
        self.vivo = False

    def join(self, timeout=None):
        pass


class _ContextoSimulado:
    def Queue(self):
        return _ColaSimulada()

    def Process(self, *, target, args):
        self.proceso = _ProcesoSimulado(target, args)
        return self.proceso


class ConversionWordTests(unittest.TestCase):
    def test_worker_real_no_bloquea_indefinidamente_si_word_no_responde(self):
        with patch.object(gestor_especificaciones, "TIEMPO_LIMITE_CONVERSION_WORD", 0.1):
            with self.assertRaisesRegex(TimeoutError, "excedió el límite"):
                gestor_especificaciones._convertir_word_com(
                    Path("origen.doc"),
                    Path("destino.docx"),
                    worker=_worker_bloqueado,
                )

    def test_termina_el_worker_si_se_agota_el_tiempo_de_conversion(self):
        contexto = _ContextoSimulado()
        with (
            patch.object(gestor_especificaciones, "_PROCESOS_WORD_PREEXISTENTES", {100}),
            patch.object(gestor_especificaciones, "_ERROR_INICIO_WORD", None),
            patch.object(gestor_especificaciones, "_procesos_word_automatizacion",
                         return_value={100, 321}),
            patch.object(gestor_especificaciones.multiprocessing, "get_context", return_value=contexto),
            patch.object(gestor_especificaciones, "TIEMPO_LIMITE_CONVERSION_WORD", 0.01),
            patch.object(gestor_especificaciones, "_terminar_proceso_word") as terminar_word,
        ):
            with self.assertRaisesRegex(TimeoutError, "excedió el límite"):
                gestor_especificaciones._convertir_word_com(
                    Path("origen.doc"),
                    Path("destino.docx"),
                )

            self.assertFalse(contexto.proceso.is_alive())
            self.assertIn("iniciar Microsoft Word", gestor_especificaciones._ERROR_INICIO_WORD)
            terminar_word.assert_called_once_with(321)

    def test_no_reintenta_inicio_word_repetidamente_en_la_misma_carga(self):
        motivo = "Word quedó bloqueado durante el inicio"
        with (
            patch.object(gestor_especificaciones, "_ERROR_INICIO_WORD", motivo),
            patch.object(gestor_especificaciones.multiprocessing, "get_context") as contexto,
            self.assertRaisesRegex(RuntimeError, "ya falló su inicio en esta carga"),
        ):
            gestor_especificaciones._convertir_word_com(
                Path("siguiente.doc"),
                Path("siguiente.docx"),
            )

        contexto.assert_not_called()

    def test_restablece_el_cortacircuitos_entre_cargas(self):
        with patch.object(
            gestor_especificaciones,
            "_ERROR_INICIO_WORD",
            "fallo anterior",
        ):
            gestor_especificaciones._restablecer_error_inicio_word()
            self.assertIsNone(gestor_especificaciones._ERROR_INICIO_WORD)

    def test_limpia_solo_instancias_automatizadas_creadas_durante_conversion(self):
        with (
            patch.object(gestor_especificaciones, "_procesos_word_automatizacion",
                         return_value={100, 321, 654}),
            patch.object(gestor_especificaciones, "_terminar_proceso_word") as terminar_word,
        ):
            gestor_especificaciones._limpiar_word_iniciado({100, 321}, None)

        terminar_word.assert_called_once_with(654)

    def test_reporta_timeout_si_word_y_libreoffice_no_convierten(self):
        with (
            patch.object(gestor_especificaciones.sys, "platform", "win32"),
            patch.object(gestor_especificaciones, "_convertir_word_com",
                         side_effect=TimeoutError("Word excedió el tiempo")),
            patch.object(gestor_especificaciones.shutil, "which", return_value=None),
            self.assertRaisesRegex(RuntimeError, "Word excedió el tiempo"),
        ):
            gestor_especificaciones.convertir_word_legacy(
                Path("origen.doc"),
                Path("temporal"),
            )


if __name__ == "__main__":
    unittest.main()
