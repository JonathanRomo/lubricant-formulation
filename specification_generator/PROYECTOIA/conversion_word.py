"""Trabajador COM reutilizable, supervisado y exclusivo de una carga."""
from __future__ import annotations

import multiprocessing
import queue
import time


class ErrorCOM(RuntimeError):
    def __init__(self, mensaje, *, inicio=False):
        super().__init__(mensaje)
        self.inicio = inicio


class TiempoCOM(TimeoutError):
    def __init__(self, mensaje, *, inicio=False):
        super().__init__(mensaje)
        self.inicio = inicio


def trabajador_word(ordenes, resultados):
    import pythoncom
    import win32com.client
    import win32process

    word = None
    documento = None
    inicializado = False
    inicio = True
    try:
        resultados.put(("stage", "iniciar automatización COM"))
        pythoncom.CoInitialize()
        inicializado = True
        resultados.put(("stage", "iniciar Microsoft Word"))
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        resultados.put(("ready", None))
        inicio = False
        while True:
            orden = ordenes.get()
            if orden is None:
                break
            origen, destino = orden
            resultados.put(("stage", "abrir documento"))
            documento = word.Documents.Open(
                origen, ConfirmConversions=False, ReadOnly=True,
                AddToRecentFiles=False, Visible=False, OpenAndRepair=False,
                NoEncodingDialog=True,
            )
            try:
                hwnd = documento.Windows.Item(1).Hwnd
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                resultados.put(("word_pid", pid))
            except Exception:
                pass
            resultados.put(("stage", "guardar documento temporal"))
            documento.SaveAs2(destino, FileFormat=16)
            resultados.put(("stage", "cerrar documento"))
            documento.Close(False)
            documento = None
            resultados.put(("ok", None))
    except Exception as error:
        resultados.put(("error", (inicio, f"{type(error).__name__}: {error}")))
    finally:
        if documento is not None:
            try:
                documento.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        if inicializado:
            pythoncom.CoUninitialize()


class SesionWord:
    def __init__(self, inspeccionar, terminar, *, worker=trabajador_word, max_documentos=50,
                 inspeccionar_preexistentes=None):
        self.inspeccionar = inspeccionar
        self.inspeccionar_preexistentes = inspeccionar_preexistentes or inspeccionar
        self.terminar = terminar
        self.worker = worker
        self.max_documentos = max_documentos
        self.proceso = None
        self.ordenes = self.resultados = None
        self.preexistentes = set()
        self.pid_word = None
        self.documentos = 0
        self.lista = False

    def _esperar(self, esperado, limite, *, inicio):
        etapa = "iniciar el proceso de conversión" if inicio else "convertir el documento"
        while True:
            restante = limite - time.monotonic()
            if restante <= 0:
                raise TiempoCOM(f"Microsoft Word excedió el tiempo al intentar {etapa}.", inicio=inicio)
            try:
                tipo, valor = self.resultados.get(timeout=min(.1, restante))
            except queue.Empty:
                if not self.proceso.is_alive():
                    raise ErrorCOM("El trabajador de Word terminó sin resultado.", inicio=inicio)
                continue
            if tipo == "stage":
                etapa = valor
            elif tipo == "word_pid":
                self.pid_word = valor
            elif tipo == "error":
                es_inicio, mensaje = valor
                raise ErrorCOM(mensaje, inicio=es_inicio)
            elif tipo == esperado:
                return

    def convertir(self, origen, destino, *, tiempo_limite):
        if self.documentos >= self.max_documentos:
            self.cerrar()
        limite = time.monotonic() + tiempo_limite
        try:
            if self.proceso is None:
                try:
                    self.preexistentes = self.inspeccionar_preexistentes()
                except Exception as error:
                    raise ErrorCOM(f"No se pudieron inspeccionar las instancias de Word: {error}", inicio=True) from error
                contexto = multiprocessing.get_context("spawn")
                self.ordenes, self.resultados = contexto.Queue(), contexto.Queue()
                self.proceso = contexto.Process(target=self.worker, args=(self.ordenes, self.resultados))
                self.proceso.start()
                self._esperar("ready", limite, inicio=True)
                self.lista = True
            self.ordenes.put((str(origen.resolve()), str(destino.resolve())))
            self._esperar("ok", limite, inicio=False)
            self.documentos += 1
        except BaseException:
            self._abortar()
            raise

    def _abortar(self):
        try:
            self._detener()
        finally:
            self._liberar()

    def _detener(self):
        if self.proceso is not None:
            if self.proceso.is_alive():
                self.proceso.terminate()
                self.proceso.join(timeout=5)
                if self.proceso.is_alive():
                    self.proceso.kill()
                    self.proceso.join(timeout=5)
            if self.pid_word is not None and self.pid_word not in self.preexistentes:
                try:
                    self.terminar(self.pid_word)
                except OSError:
                    pass
            else:
                try:
                    nuevos = self.inspeccionar() - self.preexistentes
                except Exception:
                    # Sin inspección fiable no se terminan procesos desconocidos.
                    return
                # Antes de abrir un documento puede faltar HWND. No se cierran
                # varias instancias cuya propiedad no se pueda distinguir.
                if len(nuevos) == 1:
                    try:
                        self.terminar(nuevos.pop())
                    except OSError:
                        pass

    def _liberar(self):
        for canal in (self.ordenes, self.resultados):
            if canal is not None:
                canal.close()
                canal.cancel_join_thread()
        self.proceso = None
        self.ordenes = self.resultados = None
        self.pid_word = None
        self.documentos = 0
        self.lista = False

    def cerrar(self):
        if self.proceso is None:
            return
        self.ordenes.put(None)
        self.proceso.join(timeout=5)
        if self.proceso.is_alive():
            self._abortar()
        else:
            self._liberar()
