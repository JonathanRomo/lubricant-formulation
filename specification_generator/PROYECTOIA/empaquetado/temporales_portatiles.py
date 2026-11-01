"""Temporales portatiles con limpieza al salir y recuperacion entre sesiones."""
import atexit
import os
import shutil
import sys
import tempfile
from pathlib import Path


def _bloquear(archivo):
    archivo.seek(0)
    if os.name == 'nt':
        import msvcrt
        msvcrt.locking(archivo.fileno(), msvcrt.LK_NBLCK, 1)
    else:
        import fcntl
        fcntl.flock(archivo.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)


class TemporalesPortatiles:
    """Mantiene un bloqueo del sistema operativo mientras la sesion esta viva."""

    def __init__(self, raiz):
        self.raiz = Path(raiz)
        self.raiz.mkdir(exist_ok=True, parents=True)
        self.carpeta = None
        self._candado = None
        # Serializar la creacion y la inspeccion evita borrar sesiones nacientes.
        with self._abrir_candado(self.raiz / '.limpieza.lock') as coordinador:
            import time
            while True:
                try:
                    _bloquear(coordinador)
                    break
                except OSError:
                    time.sleep(0.02)
            self._limpiar_anteriores()
            self.carpeta = Path(tempfile.mkdtemp(prefix='sesion_', dir=self.raiz))
            self._candado = self._abrir_candado(self.carpeta / '.activa.lock')
            _bloquear(self._candado)

    @staticmethod
    def _abrir_candado(ruta):
        archivo = open(ruta, 'a+b')
        if archivo.seek(0, 2) == 0:
            archivo.write(b'0')
            archivo.flush()
        return archivo

    def _limpiar_anteriores(self):
        for carpeta in self.raiz.glob('sesion_*'):
            # No seguir enlaces ni puntos de union fuera de TEMPORAL.
            if (not carpeta.is_dir() or carpeta.is_symlink()
                    or getattr(carpeta, 'is_junction', lambda: False)()):
                continue
            if carpeta.resolve().parent != self.raiz.resolve():
                continue
            try:
                with self._abrir_candado(carpeta / '.activa.lock') as candado:
                    _bloquear(candado)
                shutil.rmtree(carpeta)
            except OSError:
                # Otra instancia activa o un archivo ocupado: reintentar al iniciar.
                continue

    def cerrar(self):
        if self._candado is not None:
            self._candado.close()
            self._candado = None
        if self.carpeta is not None:
            shutil.rmtree(self.carpeta, ignore_errors=True)


def iniciar():
    sesion = TemporalesPortatiles(Path(sys.executable).resolve().parent / 'TEMPORAL')
    tempfile.tempdir = str(sesion.carpeta)
    for variable in ('TEMP', 'TMP', 'TMPDIR'):
        os.environ[variable] = str(sesion.carpeta)
    atexit.register(sesion.cerrar)
    return sesion


# PyInstaller ejecuta los runtime hooks; importar el modulo en pruebas no debe
# cambiar TEMP ni crear carpetas junto al interprete de desarrollo.
if getattr(sys, 'frozen', False):
    _sesion_temporal = iniciar()
