import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from empaquetado.temporales_portatiles import TemporalesPortatiles


class LimpiezaTemporalesTests(unittest.TestCase):
    def test_limpia_restos_y_respeta_sesiones_activas_y_otras_carpetas(self):
        with tempfile.TemporaryDirectory() as td:
            raiz = Path(td)
            vieja = raiz / 'sesion_anterior'
            vieja.mkdir()
            (vieja / 'conversion.docx').write_bytes(b'contenido')
            ajena = raiz / 'documentos'
            ajena.mkdir()
            primera = TemporalesPortatiles(raiz)
            try:
                self.assertFalse(vieja.exists())
                segunda = TemporalesPortatiles(raiz)
                try:
                    self.assertTrue(primera.carpeta.is_dir())
                    self.assertTrue(segunda.carpeta.is_dir())
                    self.assertTrue(ajena.is_dir())
                finally:
                    segunda.cerrar()
                self.assertFalse(segunda.carpeta.exists())
            finally:
                primera.cerrar()
            self.assertFalse(primera.carpeta.exists())

    def test_salida_normal_y_recuperacion_de_cierre_forzado(self):
        with tempfile.TemporaryDirectory() as td:
            codigo = '''
import atexit, os, sys
from empaquetado.temporales_portatiles import TemporalesPortatiles
s = TemporalesPortatiles(sys.argv[1])
atexit.register(s.cerrar)
(s.carpeta / 'conversion.docx').write_bytes(b'contenido')
if sys.argv[2] == 'forzado':
    os._exit(0)
'''
            for modo in ('normal', 'forzado'):
                subprocess.run([sys.executable, '-c', codigo, td, modo], check=True)
                restos = list(Path(td).glob('sesion_*'))
                self.assertEqual(len(restos), int(modo == 'forzado'))
            nueva = TemporalesPortatiles(td)
            try:
                self.assertFalse(restos[0].exists())
            finally:
                nueva.cerrar()

    def test_protege_otra_instancia_en_otro_proceso(self):
        with tempfile.TemporaryDirectory() as td:
            codigo = """
import sys
from empaquetado.temporales_portatiles import TemporalesPortatiles
s = TemporalesPortatiles(sys.argv[1])
print(s.carpeta, flush=True)
sys.stdin.readline()
s.cerrar()
"""
            proceso = subprocess.Popen(
                [sys.executable, '-c', codigo, td],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
            try:
                activa = Path(proceso.stdout.readline().strip())
                self.assertTrue(activa.is_dir())
                nueva = TemporalesPortatiles(td)
                try:
                    self.assertTrue(activa.is_dir())
                finally:
                    nueva.cerrar()
                proceso.communicate('cerrar\n', timeout=10)
                self.assertEqual(proceso.returncode, 0)
                self.assertFalse(activa.exists())
            finally:
                if proceso.poll() is None:
                    proceso.kill()
                    proceso.communicate()

    @unittest.skipUnless(hasattr(os, 'symlink'), 'Requiere enlaces simbolicos')
    def test_no_sigue_enlaces_a_datos_externos(self):
        with tempfile.TemporaryDirectory() as td:
            raiz = Path(td) / 'TEMPORAL'
            raiz.mkdir()
            datos = Path(td) / 'datos'
            datos.mkdir()
            (datos / 'original.txt').write_text('conservar')
            try:
                (raiz / 'sesion_enlace').symlink_to(datos, target_is_directory=True)
            except OSError:
                self.skipTest('El equipo no permite crear enlaces simbolicos')
            nueva = TemporalesPortatiles(raiz)
            try:
                self.assertTrue((datos / 'original.txt').exists())
            finally:
                nueva.cerrar()


if __name__ == '__main__':
    unittest.main()
