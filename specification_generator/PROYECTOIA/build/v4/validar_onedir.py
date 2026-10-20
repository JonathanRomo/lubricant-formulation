import ctypes
import json
import shutil
import subprocess
import tempfile
import time
from ctypes import wintypes
from pathlib import Path


base = Path(__file__).resolve().parents[2]
raiz = base / "build" / "v4"
paquete = base / "dist" / "v4" / "AplicacionEspecificaciones_v4"
usuario = ctypes.windll.user32
callback = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
usuario.GetWindowTextLengthW.argtypes = [wintypes.HWND]
usuario.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
usuario.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
usuario.IsHungAppWindow.argtypes = [wintypes.HWND]
usuario.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]


def texto(hwnd):
    contenido = ctypes.create_unicode_buffer(usuario.GetWindowTextLengthW(hwnd) + 1)
    usuario.GetWindowTextW(hwnd, contenido, len(contenido))
    return contenido.value


def ventanas(pid):
    resultado = []
    @callback
    def ver(hwnd, parametro):
        identidad = wintypes.DWORD()
        usuario.GetWindowThreadProcessId(hwnd, ctypes.byref(identidad))
        if identidad.value == pid:
            usuario.ShowWindow(hwnd, 0)
            titulo = texto(hwnd)
            if titulo:
                resultado.append((hwnd, titulo))
        return True
    usuario.EnumWindows(ver, 0)
    return resultado


def comprobar_dialogos(actuales):
    for hwnd, titulo in actuales:
        if titulo in {"Unhandled exception in script", "No se pudo reconstruir la base"}:
            detalles = []
            @callback
            def hijo(ventana, parametro):
                contenido = texto(ventana)
                if contenido:
                    detalles.append(contenido)
                return True
            usuario.EnumChildWindows(hwnd, hijo, 0)
            raise RuntimeError(" / ".join(detalles))


startup = subprocess.STARTUPINFO()
startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
startup.wShowWindow = 0
with tempfile.TemporaryDirectory(prefix="prueba_onedir_", dir=raiz) as carpeta:
    temporal = Path(carpeta).resolve()
    assert temporal.parent == raiz.resolve()
    destino = temporal / paquete.name
    shutil.copytree(paquete, destino)
    ejecutable = destino / (paquete.name + ".exe")
    print("Copia aislada preparada.", flush=True)
    for fase, argumentos in (("importacion", ["--help"]), ("interfaz", [])):
        proceso = subprocess.Popen([str(ejecutable), *argumentos], cwd=temporal, startupinfo=startup,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            limite = time.monotonic() + 30
            while time.monotonic() < limite:
                actuales = ventanas(proceso.pid)
                comprobar_dialogos(actuales)
                if fase == "importacion" and proceso.poll() is not None:
                    assert proceso.returncode == 0, proceso.returncode
                    break
                if fase == "interfaz":
                    assert proceso.poll() is None, proceso.returncode
                    principal = next((h for h, t in actuales if t == "Gestor de especificaciones"), None)
                    if principal:
                        time.sleep(2)
                        comprobar_dialogos(ventanas(proceso.pid))
                        assert not usuario.IsHungAppWindow(principal)
                        # El primer arranque informa que creó la base vacía.
                        # Aceptar ese aviso en esta copia de prueba antes de cerrar.
                        for _ in range(50):
                            abiertas = ventanas(proceso.pid)
                            comprobar_dialogos(abiertas)
                            for ventana, titulo in abiertas:
                                if titulo == "Base reconstruida":
                                    usuario.PostMessageW(ventana, 0x0111, 1, 0)
                            if not any(t in {"Recuperar base de datos", "Base reconstruida"} for _, t in abiertas):
                                break
                            time.sleep(.1)
                        # La comprobación termina aquí; el proceso de la copia
                        # se retira en finally sin interactuar con datos del usuario.
                        break
                time.sleep(.15)
            else:
                raise TimeoutError(f"No terminó la prueba: {fase}")
            print(f"{fase}: correcta", flush=True)
        finally:
            if proceso.poll() is None:
                proceso.terminate()
                proceso.wait(timeout=10)
    print(json.dumps({"arranque_portatil": "correcto", "ventana": "responde", "proceso_de_prueba": "retirado"}), flush=True)
