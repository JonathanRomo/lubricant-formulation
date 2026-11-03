# Compilación reproducible onedir sin datos de trabajo ni archivos JSON.
import json
import shutil
import sys
from pathlib import Path

from PyInstaller.config import CONF
from PyInstaller.utils.hooks import collect_data_files, get_package_paths
from PyInstaller.utils.hooks.tcl_tk import tcltk_info


base = Path(SPECPATH).resolve().parent
nombre = "AplicacionEspecificaciones_v4"
documentos = (
    "RESUMEN_ALINEADOR_ESPECIFICACIONES.md",
    "README_ACTUALIZACION_CODIGOS_200.md",
    "README_INTERFAZ.md",
    "README_GENERADOR_HII.md",
)
datos = collect_data_files("customtkinter", excludes=["**/*.json", "**/.DS_Store"])
binarios = []

# Algunas instalaciones no inicializan Tcl en el proceso de análisis. Recoger
# sus recursos desde la instalación local, sin depender de una distribución vieja.
if not tcltk_info.available:
    tcl = Path(sys.base_prefix) / "tcl"
    for carpeta, destino in (("tcl8.6", "_tcl_data"), ("tk8.6", "_tk_data"), ("tcl8", "tcl8")):
        origen = tcl / carpeta
        if not origen.is_dir():
            raise SystemExit(f"Faltan los recursos Tcl/Tk: {origen}")
        datos.append((str(origen), destino))
    for dll in ("tcl86t.dll", "tk86t.dll"):
        binarios.append((str(Path(sys.base_prefix) / "DLLs" / dll), "."))

a = Analysis(
    [str(base / "interfaz_aplicacion_especificaciones.py")],
    pathex=[str(base)],
    binaries=binarios,
    datas=datos,
    hiddenimports=[p.stem for p in base.glob("*.py")],
    hookspath=[str(base / "empaquetado" / "hooks")],
    hooksconfig={},
    runtime_hooks=[str(base / "empaquetado" / "temporales_portatiles.py")],
    excludes=["pytest"],
    noarchive=False,
    optimize=0,
)

modulos_necesarios = {"tkinter", "tkinter.font", "tkinter.ttk", "tkinter.scrolledtext", "optimizar_base_datos"}
faltantes = modulos_necesarios - {fila[0] for fila in a.pure}
if faltantes:
    raise SystemExit("Faltan módulos necesarios: " + ", ".join(sorted(faltantes)))

# CustomTkinter carga sus temas JSON al importar. Incorporar esos diccionarios
# al módulo compilado permite distribuir la interfaz sin archivos .json sueltos.
# La instalación del entorno virtual permanece intacta.
_, paquete_ctk = get_package_paths("customtkinter")
paquete_ctk = Path(paquete_ctk)
temas = {p.stem: json.loads(p.read_text(encoding="utf-8"))
         for p in (paquete_ctk / "assets" / "themes").glob("*.json")}
modulo_tema = "customtkinter.windows.widgets.theme.theme_manager"
fuente_tema = paquete_ctk / "windows" / "widgets" / "theme" / "theme_manager.py"
fuente = fuente_tema.read_text(encoding="utf-8")
anterior = ('with open(os.path.join(customtkinter_path, "assets", "themes", '
            'f"{theme_name_or_path}.json"), "r") as f:\n                cls.theme = json.load(f)')
if fuente.count(anterior) != 1 or "blue" not in temas:
    raise SystemExit("Cambió la carga de temas de CustomTkinter; revisar el empaquetado.")
fuente = fuente.replace(anterior, f'cls.theme = __import__("copy").deepcopy({temas!r}[theme_name_or_path])')
CONF["code_cache"][id(a.pure)][modulo_tema] = compile(fuente, str(fuente_tema), "exec")

extensiones_excluidas = {".json", ".db", ".sqlite", ".sqlite3", ".bat"}
def permitido(destino):
    ruta = Path(destino)
    archivo = ruta.name.casefold()
    return (ruta.suffix.casefold() not in extensiones_excluidas
            and not archivo.endswith((".db-shm", ".db-wal", ".bd-shm", ".bd-wal", ".db-journal"))
            and archivo not in {".gitignore", ".ds_store", "codigos_200.xlsx"}
            and not ({p.casefold() for p in ruta.parts} & {"tests", "__pycache__"}))

a.datas = [fila for fila in a.datas if permitido(fila[0])]
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [], exclude_binaries=True, name=nombre,
    debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
    console=False, disable_windowed_traceback=False, contents_directory="_internal",
)
carpeta = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name=nombre)
salida = Path(CONF["distpath"]) / nombre
for documento in documentos:
    shutil.copy2(base / documento, salida / documento)
shutil.copy2(base / "empaquetado" / "LEEME_PORTATIL_V4.txt", salida / "LEEME_PORTATIL.txt")

# Auditar recursivamente el producto final, incluidos los recursos de bibliotecas.
prohibidos = [str(p.relative_to(salida)) for p in salida.rglob("*") if p.is_file() and not permitido(p.relative_to(salida))]
if prohibidos:
    raise SystemExit("El paquete contiene archivos excluidos: " + ", ".join(prohibidos))
