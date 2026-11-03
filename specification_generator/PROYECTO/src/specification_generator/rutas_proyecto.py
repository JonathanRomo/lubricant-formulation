"""Rutas del proyecto: independientes del directorio de trabajo y portátiles al guardar."""
from __future__ import annotations

import json
import io
import os
import shutil
import sys
from functools import lru_cache
from pathlib import Path
from pathlib import PurePosixPath
from zipfile import ZipFile


def base_programa() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    origen = Path(__file__).resolve()
    # En el layout src, los datos operativos viven junto a pyproject.toml.
    # Si el modulo se copia o distribuye de forma aislada, se conserva el
    # comportamiento historico de usar la carpeta que contiene el archivo.
    for carpeta in origen.parents:
        if (carpeta / "pyproject.toml").is_file():
            return carpeta
    return origen.parent


BASE = base_programa()
CARPETAS = {
    "especificaciones_viejas": "ESPECIFICACIONES_VIEJAS",
    "json_anteriores": "ESPECIFICACIONES_JSON",
    "json_alineadas": "ESPECIFICACIONES_JSON_ALINEADAS",
    "formatos": "FORMATOS",
    "historico": "HISTORICO_ESPECIFICACIONES",
    "metodos_lab": "METODOS_LAB",
}


def resolver_ruta(ruta, base=BASE) -> Path:
    """Las rutas relativas se interpretan respecto al programa, nunca respecto a cwd."""
    base = Path(base).expanduser()
    if not base.is_absolute():
        base = BASE / base
    ruta = Path(ruta).expanduser()
    return (ruta if ruta.is_absolute() else base / ruta).resolve()


def ruta_portatil(ruta, base=BASE) -> str:
    """Guarda recursos internos como relativos; conserva ubicaciones externas explícitas."""
    base = resolver_ruta(base)
    ruta = resolver_ruta(ruta, base)
    try:
        return ruta.relative_to(base).as_posix()
    except ValueError:
        return str(ruta)


def _leer_configuracion(archivo, firma):
    try:
        datos = json.loads(Path(archivo).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    if not isinstance(datos, dict):
        raise ValueError("La configuración de carpetas debe ser un objeto JSON.")
    return datos


@lru_cache(maxsize=16)
def _rutas_resueltas(base, archivo, firma):
    datos = _leer_configuracion(archivo, firma)
    base = Path(base)
    return {
        clave: reubicar_ruta_heredada(datos[clave], base)
        if isinstance(datos.get(clave), str) and datos[clave].strip()
        else base / carpeta
        for clave, carpeta in CARPETAS.items()
    }


def rutas_configuradas(base=BASE, *, archivo=None, estricto=False) -> dict[str, Path]:
    base = resolver_ruta(base)
    archivo = resolver_ruta(archivo, base) if archivo is not None else base / "configuracion_interfaz.json"
    try:
        info = archivo.stat()
        firma = (info.st_mtime_ns, info.st_size)
    except OSError:
        firma = None
    try:
        return dict(_rutas_resueltas(str(base), str(archivo), firma))
    except (OSError, ValueError):
        if estricto:
            raise
        return {clave: base / carpeta for clave, carpeta in CARPETAS.items()}


def reubicar_ruta_heredada(ruta, base=BASE) -> Path:
    """Recupera referencias antiguas a recursos existentes tras mover la carpeta madre."""
    base = resolver_ruta(base)
    original = resolver_ruta(ruta, base)
    if not Path(ruta).is_absolute() or original.exists():
        return original
    # Solo carpetas conocidas y archivos presentes: no adivinar rutas externas.
    carpetas = {nombre.casefold() for nombre in CARPETAS.values()}
    carpetas.update({"especificaciones_viejas_eng", "revisar"})
    for indice, parte in enumerate(original.parts):
        if parte.casefold() in carpetas:
            candidata = base.joinpath(*original.parts[indice:])
            if candidata.exists():
                return candidata.resolve()
    if original.name == "CODIGOS_200.xlsx" and (base / original.name).is_file():
        return base / original.name
    return original


def carpeta_listas(base=BASE) -> Path:
    """Carpeta de métodos y traducciones según la configuración compartida."""
    return rutas_configuradas(base)["metodos_lab"]


def ruta_maestro(base=BASE) -> Path:
    return carpeta_listas(base) / "METODOS_LAB_UNIFICADO.xlsx"


def ruta_contenida(ruta, carpeta):
    """Restringe rutas derivadas de datos; no limita carpetas elegidas por el usuario."""
    carpeta = Path(carpeta).resolve()
    destino = resolver_ruta(ruta, carpeta)
    if destino == carpeta or not destino.is_relative_to(carpeta):
        raise ValueError("La ruta del documento sale de la carpeta autorizada.")
    relativos = destino.relative_to(carpeta).parts
    if any(':' in p or p.rstrip(' .') != p or Path(p).is_reserved() for p in relativos):
        raise ValueError('El nombre del archivo no es un destino permitido.')
    return destino


def referencia_documento(ruta, carpeta, permitidas=()):
    for raiz in (carpeta, *permitidas):
        destino = resolver_ruta(ruta, carpeta)
        if destino != Path(raiz).resolve() and destino.is_relative_to(Path(raiz).resolve()):
            if destino.suffix.lower() != '.json':
                break
            return ruta_contenida(destino, raiz)
    raise ValueError("La referencia de idioma no corresponde a un JSON en una carpeta autorizada.")


def ejecutable_conocido(nombre):
    import tempfile
    import time
    firma = (os.environ.get('PATH', ''), os.getcwd(), tempfile.gettempdir(),
             *(os.environ.get(k, '') for k in ('ProgramFiles', 'ProgramFiles(x86)', 'ProgramW6432', 'SystemRoot')),
             int(time.monotonic() // 30))
    resultado = _buscar_ejecutable_conocido(nombre, firma)
    return resultado if resultado and Path(resultado).is_file() else None


@lru_cache(maxsize=16)
def _buscar_ejecutable_conocido(nombre, firma):
    """No buscar programas auxiliares en cwd, temporales ni la distribucion."""
    if os.name == 'nt' and nombre.lower() == 'powershell.exe':
        ruta = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
        return str(ruta.resolve()) if ruta.is_file() else None
    candidatos = [os.environ.get(k) for k in ('ProgramFiles', 'ProgramFiles(x86)', 'ProgramW6432')]
    if nombre in {'soffice', 'libreoffice'}:
        for raiz in candidatos:
            if raiz:
                ruta = Path(raiz) / 'LibreOffice/program/soffice.exe'
                if ruta.is_file():
                    return str(ruta.resolve())
    if nombre == 'tesseract':
        for raiz in candidatos:
            if raiz:
                ruta = Path(raiz) / 'Tesseract-OCR/tesseract.exe'
                if ruta.is_file():
                    return str(ruta.resolve())
    # Compatibilidad con instalaciones explicitas fuera de Program Files:
    # solo entradas absolutas de PATH, nunca la busqueda implicita en cwd.
    import tempfile
    excluidas = (Path.cwd().resolve(), BASE.resolve(), Path(tempfile.gettempdir()).resolve())
    for entrada in os.get_exec_path():
        if not entrada or not Path(entrada).is_absolute():
            continue
        directorio = Path(entrada).resolve()
        if any(directorio == r or directorio.is_relative_to(r) for r in excluidas):
            continue
        encontrada = shutil.which(nombre, path=str(directorio))
        if encontrada:
            archivo = Path(encontrada).resolve()
            if not any(archivo.is_relative_to(r) for r in excluidas):
                return str(archivo)
    return None


def validar_archivo_office(origen):
    """Inspeccion previa sin descomprimir: 128 MiB fisicos y 256 MiB expandidos."""
    if isinstance(origen, (bytes, bytearray)):
        if len(origen) > 128 * 1024**2:
            raise ValueError('El documento supera el limite de 128 MiB.')
        _validar_paquete_office(io.BytesIO(origen))
        return
    ruta = Path(origen)
    info = ruta.stat()
    if info.st_size > 128 * 1024**2:
        raise ValueError('El documento supera el limite de 128 MiB.')
    if ruta.suffix.lower() in {'.docx', '.docm', '.xlsx', '.xlsm'}:
        _validar_paquete_office(ruta)


def _validar_paquete_office(origen):
    with ZipFile(origen) as archivo:
        entradas = archivo.infolist()
        if len(entradas) > 4096:
            raise ValueError('El documento contiene demasiadas partes internas.')
        total = 0
        nombres = set()
        for parte in entradas:
            nombre = parte.filename.replace('\\', '/')
            componentes = PurePosixPath(nombre)
            if (componentes.is_absolute() or '..' in componentes.parts or ':' in nombre
                    or nombre in nombres or (parte.external_attr >> 16) & 0o170000 == 0o120000
                    or parte.flag_bits & 1):
                raise ValueError('El documento contiene rutas, enlaces o partes no permitidas.')
            nombres.add(nombre)
            total += parte.file_size
            if (parte.file_size > 64 * 1024**2 or total > 256 * 1024**2
                    or (parte.file_size > 8 * 1024**2 and parte.file_size > max(1, parte.compress_size) * 1000)):
                raise ValueError('El contenido expandido del documento supera los limites de seguridad.')
