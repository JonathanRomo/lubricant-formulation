# Gestor v0.5.2

## Formatos Word compatibles

El gestor busca dentro de `ESPECIFICACIONES_VIEJAS` y sus subcarpetas:

- `.docx`: lectura directa.
- `.doc`: conversión temporal de solo lectura a `.docx`.
- `.docm`: conversión temporal de solo lectura a `.docx`.

Los archivos originales nunca se modifican ni se sobrescriben. Para `.doc` y `.docm`, la conversión ocurre en una carpeta temporal que se elimina al terminar. En Windows se utiliza Microsoft Word mediante `pywin32`; si no está disponible, se intenta usar LibreOffice.

`CODIGOS_200.xlsx` también permanece estrictamente en modo de solo lectura. Toda la información derivada se almacena en `especificaciones.db`.

## Estructura

```text
C:\Users\JonathanRomo\Music\PROYECTOIA\
├── gestor_especificaciones.py
├── requirements.txt
├── CODIGOS_200.xlsx
├── especificaciones.db
├── ESPECIFICACIONES_VIEJAS\
└── ESPECIFICACIONES_JSON\
```

## Instalación y ejecución con Python 3.11

```powershell
cd "C:\Users\JonathanRomo\Music\PROYECTOIA"
py -3.11 -m venv ".venv"
.\.venv\Scripts\python.exe -m pip install -r ".\requirements.txt"
.\.venv\Scripts\python.exe ".\gestor_especificaciones.py"
```

Para consultar el registro:

```powershell
.\.venv\Scripts\python.exe ".\gestor_especificaciones.py" --ver-registro
```

## Unificador de métodos de laboratorio

El unificador también requiere Python 3.11. Instala su dependencia y ejecútalo
desde la carpeta del proyecto:

```powershell
.\.venv\Scripts\python.exe -m pip install -r ".\requirements_metodos_lab.txt"
.\.venv\Scripts\python.exe ".\unificar_metodos_lab_v3.py" --base "."
```

También puedes ejecutar `EJECUTAR_UNIFICAR_METODOS_LAB.bat`. Este crea el
entorno virtual `.venv` si hace falta e instala allí `openpyxl`, sin modificar
la instalación de Python administrada por uv.

No ejecutes `pip install` directamente sobre el Python administrado por uv. Si
Python 3.11 no está instalado, instálalo y luego crea el entorno con
`py -3.11 -m venv .venv`.
