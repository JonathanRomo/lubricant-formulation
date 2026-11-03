# Gestor de especificaciones

Aplicacion de escritorio para extraer, editar, alinear y publicar especificaciones tecnicas bilingues. Incluye el editor de especificaciones, el alineador de metodos, el selector, la generacion de documentos Word y HII, y herramientas de mantenimiento del catalogo.

## Requisitos

- Windows 10 u 11.
- Python 3.11.
- Microsoft Word y Excel para las funciones que automatizan Office.

## Instalacion

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Ejecucion

La forma mas sencilla en Windows es abrir `EJECUTAR_APLICACION_ESPECIFICACIONES.bat`.

Desde PowerShell:

```powershell
$env:PYTHONPATH = "$PWD\src"
.\.venv\Scripts\python.exe -m specification_generator
```

## Pruebas

```powershell
$env:PYTHONPATH = "$PWD\src;$PWD\src\specification_generator;$PWD\tests"
.\.venv\Scripts\python.exe -m unittest discover -s tests -p "test_*.py"
```

## Estructura

```text
PROYECTO/
|-- src/specification_generator/   Codigo de la aplicacion
|-- tests/                         Pruebas automatizadas
|-- docs/                          Documentacion tecnica y funcional
|-- empaquetado/                   Configuracion del ejecutable portable
|-- logs/                          Registros locales (no versionados)
|-- pyproject.toml                 Metadatos y herramientas Python
|-- requirements.txt              Dependencias de ejecucion
`-- EJECUTAR_APLICACION_ESPECIFICACIONES.bat
```

Los directorios de JSON, formatos, metodos y bases de datos son datos operativos locales y no se publican en Git. Consulta [docs/arquitectura.md](docs/arquitectura.md) para conocer las responsabilidades de cada modulo y el plan de separacion por capas.
