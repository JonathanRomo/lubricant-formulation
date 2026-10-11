# Gestor v0.8.4

## Formatos Word compatibles

El gestor busca dentro de `ESPECIFICACIONES_VIEJAS` y
`ESPECIFICACIONES_VIEJAS_ENG`, incluidas sus subcarpetas:

- `.docx`: lectura directa.
- `.doc`: conversión temporal de solo lectura a `.docx`.
- `.docm`: conversión temporal de solo lectura a `.docx`.

Los originales `.doc` y `.docm` no se modifican: se convierten en una carpeta
temporal que se elimina al terminar. En Windows se utiliza Microsoft Word
mediante `pywin32`; si no está disponible, se intenta usar LibreOffice. Los
documentos ingleses detectados dentro de `ESPECIFICACIONES_VIEJAS` se mueven a
`ESPECIFICACIONES_VIEJAS_ENG` antes de procesarse.

Los documentos de `ESPECIFICACIONES_VIEJAS_ENG` también se extraen al mismo
esquema JSON. Las etiquetas inglesas de columnas y control documental se
normalizan a los mismos campos; el idioma del JSON se marca como `en`, aunque se
haya usado una plantilla visual en español.

`CODIGOS_200.xlsx` también permanece estrictamente en modo de solo lectura. Toda la información derivada se almacena en `especificaciones.db`.

## Documentos para revisión manual

Antes de extraer la especificación, el gestor busca las palabras completas
`Yves` y `YOI` en el contenido Word, sin distinguir mayúsculas y minúsculas.
Si encuentra alguna, mueve el documento a `REVISAR/`, registra
`REVISAR_YVES_YOI` en la base de datos y continúa con los demás archivos sin
generar un JSON para ese documento.

## Numeración cronológica de revisiones

`revision_json` es la revisión normalizada que asigna el gestor por producto:
la especificación con la fecha documental más antigua queda como `00`, y las
siguientes se numeran `01`, `02`, etc. Se reconocen fechas ISO, numéricas y con
meses escritos en español o inglés. En fechas numéricas ambiguas, como
`05/06/2022`, se interpreta día/mes/año. Si falta una fecha reconocible, esa
versión queda después de las que sí tienen fecha y conserva un orden estable.
Cuando varias versiones tienen la misma fecha, también se conserva su orden
previo porque el documento no permite distinguir cuál fue primero.
Cuando solo se conoce el mes y el año, se toma el primer día de ese mes.

El valor original declarado por el autor se conserva por separado en
`control_documental.revision` y en `revision_especificacion`; no se usa para
decidir la numeración normalizada. Si el recuadro de revisión contiene
solamente `-`, `--` o `---`, el gestor sustituye ese marcador por el número
cronológico en `control_documental.revision` y conserva el valor escrito en
`control_documental.revision_fuente`; `revision_especificacion` en SQLite
mantiene el original. El nombre del JSON también usa la revisión cronológica,
no el marcador de guiones. Al ejecutar el gestor, se normalizan estos nombres,
se reordena el historial ya guardado en `especificaciones.db` y se actualiza
el campo `revision_json` de cada JSON.

## Resolución del nombre y del código de producto

Antes de clasificar un producto, el gestor busca descripciones del catálogo en
todo el contenido del JSON, sin distinguir mayúsculas, acentos ni signos de
puntuación y respetando los límites de las palabras. Entre las coincidencias
elige la descripción más específica (la de mayor longitud); así, por ejemplo,
`SCHUTZ PRO HVI 40` no coincide con `SCHUTZ PRO HVI 400-10`.

Si esa descripción tiene varios IDs y exactamente uno pertenece a la familia
de códigos 200 (`2` seguido de siete dígitos), ese ID se usa como código
canónico. Un código de catálogo ya válido se conserva; no se cambia un ID
existente de la familia `8` solo porque haya otro ID para la misma descripción.
El JSON guarda el nombre y el código canónicos en sus campos normales, y
conserva el dato original en `producto.nombre_fuente` y
`control_documental.codigo_unico_producto_fuente` cuando difiere. La base de
datos también registra esos valores originales, el código final, el método de
resolución y el estado de validación. Al ejecutar el gestor se actualizan con
esta regla los JSON históricos vinculados a la base de datos y se reagrupan
sus revisiones por el nombre canónico. Si hay JSON históricos ilegibles, se
evita reagrupar nombres para no fusionar historiales incompletos.

Un **nombre ambiguo** es uno que no puede asignarse a un producto con evidencia
única: quedan varias descripciones igualmente específicas, hay más de un ID
200 para el nombre coincidente, o no existe un ID 200 y quedan varios IDs de
catálogo posibles. En esos casos no se elige un código arbitrariamente: el
JSON queda marcado `CONFLICTO_NOMBRE_AMBIGUO` y el archivo mantiene el prefijo
de revisión `--- `. Un Part ID de catálogo válido que contradiga el nombre se
marca por separado como `CONFLICTO_NOMBRE_PART_ID`.

El nombre del JSON usa el código final si está reconocido en `CODIGOS_200`,
aunque se haya recuperado buscando el nombre del producto en el contenido.
`SIN_CODIGO_200` se reserva para cuando no se pudo confirmar un código final en
el catálogo. La falta o invalidez del código escrito en el Word, por sí sola,
no convierte el documento en duplicado. Los JSON históricos con prefijo
`DUPLICADO` se renombran al código final confirmado (o a `SIN_CODIGO_200` si no
hay código final reconocido), y se actualizan las referencias en SQLite. El
prefijo `DUPLICADO` ya no se asigna automáticamente.

## Estructura

```text
C:\Users\JonathanRomo\Music\PROYECTOIA\
├── gestor_especificaciones.py
├── requirements.txt
├── CODIGOS_200.xlsx
├── especificaciones.db
├── ESPECIFICACIONES_VIEJAS\
├── ESPECIFICACIONES_VIEJAS_ENG\
└── ESPECIFICACIONES_JSON\
```

## Instalación y ejecución con Python 3.11

En Linux, puedes ejecutar el lanzador desde cualquier carpeta. Siempre usa
`.venv/bin/python` y crea el entorno e instala los requisitos si hacen falta:

```bash
/home/jonathan/Documents/lubricant-formulation/specification_generator/PROYECTOIA/EJECUTAR_GESTOR.sh
```

En VS Code, abre `PROYECTOIA` como carpeta (no solamente `Documents`) para que
se apliquen los ajustes locales de `.vscode`.

En Windows:

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
