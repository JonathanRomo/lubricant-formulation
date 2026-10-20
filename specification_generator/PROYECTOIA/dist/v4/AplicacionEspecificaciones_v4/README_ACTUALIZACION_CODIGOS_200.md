# Gestor v0.8.7

Guía actualizada el 19 de octubre de 2026. Interfaz 1.1.4 y alineador 2.1.35.

## Inicio portátil

Extrae la distribución completa y ejecuta `AplicacionEspecificaciones_v4.exe`.
Mantén la carpeta `_internal` junto al ejecutable; no necesitas instalar Python
ni ejecutar un BAT. Selecciona tus carpetas en **Configuración** y proporciona
`CODIGOS_200.xlsx` junto al ejecutable cuando necesites extraer documentos Word.
Los formatos y las listas maestras se administran desde **Formatos** y **Métodos**.

El paquete se distribuye sin carpetas de trabajo, catálogo `CODIGOS_200.xlsx`,
bases SQLite ni sus diarios, archivos JSON, BAT o `.gitignore`. La configuración,
las preferencias y la base se crean o recuperan durante el uso. Los temporales
del programa se alojan en `TEMPORAL/sesion_*`, junto al ejecutable, y se limpian
al terminar normalmente. Word y otras herramientas externas pueden crear sus
propios temporales.

En **Instrucciones**, la guía del alineador es la selección inicial; también
puedes consultar este documento y `README_INTERFAZ.md`.

## Datos que conserva la base

`archivos` y `versiones` mantienen las extracciones y su historial necesario
para detectar cambios y evitar reprocesar documentos. La caché de códigos
conserva una sola edición del catálogo, no una copia completa por actualización.
Los estados y firmas auxiliares del gestor evitan revisar JSON sin cambios.

Los catálogos y relaciones HII se calculan en memoria. Los vínculos confirmados
y el historial de Excel permanecen en los JSON. Las tablas antiguas `hii_*` y
`recuperacion_documentos` se retiran con un respaldo previo, rescatando primero
los registros que solo estaban en SQLite. La reconstrucción no vuelve a crearlas.

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

Durante una carga se reutiliza una instancia propia de Word, iniciada solo si
hay documentos que convertir. Un proceso supervisor mantiene el límite de
tiempo por documento y reinicia el trabajador tras un error. La sesión se
recicla cada 50 conversiones y se cierra al terminar la carga. Las instancias
de Word que ya estaban abiertas quedan protegidas.

## Actualización del lector y caché

`VERSION_LECTOR` identifica las reglas de extracción, independientemente de
`APP_VERSION` y de `VERSION_NORMALIZACION`. La caché comprueba la huella del
Word, la del catálogo, la versión del lector y la existencia del JSON.

Si el Word no cambió y hay que renovar la extracción, se actualizan el mismo
JSON y registro de versión, conservando la revisión y la fecha de creación.
Una mejora del lector o un cambio de catálogo no inventan otra revisión del
producto. Un cambio real en los bytes del Word conserva el flujo de versiones
documentales habitual.

La base guarda una instantánea de los campos extraídos para distinguir los
datos derivados de cambios manuales al actualizar. Los campos editados, listas
de pruebas modificadas y campos adicionales del JSON se conservan. En JSON
heredados sin esa instantánea, se aplica una regla conservadora: rellenar datos
ausentes o vacíos y conservar los valores existentes. Las diferencias que no
se pueden renovar con seguridad se muestran en el resumen de la carga.

Si la relectura falla, se conservan el Word y la extracción anterior. Estas
comprobaciones están en `tests/test_actualizacion_extraccion.py`; la sesión de
Word se verifica en `tests/test_sesion_word.py`. La prueba con Office real se
activa mediante `WORD_INTEGRATION=1` y usa únicamente documentos de solo lectura
y conversiones temporales.

## Documentos para revisión manual

Si Word o LibreOffice no pueden convertir un `.doc` o `.docm`, el documento
permanece en su carpeta para reintentarlo en la siguiente carga. Estos fallos
no crean entradas de error ni versiones por archivo en SQLite y no mueven los
originales a `REVISAR`. El resumen de la carga muestra cuántos documentos
quedaron pendientes de conversión. El inicio de Word se reintenta una sola
vez; si vuelve a fallar, se evita reiniciarlo para cada archivo del lote.
Los `.docx` continúan procesándose y los errores de estructura detectados
después de la conversión siguen enviándose a revisión.

Las comprobaciones de estos casos están en
`tests/test_conversion_pendiente.py`, usan datos sintéticos y carpetas
temporales, y no generan informes, inventarios ni muestras en la aplicación.

Antes de extraer la especificación, el gestor busca las palabras completas
`Yves` y `YOI` en el contenido Word, sin distinguir mayúsculas y minúsculas.
Si encuentra alguna, mueve el documento a `REVISAR/`, registra
`REVISAR_YVES_YOI` en la base de datos y continúa con los demás archivos sin
generar un JSON para ese documento.

## Lectura de formatos antiguos

El extractor usa `lectura_estructuras_word.py` para reconocer los encabezados
en cualquier fila y dentro de tablas anidadas. Busca el título también en
celdas y cuadros de texto, y contrasta los candidatos con el catálogo. Admite
etiquetas como `PRUEBAS`, `PREUBA`, `TEST DESCRIPTION`, `ASTM METHOD`,
`TYPICAL RESULTS`, `Elaboró` y nombres de cliente con apóstrofes.

Las notas fuera de las tablas, los símbolos de característica crítica y los
subtítulos de análisis se conservan en `notas`. Las filas `NOTAS` y los códigos
de producto incrustados no se cuentan como pruebas. La columna `TEST TYPE`,
cuando existe, se conserva en el campo opcional `tipo_prueba` de cada prueba.
Un texto como `15.86 TYPICAL` permanece literal en `especificacion`.

El registro de calidad del pie identifica el código y la revisión del formato
sin sustituir la fecha, revisión o código del producto. Los metadatos que estén
únicamente en imágenes requieren el OCR opcional ya existente. Una plantilla
vacía, una tabla sin encabezados o columnas `PRUEBA` duplicadas con contenidos
distintos requieren revisión; no se adivina su significado.

Las pruebas automáticas están en `tests/test_estructuras_word.py`. La utilidad
`tests/validar_formatos_word_reales.py` convierte el corpus real en modo de solo
lectura y compara versiones usando exclusivamente una carpeta temporal. Con
`WORD_CORPUS_TEMP` apuntando a esa carpeta se pueden ejecutar también los tests
optativos `tests/test_word_reales.py`. Estas herramientas no se ejecutan desde
la aplicación ni actualizan su base de datos.

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
AplicacionEspecificaciones_v4\
├── AplicacionEspecificaciones_v4.exe
├── _internal\
├── RESUMEN_ALINEADOR_ESPECIFICACIONES.md
├── README_ACTUALIZACION_CODIGOS_200.md
├── README_INTERFAZ.md
└── LEEME_PORTATIL.txt
```

Este es el contenido inicial de la distribución. Añade tus recursos de trabajo
cuando vayas a utilizarlos. Las rutas internas se guardan relativas a la carpeta
del programa; si eliges recursos externos, copia esos recursos también al trasladarlo.

## Instalación y ejecución con Python 3.11

Esta sección corresponde al código fuente; el ejecutable portátil ya incluye
Python y sus bibliotecas. Desde la carpeta del proyecto, instala las dependencias
en un entorno virtual y abre la interfaz:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe interfaz_aplicacion_especificaciones.py
```

Para utilizar únicamente el gestor de consola desde el código fuente:

```powershell
.venv\Scripts\python.exe gestor_especificaciones.py
```

Para consultar el registro:

```powershell
.\.venv\Scripts\python.exe ".\gestor_especificaciones.py" --ver-registro
```

## Flujo de extraccion, alineacion y generacion

El alineador lee `ESPECIFICACIONES_JSON` y `METODOS_LAB\METODOS_LAB_UNIFICADO.xlsx`
(hoja `METODOS`), y escribe los resultados en
`ESPECIFICACIONES_JSON_ALINEADAS`. Usa pandas para validar/leer las columnas del
maestro y RapidFuzz para comparar nombres; solo sustituye automaticamente una
prueba cuando supera el umbral configurado y no hay ambiguedad. Selecciona el
nombre, el método y las unidades segun `documento.idioma`; para documentos en
ingles usa `MÉTODO INGLÉS` y recurre a `METODO` solo cuando falta una traduccion.
Si falta el nombre maestro para ese idioma, conserva el original y exige
revision. Los asteriscos
al inicio de los nombres en la lista maestra son marcas de acreditacion y se
quitan del nombre canonico que se escribe en el JSON alineado; la auditoria
mantiene el nombre original de la lista maestra. En los campos de unidades y
valor tipico, los rangos numericos se escriben con espacios alrededor del
guion, los decimales usan coma y los millares se agrupan con espacio fino.
Cuando un punto o una coma seguidos por tres digitos sean ambiguos, se conserva
la interpretacion decimal para no cambiar el valor medido; los grupos de miles
se convierten cuando su estructura los identifica claramente. En todas las
celdas de contenido de la tabla se separa el numero del simbolo de unidad con
un espacio, por ejemplo `0,6 mL` y `40 °C`; tambien aplica a nombres,
metodos, especificaciones y valores tipicos. Los codigos de metodo sin unidad
reconocida, como `ASTM D445` e `ISO 4406`, se conservan. Los metodos
ASTM se comparan por designacion, ignorando espacios, guiones tipograficos,
ediciones/anios y variantes como `D445`, `D 445`, `D-445`, `ASTM D 445`,
`ASTM D-445` o `ASTM 445`. En conjugaciones detecta referencias posteriores
sin repetir `ASTM` (por ejemplo, `ASTM D445 / D7042`). Si se alinea, el metodo
ASTM se toma de la fila elegida del maestro, se elimina su edicion y se escribe
como `ASTM D445`; una conjugacion ASTM queda como `ASTM D445 / ASTM D7042`.
Las farmacopeas USP, FEUM y Ph. Eur. se comparan sin tomar en cuenta ediciones,
capitulos USP, claves MGA ni textos como "vigente"; se escriben como `USP`,
`FEUM` y `Ph Eur`. Las combinaciones se conservan con esos nombres canonicos
(por ejemplo, `USP / FEUM`), y las citas ASTM reconocidas se conservan tambien
si aparecen junto con una farmacopea. Las normas no reconocidas se omiten
cuando el maestro incluye referencias ASTM o farmacopeicas canonizables.
Designaciones ASTM distintas no se consideran equivalentes; los empates entre
metodos maestros realmente distintos o con unidades conflictivas se dejan para
revision. Si un metodo original contiene referencias ASTM ausentes de todas las
filas del maestro, el JSON alineado las informa en
`alineacion_metodos_lab.mensajes_astm_no_encontradas` y en el detalle de la
prueba, aun cuando el nombre se haya alineado y el metodo final se haya
sustituido por el canonico de la lista. Cada JSON incluye
`alineacion_metodos_lab` con los valores originales y normalizados, las
referencias ASTM y farmacopeicas detectadas, puntajes, alternativas y conteos;
los casos no alineados se detallan en `REPORTE_ALINEACION.txt`.

En la aplicación portátil, **SUBIR ESPEC** realiza la extracción y alineación
de los documentos seleccionados. Después puedes revisar los resultados en
**SELECCIONAR ESPEC** o generar documentos desde el editor y **GENERAR HII**.
Coloca las plantillas en la carpeta configurada de formatos; los documentos
generados se guardan en la ubicación que elijas.

Desde el código fuente, para ejecutar una sola etapa:

```powershell
.venv\Scripts\python.exe gestor_especificaciones.py
.venv\Scripts\python.exe alinear_especificaciones_json.py --base .
```

Estos comandos son alternativas de desarrollo y no se necesitan para abrir
el ejecutable portátil. No se distribuyen lanzadores BAT ni scripts de pipeline.

## Aplicación gráfica

En Windows, inicia `AplicacionEspecificaciones_v4.exe` para abrir el
menú principal. `SUBIR ESPEC` permite elegir varios archivos `.doc`, `.docm` y
`.docx` y copiarlos a `ESPECIFICACIONES_VIEJAS`; los formatos no compatibles se
omiten individualmente. La extracción y alineación se limitan a los documentos
seleccionados, y la base de datos permite omitir los que no cambiaron. Durante
el proceso muestra una barra lineal en un pie gris permanente; `Ver más` abre
los mensajes de copia, extracción y alineación que también se envían a la
terminal.
`GENERAR ESPEC` abre un formulario para elegir idioma, capturar control
documental y administrar las filas de pruebas antes de guardar un JSON
compatible con la vista alineada en la carpeta correspondiente. Las sugerencias
de pruebas y unidades aparecen en gris y se aceptan explícitamente; los métodos
ASTM conjugados se sugieren como una combinación completa cuando existe una
coincidencia exacta en la lista maestra. Al confirmar una prueba que coincide
con la lista maestra, se completan automáticamente el método y la unidad si se
pueden determinar sin ambigüedad. El formulario ofrece hasta cinco Part ID del
catálogo y bloquea el guardado si el código no existe. También precarga las
notas de la plantilla y «Investigación y desarrollo» (traducido al inglés
cuando corresponde).
La periodicidad es obligatoria por fila: se elige de uno a tres símbolos
distintos en desplegables separados; el JSON los concatena en un mismo valor
(por ejemplo, `QIP`) y el encabezado abre su leyenda localizada.
Ese campo se conserva en `PERIODIDICAD DE ANALISIS` dentro del JSON y no se
agrega a las columnas de la tabla Word. `SELECCIONAR
SPEC` abre el selector y permite enviar una especificación marcada al editor.
El botón «Editar» en el selector o en cada panel del comparador abre esa
especificación en el generador con sus datos cargados; al guardar, permite
confirmar la actualización del mismo JSON. Si se abrió un Word solo para
compararlo, «Editar» carga su extracción temporal en el generador para crear un
JSON persistente.
Si la traducción vinculada tiene otro código, cliente o revisión, falta o no
puede emparejarse con las pruebas, el editor abre la especificación y prepara
un borrador del otro idioma con un aviso dentro de la ventana. La versión
vinculada se conserva y la traducción nueva requiere otro archivo. Al guardar
la pareja se comprueba que coincidan código, cliente y revisión, que los
idiomas sean distintos y que no queden traducciones pendientes.
Al modificar una especificación existente y aplicar la siguiente revisión, el
editor sustituye su fecha por la fecha local del sistema (`DD/MM/YY`). Ambos
idiomas comparten la nueva revisión y fecha; abrir sin cambios conserva las originales.
Cuando faltan nombre o revisión del cliente al abrir una especificación, el
editor propone un guion que se puede borrar. Desarrollador es obligatorio al
guardar tanto en progreso como terminado. Junto a Revisión aparece Fecha de
reanálisis, con la unidad meses: carga y edita `FECHA DE REANALISIS` del JSON y
se comparte entre idiomas. Puede quedar incompleta en progreso; el guardado
terminado exige exactamente dos dígitos (por ejemplo, `06` o `12`).
El JSON guarda la unidad junto al valor (`"FECHA DE REANALISIS": "06 meses"`);
al reabrirlo, el campo numérico del editor muestra solamente `06`.
La fecha inicial usa la fecha local del sistema en formato `DD/MM/YY`; Folio
VAP, nombre de cliente y revisión del cliente empiezan con `-`. En la tabla se
puede pulsar Suprimir para vaciar la celda; el desplegable de periodicidad
agrega símbolos de uno en uno y permite hasta cuatro. Las sugerencias de
pruebas, unidades, códigos y nombres de producto se actualizan al escribir y se
pueden elegir con el ratón o el teclado; los campos conservan la escritura
manual. Elegir un Part ID completa el nombre del catálogo, y escribir o elegir
un nombre con una relación única completa el código. Un nombre que no exista en
el catálogo se puede conservar como texto libre. Al cambiar el idioma, los
nombres de prueba que tengan traducción única en la lista maestra se traducen
automáticamente.
`ACTUALIZAR MÉTODOS` permite consultar y editar manualmente las siete columnas
de la lista maestra o abrir la ubicación del Excel.
`COMPARAR ESPECS` abre dos paneles de comparación, distribuidos inicialmente
con el mismo ancho. El botón «＋ Agregar especificación» permite añadir un
tercer panel; cada panel tiene zoom independiente y puede abrir su vista en una
ventana emergente (hasta seis ventanas flotantes). Al añadir el tercer panel,
los tres se distribuyen con el mismo ancho. En las vistas, la tabla de pruebas
aparece antes que los datos complementarios del producto; el botón «×» de un
panel limpia esa visualización sin cerrar la ventana de comparación. El divisor
horizontal de cada vista permite ajustar verticalmente el espacio de la tabla
frente a los datos complementarios; el divisor vertical entre especificaciones
permite cambiar cuánto espacio ocupa cada tabla.
Cada panel permite elegir un JSON existente o un Word `.doc`, `.docx` o `.docm`.
El botón «ⓘ» junto al selector explica que se pueden comparar JSON limpios o
alineados, JSON viejos o extraídos y documentos Word propios.
Los documentos Word se extraen para la comparación con las funciones de lectura
del gestor, sin ejecutar el procesamiento que los registra ni mover el original.
El JSON y las conversiones auxiliares viven en una carpeta temporal: se eliminan
al limpiar ese panel o cerrar el comparador. Si se abrió una vista emergente de
un Word temporal, sus datos se conservan hasta cerrar también esa ventana.

`VER HISTÓRICO` permite consultar las tablas de `especificaciones.db` desde la
interfaz, buscar en todos los campos y recorrer los resultados por páginas; la
base se abre en modo de solo lectura. Desde esa ventana también se puede abrir
la carpeta de documentos generados o exportar la tabla seleccionada a un archivo
Excel `.xlsx` en la ruta que elija el usuario; si hay una búsqueda activa, se
exportan todos los registros coincidentes, no solo la página visible. La
aplicación permite además copiar
plantillas Word a `FORMATOS`, editar manualmente la lista maestra y
configurar las carpetas usadas por la interfaz. Los controles `A−` y `A+` ajustan el zoom
general; los controles `−` y `+` de cada panel ajustan solo esa vista previa.
El selector y la comparación incluyen JSON en español e inglés; las vistas
inglesas se identifican como «English version» y muestran en inglés los
encabezados de la tabla y las etiquetas del control documental.

## Selector y vista previa

Pulsa **SELECCIONAR ESPEC** en el menú principal para abrir el selector visual.
La pestaña `Anteriores · JSON extraídos` busca especificaciones extraídas desde
los documentos fuente en `ESPECIFICACIONES_JSON`; `Nuevas · JSON alineados`
busca la versión normalizada por la lista maestra en
`ESPECIFICACIONES_JSON_ALINEADAS`. Cada pestaña tiene su propia búsqueda por
codigo, producto, cliente o prueba, muestra todos los documentos coincidentes
y tiene su vista previa. En la parte superior de cada vista aparecen el nombre
del producto y el contador de pruebas, seguidos inmediatamente por la tabla.
Los desplazamientos de la lista, la tabla de pruebas
y el resto de la vista previa son independientes. Cada lista incluye casillas
de selección independientes y cada vista tiene el botón «Actualizar
especificacion», reservado por ahora sin acción. En ambas pestañas, «×» cierra
la visualización y vuelve al mensaje de selección; la tabla aparece antes del
control documental y los avisos. El divisor debajo de la tabla se puede
arrastrar verticalmente para cambiar su altura y el espacio de los datos
complementarios.
La pestaña anterior aplica el incremento de revisión y la fecha que usaría el generador,
pero no modifica el JSON ni genera un Word. Si la revisión fuente no es
numérica, conserva la revisión original, muestra el error en «Avisos del
control documental» y continúa presentando la especificación. En cambio, la
pestaña nueva toma la tabla exclusivamente de
`alineacion_metodos_lab.detalle[].valores_alineados` en el JSON alineado
(`prueba`, `unidades`, `metodo`, `especificacion` y `valor_tipico`), no de la
lista `pruebas` de la raíz. Si esa instantánea todavía no existe, la ventana
indica que hay que volver a ejecutar la alineación; no presenta los valores
antiguos como si fueran los alineados. La revisión y fecha se leen del control
documental sin incrementarlos en esta pestaña. Esto no altera la vista de los
JSON anteriores.

El proyecto aún no contiene la carpeta `FORMATOS` ni una plantilla Word
española. Por eso, la ventana presenta una vista de los campos y tablas que
rellena el generador, no una reproducción exacta del diseño del archivo Word.

## Mantenimiento manual de métodos de laboratorio

La lista maestra se mantiene manualmente en `METODOS_LAB_UNIFICADO.xlsx`, dentro
de la carpeta de métodos configurada. La aplicación utiliza ese archivo existente
para la alineación y el editor; no lo reconstruye ni combina archivos fuente.

Desde **Métodos**, usa **Editar lista maestra** o **Ubicación y edición manual**
para consultar y modificar el Excel. Para agregar o limpiar registros, trabaja
directamente en Excel y conserva la hoja `METODOS` y sus siete encabezados:
`METODO`, `MÉTODO INGLÉS`, `NOMBRE IQM`, `NOMBRE IQM INGLES`, `UNIDADES`, `UNITS`
y `ACREDITACION NOMBRE`.

La fecha mostrada en la tarjeta de métodos corresponde a la modificación del
propio Excel, incluidas las ediciones externas. Las dependencias necesarias para
la aplicación siguen en `requirements.txt`.

## Optimización del gestor

El gestor sigue calculando SHA-256 del catálogo y de los Word para detectar
cambios de contenido. Si el Excel coincide con una carga completa registrada,
recupera sus registros de SQLite sin interpretar el libro otra vez. Si cambia
o la caché está incompleta, lee el Excel. Nunca modifica el catálogo original.
Al registrar una edición, descarta las copias anteriores de la caché; las huellas
históricas de las extracciones permanecen en `archivos` y `versiones`.

La validación prepara una vez por carga los índices por nombre y Part ID,
respetando duplicados, prioridades y conflictos del catálogo. El texto OOXML
de cada Word se comparte entre detección de idioma, marcadores y metadatos;
esa caché se elimina al terminar el documento, incluso ante errores.

La primera carga con esta optimización revisa todo el historial. También se
revisa completo cuando cambia el catálogo o la versión del gestor. Después
se reindexan solamente los productos afectados y los posibles historiales de
destino de una normalización. Los cambios de tamaño o fecha de modificación
de un JSON histórico, y los archivos faltantes, también activan su revisión.
Las tablas auxiliares `gestor_estado` y `gestor_historial_cache` se crean en
SQLite sin borrar el historial existente.

Se mantienen las conversiones Word, el OCR y los controles de bloqueo actuales.

## Relaciones entre español e inglés

Los JSON pueden contener `versiones_idioma`, con listas `ESPANOL` o `INGLES`.
Cada referencia incluye `archivo_json` (ruta relativa al JSON que la contiene),
código, revisión y cliente de la contraparte. La relación se escribe en ambos
archivos. El índice HII se calcula en memoria y no se duplica en SQLite.

Solo se infiere automáticamente cuando hay una única especificación por idioma
con el mismo Part ID, revisión y cliente dentro de la carpeta. Estas coincidencias
no se guardan como vínculos permanentes. Las relaciones confirmadas
explícitamente en el generador HII se conservan, incluso con distinta revisión.
Los archivos faltantes no se exportan como referencias.

El gestor, el alineador, el editor y la recarga del generador HII actualizan
estas referencias. La escritura es atómica y solo se realiza cuando cambian
los enlaces; no modifica las pruebas ni el control documental.
