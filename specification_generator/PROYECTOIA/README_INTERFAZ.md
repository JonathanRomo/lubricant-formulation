# Interfaz 1.1.4

En la distribución portátil, ejecuta `AplicacionEspecificaciones_v4.exe` y
mantén `_internal` junto al ejecutable. No requiere instalar Python ni usar un BAT.
Desde el código fuente también puede iniciarse con
`.venv\Scripts\python.exe interfaz_aplicacion_especificaciones.py`.

La pestaña **Instrucciones** incluye las guías del alineador, del gestor y esta
guía de interfaz. Al abrirla muestra `RESUMEN_ALINEADOR_ESPECIFICACIONES.md`.

## Carpetas y portabilidad

`rutas_proyecto.py` centraliza la carpeta del programa y las rutas de listas,
métodos y traducciones mediante `carpeta_listas()` y `ruta_maestro()`. El menú, selector,
generadores, extractor, alineador, métodos y traductor usan esa ubicación,
independientemente de la carpeta desde la que se ejecuten. Los argumentos de
consola relativos también parten de esa carpeta, incluido `--base`.

En Configuración, las carpetas internas se muestran y guardan como rutas
relativas en `configuracion_interfaz.json`, por ejemplo `FORMATOS` y
`METODOS_LAB`. Los programas de consola consultan la misma configuración.
Las carpetas externas elegidas explícitamente conservan su ruta absoluta.
La configuración se mantiene en memoria y solo se relee si cambia el archivo.

Para trasladar el proyecto, conserva juntos los programas, sus carpetas de
consulta, `CODIGOS_200.xlsx`, `especificaciones.db` y los archivos de configuración
y preferencias. Las referencias bilingües y los estados WIP siguen siendo
relativos a sus JSON; la caché de alineación y las rutas internas registradas en
SQLite también son relativas. Los registros anteriores se convierten una vez
al abrir la base, conservando los enlaces. Si una ubicación absoluta antigua ya
no existe, se recupera su equivalente presente en las carpetas conocidas del
proyecto. No se cambia el esquema de las especificaciones.

La distribución onedir incluye el ejecutable, sus bibliotecas y la documentación.
Se entrega sin carpetas de trabajo, `CODIGOS_200.xlsx`, bases SQLite y diarios,
archivos JSON, BAT ni `.gitignore`. Incorpora tus recursos al utilizarla y
configura sus ubicaciones. La base y las preferencias se generan durante el uso.

## Recuperación de la base

Si falta `especificaciones.db`, la aplicación reconstruye su registro desde las
carpetas configuradas de JSON originales y alineados. También se recupera una
base existente sin registros de especificaciones cuando hay JSON disponibles:
puede ser un archivo de cero bytes, una SQLite sin tablas o con solo estructura,
catálogos y datos de control. No se utiliza un umbral de tamaño del archivo.
Una base con registros, bloqueada o ilegible no se considera vacía ni se
sustituye automáticamente. Si no hay JSON, una base vacía se conserva intacta.
La recuperación de una base vacía guarda un respaldo de la base anterior.
El gestor aplica este control. HII utiliza los JSON y no necesita abrir esa base.
La reconstrucción automática de la interfaz se ejecuta en segundo plano.

En **Personalizar → Reconstruir base desde JSON** se puede repetir la recuperación
manualmente. También existe el programa independiente:

```powershell
.venv\Scripts\python.exe reconstruir_base_datos.py
.venv\Scripts\python.exe reconstruir_base_datos.py --reemplazar
```

Antes de publicar la base, el programa revisa duplicados confirmados y los
mueve directamente a `REVISAR/DUPLICADOS`. Conserva los nombres y datos de las
especificaciones restantes y sus auxiliares WIP; ajusta las referencias de
idioma que apuntaban a una copia trasladada.
Recupera datos del producto, versiones, idiomas, huellas del Word y del catálogo
cuando existen, y relaciones bilingües explícitas o inequívocas. No inventa
parejas entre varias traducciones posibles. Una copia original y su alineada
equivalente se registran como una sola extracción y dos documentos disponibles.
Los alineados que no tienen original también permanecen registrados.

Las extracciones y sus versiones se conservan en `archivos` y `versiones`.
La recuperación ordena por fecha del documento, revisión y fecha de creación
del JSON; no usa la fecha de modificación del archivo. Los antiguos índices
`recuperacion_documentos` y `recuperacion_actuales` se retiraron porque duplicaban
metadatos sin ser consultados por la aplicación.

## Base portátil sin índices redundantes

HII calcula sus índices en memoria y conserva los vínculos confirmados y el
historial de Excel en los JSON. No acumula copias del catálogo de especificaciones
ni coincidencias automáticas en SQLite. La caché de `CODIGOS_200.xlsx` conserva
solo la edición vigente; las huellas históricas de extracciones permanecen.

Al abrir una base antigua desde el gestor se guarda una copia en
`RESPALDOS_BASE_DATOS/optimizacion_*` y se retiran las cinco tablas `hii_*` y
`recuperacion_documentos`. Los datos exclusivos de HII se rescatan a sus JSON
antes de retirarlos. Los archivos ausentes o vínculos incompatibles quedan
conservados en el respaldo. Una base ya migrada no genera otro respaldo.
Para migrar y compactar manualmente:

```powershell
.venv\Scripts\python.exe optimizar_base_datos.py
```

La compactación puede requerir cerrar los visores de SQLite que mantengan una
transacción abierta. Los archivos `-wal` y `-shm` son diarios de SQLite en uso;
no se eliminan manualmente.

## Revisión de duplicados

En **Personalizar**, junto a **Reconstruir base desde JSON**, el botón **Revisar duplicados**
revisa los JSON ya extraídos y los Word recibidos en segundo plano. También
se puede ejecutar sin la interfaz:

```powershell
.venv\Scripts\python.exe reconstruir_base_datos.py --revisar-duplicados
```

Al terminar una revisión o reconstrucción desde la aplicación, se recargan
automáticamente las listas abiertas del selector, la lupa del comparador y el
generador HII. Se conserva el filtro de búsqueda y se descartan las selecciones
y vistas previas de archivos que ya no estén disponibles.

Compara originales entre sí y alineados entre sí, sin confundir una extracción
con su versión alineada. Ignora nombres de archivos, rutas, fechas de extracción
y auditorías; normaliza espacios y saltos de línea. Conserva diferencias en
producto, código, idioma, revisión, cliente, control documental, pruebas,
unidades, métodos, cifras, notas, pie e instantáneas de valores alineados.
Los JSON incompletos y los borradores WIP no se apartan automáticamente.

Se conserva preferentemente la extracción ya registrada y, entre las demás,
la de fecha de creación anterior. Las copias JSON y sus Word originales se
mueven directamente a `REVISAR/DUPLICADOS`, sin carpetas por grupo ni un
historial adicional de copias. Solo se mueve un Word asociado si su huella
coincide con la extracción y no lo necesita otro JSON conservado. Las copias
Word idénticas que todavía no tienen JSON también se detectan por tamaño y
SHA-256. Un Word abierto o modificado desde su extracción se conserva.

La revisión independiente quita los registros de esas copias de SQLite y de
la caché de alineación, conservando el resto del historial. Si un nombre ya
existe en el destino se añade un sufijo, sin sobrescribir. Si falla un traslado
o la actualización de la base, se revierten los movimientos realizados.

Las huellas recuperadas permiten omitir los Word sin cambios al ejecutar el
gestor con el mismo catálogo. Un Word modificado, un catálogo diferente o un
JSON sin huellas suficientes puede requerir nueva extracción.

La nueva base se prepara aparte y se valida antes de publicarse. La recuperación
manual guarda la base anterior en `RESPALDOS_BASE_DATOS`. Los JSON inválidos y
referencias no recuperables se detallan en `REPORTE_RECONSTRUCCION_BASE.json`.
Si todos los JSON encontrados son inválidos o no se encuentran JSON para una
recuperación manual, no se sustituye la base anterior.
Los eventos y registros de HII que solo existían en SQLite no se pueden recuperar
con certeza desde estos JSON.

Si faltó el `.db` pero quedaron archivos `-wal`, `-shm` o `-journal`, se detiene
la publicación para evitar reutilizar diarios antiguos. Cierra los programas y
conserva esos archivos junto con cualquier base dañada antes de retirarlos y
volver a intentar. Una base existente no se reemplaza automáticamente por un
error de lectura; se utiliza la opción manual.

Prueba aislada con veinte JSON reales del proyecto:

```powershell
.venv\Scripts\python.exe tests\validar_recuperacion_20_json.py
```

Esta prueba trabaja con copias temporales y deja sus resultados en
`tests/REPORTE_RECUPERACION_20_JSON.json`.

El menú principal organiza las funciones en Inicio, Instrucciones, Histórico,
Formatos, Métodos, Configuración y Personalizar. Histórico se consulta dentro
de su pestaña; las operaciones de edición y comparación abren sus ventanas.
Las pestañas se reparten en el ancho disponible sin barra de desplazamiento.
Inicio muestra cinco
botones apilados en aproximadamente un tercio del ancho; el resto queda libre.
Sus descripciones aparecen al pasar el mouse.
Las ventanas nuevas se abren al frente. Selector y comparador comienzan con
zoom de vista previa de 144 %; cambiar el tema conserva ese zoom.

En Personalizar puedes guardar:

- Apariencia clara, oscura o del sistema.
- Acento azul, petróleo, índigo o verde.
- Escala del texto y los controles entre 80 y 140 %.
- Espaciado cómodo o compacto y tipografía.
- Apertura del menú maximizada.
- Adaptación automática del editor para pantallas pequeñas.
- Apertura del control documental colapsado.

Las preferencias visuales se guardan por separado en `preferencias_interfaz.json`;
las carpetas continúan en su configuración existente. Aplicar y guardar actualiza
el estilo de las ventanas abiertas; las opciones de apertura se usan en ventanas
nuevas.

Todas las ventanas usan CustomTkinter y sus botones; las tablas Treeview,
desplegables y el editor de notas protegido se conservan y reciben un tema
coherente. Las vistas previas de documentos mantienen el fondo de página para
representar sus archivos. No cambia la extracción, alineación ni generación.

Las ventanas limitan sus dimensiones y su tamaño mínimo al espacio de pantalla,
teniendo en cuenta el DPI. Notas y los botones inferiores del editor reservan
su espacio antes de asignar altura a la tabla. Con adaptación automática y poco
espacio vertical, el control documental comienza colapsado y puede desplegarse.

Validación de lógica:

```powershell
.venv\Scripts\python.exe -B -m unittest discover -s tests
```

La validación gráfica de Windows construye las ventanas fuera del área visible,
comprueba los temas y simula pantallas de 1280×720, 1920×1080 y 2560×1440:

```powershell
.venv\Scripts\python.exe -B tests\validar_interfaz_customtkinter.py
```

El alineador distingue viscosidad cinemática y dinámica mediante las unidades
y el nombre. Desde 2.1.4 se retiró la regla de temperaturas: no bloquea por
diferencias de temperatura o C/F, no sustituye X y no añade parámetros al
nombre del catálogo.

Incluye las equivalencias confirmadas del laboratorio para IV con ASTM D2270,
compuestos carbonizables y polinucleares, apariencia/aspecto/aspcto con VISUAL,
COLOR LOVIBOND con LOVIBOND, secuencias de espuma con ASTM D892 y humedad con
INS-CCA-07-00 o 91300092 hacia crepitación con 91300029. La auditoría conserva
los nombres y métodos originales y enumera las equivalencias aplicadas.
Si un nombre está confirmado pero las unidades A/NA siguen siendo ambiguas,
normaliza el nombre y conserva la unidad original para revisión.

Los compuestos ASTM/NMX requieren referencias y apartados presentes para la
misma prueba, unidades y condiciones. La salida conserva el apartado NMX de
origen. Los métodos ausentes siguen conservándose y resaltándose en rojo;
la auditoría distingue referencias ausentes, conjuntos no aprobados, métodos
sin equivalencia y condiciones incompatibles.

La versión 2.1.1 añade VISUAL APPEARANCE, las secuencias de espuma en inglés
y los nombres generales de espuma confirmados para ASTM D892. Si la secuencia
no se proporciona, reconoce el método y deja la secuencia pendiente.
Los carbonos aromáticos, nafténicos y parafínicos se relacionan con sus tres
filas ASTM D2140 y no se intercambian. Un título genérico de composición no
selecciona un componente por defecto. Las ediciones y años de las citas ASTM
siguen ignorándose al comparar el número de norma.

La versión 2.1.2 reconoce por identificador ASTM D2272 como estabilidad a la
oxidación, método B, salvo indicación explícita de método/procedimiento A.
ASTM E1252 identifica espectro infrarrojo y 91300092 identifica humedad por
crepitación con el método vigente 91300029, aunque el nombre sea distinto o
esté vacío. Las equivalencias quedan registradas junto con los datos originales
y funcionan también con INTERNAL METHOD en el maestro inglés.

## Referencia del alineador 2.1.35

Las reglas se aplican en el alineador JSON y en el editor; el formato numérico
Brookfield conserva su tratamiento específico en el alineador JSON. Los nombres
de salida provienen del maestro en el idioma del documento, con las equivalencias
confirmadas que se describen a continuación.

| Entrada confirmada | Resultado |
| --- | --- |
| IV con ASTM D2270 | Índice de viscosidad |
| Compuestos carbonizables con farmacopea | Sustancias fácilmente carbonizables |
| Compuestos polinucleares con farmacopea | Límite de hidrocarburos policíclicos aromáticos |
| Apariencia, aspecto, aspcto, appearance, apariencia visual, visual appearance o visual appeareance con VISUAL/ORGANOLÉPTICO o método vacío/NA | Apariencia visual / visual appearance, ORGANOLÉPTICO / ORGANOLEPTIC |
| Color Lovibond / LOVIBOND | Color Lovibond RYBN / IP17 A |
| INS-CCA-07, INS-CCA-07-00, 91300092 o 91300029 en el método, incluido CLIENT METHOD (91300029) | Humedad por crepitación / método interno 91300029 |
| ASTM E1252, aunque el nombre sea distinto o vacío | Espectro infrarrojo |
| ASTM D2272, aunque el nombre sea distinto o vacío | Estabilidad a la oxidación, método B; respeta indicación explícita de método/procedimiento A |
| Carbonos aromáticos, nafténicos o parafínicos con ASTM D2140 | Filas separadas de %CA, %CN y %CP del maestro; un título genérico queda pendiente de definir componente |
| SEC/SECUENCIA/SEQUENCE I, II o III, TENDENCIA/ESTABILIDAD o TENDENCY/STABILITY | Formación de espuma, secuencia correspondiente, ASTM D892; puede recuperar el método vacío para estos nombres confirmados |
| Tendencia a la espuma, espuma, estabilidad a la espumación o prueba a la espuma | ASTM D892, sin inventar secuencia; queda pendiente de definirla |
| Saybolt con ASTM D2161, 37,8 °C/100 °F o 98,9 °C/210 °F | Viscosidad Saybolt a 100 °F o 210 °F: equivalencias nominales limitadas a esta prueba |

Los identificadores no se corrigen por parecido: no se eliminan otras normas de
una cita compuesta para forzar una equivalencia. Las reglas por identificador
de crepitación, D2272 y E1252 no requieren similitud del nombre.

Unidades:

- N/A y NA se escriben NA en ambos idiomas.
- A/D, AD, ADIMENSIONAL y DIMENSIONLESS se escriben ADIMENSIONAL en español
  o DIMENSIONLESS en inglés. Se toleran diferencias de mayúsculas y espacios.
- Estas declaraciones se normalizan incluso en filas sin coincidencia y con
  método ausente. No se reemplaza una declaración NA por AD, ni viceversa,
  por el valor predeterminado del maestro.
- cSt y mm²/s identifican viscosidad cinemática; cP y mPa·s, dinámica. Nombres
  y unidades incompatibles impiden la alineación automática.
- Si el maestro ofrece varias unidades, conserva la original cuando figura
  entre las alternativas; en otro caso usa la primera propuesta. Los alias
  NA/AD se normalizan como respaldo en filas sin coincidencia aceptada;
  en filas alineadas tiene prioridad la unidad del maestro. No se convierten
  numéricamente resultados.

Temperaturas y subtipos:

- Se retiró la regla general de temperaturas: no se bloquean diferencias de
  temperatura o C/F, no se completan X ni se añaden temperaturas al catálogo.
- Se mantiene la equivalencia nominal específica de Saybolt/D2161.
- No se intercambian secuencias de espuma ni componentes de carbono.
- Peso específico/densidad relativa se mantienen en su familia y no se
  sustituyen por una densidad absoluta solo por parecido de nombre.

Identidad de métodos y maestro:

- ASTM reconoce guiones, espacios, letras de designación y ediciones/años;
  D445 -25 (2025) sigue siendo D445. Conserva referencias de métodos compuestos
  y recupera ASTM/D omitidos en las formas conjugadas reconocidas.
- Distingue números y familias ASTM, NOM, NMX, ISO, IEC, DIN, EN, IP, SAE,
  JIS y API. Una NOM ausente no se sustituye por USP/FEUM.
- USP, USP/NF, FEUM y Ph Eur se reconocen con ediciones, capítulos o MGA en
  las variantes admitidas. La salida omite esos accesorios y no añade FEUM
  si el original solo cita USP. NF aislada no tiene equivalencia confirmada.
- VIGENTE, VIGENTES, CURRENT y CURRENCY no cambian la farmacopea. Se limpia
  la cita reconocida aunque el nombre de prueba quede pendiente.
- Los compuestos ASTM y ASTM/NMX se validan en una misma prueba y unidades
  del maestro. Los apartados NMX deben existir en ese grupo y se conservan
  en la salida. Otras normas compuestas requieren un registro compatible.
- Se comparan nombres españoles e ingleses. Se utiliza el método/unidad del
  idioma solicitado, con respaldo al español cuando falta su traducción.
- Para comparar se ignoran mayúsculas, acentos, puntuación y espacios
  repetidos. Se retiran los asteriscos iniciales de acreditación del nombre
  canónico. La similitud textual combina 75 % orden, 15 % palabras ordenadas
  y 10 % conjunto de palabras; las referencias normativas tienen su comparación
  específica por identidad.
- El lector elimina filas duplicadas del maestro y omite filas vacías.
  Tolera un encabezado inglés sobrescrito solo en el diseño fijo conocido
  de siete columnas con las otras seis posiciones intactas.

Decisión y pendientes:

- Puntaje normal: 68 % nombre y 32 % método; sin método, 88 % y 12 %.
  Para nombre/método muy coincidentes se garantiza 0,96. Una coincidencia
  ASTM compatible con método >= 0,90 y nombre >= 0,45 puede alcanzar 0,84.
  Las restricciones de identidad, unidades y subtipos siguen siendo obligatorias;
  no hay bonificación adicional por temperatura.
- ALINEADO: puntaje >= 0,84 sin ambigüedad; REVISION_REQUERIDA: >= 0,72 o
  pendiente específico; SIN_COINCIDENCIA: los demás. La diferencia < 0,035
  entre candidatos potencialmente ambiguos puede obligar a revisión.
- Los métodos alternativos del mismo grupo ASTM no son pruebas diferentes.
  Una unidad explícita compatible puede resolver un empate del mismo nombre.
- Un nombre confirmado con unidades A/NA aún ambiguas puede normalizarse
  manteniendo la unidad original y la revisión pendiente.
- Los métodos ausentes se conservan y se muestran en rojo. Se distinguen
  referencia ausente, conjunto no aprobado y método sin equivalencia.
  Las condiciones incompatibles, unidades dudosas, nombre sin equivalencia,
  secuencia o componente sin definir quedan identificados para revisión.
- En filas con método ausente se conservan nombre, método y valores; se
  normalizan únicamente las declaraciones NA/AD confirmadas.

Formato de contenido:

- Se separan valores y símbolos de unidades: 25°C pasa a 25 °C. Se unifican
  formas reconocidas de mL, µL, mm², cm³, kg/m³, mPa·s, Pa·s, N·m, s⁻¹,
  cm⁻¹ y otras unidades admitidas, sin conversiones de magnitud.
- Se usa coma decimal y agrupación en bloques de tres, también en la parte
  decimal: 0,8750 pasa a 0,875 0. Se conserva la precisión significativa.
- En rangos con contexto decimal claro se iguala el número de decimales
  rellenando ceros: 7,00 - 8,000 pasa a 7,000 - 8,000.
- Se normalizan separadores de rango, signos y espacios. La notación
  científica e/x10 se escribe con × 10 y exponente en superíndice.
- Una coma con tres cifras y parte entera distinta de cero puede ser ambigua:
  se conserva para revisión cuando no hay contexto suficiente para decidir.
- Brookfield/D2983 tiene la excepción de coma de miles: 3,610 pasa a 3 610;
  un punto único sigue siendo decimal. Agrupaciones inválidas quedan pendientes.
- Se protegen referencias normativas, fechas, versiones, direcciones IP,
  códigos y números de serie para no reformatearlos como mediciones.
- MAX/MÁX y MIN se normalizan como MÁXIMO/MÍNIMO; min minúsculo se conserva
  como minutos. Se normalizan mayúsculas/acentos de típico, pasa/pass y
  reportar/report; negativo/negativa se unifica como NEGATIVA y negative como
  NEGATIVE, sin traducir indiscriminadamente el contenido.

Limpieza y trazabilidad:

- En JSON se excluyen los encabezados reconocidos de la tabla y se conserva
  la fila omitida en auditoría. Las notas estándar reconocidas en español o
  inglés se separan de las celdas y no se duplican en la lista de notas.
- El editor conserva los comentarios y marcadores de celda; los cambios
  normales se resaltan en azul y los métodos ausentes en rojo.
- La auditoría registra fuente, propuesta, regla aplicada, unidades originales,
  puntajes, alternativas, referencias ausentes, versión y motivo de revisión.
- Al realinear se reutilizan los originales auditados si la fila sigue
  correspondiendo a la salida anterior; se conservan las modificaciones del
  usuario que ya no coinciden con esa salida.
- Se conserva el historial HII. Se validan la raíz JSON y la lista de pruebas;
  la escritura se publica mediante archivo temporal. Los campos auxiliares
  heredados FECHA DE REANALISIS y Desarrollador se dejan vacíos.
- La periodicidad recupera los campos PERIODIDICAD DE ANALISIS (campo heredado),
  periodicidad o periodicidad_analisis; si no existe, se asigna ▶ en el JSON.
- La caché se invalida al cambiar fuente, maestro, salida o versión del
  alineador. Las rutas internas se guardan de forma portátil.

En ASTM D2140, las familias aromatic/naften/naphten/naphthen/parafin/paraffin
y los tokens completos CA/CN/CP identifican el componente, incluso sin la
palabra carbono. Se usa la fila oficial del maestro en el idioma del documento.
Los nombres que indican varios componentes o ninguno no se asignan a un
componente por esta regla.

FIXED OILS, FATS AND ROSIN/RESIN se vincula a ACEITES FIJOS, GRASAS Y RESINA
del maestro, admitiendo el texto NO OILY OR SOLID MATTER SEPARATES/SAPARATES
y la marca B final. Se conserva la validacion independiente del metodo.

HUMEDAD/HUMIDITY, HUMIDITY (SPLATTER TEST), HUMIDITY CREPITATION y
CREPITATION HUMIDITY se reconocen como humedad por crepitacion cuando el
metodo esta vacio, NA, metodo del cliente/CLIENT METHOD o lleva los
identificadores confirmados INS-CCA-07(-00) y 91300092. No se sustituye una
norma explicita distinta ni se infiere una prueba de una fila sin nombre.

COLOR VISUAL, VISUAL COLOR y DETERMINACION DE COLOR con metodo VISUAL
se vinculan a COLOR del maestro. La equivalencia no cambia otros metodos.

Los nombres con el token LOVIBOND se vinculan a COLOR LOVIBOND RYBN /
IP17 A si el metodo es LOVIBOND, vacio, NA o solo USP (incluida su edicion).
Se admite texto adicional de color, celda y temperatura sin completar X.
Otras normas explicitas conservan su validacion.

Desde 2.1.10, AD/A/D de la fuente no reemplaza la unidad del maestro en
coincidencias aceptadas. Sin coincidencia aceptada, se normaliza como
ADIMENSIONAL/DIMENSIONLESS segun el idioma.

ASTM D6304 con PROCEDIMIENTO A/PROCEDURE A vincula nombres de AGUA,
HUMEDAD o WATER a CONTENIDO DE AGUA POR KARL FISCHER (COULOMETRICO).
La equivalencia admite ediciones de la norma y no se aplica a procedimientos
B/C explicitos ni conjuntos con otras normas. Se usa la unidad del maestro.

ACIDEZ/ACIDITY, incluido ACIDITY AND ALKALINITY y la nota NO RED OR PINK
COLOR IS PRODUCED, con USP/FEUM o sin metodo se vincula a ACIDEZ, USP / FEUM.
Esta equivalencia confirmada usa ambas farmacopeas como excepcion a conservar
solo las referencias originales. Otras normas explicitas no se sustituyen.

Con ISO 4406, las variantes CODIGO/NIVEL DE LIMPIEZA, PUREZA DE PARTICULA,
AUTOMATIC CLEANING CODE y CODING THE LEVEL OF CONTAMINATION se vinculan
a CODIGO DE LIMPIEZA / CLEANLINESS CODE del maestro. Se admiten detalles
de control interno, particulas y volumen. No se infiere desde un nombre vacio
ni se aplica a otras normas o conjuntos normativos.

COMPUESTOS POLINUCLEARES admite el prefijo LIMITE DE y marcas finales.
Por confirmacion del laboratorio, los casos con USP, FEUM o ambos se vinculan
a LIMITE DE HIDROCARBUROS POLICICLICOS AROMATICOS, metodo USP del maestro.
El metodo fuente permanece en la auditoria; esta equivalencia no incorpora
FEUM al catalogo ni modifica otras normas.

ALKALINITY (THE SOLUTION DOES NOT AQUIRE/ACQUIRE A PINK COLOR), con
o sin marca B final, se vincula a ALCALINIDAD cuando cita USP/FEUM.
Se mantiene la regla general de conservar las farmacopeas del original.

La confirmacion de CONTENIDO DE AGUA, PROCEDIMIENTO A tambien admite
metodo vacio o NA: se vincula a Karl Fischer coulometrico / ASTM D6304.
No se infiere D6304 de CONTENIDO DE AGUA sin procedimiento ni se reemplaza
una norma distinta explicitamente indicada.

CONTENIDO DE AGUA con ASTM D1533 se vincula a Karl Fischer (LIQUIDOS
AISLANTES). Tambien admite conjuntos con NMX si el catalogo autoriza la
misma prueba y apartado; conserva la cita NMX original y valida sus apartados.

La equivalencia confirmada de ACIDEZ admite tambien Ph. Eur con edicion y
ACIDEZ % sin metodo: se vinculan a ACIDEZ / USP-FEUM del maestro actual.
El porcentaje final no cambia la unidad maestra ni convierte resultados.
La referencia original se conserva en la auditoria.

AZUFRE, AZUFRE Y SULFUROS y LIMITE DE COMPUESTOS DE AZUFRE con USP/FEUM
se vinculan a LIMITE DE COMPUESTOS DE AZUFRE / USP del maestro actual,
por confirmacion del laboratorio. Se conserva la referencia fuente en la
auditoria. No se aplica a ASTM ni a nombres sin metodo.

PUNTO DE INFLAMACION sin condiciones adicionales, con ASTM D92, se vincula
a PUNTO DE INFLAMACION A 101,3 kPa, COPA ABIERTA CLEVELAND. No reemplaza
nombres que ya especifican otra presion ni otras normas como ASTM D93.

PUNTODE ANILINA con ASTM D611 y METODO E/METHOD E se vincula a
PUNTO DE ANILINA, METODO E. La regla admite ediciones y reaprobaciones
y no se activa para otro metodo de ensayo explicitamente indicado.

PRUEBA CADA LOTE, incluido el simbolo de periodicidad, se identifica como
separador y no se alinea a una prueba. Se conserva para revision y se marca
en rojo en el editor y vista previa con el motivo correspondiente.

PUNTO DE ESCURRIMIENTO con ASTM D97 se vincula a PUNTO DE FLUIDEZ /
POUR POINT, sin asignarlo a las variantes superior o inferior del maestro.

VISCOSIDAD A 40 C con ASTM D445 se identifica como VISCOSIDAD CINEMATICA
A 40 C por confirmacion del laboratorio, incluso sin unidad. Las unidades
explicitas de viscosidad dinamica mantienen prioridad para distinguir magnitudes.

La equivalencia de viscosidad sin magnitud declarada con ASTM D445 admite
40 C y 100 C, usando la fila cinematica de la temperatura correspondiente.
Las unidades explicitas siguen distinguiendo viscosidad cinematica y dinamica.

TENDENCIA A LA ESPUMA con metodo --- es un separador: se marca en rojo
y no se infiere ASTM D892. Con ASTM D892 explicito mantiene su tratamiento
de prueba de espuma pendiente de secuencia.

SEPARACION AGUA ACEITE A 54 C con ASTM D1401 se vincula a
SEPARABILIDAD AGUA-ACEITE A 54 C. Esta equivalencia confirmada no se
extiende a otras temperaturas, escalas o normas.

FLASH POINT con ASTM D92 comparte la equivalencia confirmada de PUNTO
DE INFLAMACION: copa abierta Cleveland a 101,3 kPa del maestro.

DEMULSIBILIDAD con ASTM D1401 se vincula, por confirmacion del laboratorio,
a SEPARABILIDAD AGUA-ACEITE A 54 C del maestro, sin asumir la variante 82 C.

SEPARABILDAD/SEPARABILIDAD DEL AGUA Y DEL ACEITE Y FLUIDOS SINTETICOS
A 54 C se vincula a SEPARABILIDAD AGUA-ACEITE A 54 C / ASTM D1401,
con metodo vacio, NA o D1401. No reemplaza otra norma explicita.

La equivalencia de punto de inflamacion con ASTM D92 admite la unidad
Celsius al final del nombre, como PUNTO DE INFLAMACION, C.

DENSIDAD RELATIVA A 15.6/15.6 C con el conjunto D1298/D7042 se vincula
a la fila D1298 de esas condiciones por confirmacion del laboratorio.
Se conservan ambas normas; no se autoriza este conjunto para otras pruebas
o condiciones ni se extiende a normas adicionales.

SEPARACION AGUA ACEITE con ASTM D1401 admite tambien 82 C, seleccionando
la fila SEPARABILIDAD AGUA-ACEITE A 82 C, sin confundirla con 54 C.

SECUENCIA I/II/III, sola o precedida de TENDENCIA A LA ESPUMA, admite
marcas finales y se vincula a FORMACION DE ESPUMA de la misma secuencia
con ASTM D892, metodo vacio, NA o ---. La tendencia sin secuencia y con
--- permanece como separador rojo. No se reemplazan otras normas explicitas.

Con ASTM D972, VOLATILIDAD/VOLATILITY y PERDIDA/PERDIDAD POR EVAPORACION
se vinculan a PERDIDA POR EVAPORACION DE GRASAS LUBRICANTES Y ACEITES.
Por confirmacion del laboratorio, el nombre generico del maestro sustituye
los detalles de tiempo y temperatura del titulo; el original queda auditado.
No se extiende la equivalencia a otras normas como ASTM D5800.
