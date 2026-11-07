# Resumen del alineador de especificaciones

**Versión documentada: 2.5.0** (`alinear_especificaciones_json.py`).
Actualizado el 6 de noviembre de 2026.

Esta guía aparece en la pestaña **Instrucciones**. El alineador recibe JSON de
especificaciones, compara sus pruebas con la lista maestra de métodos y genera
JSON alineados para su revisión, edición y uso en los documentos de la aplicación.

## Uso y archivos

En la aplicación portátil, ejecuta `AplicacionEspecificaciones_v4.exe` desde su
carpeta completa, junto a `_internal`. En **Configuración** puedes seleccionar las
carpetas de trabajo y de métodos. El maestro es `METODOS_LAB_UNIFICADO.xlsx`,
hoja `METODOS`, dentro de la carpeta de métodos configurada.

Por defecto, el alineador lee `ESPECIFICACIONES_JSON` y escribe en
`ESPECIFICACIONES_JSON_ALINEADAS`, conservando la ruta relativa del archivo.
La carga de Word puede enviar al alineador los JSON recién extraídos. Desde el
selector también puedes alinear originales o seleccionar alineados y utilizar
**Realinear a versión vigente**.

En un entorno de desarrollo con las dependencias instaladas y Python 3.11, el
alineador puede ejecutarse desde la raíz del proyecto:

```powershell
.venv\Scripts\python.exe src\specification_generator\alinear_especificaciones_json.py
```

Se puede indicar otra carpeta de proyecto con `--base RUTA`. La opción `--forzar`
solicita procesar también los archivos sin cambios, pero no desactiva la
protección de destinos existentes descrita más adelante.

## Cómo busca coincidencias

Compara el nombre de origen con las columnas española e inglesa del maestro.
Para el método y el nombre de salida utiliza el idioma detectado en el JSON;
si no está declarado, aplica las reglas de idioma disponibles en el documento.
El nombre se compara con similitud de secuencia, orden y conjunto de palabras.

Con método declarado, el puntaje base combina 68 % de nombre y 32 % de método;
sin método combina 88 % y 12 %. Existen ajustes para coincidencias de nombre y
método muy altas, referencias ASTM compatibles y equivalencias confirmadas.
El puntaje por sí solo no garantiza que una propuesta sea técnicamente correcta.

Los umbrales generales son:

- **0.84 o más:** puede marcarse `ALINEADO` si no existe ambigüedad u otra
  condición que obligue a revisar.
- **Desde 0.72 y menos de 0.84:** `REVISION_REQUERIDA`.
- **Menos de 0.72:** `SIN_COINCIDENCIA`.

Una diferencia menor de 0.035 entre las mejores opciones puede indicar
ambigüedad cuando sus métodos o unidades difieren. Hay reglas específicas para
resolver alternativas confirmadas. También pueden requerir revisión referencias
incompatibles, métodos parciales, unidades no verificables o secuencias ausentes.

## Métodos y equivalencias

El alineador reconoce variantes de escritura ASTM, por ejemplo `D 92`, `D-92` y
`ASTM D92`, y normaliza citas identificables aun cuando no se encuentre la prueba
en el maestro. Reconoce `E 1252` como ASTM E1252 y algunos segundos números sin
letra dentro de métodos combinados, como `ASTM D445 / 7042`. La identificación
se limita a los patrones reconocidos: no todo texto con números es una norma ASTM.

Compara referencias normativas antes de aceptar candidatos y registra los métodos
ausentes o conjuntos no confirmados. En referencias combinadas con NMX puede
identificar la falta de `/` antes de `NMX`. Para NMX-J-123-ANCE, si se encuentra
una coincidencia y el año o apartado falta en el origen, puede completar el dato
que figura en la fila maestra. Los apartados expresamente indicados participan
en la comparación.

Reconoce referencias a USP, FEUM y Ph. Eur. Al alinear, conserva las farmacopeas
identificadas en el origen cuando corresponden a la comparación; no debe añadir
una farmacopea simplemente porque la fila maestra la enumere junto con otra.
Las reglas específicas confirmadas para determinadas pruebas se registran en la
auditoría del JSON.

Las versiones numéricas de USP se eliminan de la cita visible: `USP 32` y
`USP32` se expresan como `USP`, incluso al realinear campos editados. Las
variantes reconocidas de farmacopea europea se presentan como `PH. EUR`, igual
que en el maestro. En citas pendientes se conservan las ediciones, apartados y
otras normas presentes; estandarizar la escritura no confirma por sí solo una
combinación que el maestro no respalda.

Entre las equivalencias particulares se incluyen, según nombre y método, azufre
ASTM D2622/D4294, zinc y fósforo ASTM D4951, infrarrojo ASTM E1252, viscosidad
ASTM D445, limpieza ISO 4406 y secuencias de espuma ASTM D892. Para ASTM D92,
la regla de punto de inflamación prefiere la denominación a 101,3 kPa, salvo
que el nombre de origen indique expresamente 760 mmHg. Estas reglas dependen de
sus condiciones concretas; no se extienden automáticamente a otras normas.
Las denominaciones de viscosidad con `SUS` y 100 o 210 °F pueden vincularse a
las filas Saybolt del maestro, conservando las normas presentes en el origen.
Una cita de 200 °F, o de 100/210 °C con `SUS`, requiere revisión y no se trata
como si indicara 210 o 100 °F. La humedad declarada como Karl Fischer
coulométrico y el contenido de agua ASTM D6304 pueden vincularse a la fila
coulométrica cuando no indican un procedimiento incompatible.

## Unidades, valores y limpieza

Las unidades del maestro pueden contener alternativas. Se prioriza una unidad
declarada expresamente en ellas o un alias de la misma escala reconocido, como
`cSt` y `mm²/s`, o `cP` y `mPa·s`. Las unidades vacías, `NA` y adimensionales
mantienen las reglas existentes para completar o normalizar la unidad declarada.

Si la unidad cuantitativa original no tiene una equivalencia de escala confirmada
con el maestro, se conserva la unidad y las cifras, y se marca
`UNIDADES_POR_REVISAR`. Por ejemplo, `kg/m³` no se cambia a `g/cm³` dejando los
mismos números. El alineador **no convierte magnitudes físicas**: un cambio de
escala requiere revisión y conversión técnica explícita.

La unidad `cSt` se expresa como `mm²/s` incluso si la prueba queda pendiente o
se realinea una edición manual, ya que representan la misma escala. No se aplica
esta sustitución a unidades de viscosidad dinámica como `cP` o `mPa·s`.

Para métodos no reconocidos y pruebas pendientes se utiliza además el catálogo
`ALIAS_UNIDADES_SIN_CONVERSION`. Incluye variantes completas como `mg/Kg` →
`mg/kg`, `% MASA` → `% masa`, `PPM` → `ppm` y `g/cm3` → `g/cm³`. En los casos
reconocidos, el catálogo también ayuda a comparar alternativas de igual escala
con la lista maestra. Los textos que no figuran en el catálogo se conservan;
no se eliminan prefijos, denominadores ni contenido dudoso. La auditoría guarda
la unidad original y registra los cambios de escritura realizados.
Las unidades formadoras de colonias usan `UFC` o `UFC/g` en español y `CFU`
o `CFU/g` en inglés. Si no se conoce el idioma, se conserva la sigla original;
la comparación reconoce las equivalencias sin añadir ni quitar `/g`.

Las especificaciones y valores típicos se normalizan en su escritura: espacios,
separadores, agrupación de millares, operadores, símbolos y algunos términos
como `MÁXIMO`, `MAXIMUM`, `PASA` o `PASS`. No se redondean intencionalmente los
valores. Si un separador numérico es ambiguo, se conserva el valor y se registra
la observación para revisión. Cuando no hay coincidencia, también pueden
normalizarse la escritura del método, unidades inequívocas y valores; esto no
significa que la prueba haya sido aceptada como equivalente.

La periodicidad declarada se conserva; si falta, el alineador usa `▶` como valor
predeterminado. Detecta encabezados y notas estándar sobre la vigencia del método
para que no se traten como pruebas, conservando las notas correspondientes.

## Protección de archivos y realineación

El resultado se escribe mediante un archivo temporal y reemplazo del destino.
Cuando el origen y el destino son archivos distintos y ya existe un alineado,
solo se sustituye automáticamente si su huella coincide con la registrada para
esa fuente. Si fue modificado, falta la huella o cambia durante el proceso, se
conserva el alineado y se informa como **protegido** en el resumen y en
`REPORTE_ALINEACION.txt`. La sincronización de idiomas no reescribe los destinos
protegidos durante esa ejecución.

Antes de sustituir un destino verificable se guarda y comprueba una copia exacta
`.bak` en `.respaldos_alineacion`, junto al JSON. Si no puede verificarse el
respaldo, no se sustituye el destino. Estas protecciones también se aplican con
`--forzar`.

Para actualizar nombres del maestro en alineados existentes, selecciona sus
casillas y usa **Realinear a versión vigente**. Esta operación trabaja sobre los
alineados seleccionados y procura conservar los campos editados manualmente, las
notas y el historial HII. Si no puede conservar las filas con seguridad, informa
el error. La realineación no fusiona automáticamente una nueva extracción con
las ediciones de un alineado protegido: esa diferencia requiere revisión.

El historial HII se conserva y, cuando corresponde, se ajustan sus referencias.
Los vínculos de idioma confirmados se exportan al JSON con rutas relativas; las
coincidencias automáticas entre idiomas se calculan mediante el catálogo HII.
La asignación y conservación de revisiones corresponden a los flujos del gestor,
el editor y las herramientas de revisiones, no a un nuevo número calculado por
cada alineación.

## Auditoría y rendimiento

Cada JSON procesado incorpora `alineacion_metodos_lab`: versión, fecha, resumen y
detalle por prueba. El detalle incluye datos originales, propuestas, puntajes,
reglas aplicadas, estado, motivos de revisión y referencias no encontradas.
`REPORTE_ALINEACION.txt` resume los archivos procesados, omitidos, protegidos y
con error. Consulta estos datos al investigar una propuesta o una recarga.

El maestro se prepara una vez por lote. Las coincidencias repetidas se reutilizan
en memoria según nombre, método, idioma y unidades; esta memoria dura la
operación. La salida contiene `.cache_alineacion`, un índice auxiliar sin
extensión JSON. Un archivo se omite como `SIN CAMBIOS` cuando coinciden las
rutas, las huellas de origen, maestro y destino, la versión del alineador y la
versión interna de las reglas de unidades. Si el índice falta o el destino cambió,
la protección de archivos puede dejar el alineado pendiente en vez de sobrescribirlo.

## Limitación conocida: temperaturas

El filtro actual comprueba ciertas magnitudes, subtipos y secuencias, pero **no
bloquea de forma general las diferencias de temperatura** entre la prueba de
origen y la candidata. Por ello una coincidencia textual puede aceptar una prueba
a otra temperatura. Revisa esa condición y los valores antes de utilizar el JSON
alineado. Este comportamiento se registró como A03 en la auditoría y requiere
una corrección independiente.
