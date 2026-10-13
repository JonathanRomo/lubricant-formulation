# Resumen del alineador de especificaciones

**Versión documentada: 1.12.0** (`alinear_especificaciones_json.py`)

El alineador toma los JSON de origen seleccionados y genera o reemplaza sus
versiones alineadas en `ESPECIFICACIONES_JSON_ALINEADAS`, conservando la ruta
relativa de cada archivo. Usa la hoja `METODOS` del archivo maestro
`METODOS_LAB/METODOS_LAB_UNIFICADO.xlsx`.

## Cómo encuentra coincidencias

- Compara el nombre de cada prueba con los nombres del maestro en español e
  inglés. Si el JSON declara idioma, usa también el método maestro de ese idioma.
- Combina similitud de texto, orden y conjunto de palabras. Cuando existe un
  método en el JSON, el puntaje pondera 68 % el nombre y 32 % el método; si no
  existe, pondera 88 % el nombre y 12 % el método.
- Las referencias ASTM se comparan por designación, ignorando diferencias de
  espacios, guiones y edición/año. También reconoce métodos conjugados y ciertas
  referencias de farmacopea (USP, FEUM y Ph. Eur.).
- Un puntaje de 0.84 o más permite alinear automáticamente; entre 0.72 y 0.84
  requiere revisión; por debajo de 0.72 se clasifica sin coincidencia.
- Si la diferencia entre las dos mejores opciones es menor que 0.035, considera
  ambigua la coincidencia cuando los métodos difieren o cuando el mismo nombre
  tiene unidades distintas. Una coincidencia ambigua no se aplica
  automáticamente.
- Si una referencia ASTM del método de origen no aparece en el maestro, la
  registra para revisión. Si el nombre del maestro no está disponible, tampoco
  reemplaza automáticamente el nombre de la prueba.

## Qué modifica y qué conserva

- Para coincidencias automáticas, sustituye el nombre y método por sus formas
  canónicas del maestro y usa las unidades correspondientes. Las unidades del
  maestro separadas por comas son alternativas: conserva la unidad original si
  coincide con cualquiera de ellas; si no coincide, propone solo la primera
  unidad alternativa, nunca la lista completa.
- Conserva los valores de especificación y típico, sin convertir unidades ni
  redondear. Normaliza separadores decimales, agrupación de millares, ceros
  iniciales, separación y símbolos comunes de unidades, operadores y notación
  científica. Protege fechas, identificadores y referencias normativas.
- Si un separador decimal puede confundirse con agrupación de millares, conserva
  el valor original y marca la prueba como `REVISION_REQUERIDA`, indicando el
  campo y el valor ambiguo en la auditoría.
- Normaliza variantes de `MINIMO` y `MAXIMO` en mayúsculas a `MÍNIMO` y
  `MÁXIMO`; preserva `min` en minúsculas para no confundirlo con minutos.
- Normaliza términos frecuentes de especificación y valor típico, como
  `TÍPICO/TYPICAL`, `PASA/PASS`, `REPORTAR/REPORT` y
  `NEGATIVO/NEGATIVA/NEGATIVE`, con ortografía correcta y en mayúsculas.
  Las variantes de género y número en español se estandarizan a `NEGATIVA`.
  Evita que las unidades `m` o `M` alteren palabras con letras acentuadas.
- No aplica automáticamente propuestas a coincidencias que requieren revisión
  o no tienen coincidencia.
- Normaliza espacios entre cantidades y unidades y el formato de números en
  unidades, valores típicos y notas. Los separadores numéricos ambiguos se
  conservan para evitar reinterpretarlos.
- Reinicia los campos `FECHA DE REANALISIS`, `Desarrollador` y
  `PERIODIDICAD DE ANALISIS`.
- Reconoce las notas estándar en español e inglés sobre la versión vigente del
  método. Si vienen como una fila de la tabla, las retira de `pruebas` y las
  agrega a `notas`; si están junto al contenido de una prueba, separa la nota
  sin descartar la prueba. Evita duplicarlas cuando ya estaban en `notas`.
- Guarda en cada JSON una auditoría con versión, fecha, resumen, propuestas,
  puntajes, estado de cada prueba y referencias ASTM no encontradas. También
  genera `REPORTE_ALINEACION.txt` con el resultado por archivo.

Los resultados dudosos, sin coincidencia o con errores deben revisarse; el
puntaje de similitud no sustituye la validación técnica de la especificación.
