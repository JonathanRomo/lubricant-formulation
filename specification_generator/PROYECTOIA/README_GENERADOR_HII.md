# Generador HII

`generador_hii.py` reúne la generación de Excel, el zoom, la trazabilidad, el
índice de trabajo en memoria y la caché del catálogo. El alineador, el editor y la recuperación
importan de ese archivo las funciones HII que necesitan.

En el gestor, pulsa **GENERAR HII**. Busca por código, nombre, cliente o revisión
y marca las casillas de las especificaciones alineadas para agregarlas al lote.
La lista permanece completa al marcar; **Ø** alterna manualmente entre todas
las especificaciones y solo las marcadas. La barra espaciadora también marca
o desmarca una fila. **Guardar seleccionadas (N)** indica el total del lote.

Al seleccionar cualquier fila se muestra su resumen con tabla de pruebas
y control documental, sin cambiar las casillas del lote. La vista previa del Excel inicia colapsada. Cada vista
conserva su encabezado con dos botones: **+** muestra solo esa vista y cambia
a **−** para restaurar la disposición anterior; **∅** oculta o muestra su propia
vista. Se pueden colapsar ambas. Al ocultar una vista se liberan sus widgets;
al ocultar Excel también se liberan sus imágenes y el archivo preparado.
Los separadores permiten ajustar el ancho cuando ambas están visibles.
Guarda la hoja actual con **Guardar HII**
o todo el lote con **Guardar seleccionadas**. Este último usa el nombre del JSON
e idioma para distinguir revisiones y pide confirmación antes de reemplazar
archivos existentes. También puede iniciarse directamente:

```powershell
.venv\Scripts\python.exe generador_hii.py
```

## Apertura y búsqueda

El encabezado compacto reserva más altura para las vistas previas. Sus botones
**−** y **+** ajustan el zoom de todo el contenido inferior entre 60% y 160%:
controles, lista, tabla, detalles y Excel dibujado. El zoom es local a esta
ventana y conserva las preferencias generales de la aplicación. El dibujo del
Excel se escala sin volver a leer ni generar el archivo.

La ventana aparece con el mensaje **Cargando especificaciones…**. El catálogo
se prepara en segundo plano; la lista se inserta en bloques de 100 filas para
permitir que la ventana responda mientras termina. La búsqueda se aplica 250 ms
después de la última tecla.

La primera carga lee los JSON y crea `.cache_catalogo_hii` en la carpeta de
especificaciones alineadas. Las siguientes cargas verifican la ruta, fecha de
modificación, tamaño y presencia del estado WIP de cada archivo, y releen los
documentos nuevos o modificados. También detectan archivos eliminados o
renombrados. **Actualizar lista** vuelve a comprobar la carpeta.

La caché es independiente de SQLite, se reconstruye si falta o está dañada y se
puede borrar para forzar una lectura completa. Si la carpeta es de solo lectura,
el catálogo sigue funcionando sin guardar la caché. Los documentos originales
permanecen como fuente de información.

Abrir el generador carga el catálogo sin sincronizar SQLite ni reescribir los
JSON. Al solicitar otro idioma, se buscan las relaciones `versiones_idioma`
(también en sentido inverso) y, si no hay enlaces explícitos, las coincidencias
de código, revisión y cliente del catálogo. Las opciones ambiguas siguen
requiriendo una elección. Se comprueba el JSON elegido antes de usarlo; el
registro del vínculo y de las HII continúa haciéndose cuando corresponde.

Excel se construye al abrir su vista previa o guardar la HII;
las vistas ocultas no se dibujan al cambiar de especificación.

Cada selección comparte una sola lectura del JSON entre el resumen y los datos
HII. La construcción individual de Excel y la preparación del plano se ejecutan
en un único trabajador por ventana; Tk solo recibe los resultados y dibuja.
El dibujo avanza de arriba hacia abajo en bloques de hasta 40 órdenes.
El zoom también funciona durante el
dibujo. Ocultar Excel o cambiar de selección descarta los resultados anteriores;
un guardado ya solicitado conserva su propia especificación de origen. Cerrar
la ventana cancela los trabajos pendientes y sus callbacks.

Las plantillas preparadas se reutilizan mediante una caché en memoria de hasta
cuatro entradas, identificadas por ruta, tamaño y marcas de tiempo del archivo.
Cada HII modifica una copia de las partes XML, de modo que sus datos no pasan a
la siguiente salida. Si cambia la plantilla se vuelve a leer. Los lotes también
aprovechan esta caché.

Se elige `91400005` para español y `91400019` para inglés en la carpeta de
formatos configurada. Cada código debe tener una sola plantilla vigente. Se
admiten `.xlsm`, `.xlsx`, `.xltm` y `.xltx`; cada archivo generado conserva la
extensión de su plantilla. El formato binario antiguo `.xls` no está incluido.
Se conserva el idioma del JSON; no se traduce contenido.

| Destino | Fuente |
| --- | --- |
| Objeto del nombre | `producto.nombre` |
| Objeto del código | `control_documental.codigo_unico_producto` |
| A: No. | Número consecutivo |
| B: Prueba | `valores_alineados.prueba` |
| C: Método | `valores_alineados.metodo` |
| D: Periodicidad | Periodicidad alineada o de la prueba original |
| E: Unidades | `valores_alineados.unidades` |
| F: Especificación | `valores_alineados.especificacion` |
| Comentarios combinados | Notas de la especificación y notas separadas de las pruebas |
| Reanálisis | `fecha_reanalisis` en el JSON, control o producto; alternativamente una nota `Fecha de reanálisis: …` |
| Todos los `DD/MM/YY` | Fecha de generación |

Los valores típicos se muestran en la especificación, pero no se exportan a
la HII. Si faltan periodicidad o reanálisis se dejan vacíos. Los comentarios se
copian como texto simple; `comentarios_simples` contiene el TODO para agregar
comportamientos especiales. Si no hay notas se conserva el comentario base de
la plantilla.

La tabla se amplía después de sus seis filas iniciales. Se desplazan los
comentarios, reanálisis, objetos posteriores y catálogo auxiliar. El archivo
se modifica como paquete XML para conservar imágenes, macros y objetos. Las
macros se conservan sin ejecutarse. Las plantillas no se sobrescriben.

La previsualización dibuja celdas, combinaciones, logo y textos de los objetos
del archivo generado. Es una representación aproximada: Excel puede ajustar
tipografía y paginación de otra manera.

## Versiones por idioma

Los botones **Español** e **Inglés** eligen el idioma para
la vista previa y para el guardado individual o de las seleccionadas. Aparecen
deshabilitados y atenuados si alguna selección no tiene esa versión. Al marcar
un documento se muestra inicialmente su idioma, sin un botón adicional. Se busca
una especificación alineada con el mismo código, revisión y cliente en el idioma
solicitado. Si hay varias candidatas se pide elegir, y esa relación se recuerda.

La casilla **Ambos** se habilita cuando todas las selecciones tienen español e
inglés. Al marcarla, **Guardar ambas** pide una carpeta y genera las dos HII de
la especificación actual; **Guardar seleccionadas** genera ambos idiomas de todo
el lote. Si las dos versiones ya estaban marcadas, se guardan una sola vez.
Se conserva la confirmación antes de reemplazar archivos existentes.

Si falta una versión, se debe crear y guardar desde el editor de especificaciones
y pulsar **Actualizar lista**. El generador no abre automáticamente una traducción
al pulsar un idioma ausente. Los enlaces explícitos y la trazabilidad de cada
archivo se guardan en los JSON, con rutas relativas cuando están en la misma unidad.

HII utiliza un índice temporal en memoria que se descarta al terminar cada
operación. Generar Excel o vincular traducciones no crea ni actualiza
`especificaciones.db`. Las coincidencias automáticas se calculan al consultar;
solo los vínculos confirmados se guardan en `versiones_idioma` de los JSON.

## Historial de HII en el JSON

Desde el alineador **1.13.0**, los JSON alineados incluyen `hii_generadas: []`.
El generador **1.1.0** añade un registro al guardar cada Excel, individualmente
o en lote. También admite JSON anteriores sin ese campo:

```json
"hii_generadas": [
  {
    "id": "f626d445-55b7-4424-9a1e-3253562e053a",
    "archivo": "../HII/20100001 HII ESPANOL.xlsm",
    "fecha_generacion": "2026-10-14T10:30:00-06:00",
    "version_generador": "1.1.0",
    "hash_especificacion": "sha256 de los valores utilizados"
  }
]
```

Cada guardado recibe un UUID nuevo, incluso al reemplazar el mismo archivo. El
Excel contiene ese valor en la propiedad personalizada **HII_ID**, visible desde
las propiedades avanzadas del archivo. El historial conserva los guardados
anteriores; al sobrescribir un Excel, solo existe en esa ruta la última
generación. No se crean copias históricas de los Excel.

`archivo` es relativo a la carpeta del JSON. Si se guarda en otra unidad de
Windows, se usa una ruta absoluta. Mover toda la carpeta del proyecto conserva
las referencias relativas. Renombrar o mover un Excel por separado no actualiza
automáticamente el JSON; su HII_ID permite identificarlo al inspeccionarlo.

La huella SHA-256 corresponde al nombre, código, idioma, las seis columnas de las
pruebas, comentarios y reanálisis utilizados. Excluye la fecha de impresión, el
historial, los valores típicos y los metadatos de auditoría. Registrar una HII no
cambia esa huella. Si los valores efectivos cambiaron desde que se preparó la
vista previa, se solicita cargarla de nuevo antes de guardar.

El alineador conserva el historial existente al realinear desde el original o
sobre el propio JSON. El editor conserva los registros añadidos después de abrir
un documento; una traducción nueva empieza con su propio historial.

Los archivos se reemplazan mediante temporales. Se guarda primero el Excel con
su ID y después el historial del JSON. Si falla este segundo paso, se informa del
guardado parcial, incluyendo el ID. Los dos archivos no forman una transacción
única; editar el mismo JSON simultáneamente desde varios programas requiere
coordinación externa.

La reconstrucción comprueba el historial desde los JSON sin modificarlos y
avisa de Excel no localizados o registros incompletos. No duplica ese historial
en tablas permanentes. La migración de bases anteriores guarda un respaldo y
rescata los vínculos exclusivos de SQLite. Un registro antiguo de Excel que no
tenga historial JSON se conserva como `hii_registros_heredados`, con archivo,
idioma y fecha originales; no se inventan ID, huellas ni horas de generación.
