# Generador HII

En el gestor, pulsa **GENERAR HII**. Busca por código, nombre, cliente o revisión
y marca las casillas de las especificaciones alineadas. La lista se compacta
automáticamente y muestra las seleccionadas; **Ampliar lista** permite agregar
más. La barra espaciadora también marca o desmarca una fila.

Al pulsar una especificación marcada se muestra su resumen con tabla de pruebas
y control documental junto a la vista previa del Excel. Los separadores permiten
ajustar el ancho de ambos paneles. Guarda la hoja actual con **Guardar HII (.xlsm)**
o todo el lote con **Guardar seleccionadas**. Este último usa el nombre del JSON
e idioma para distinguir revisiones y pide confirmación antes de reemplazar
archivos existentes. También puede iniciarse directamente:

```powershell
.venv\Scripts\python.exe generador_hii.py
```

Se elige `91400005` para español y `91400019` para inglés en la carpeta de
formatos configurada. Se conserva el idioma del JSON; no se traduce contenido.

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

Los botones **Español**, **English** e **Idioma original** eligen el idioma para
la vista previa y para el guardado individual o de las seleccionadas. Se busca
una especificación alineada con el mismo código, revisión y cliente en el idioma
solicitado. Si hay varias candidatas se pide elegir, y esa relación se recuerda.

Si falta la versión se abre una copia en el editor, con un nombre nuevo y el
idioma solicitado. El editor ayuda con nombres del maestro; el usuario debe
terminar la traducción de valores, notas y controles. La casilla de revisión es
obligatoria para guardar esta traducción. Al guardar se vinculan las versiones
y se actualiza la previsualización HII. Si se cierra sin guardar, no se genera
una HII en el idioma pendiente.

Las tablas `hii_especificaciones`, `hii_traducciones` y `hii_generadas` se crean
de forma aditiva en `especificaciones.db`; conservan las relaciones y los archivos
HII guardados. La vista `hii_versiones_idiomas` permite consultar cada HII junto
a las HII guardadas de su traducción. No se modifican las tablas del extractor.
