# Interfaz 1.1.4

Inicia `EJECUTAR_APLICACION_ESPECIFICACIONES.bat`. El lanzador instala
CustomTkinter 5.2.2 si falta en el entorno del proyecto. El selector y el
generador HII también pueden iniciarse directamente con sus archivos Python.

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

Estos cambios preparan las rutas para el traslado. Actualmente se sigue
ejecutando con Python y sus dependencias; todavía no se ha compilado la aplicación.

## Recuperación de la base

Si falta `especificaciones.db`, la aplicación reconstruye su registro desde las
carpetas configuradas de JSON originales y alineados. También se recupera una
base existente sin registros de especificaciones cuando hay JSON disponibles:
puede ser un archivo de cero bytes, una SQLite sin tablas o con solo estructura,
catálogos y datos de control. No se utiliza un umbral de tamaño del archivo.
Una base con registros, bloqueada o ilegible no se considera vacía ni se
sustituye automáticamente. Si no hay JSON, una base vacía se conserva intacta.
La recuperación de una base vacía guarda un respaldo de la base anterior.
El gestor y los procesos que abren el registro HII también aplican este control.
La reconstrucción automática de la interfaz se ejecuta en segundo plano.

En **Configuración → Reconstruir base desde JSON** se puede repetir la recuperación
manualmente. También existe el programa independiente:

```powershell
.venv\Scripts\python.exe reconstruir_base_datos.py
.venv\Scripts\python.exe reconstruir_base_datos.py --reemplazar
```

El programa conserva nombres y contenido de los JSON y sus auxiliares WIP.
Recupera datos del producto, versiones, idiomas, huellas del Word y del catálogo
cuando existen, y relaciones bilingües explícitas o inequívocas. No inventa
parejas entre varias traducciones posibles. Una copia original y su alineada
equivalente se registran como una sola extracción y dos documentos disponibles.
Los alineados que no tienen original también permanecen registrados.

La base incluye `recuperacion_documentos` con todos los archivos válidos y
`recuperacion_actuales` con las versiones terminadas más recientes por producto,
cliente, idioma y tipo. Su orden usa fecha del documento, revisión y fecha de
creación del JSON; no usa la fecha de modificación del archivo. Este criterio
solo reconstruye el índice y no oculta ni elimina archivos del selector.

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
nuevas. Restaurar valores vuelve a los valores predeterminados.

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
