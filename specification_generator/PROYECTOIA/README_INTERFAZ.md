# Interfaz 1.1.4

Inicia `EJECUTAR_APLICACION_ESPECIFICACIONES.bat`. El lanzador instala
CustomTkinter 5.2.2 si falta en el entorno del proyecto. El selector conserva
su lanzador independiente y el generador HII puede iniciarse directamente.

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
