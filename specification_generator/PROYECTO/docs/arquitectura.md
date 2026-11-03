# Arquitectura del gestor de especificaciones

## Objetivo de la reorganizacion

El proyecto utiliza un layout `src`, recomendado para evitar que Python importe accidentalmente archivos desde la raiz del repositorio. La aplicacion, las pruebas, la documentacion, el empaquetado y los datos operativos tienen ubicaciones y responsabilidades distintas.

La estructura del ejemplo academico no se copia literalmente: `api/` no aplica a una interfaz de escritorio y crear carpetas vacias de modelos o repositorios solo aparentaria una separacion que el codigo aun no tiene. Esta migracion prioriza una aplicacion ejecutable y comprobable.

## Estructura actual

```text
PROYECTO/
|-- src/
|   `-- specification_generator/
|       |-- __init__.py
|       |-- __main__.py
|       |-- main.py
|       |-- interfaz_aplicacion_especificaciones.py
|       |-- interfaz_seleccion_especificaciones.py
|       |-- menu_principal.py
|       |-- gestor_especificaciones.py
|       |-- generador_hii.py
|       `-- ...
|-- tests/
|-- docs/
|-- empaquetado/
|-- logs/
|-- pyproject.toml
|-- requirements.txt
|-- .env.example
`-- README.md
```

## Responsabilidades logicas

Aunque los modulos siguen planos durante la fase de compatibilidad, sus responsabilidades ya se pueden leer como capas:

| Capa | Modulos principales | Responsabilidad |
| --- | --- | --- |
| Presentacion | `interfaz_aplicacion_especificaciones`, `interfaz_seleccion_especificaciones`, `menu_principal`, `estilo_interfaz` | Ventanas, controles y navegacion |
| Aplicacion | `editor_comportamientos`, `analisis_productos`, `generador_hii` | Casos de uso del editor y generadores |
| Dominio | `alinear_especificaciones_json`, `edicion_bilingue`, `cambios_alineacion`, `resumen_comparacion` | Reglas de negocio y transformaciones |
| Persistencia | `gestor_especificaciones`, `reconstruir_base_datos`, `cache_extraccion` | Catalogo, SQLite, recuperacion y cache |
| Infraestructura | `rutas_proyecto`, `conversion_word`, `generador_especs_hii`, `lectura_estructuras_word` | Rutas, Word, Excel y sistema operativo |

## Datos operativos

Los directorios `ESPECIFICACIONES_JSON`, `ESPECIFICACIONES_JSON_ALINEADAS`, `FORMATOS`, `METODOS_LAB` y las bases SQLite permanecen fuera de `src`. Son datos de trabajo, no codigo Python, y se excluyen de Git para evitar publicar informacion interna o archivos generados.

## Configuracion y secretos

`.env` esta ignorado. El repositorio solo contiene `.env.example`, que documenta variables sin exponer secretos. Actualmente las rutas elegidas por el usuario se guardan en `configuracion_interfaz.json`, tambien ignorado por ser configuracion local.

## Siguiente fase recomendada

La separacion fisica en `ui/`, `services/`, `repositories/`, `models/` y `utils/` debe hacerse modulo por modulo. Antes es necesario reducir los ciclos existentes entre interfaz, gestor, generador HII y menu principal. Cada extraccion debe conservar compatibilidad y ejecutar las pruebas automatizadas; no conviene mover los archivos grandes solo para obtener una estructura visual.
