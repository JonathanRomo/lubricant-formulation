"""Planificación de revisiones documentales sin perder el valor de origen."""
from __future__ import annotations

import copy
import hashlib
import json
import os
import re
import sqlite3
import tempfile
import zipfile
from collections import defaultdict
from contextlib import closing
from datetime import date
from datetime import datetime
from pathlib import Path

from especificaciones_wip import es_wip
from gestor_especificaciones import clave, parsear_fecha_documento


def clave_historial(datos, ruta):
    control = datos.get('control_documental') or {}
    codigo = str(control.get('codigo_unico_producto') or '').strip()
    if codigo and clave(codigo) not in {'TBD', 'PTE', 'NA', 'N A', 'SIN CODIGO', 'PENDIENTE'} and re.search(r'[A-Z0-9]', codigo, re.I):
        return 'CODIGO:' + clave(codigo)
    producto = datos.get('producto') or {}
    nombre = clave(str(producto.get('nombre') or ''))
    return 'NOMBRE:' + nombre if nombre else 'ARCHIVO:' + str(ruta).casefold()


def planificar_revisiones(carpeta):
    """Plan por JSON alineado; no escribe ni inventa parejas de idioma."""
    carpeta = Path(carpeta).resolve()
    grupos = defaultdict(list)
    for ruta in sorted(carpeta.rglob('*')):
        if not ruta.is_file() or ruta.suffix.lower() != '.json' or ruta.name.startswith('.') or es_wip(ruta):
            continue
        contenido = ruta.read_bytes()
        datos = json.loads(contenido.decode('utf-8-sig'))
        if not isinstance(datos, dict) or not isinstance(datos.get('control_documental'), dict):
            raise ValueError(f'JSON sin control documental válido: {ruta}')
        control = datos['control_documental']
        fecha = parsear_fecha_documento(str(control.get('fecha') or ''))
        grupos[clave_historial(datos, ruta)].append({'ruta': ruta, 'datos': datos, 'fecha': fecha,
                                                    'sha256': hashlib.sha256(contenido).hexdigest()})

    plan = []
    for grupo, versiones in sorted(grupos.items()):
        por_ruta = {v['ruta']: i for i, v in enumerate(versiones)}
        parejas = defaultdict(set)
        for i, version in enumerate(versiones):
            relaciones = version['datos'].get('versiones_idioma') or {}
            if not isinstance(relaciones, dict):
                continue
            for refs in relaciones.values():
                for ref in refs if isinstance(refs, list) else []:
                    if not isinstance(ref, dict) or not isinstance(ref.get('archivo_json'), str):
                        continue
                    destino = (version['ruta'].parent / ref['archivo_json']).resolve()
                    j = por_ruta.get(destino)
                    if j is None or i == j:
                        continue
                    from generador_especs_hii import leer_idioma_datos
                    a, b = version['datos'], versiones[j]['datos']
                    ca, cb = a['control_documental'], b['control_documental']
                    if leer_idioma_datos(a) != leer_idioma_datos(b) and clave(str(ca.get('nombre_cliente') or '')) == clave(str(cb.get('nombre_cliente') or '')):
                        parejas[i].add(j); parejas[j].add(i)
        from reconstruir_base_datos import contenido_comparable
        padres = list(range(len(versiones)))
        def raiz(i):
            while padres[i] != i:
                padres[i] = padres[padres[i]]
                i = padres[i]
            return i
        def unir(i, j):
            padres[raiz(j)] = raiz(i)
        firmas = {}
        for i, version in enumerate(versiones):
            comparable = contenido_comparable(version['datos'])
            if comparable is not None:
                firma = json.dumps(comparable, ensure_ascii=False, sort_keys=True, allow_nan=False)
                if firma in firmas:
                    unir(i, firmas[firma])
                else:
                    firmas[firma] = i
        for i, version in enumerate(versiones):
            opciones = parejas.get(i, set())
            if len(opciones) == 1:
                j = next(iter(opciones))
                if parejas.get(j) == {i}:
                    unir(i, j)
        agrupadas = defaultdict(list)
        for i, version in enumerate(versiones):
            agrupadas[raiz(i)].append(version)
        unidades = []
        for miembros in agrupadas.values():
            fechas = [v['fecha'] for v in miembros if v['fecha']]
            unidades.append((min(fechas) if fechas else None, miembros))
        unidades.sort(key=lambda u: (u[0] is None, u[0] or date.max,
                                     min(str(v['ruta']).casefold() for v in u[1])))
        if len(unidades) > 100:
            raise ValueError(f'El historial {grupo} excede 100 revisiones; requiere revisión manual.')
        for indice, (fecha, miembros) in enumerate(unidades):
            for version in miembros:
                datos = copy.deepcopy(version['datos']); control = datos['control_documental']
                original = control.get('revision')
                control.setdefault('revision_fuente', original if original is not None else '')
                control['revision'] = f'{indice:02d}'
                plan.append({'ruta': version['ruta'], 'grupo': grupo, 'fecha': fecha.isoformat() if fecha else '',
                             'revision_anterior': original, 'revision_nueva': control['revision'],
                             'sha256': version['sha256'], 'datos': datos})
    return plan


class IndiceRevisiones:
    """Índice en memoria para una carga; se reconstruye ante cambios externos."""
    def __init__(self, carpetas):
        self.carpetas = tuple(dict.fromkeys(Path(p).resolve() for p in carpetas))
        self._reconstruir()

    @staticmethod
    def _firma(ruta):
        try:
            info = ruta.stat()
            return info.st_mtime_ns, info.st_size
        except OSError:
            return None

    def _reconstruir(self):
        self.grupos = defaultdict(dict)
        self.archivos = {}
        self.directorios = {}
        for carpeta in self.carpetas:
            if not carpeta.is_dir():
                self.directorios[carpeta] = (None, set())
                continue
            for raiz, directorios, nombres in os.walk(carpeta):
                directorio = Path(raiz)
                self.directorios[directorio] = (self._firma(directorio), set(directorios) | set(nombres))
                for nombre in nombres:
                    ruta = directorio / nombre
                    if ruta.suffix.lower() != '.json' or es_wip(ruta):
                        continue
                    firma = self._firma(ruta)
                    self.archivos[ruta] = firma
                    try:
                        datos = json.loads(ruta.read_text(encoding='utf-8-sig'))
                        grupo = clave_historial(datos, ruta)
                        relativo = ruta.relative_to(carpeta).as_posix().casefold()
                        self.grupos[grupo][relativo] = ruta
                    except (OSError, ValueError, AttributeError):
                        continue

    def _vigente(self):
        return (all(self._firma(ruta) == firma for ruta, (firma, _) in self.directorios.items())
                and all(self._firma(ruta) == firma for ruta, firma in self.archivos.items()))

    def asignar(self, datos):
        if not self._vigente():
            self._reconstruir()
        grupo = clave_historial(datos, Path('documento_nuevo'))
        documentos = []
        mayor = -1
        for ruta in self.grupos.get(grupo, {}).values():
            anterior = json.loads(ruta.read_text(encoding='utf-8-sig'))
            documentos.append((ruta, anterior))
            valor = (anterior.get('control_documental') or {}).get('revision')
            revision = str(valor if valor is not None else '').strip()
            if re.fullmatch(r'\d{1,2}', revision):
                mayor = max(mayor, int(revision))
        from editor_comportamientos import _versiones_distintas
        siguiente = max(_versiones_distintas(documentos), mayor + 1)
        if siguiente > 99:
            raise ValueError('El historial alcanza el límite de revisión 99.')
        control = datos['control_documental']
        control.setdefault('revision_fuente', control.get('revision') or '')
        control['revision'] = f'{siguiente:02d}'
        return control['revision']

    def registrar(self, ruta, datos):
        ruta = Path(ruta).resolve()
        carpeta = next((p for p in self.carpetas if ruta.is_relative_to(p)), None)
        anterior = self.directorios.get(ruta.parent)
        if carpeta is None or anterior is None:
            self._reconstruir()
            return
        try:
            nombres = set(os.listdir(ruta.parent))
        except OSError:
            self._reconstruir()
            return
        if nombres != anterior[1] | {ruta.name}:
            self._reconstruir()
            return
        self.directorios[ruta.parent] = (self._firma(ruta.parent), nombres)
        self.archivos[ruta] = self._firma(ruta)
        self.grupos[clave_historial(datos, ruta)][ruta.relative_to(carpeta).as_posix().casefold()] = ruta


def asignar_revision_nueva(datos, carpetas, *, indice=None):
    """Asigna el siguiente número sin alterar la revisión del documento fuente."""
    return (indice or IndiceRevisiones(carpetas)).asignar(datos)


def revisar_ingresos_json(carpeta, base_datos, carpeta_respaldos):
    """Detecta ingresos y resuelve colisiones inequívocas sin renumerar el historial local."""
    from reconstruir_base_datos import contenido_comparable, _escribir_atomico
    from generador_especs_hii import leer_idioma_datos
    from rutas_proyecto import ruta_portatil

    carpeta, base_datos = Path(carpeta).resolve(), Path(base_datos).resolve()
    if not carpeta.is_dir():
        raise FileNotFoundError(f'No existe la carpeta de JSON alineados: {carpeta}')
    if not base_datos.is_file():
        raise FileNotFoundError(f'No existe la base de datos: {base_datos}')
    documentos = {}
    fotografia = {}
    avisos = []
    for ruta in sorted(carpeta.rglob('*')):
        if not ruta.is_file() or ruta.suffix.lower() != '.json' or ruta.name.startswith('.') or es_wip(ruta):
            continue
        try:
            estado_antes = ruta.stat()
            bruto = ruta.read_bytes()
            estado_despues = ruta.stat()
            firma_archivo = (estado_antes.st_mtime_ns, estado_antes.st_size)
            if firma_archivo != (estado_despues.st_mtime_ns, estado_despues.st_size):
                raise RuntimeError(f'El archivo cambió durante la lectura: {ruta}')
            fotografia[ruta] = firma_archivo
            datos = json.loads(bruto.decode('utf-8-sig'))
            if not isinstance(datos, dict) or not isinstance(datos.get('control_documental'), dict):
                raise ValueError('falta control documental válido')
            relativo = ruta.relative_to(carpeta).as_posix()
            comparable = contenido_comparable(datos)
            firma = (hashlib.sha256(json.dumps(comparable, ensure_ascii=False, sort_keys=True,
                                              separators=(',', ':'), allow_nan=False).encode('utf-8')).hexdigest()
                     if comparable is not None else None)
            if comparable is not None:
                comparable_sin_revision = copy.deepcopy(comparable)
                comparable_sin_revision['control_documental'].pop('revision', None)
                comparable_sin_revision['control_documental'].pop('revision_recibida', None)
                comparable_sin_revision['control_documental'].pop('historial_revisiones_recibidas', None)
                firma_fuente = hashlib.sha256(json.dumps(comparable_sin_revision, ensure_ascii=False, sort_keys=True,
                    separators=(',', ':'), allow_nan=False).encode('utf-8')).hexdigest()
            else:
                firma_fuente = None
            metadatos = {k: datos[k] for k in ('producto', 'documento', 'control_documental', 'versiones_idioma') if k in datos}
            documentos[relativo] = {'ruta': ruta, 'datos': metadatos,
                                    'huella': hashlib.sha256(bruto).hexdigest(),
                                    'grupo': clave_historial(datos, ruta), 'firma': firma, 'firma_fuente': firma_fuente}
        except (OSError, ValueError, TypeError, AttributeError) as exc:
            avisos.append(f'{ruta.name}: {exc}')

    resumen = {'nuevos': 0, 'duplicados': 0, 'ajustados': 0, 'pendientes': 0,
               'modificados': 0, 'avisos': avisos, 'respaldo': None, 'inicializado': False}
    con = sqlite3.connect(base_datos, timeout=30)
    ambito = ruta_portatil(carpeta, base_datos.parent)
    escritos = []
    try:
        con.execute('BEGIN IMMEDIATE')
        existe_v2 = con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='registro_json_revision_v2'").fetchone()
        legado = con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='registro_json_revision'").fetchone()
        con.execute('''CREATE TABLE IF NOT EXISTS registro_json_revision_v2 (
                       ambito TEXT NOT NULL, ruta TEXT NOT NULL, huella TEXT NOT NULL, estado TEXT NOT NULL,
                       PRIMARY KEY(ambito, ruta))''')
        con.execute('CREATE TABLE IF NOT EXISTS registro_json_revision_ambitos (ambito TEXT PRIMARY KEY)')
        if not existe_v2 and legado:
            con.execute('INSERT INTO registro_json_revision_v2 SELECT ?, ruta, huella, estado FROM registro_json_revision', (ambito,))
            con.execute('INSERT OR IGNORE INTO registro_json_revision_ambitos VALUES (?)', (ambito,))
        existe = con.execute('SELECT 1 FROM registro_json_revision_ambitos WHERE ambito=?', (ambito,)).fetchone()
        hay_otros_ambitos = con.execute('SELECT 1 FROM registro_json_revision_ambitos LIMIT 1').fetchone()
        registros = {r[0]: (r[1], r[2]) for r in con.execute('SELECT ruta, huella, estado FROM registro_json_revision_v2 WHERE ambito=?', (ambito,))}
        if not existe:
            resumen['inicializado'] = True
            try:
                conocidas = ({str(r[0]).replace('\\', '/').casefold() for r in con.execute('SELECT archivo_json FROM versiones') if r[0]}
                             if not hay_otros_ambitos else set())
            except sqlite3.OperationalError:
                conocidas = set()
            for ruta, item in documentos.items():
                estado = 'conocido' if ruta.casefold() in conocidas else 'pendiente'
                registros[ruta] = (item['huella'], estado)
                con.execute('INSERT INTO registro_json_revision_v2 VALUES (?, ?, ?, ?)', (ambito, ruta, item['huella'], estado))
                if estado == 'pendiente':
                    avisos.append(f'{ruta}: no se pudo determinar si ya era local; revisar antes de renumerar.')
                    resumen['pendientes'] += 1
            con.execute('INSERT INTO registro_json_revision_ambitos VALUES (?)', (ambito,))

        nuevos = {ruta for ruta in documentos if ruta not in registros}
        resumen['nuevos'] = len(nuevos)
        # La comparación existente incluye revisión, idioma y valores alineados.
        firmas = defaultdict(list)
        for ruta, item in documentos.items():
            if item['firma']:
                firmas[(item['grupo'], item['firma'])].append(ruta)
        duplicadas = {ruta for grupo in firmas.values() if len(grupo) > 1
                      for ruta in sorted(grupo, key=lambda r: (r in nuevos or registros.get(r, ('', ''))[1] != 'conocido', r))[1:]}
        ingresos_previos = set()
        for ruta, item in documentos.items():
            if ruta in nuevos or registros.get(ruta, ('', ''))[1] != 'conocido' or not item['firma_fuente']:
                continue
            recibida = item['datos']['control_documental'].get('revision_recibida')
            if recibida is not None:
                ingresos_previos.add((item['grupo'], str(recibida).strip(), item['firma_fuente']))
        recibidos_revisar = nuevos | {ruta for ruta, (_, estado) in registros.items() if estado == 'duplicado' and ruta in documentos}
        for ruta in recibidos_revisar:
            item = documentos[ruta]
            revision = item['datos']['control_documental'].get('revision')
            if (item['grupo'], str(revision if revision is not None else '').strip(), item['firma_fuente']) in ingresos_previos:
                duplicadas.add(ruta)
        resumen['duplicados'] = len(duplicadas)
        resumen['archivos_duplicados'] = sorted(duplicadas)
        for ruta, (huella, estado) in registros.items():
            if ruta not in documentos:
                continue
            if estado == 'duplicado' and ruta not in duplicadas:
                estado = 'pendiente'
                con.execute("UPDATE registro_json_revision_v2 SET estado='pendiente' WHERE ambito=? AND ruta=?", (ambito, ruta))
            if estado == 'pendiente' and documentos[ruta]['huella'] != huella:
                item = documentos[ruta]
                valor = item['datos']['control_documental'].get('revision')
                rev = str(valor if valor is not None else '').strip()
                colision = any(otra != ruta and candidato['grupo'] == item['grupo'] and
                               str(candidato['datos']['control_documental'].get('revision')) == rev
                               for otra, candidato in documentos.items())
                if item['grupo'].startswith('CODIGO:') and item['firma'] and re.fullmatch(r'\d{2}', rev) and not colision:
                    con.execute("UPDATE registro_json_revision_v2 SET huella=?, estado='conocido' WHERE ambito=? AND ruta=?", (item['huella'], ambito, ruta))
                    continue
            if estado == 'pendiente' and not resumen['inicializado']:
                resumen['pendientes'] += 1
                avisos.append(f'{ruta}: procedencia pendiente de revisión.')
            if documentos[ruta]['huella'] != huella:
                resumen['modificados'] += 1
                avisos.append(f'{ruta}: archivo conocido modificado o reemplazado; no se cambió su revisión ni se puede recuperar aquí el contenido anterior.')
                con.execute('UPDATE registro_json_revision_v2 SET huella=? WHERE ambito=? AND ruta=?', (documentos[ruta]['huella'], ambito, ruta))

        ocupadas = defaultdict(set)
        for ruta, item in documentos.items():
            if ruta in nuevos:
                continue
            valor_revision = item['datos']['control_documental'].get('revision')
            rev = str(valor_revision if valor_revision is not None else '').strip()
            if re.fullmatch(r'\d{1,2}', rev):
                ocupadas[item['grupo']].add(int(rev))

        # Vinculaciones explícitas entre ingresos, solo si son inequívocas.
        relaciones = defaultdict(set)
        enlaces_locales = defaultdict(set)
        enlaces_incompatibles = set()
        por_absoluta = {item['ruta'].resolve(): ruta for ruta, item in documentos.items()}
        def revision_normalizada(control):
            valor = control.get('revision')
            revision = str(valor if valor is not None else '').strip()
            return f'{int(revision):02d}' if re.fullmatch(r'\d{1,2}', revision) else revision
        for ruta in nuevos:
            item = documentos[ruta]
            for lista in (item['datos'].get('versiones_idioma') or {}).values() if isinstance(item['datos'].get('versiones_idioma') or {}, dict) else ():
                for ref in ([lista] if isinstance(lista, dict) else lista if isinstance(lista, list) else []):
                    if not isinstance(ref, dict) or not isinstance(ref.get('archivo_json'), str):
                        continue
                    otra = por_absoluta.get((item['ruta'].parent / ref['archivo_json']).resolve())
                    if otra is None:
                        enlaces_incompatibles.add(ruta)
                        for posible in nuevos:
                            if (posible != ruta and documentos[posible]['grupo'] == item['grupo'] and
                                    revision_normalizada(documentos[posible]['datos']['control_documental']) ==
                                    revision_normalizada(item['datos']['control_documental'])):
                                enlaces_incompatibles.add(posible)
                        continue
                    if otra == ruta:
                        continue
                    op = documentos[otra]
                    a, b = item['datos']['control_documental'], op['datos']['control_documental']
                    try:
                        misma_pareja = (item['grupo'] == op['grupo'] and
                            clave(str(a.get('nombre_cliente') or '')) == clave(str(b.get('nombre_cliente') or '')) and
                            leer_idioma_datos(item['datos']) != leer_idioma_datos(op['datos']))
                    except (ValueError, TypeError):
                        misma_pareja = False
                    if misma_pareja:
                        if otra in nuevos and revision_normalizada(a) == revision_normalizada(b):
                            relaciones[ruta].add(otra); relaciones[otra].add(ruta)
                        elif otra in nuevos:
                            enlaces_incompatibles.update((ruta, otra))
                        else:
                            enlaces_locales[ruta].add(otra)
                    else:
                        enlaces_incompatibles.add(ruta)
                        if otra in nuevos:
                            enlaces_incompatibles.add(otra)

        cambios = {}
        procesados = set()
        for ruta in sorted(nuevos):
            if ruta in procesados:
                continue
            item = documentos[ruta]
            miembros = [ruta]
            if len(relaciones[ruta]) == 1:
                otra = next(iter(relaciones[ruta]))
                if relaciones[otra] == {ruta}:
                    miembros.append(otra)
            procesados.update(miembros)
            if (any(len(relaciones[r]) > 1 or enlaces_locales[r] or r in enlaces_incompatibles or not documentos[r]['firma'] for r in miembros)
                    or any(r in duplicadas for r in miembros) and not all(r in duplicadas for r in miembros)
                    or not item['grupo'].startswith('CODIGO:')):
                for r in miembros:
                    con.execute('INSERT INTO registro_json_revision_v2 VALUES (?, ?, ?, ?)', (ambito, r, documentos[r]['huella'], 'pendiente'))
                    avisos.append(f'{r}: Part ID, duplicado parcial o vínculo de idioma requiere revisión manual.')
                    resumen['pendientes'] += 1
                continue
            if all(r in duplicadas for r in miembros):
                for r in miembros:
                    con.execute('INSERT INTO registro_json_revision_v2 VALUES (?, ?, ?, ?)', (ambito, r, documentos[r]['huella'], 'duplicado'))
                continue
            valor_revision = item['datos']['control_documental'].get('revision')
            rev = str(valor_revision if valor_revision is not None else '').strip()
            comparacion_incompleta = any(
                r not in nuevos and existente['grupo'] == item['grupo'] and not existente['firma']
                and str(existente['datos']['control_documental'].get('revision')) == rev
                for r, existente in documentos.items()
            )
            if comparacion_incompleta:
                for r in miembros:
                    con.execute('INSERT INTO registro_json_revision_v2 VALUES (?, ?, ?, ?)', (ambito, r, documentos[r]['huella'], 'pendiente'))
                    avisos.append(f'{r}: no se pudo comparar con una versión local de la misma revisión.')
                    resumen['pendientes'] += 1
                continue
            numero = int(rev) if re.fullmatch(r'\d{1,2}', rev) else None
            if numero is None or numero in ocupadas[item['grupo']]:
                numero = max(ocupadas[item['grupo']], default=-1) + 1
            if numero > 99:
                for r in miembros:
                    con.execute('INSERT INTO registro_json_revision_v2 VALUES (?, ?, ?, ?)', (ambito, r, documentos[r]['huella'], 'pendiente'))
                    avisos.append(f'{r}: historial sin revisiones disponibles de dos dígitos.')
                    resumen['pendientes'] += 1
                continue
            ocupadas[item['grupo']].add(numero)
            for r in miembros:
                entrada = documentos[r]
                bruto = entrada['ruta'].read_bytes()
                if hashlib.sha256(bruto).hexdigest() != entrada['huella']:
                    raise RuntimeError(f"El archivo cambió durante la revisión: {entrada['ruta']}")
                datos = json.loads(bruto.decode('utf-8-sig')); control = datos['control_documental']
                previa = control.get('revision')
                if str(previa if previa is not None else '').strip() != f'{numero:02d}':
                    historial = control.setdefault('historial_revisiones_recibidas', [])
                    if not isinstance(historial, list):
                        raise ValueError(f"Historial de revisiones recibidas inválido: {entrada['ruta']}")
                    if control.get('revision_recibida') is not None and not historial:
                        historial.append({'revision_recibida': control['revision_recibida'],
                                          'revision_asignada': previa if previa is not None else ''})
                    historial.append({'revision_recibida': previa if previa is not None else '',
                                      'revision_asignada': f'{numero:02d}'})
                    control['revision_recibida'] = previa if previa is not None else ''
                    control['revision'] = f'{numero:02d}'
                    contenido = (json.dumps(datos, ensure_ascii=False, indent=2) + '\n').encode('utf-8')
                    cambios[entrada['ruta']] = (bruto, contenido)
                    resumen['ajustados'] += 1
                    huella = hashlib.sha256(contenido).hexdigest()
                else:
                    huella = entrada['huella']
                con.execute('INSERT INTO registro_json_revision_v2 VALUES (?, ?, ?, ?)', (ambito, r, huella, 'conocido'))

        # Los conflictos entre archivos ya conocidos se informan, nunca se
        # decide cuál de ellos renumerar sin conocer su procedencia.
        por_revision = defaultdict(list)
        for ruta, item in documentos.items():
            control = item['datos']['control_documental']
            valor = control.get('revision')
            revision = str(valor if valor is not None else '').strip()
            if ruta not in nuevos and item['grupo'].startswith('CODIGO:') and re.fullmatch(r'\d{1,2}', revision):
                por_revision[(item['grupo'], int(revision))].append(ruta)
        conflictos = []
        for (grupo, revision), rutas in por_revision.items():
            if len(rutas) < 2:
                continue
            padres = {ruta: ruta for ruta in rutas}
            def raiz(ruta):
                while padres[ruta] != ruta:
                    ruta = padres[ruta]
                return ruta
            def unir(a, b):
                padres[raiz(b)] = raiz(a)
            vistas = {}
            for ruta in rutas:
                firma = documentos[ruta]['firma']
                if firma and firma in vistas:
                    unir(ruta, vistas[firma])
                elif firma:
                    vistas[firma] = ruta
            candidatos = defaultdict(set)
            for ruta in rutas:
                datos = documentos[ruta]['datos']; referencias = datos.get('versiones_idioma') or {}
                if not isinstance(referencias, dict):
                    continue
                for lista in referencias.values():
                    for ref in ([lista] if isinstance(lista, dict) else lista if isinstance(lista, list) else []):
                        if not isinstance(ref, dict) or not isinstance(ref.get('archivo_json'), str):
                            continue
                        otra = por_absoluta.get((documentos[ruta]['ruta'].parent / ref['archivo_json']).resolve())
                        if otra not in padres or otra == ruta:
                            continue
                        try:
                            d2 = documentos[otra]['datos']
                            if (leer_idioma_datos(datos) != leer_idioma_datos(d2) and
                                    clave(str(datos['control_documental'].get('nombre_cliente') or '')) ==
                                    clave(str(d2['control_documental'].get('nombre_cliente') or ''))):
                                candidatos[ruta].add(otra); candidatos[otra].add(ruta)
                        except (ValueError, TypeError):
                            pass
            for ruta, opciones in candidatos.items():
                if len(opciones) == 1:
                    otra = next(iter(opciones))
                    if candidatos[otra] == {ruta}:
                        unir(ruta, otra)
            if len({raiz(ruta) for ruta in rutas}) > 1:
                conflictos.append(f'{grupo[7:]} · revisión {revision:02d}: ' + ', '.join(rutas))
        resumen['conflictos_existentes'] = len(conflictos)
        if conflictos:
            avisos.extend('Conflicto de revisión existente: ' + texto for texto in conflictos)

        actuales = {ruta for ruta in carpeta.rglob('*') if ruta.is_file() and ruta.suffix.lower() == '.json'
                    and not ruta.name.startswith('.') and not es_wip(ruta)}
        if actuales != set(fotografia):
            raise RuntimeError('Los JSON cambiaron durante la revisión. Vuelve a pulsar actualizar.')
        for ruta, firma_archivo in fotografia.items():
            estado = ruta.stat()
            if (estado.st_mtime_ns, estado.st_size) != firma_archivo:
                raise RuntimeError(f'El archivo cambió durante la revisión: {ruta}')

        if cambios:
            Path(carpeta_respaldos).mkdir(parents=True, exist_ok=True)
            respaldo = Path(carpeta_respaldos) / ('INGRESOS_JSON_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.zip')
            with zipfile.ZipFile(respaldo, 'x', zipfile.ZIP_DEFLATED) as archivo:
                for ruta, (original, _) in cambios.items():
                    archivo.writestr(ruta.relative_to(carpeta).as_posix(), original)
            with zipfile.ZipFile(respaldo) as archivo:
                if archivo.testzip() is not None:
                    raise ValueError('No se pudo verificar el respaldo de los JSON recibidos.')
            resumen['respaldo'] = str(respaldo)
            for ruta, (original, _) in cambios.items():
                if hashlib.sha256(ruta.read_bytes()).digest() != hashlib.sha256(original).digest():
                    raise RuntimeError(f'El archivo cambió durante la revisión: {ruta}')
            for ruta, (_, nuevo) in cambios.items():
                _escribir_atomico(ruta, nuevo); escritos.append(ruta)
        con.commit()
    except BaseException:
        con.rollback()
        for ruta in reversed(escritos):
            _escribir_atomico(ruta, cambios[ruta][0])
        raise
    finally:
        con.close()
    return resumen


def aplicar_plan(plan, alineadas, anteriores, base_datos, carpeta_respaldos):
    """Respalda y actualiza JSON y SQLite; revierte escrituras si ocurre un error."""
    alineadas, anteriores = Path(alineadas).resolve(), Path(anteriores).resolve()
    base_datos = Path(base_datos).resolve()
    if not base_datos.is_file():
        raise FileNotFoundError(f'No existe la base de datos: {base_datos}')
    cambios = {}
    reporte = []
    for item in plan:
        ruta = item['ruta'].resolve()
        if hashlib.sha256(ruta.read_bytes()).hexdigest() != item['sha256']:
            raise RuntimeError(f'El archivo cambió después de crear el plan: {ruta}')
        relativo = ruta.relative_to(alineadas)
        destinos = [ruta]
        extraido = anteriores / relativo
        if extraido.is_file():
            destinos.append(extraido)
        for destino in destinos:
            original = destino.read_bytes()
            datos = json.loads(original.decode('utf-8-sig'))
            if clave_historial(datos, destino) != item['grupo']:
                raise ValueError(f'El original y el alineado no pertenecen al mismo producto: {destino}')
            control = datos['control_documental']
            control.setdefault('revision_fuente', control.get('revision') if control.get('revision') is not None else '')
            control['revision'] = item['revision_nueva']
            nuevo = (json.dumps(datos, ensure_ascii=False, indent=2) + '\n').encode('utf-8')
            cambios[destino] = (original, nuevo)
        reporte.append({k: (str(v) if isinstance(v, Path) else v) for k, v in item.items() if k != 'datos'})
    Path(carpeta_respaldos).mkdir(parents=True, exist_ok=True)
    respaldo = Path(carpeta_respaldos) / ('REVISIONES_CRONOLOGICAS_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.zip')
    con = sqlite3.connect(base_datos, timeout=30)
    db_backup = None
    escritos = []
    try:
        con.execute('BEGIN IMMEDIATE')
        with tempfile.NamedTemporaryFile(suffix='.db', delete=False) as temporal:
            db_backup = Path(temporal.name)
        with closing(sqlite3.connect(base_datos)) as origen, closing(sqlite3.connect(db_backup)) as copia:
            origen.backup(copia)
        with zipfile.ZipFile(respaldo, 'x', zipfile.ZIP_DEFLATED) as archivo:
            for ruta, (original, _) in cambios.items():
                prefijo, raiz = ('alineadas', alineadas) if ruta.is_relative_to(alineadas) else ('anteriores', anteriores)
                archivo.writestr(prefijo + '/' + ruta.relative_to(raiz).as_posix(), original)
            archivo.write(db_backup, 'especificaciones.db')
            archivo.writestr('plan.json', json.dumps(reporte, ensure_ascii=False, indent=2))
        with zipfile.ZipFile(respaldo) as archivo:
            if archivo.testzip() is not None:
                raise ValueError('El respaldo no superó la verificación.')
        for ruta, (original, _) in cambios.items():
            if hashlib.sha256(ruta.read_bytes()).digest() != hashlib.sha256(original).digest():
                raise RuntimeError(f'El archivo cambió durante la preparación: {ruta}')
        for ruta, (original, nuevo) in cambios.items():
            nombre = None
            try:
                with tempfile.NamedTemporaryFile(dir=ruta.parent, prefix='.revision_', suffix='.tmp', delete=False) as f:
                    nombre = Path(f.name); f.write(nuevo)
                os.replace(nombre, ruta); escritos.append(ruta)
            finally:
                if nombre and nombre.exists(): nombre.unlink()
        for item in plan:
            relativo = item['ruta'].resolve().relative_to(alineadas).as_posix()
            actualizado = con.execute('UPDATE versiones SET revision_especificacion=? WHERE replace(archivo_json, ?, ?) = ?',
                                      (item['revision_nueva'], '\\', '/', relativo))
            if actualizado.rowcount != 1:
                raise ValueError(f'No hay exactamente una versión en SQLite para {relativo}.')
        con.execute('''UPDATE archivos SET revision_especificacion=(SELECT v.revision_especificacion
                       FROM versiones v WHERE v.archivo_id=archivos.id AND v.archivo_json=archivos.ultimo_json LIMIT 1)
                       WHERE EXISTS (SELECT 1 FROM versiones v WHERE v.archivo_id=archivos.id AND v.archivo_json=archivos.ultimo_json)''')
        if con.execute('PRAGMA integrity_check').fetchone()[0] != 'ok':
            raise RuntimeError('La base no superó la verificación de integridad.')
        con.commit()
    except BaseException:
        con.rollback()
        for ruta in reversed(escritos):
            with tempfile.NamedTemporaryFile(dir=ruta.parent, prefix='.revision_revertir_', suffix='.tmp', delete=False) as f:
                temporal = Path(f.name); f.write(cambios[ruta][0])
            os.replace(temporal, ruta)
        raise
    finally:
        con.close()
        if db_backup: db_backup.unlink(missing_ok=True)
    return respaldo, reporte
