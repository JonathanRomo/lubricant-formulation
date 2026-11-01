"""Protección de notas de formato y consulta del historial fuera del editor."""
import json
import re
import tkinter as tk


def es_nota_formato(texto):
    texto = " ".join(texto.strip().lstrip("-• ").casefold().split())
    return texto.startswith((
        "the active version of the applicable method shall be",
        "la versión vigente del método aplicable corresponderá",
    ))


class NotasProtegidas(tk.Text):
    def __init__(self, padre, **kwargs):
        super().__init__(padre, **kwargs)
        self._interno = False
        self._desbloqueadas = False
        self._original = self._w + "_original"
        self.tk.call("rename", self._w, self._original)
        self.tk.createcommand(self._w, self._proxy)
        self.tag_configure("nota_formato", foreground="#9aa0a6")
        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label="Editar nota del formato", command=self._desbloquear)
        self.bind("<Button-3>", self._menu)
        self.bind("<<Copy>>", self._copiar)
        self.bind("<Control-c>", self._copiar)
        self.bind("<Control-C>", self._copiar)
        self.bind("<<Paste>>", self._pegar)

    def _copiar(self, _evento=None):
        # Lee la selección sin pasar por el manejador de copia nativo de Text.
        seleccion = self.tag_ranges("sel")
        if not seleccion:
            return "break"
        contenido = self.tk.call(self._original, "get", seleccion[0], seleccion[1])
        try:
            self.clipboard_clear()
            self.clipboard_append(contenido)
        except tk.TclError:
            self.bell()
        return "break"

    def _pegar(self, _evento=None):
        try:
            contenido = self.clipboard_get()
        except tk.TclError:
            return "break"
        # Usa una sola edición: el pegado nativo elimina e inserta por separado.
        # Así una selección que cruza una nota protegida se conserva completa.
        contenido = contenido.replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "")
        if not contenido or not self.winfo_exists():
            return "break"
        seleccion = self.tag_ranges("sel")
        if seleccion:
            self.tk.call(self._w, "replace", seleccion[0], seleccion[1], contenido)
        else:
            self.tk.call(self._w, "insert", "insert", contenido)
        self.see("insert")
        return "break"

    def _menu(self, evento):
        self.menu.tk_popup(evento.x_root, evento.y_root)
        self.menu.grab_release()
        return "break"

    def _desbloquear(self):
        self._desbloqueadas = True
        self.tag_configure("nota_formato", foreground=self.cget("foreground"))
        self.focus_set()

    def _marcar(self):
        self.tag_remove("nota_formato", "1.0", "end")
        for fila, linea in enumerate(self.get("1.0", "end-1c").splitlines(), 1):
            if es_nota_formato(linea):
                self.tag_add("nota_formato", f"{fila}.0", f"{fila}.end")

    def _proxy(self, operacion, *args):
        if operacion in {"insert", "delete", "replace"} and not self._interno and not self._desbloqueadas:
            inicio = self.index(args[0])
            fin = inicio if operacion == "insert" else self.index(args[1] if len(args) > 1 else inicio + "+1c")
            rangos = self.tag_ranges("nota_formato")
            for a, b in zip(rangos[::2], rangos[1::2]):
                if operacion == "insert":
                    bloqueada = self.compare(inicio, ">=", a) and self.compare(inicio, "<", b)
                    bloqueada |= self.compare(inicio, "==", b) and not str(args[1]).startswith("\n")
                else:
                    bloqueada = self.compare(inicio, "<", b) and self.compare(fin, ">", a)
                    # Impide unir la nota protegida a la línea anterior.
                    bloqueada |= self.compare(fin, "==", a) and self.compare(inicio, "<", fin)
                    bloqueada |= self.compare(inicio, "==", b) and self.compare(fin, ">", b)
                if bloqueada:
                    self.bell()
                    return ""
        resultado = self.tk.call(self._original, operacion, *args)
        if operacion in {"insert", "delete", "replace"}:
            self._marcar()
        return resultado

    def insert(self, *args):
        self._interno = True
        try:
            return super().insert(*args)
        finally:
            self._interno = False

    def delete(self, *args):
        self._interno = True
        try:
            return super().delete(*args)
        finally:
            self._interno = False

    def destroy(self):
        self.tk.deletecommand(self._w)
        super().destroy()


class RevisionCalculada(str):
    """Consecutivo compatible con texto, con su explicación solo en memoria."""
    def __new__(cls, valor, explicacion):
        resultado = super().__new__(cls, valor)
        resultado.explicacion = explicacion
        return resultado


def _versiones_distintas(documentos):
    """Reutiliza el verificador sin mover archivos ni modificar el historial."""
    from reconstruir_base_datos import contenido_comparable
    from generador_especs_hii import leer_idioma_datos
    from rutas_proyecto import resolver_ruta

    padres = list(range(len(documentos)))
    def raiz(indice):
        while padres[indice] != indice:
            padres[indice] = padres[padres[indice]]
            indice = padres[indice]
        return indice
    def unir(a, b):
        padres[raiz(b)] = raiz(a)

    por_contenido = {}
    por_ruta = {ruta.resolve(): i for i, (ruta, _) in enumerate(documentos)}
    for i, (_, datos) in enumerate(documentos):
        comparable = contenido_comparable(datos)
        if comparable is None:
            continue  # Sin evidencia suficiente, no fusionar documentos.
        try:
            firma = json.dumps(comparable, sort_keys=True, ensure_ascii=False, allow_nan=False)
        except (ValueError, TypeError):
            continue
        if firma in por_contenido:
            unir(i, por_contenido[firma])
        else:
            por_contenido[firma] = i

    def revision(datos):
        valor = str(datos["control_documental"].get("revision", "")).strip()
        return str(int(valor)) if valor.isdigit() else valor
    def cliente(datos):
        return " ".join(str(datos["control_documental"].get("nombre_cliente") or "").casefold().split())

    traducciones = set()
    for i, (ruta, datos) in enumerate(documentos):
        relaciones = datos.get("versiones_idioma", {})
        if not isinstance(relaciones, dict):
            continue
        for referencias in relaciones.values():
            if isinstance(referencias, dict):
                referencias = [referencias]
            if not isinstance(referencias, list):
                continue
            for referencia in referencias:
                if not isinstance(referencia, dict) or not isinstance(referencia.get("archivo_json"), str):
                    continue
                j = por_ruta.get(resolver_ruta(referencia["archivo_json"], ruta.parent))
                if j is None:
                    continue
                otra = documentos[j][1]
                if (revision(datos) == revision(otra) and cliente(datos) == cliente(otra)
                        and leer_idioma_datos(datos) != leer_idioma_datos(otra)):
                    traducciones.add(tuple(sorted((raiz(i), raiz(j)))))
    candidatas = {}
    for a, b in traducciones:
        candidatas.setdefault(a, set()).add(b)
        candidatas.setdefault(b, set()).add(a)
    for a, b in traducciones:
        if len(candidatas[a]) == len(candidatas[b]) == 1:
            unir(a, b)
    return len({raiz(i) for i in range(len(documentos))})


def siguiente_revision(codigo, carpetas, base_datos=None, revision_actual="00", *, exigir_historial=False):
    """Calcula el consecutivo desde las revisiones terminadas del Part ID.

    Los parámetros de SQLite y revisión inicial se conservan para llamadas antiguas;
    el cálculo usa exclusivamente los JSON actuales de estas carpetas. Los años
    de origen identifican versiones, pero nunca se incrementan como consecutivos.
    """
    from rutas_proyecto import resolver_ruta
    from especificaciones_wip import es_wip

    codigo = str(codigo).strip().casefold()
    documentos = {}
    consecutivos = set()
    invalidos = []
    for carpeta in dict.fromkeys(resolver_ruta(p) for p in carpetas):
        if not carpeta.exists():
            continue
        for ruta in carpeta.rglob("*"):
            if not ruta.is_file() or ruta.suffix.lower() != ".json" or es_wip(ruta):
                continue
            try:
                datos = json.loads(ruta.read_text(encoding="utf-8-sig"))
            except (OSError, ValueError):
                continue
            if not isinstance(datos, dict):
                continue
            control = datos.get("control_documental", {})
            if not isinstance(control, dict):
                continue
            if str(control.get("codigo_unico_producto") or "").strip().casefold() != codigo:
                continue
            revision = str(control.get("revision", "")).strip()
            if re.fullmatch(r"[0-9]{1,2}", revision):
                consecutivos.add(int(revision))
            elif re.fullmatch(r"(?:19|20)[0-9]{2}", revision):
                pass  # El año es válido como revisión de origen, no como consecutivo.
            else:
                invalidos.append(ruta.name)
            documentos[ruta.resolve()] = datos
    if invalidos:
        raise ValueError("Hay alineados del Part ID con revisión vacía o inválida. "
                         "Corrige estos archivos antes de calcular la siguiente revisión: " + ", ".join(sorted(invalidos)))
    if not documentos:
        return RevisionCalculada("00", "No hay revisiones terminadas para este Part ID; la primera es 00.")
    # El conteo permite comenzar tras históricos con año. Si hay consecutivos
    # ya emitidos, nunca retroceder ni reutilizar uno por huecos en el historial.
    cantidad = _versiones_distintas(list(documentos.items()))
    siguiente = max(cantidad, max(consecutivos, default=-1) + 1)
    if siguiente > 99:
        raise ValueError("El historial alcanza el límite 99; no se puede asignar otra revisión de dos dígitos automáticamente.")
    cantidad_texto = "Se encontró 1 revisión distinta" if cantidad == 1 else f"Se encontraron {cantidad} revisiones distintas"
    explicacion = (f"{cantidad_texto} del Part ID. "
                   "Solo se descuentan copias confirmadas por contenido y traducciones con vínculo confirmado. "
                   "Un mismo número con contenido diferente cuenta por separado.\n")
    if consecutivos:
        ultima = max(consecutivos)
        explicacion += f"La mayor revisión numérica es {ultima:02d}; su consecutivo es {ultima + 1:02d}."
        if siguiente > ultima + 1:
            explicacion += f" El historial completo tiene {cantidad} versiones, por lo que corresponde {siguiente:02d}."
    else:
        explicacion += (f"Las revisiones encontradas son años del documento de origen. "
                        f"Por el conteo del historial corresponde {siguiente:02d}; el año no se incrementa.")
    return RevisionCalculada(f"{siguiente:02d}", explicacion)
