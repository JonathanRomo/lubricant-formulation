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


def siguiente_revision(codigo, carpetas, base_datos=None, revision_actual="00", *, exigir_historial=False):
    """Busca la mayor revisión terminada en las carpetas de alineados indicadas.

    Los parámetros de SQLite y revisión inicial se conservan para llamadas antiguas;
    el cálculo usa exclusivamente los JSON actuales de estas carpetas.
    """
    from rutas_proyecto import resolver_ruta
    from especificaciones_wip import es_wip

    codigo = str(codigo).strip().casefold()
    revisiones = set()
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
            if not re.fullmatch(r"[0-9]{1,2}", revision):
                invalidos.append(ruta.name)
            else:
                revisiones.add(int(revision))
    if invalidos:
        raise ValueError("Hay alineados del Part ID con revisión vacía o inválida. "
                         "Corrige estos archivos antes de calcular la siguiente revisión: " + ", ".join(sorted(invalidos)))
    if not revisiones:
        return "00"
    siguiente = max(revisiones) + 1
    if siguiente > 99:
        raise ValueError("La última revisión es 99; no se puede asignar otra de dos dígitos automáticamente.")
    return f"{siguiente:02d}"
