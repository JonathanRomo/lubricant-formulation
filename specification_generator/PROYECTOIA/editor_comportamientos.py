"""Protección de notas de formato y consulta del historial fuera del editor."""
import json
import sqlite3
import tkinter as tk
from pathlib import Path


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


def siguiente_revision(codigo, carpetas, base_datos, revision_actual="00"):
    """Deduplica originales/alineados por nombre del JSON; solo consulta SQLite."""
    nombres = set()
    if Path(base_datos).exists():
        con = sqlite3.connect(Path(base_datos).resolve().as_uri() + "?mode=ro", uri=True)
        try:
            for nombre, revision in con.execute(
                "SELECT archivo_json, revision_especificacion FROM versiones WHERE part_id_final=?", (codigo,)
            ):
                nombres.add(Path(nombre).name.casefold())
        except sqlite3.OperationalError:
            # Una base nueva puede no tener todavía las tablas del extractor.
            pass
        finally:
            con.close()
    for carpeta in dict.fromkeys(Path(p) for p in carpetas):
        if not carpeta.is_dir():
            continue
        for ruta in carpeta.rglob("*"):
            if ruta.suffix.lower() != ".json":
                continue
            try:
                datos = json.loads(ruta.read_text(encoding="utf-8-sig"))
                control = datos.get("control_documental", {})
                if str(control.get("codigo_unico_producto") or "").strip() != codigo:
                    continue
                nombres.add(ruta.name.casefold())
            except (OSError, ValueError, AttributeError):
                continue
    return f"{len(nombres) + 1:02d}"
