"""Comparación manual del corpus real; usa únicamente una carpeta temporal.

--preparar conserva en TEMP la lectura anterior y conversiones de solo lectura.
--comparar vuelve a leerlas con el extractor actual. No llama a procesar(),
no actualiza SQLite y no mueve fuentes. Esta utilidad pertenece a los tests.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import multiprocessing
import queue
import sys
import time
from collections import Counter
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))
import gestor_especificaciones as gestor


def convertir_lote(casos, temporal, eventos):
    import pythoncom
    import win32com.client
    import win32process

    pythoncom.CoInitialize()
    word = None
    try:
        eventos.put(("inicio", None))
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        for caso in casos:
            documento = None
            eventos.put(("abriendo", caso["indice"]))
            try:
                documento = word.Documents.Open(
                    caso["fuente"], ConfirmConversions=False, ReadOnly=True,
                    AddToRecentFiles=False, Visible=False, OpenAndRepair=False,
                    NoEncodingDialog=True,
                )
                try:
                    _, pid = win32process.GetWindowThreadProcessId(documento.Windows.Item(1).Hwnd)
                    eventos.put(("pid", pid))
                except Exception:
                    pass
                documento.SaveAs2(str(Path(temporal) / f"{caso['indice']}.docx"), FileFormat=16)
                documento.Close(False)
                documento = None
                eventos.put(("convertido", caso["indice"]))
            except Exception as error:
                eventos.put(("error", (caso["indice"], str(error))))
            finally:
                if documento is not None:
                    documento.Close(False)
        eventos.put(("cerrando", None))
    except Exception as error:
        eventos.put(("inicio_error", str(error)))
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
    eventos.put(("fin", None))


def lectura(ruta, *, catalogo=None):
    try:
        with gestor.lectura_word_compartida():
            idioma = "en" if gestor.contiene_marcadores_ingles(ruta) else "es"
            datos = (gestor.extraer(ruta, idioma) if catalogo is None else
                     gestor.extraer(ruta, idioma, catalogo=catalogo))
            return {"estado": "OK", "datos": datos, "yves_yoi": gestor.menciones_yves_yoi(ruta)}
    except Exception as error:
        return {"estado": "ERROR", "error": f"{type(error).__name__}: {error}"}


def preparar(temporal):
    casos = []
    for grupo, carpeta in (("revisar", BASE / "REVISAR"),
                           ("buenos", BASE / gestor.ENTRADA),
                           ("buenos", BASE / gestor.ENTRADA_ENG)):
        for fuente in sorted(carpeta.rglob("*")):
            if (not fuente.is_file() or fuente.suffix.lower() not in gestor.EXTENSIONES
                    or fuente.name.startswith("~$") or "DUPLICADOS" in fuente.parts):
                continue
            casos.append({"indice": len(casos), "grupo": grupo, "fuente": str(fuente.resolve()),
                          "sha256": gestor.sha256(fuente)})
    for caso in casos:
        if Path(caso["fuente"]).suffix.lower() == ".docx":
            caso["lectura"] = caso["fuente"]
            caso["anterior"] = lectura(Path(caso["lectura"]))
    pendientes = [c for c in casos if "lectura" not in c]
    print("CORPUS", dict(Counter(c["grupo"] for c in casos)), flush=True)
    contexto = multiprocessing.get_context("spawn")
    fallos_inicio = 0
    while pendientes and fallos_inicio < 2:
        antes = gestor._procesos_word_automatizacion()
        eventos = contexto.Queue()
        proceso = contexto.Process(target=convertir_lote, args=(pendientes, str(temporal), eventos))
        proceso.start()
        limite = time.monotonic() + 45
        indice_actual = None
        pid = None
        terminado = False
        while not terminado:
            try:
                tipo, valor = eventos.get(timeout=0.2)
            except queue.Empty:
                if time.monotonic() < limite and proceso.is_alive():
                    continue
                if indice_actual is not None:
                    casos[indice_actual]["anterior"] = {"estado": "CONVERSION_ERROR", "error": "Tiempo agotado en Word"}
                else:
                    fallos_inicio += 1
                break
            limite = time.monotonic() + 45
            if tipo == "pid":
                pid = valor
            elif tipo == "abriendo":
                indice_actual = valor
            elif tipo == "convertido":
                caso = casos[valor]
                caso["lectura"] = str(temporal / f"{valor}.docx")
                caso["anterior"] = lectura(Path(caso["lectura"]))
                completados = sum("anterior" in c for c in casos)
                print(f"LEÍDOS {completados}/{len(casos)}: {Path(caso['fuente']).name}", flush=True)
            elif tipo == "error":
                indice, error = valor
                casos[indice]["anterior"] = {"estado": "CONVERSION_ERROR", "error": error}
            elif tipo == "inicio_error":
                fallos_inicio += 1
                print("INICIO WORD:", valor, flush=True)
            elif tipo == "fin":
                terminado = True
        proceso.join(timeout=2)
        if proceso.is_alive():
            proceso.terminate()
            proceso.join(timeout=5)
            gestor._limpiar_word_iniciado(antes, pid)
        eventos.close()
        eventos.cancel_join_thread()
        pendientes = [c for c in pendientes if "anterior" not in c]
    for caso in pendientes:
        caso["anterior"] = {"estado": "CONVERSION_ERROR", "error": "Word no disponible"}
    (temporal / "casos.json").write_text(json.dumps(casos, ensure_ascii=False), encoding="utf-8")
    print("ANTERIOR", dict(Counter((c["grupo"], c["anterior"]["estado"]) for c in casos)), flush=True)


def comparar(temporal):
    casos = json.loads((temporal / "casos.json").read_text(encoding="utf-8"))
    catalogo = gestor.CatalogoPreparado(gestor.leer_catalogo(BASE / gestor.CATALOGO)[0])
    resumen = Counter()
    cambios = []
    for caso in casos:
        if gestor.sha256(Path(caso["fuente"])) != caso["sha256"]:
            raise AssertionError("Fuente modificada: " + caso["fuente"])
        anterior = caso["anterior"]
        actual = lectura(Path(caso["lectura"]), catalogo=catalogo) if "lectura" in caso else anterior
        resumen[(caso["grupo"], actual["estado"])] += 1
        cambio = {"archivo": Path(caso["fuente"]).name, "grupo": caso["grupo"]}
        if actual["estado"] != "OK":
            cambio["error"] = actual.get("error")
        else:
            datos = actual["datos"]
            if any(not p["prueba"] for p in datos["pruebas"]):
                raise AssertionError("Pruebas sin nombre: " + caso["fuente"])
            previo = anterior.get("datos", {})
            for campo in ("producto", "pruebas", "control_documental", "notas", "documento"):
                if previo.get(campo) != datos.get(campo):
                    cambio[campo] = {"anterior": previo.get(campo), "actual": datos.get(campo)}
            if actual["yves_yoi"]:
                resumen[(caso["grupo"], "REGLA_YVES_YOI")] += 1
            resumen[(caso["grupo"], "PRUEBAS")] += len(datos["pruebas"])
            if any(p.get("especificacion", "").endswith("TYPICAL") for p in datos["pruebas"]):
                resumen[(caso["grupo"], "TYPICAL_LITERAL")] += 1
        if len(cambio) > 2:
            cambios.append(cambio)
    resultado = {"resumen": {"/".join(k): v for k, v in resumen.items()}, "cambios": cambios}
    (temporal / "comparacion.json").write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(resultado["resumen"], ensure_ascii=False, indent=2))
    for cambio in cambios:
        campos = {k: v for k, v in cambio.items() if k not in {"pruebas"}}
        if "pruebas" in cambio:
            campos["filas"] = {k: len(v or []) for k, v in cambio["pruebas"].items()}
        print(json.dumps(campos, ensure_ascii=False))


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("accion", choices=("preparar", "comparar"))
    parser.add_argument("temporal", type=Path)
    args = parser.parse_args()
    args.temporal.mkdir(parents=True, exist_ok=True)
    (preparar if args.accion == "preparar" else comparar)(args.temporal.resolve())
