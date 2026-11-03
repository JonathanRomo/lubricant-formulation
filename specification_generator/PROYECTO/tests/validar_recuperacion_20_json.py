"""Prueba aislada con 20 JSON reales; solo lee los originales del proyecto."""
import hashlib
import json
import shutil
import sqlite3
import sys
import tempfile
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from rutas_proyecto import BASE
from reconstruir_base_datos import asegurar_base, motivo_recuperacion
import gestor_especificaciones as gestor


def huella(ruta):
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


def main():
    seleccion = []
    for carpeta in ("ESPECIFICACIONES_JSON", "ESPECIFICACIONES_JSON_ALINEADAS"):
        seleccion.extend(sorted(p for p in (BASE / carpeta).rglob("*") if p.suffix.lower() == ".json" and not p.name.startswith("."))[:10])
    if len(seleccion) != 20:
        raise RuntimeError("Se necesitan diez JSON originales y diez alineados para esta prueba.")
    originales = {p: huella(p) for p in seleccion}
    base_real = BASE / "especificaciones.db"
    huella_base_real = huella(base_real) if base_real.exists() else None
    with tempfile.TemporaryDirectory(prefix="recuperacion_20_json_") as temporal:
        base = Path(temporal).resolve()
        copias = []
        for ruta in seleccion:
            destino = base / ruta.relative_to(BASE)
            destino.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(ruta, destino)
            copias.append(destino)
            auxiliar = ruta.with_name("." + ruta.name + ".wip")
            if auxiliar.is_file():
                shutil.copy2(auxiliar, destino.with_name(auxiliar.name))
        if (BASE / "CODIGOS_200.xlsx").exists():
            shutil.copy2(BASE / "CODIGOS_200.xlsx", base / "CODIGOS_200.xlsx")
        antes = {p: huella(p) for p in copias}
        (base / "especificaciones.db").touch()
        assert motivo_recuperacion(base) == "vacia"
        resultado = asegurar_base(base)
        assert resultado["motivo"] == "vacia" and resultado["respaldo"] is not None
        assert resultado["documentos"] == 20, resultado
        assert resultado["omitidos"] == [], resultado
        assert {p: huella(p) for p in copias} == antes, "La reconstrucción modificó una copia JSON."
        with closing(sqlite3.connect(base / "especificaciones.db")) as con:
            assert con.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
            assert con.execute("PRAGMA foreign_key_check").fetchall() == []
            archivos = {Path(f[0]).name for f in con.execute("SELECT archivo_json FROM versiones")}
            assert archivos == {p.name for p in copias}, "Se alteraron los nombres registrados."
            actuales = con.execute("SELECT COUNT(*) FROM archivos").fetchone()[0]

        # Copiar solo Word presentes cuya huella y catálogo coincidan con un registro
        # recuperado. No se ejecuta Word: cualquier intento de extraer hace fallar la prueba.
        words = set()
        digest_catalogo = huella(base / "CODIGOS_200.xlsx") if (base / "CODIGOS_200.xlsx").exists() else ""
        with closing(sqlite3.connect(base / "especificaciones.db")) as con:
            for ruta, digest, catalogo in con.execute("SELECT ruta_relativa, hash_actual, hash_catalogo FROM archivos WHERE estado='OK'"):
                origen = Path(ruta)
                if origen.is_absolute() or ".." in origen.parts:
                    continue
                real = BASE / origen
                if real.is_file() and huella(real) == digest and catalogo == digest_catalogo:
                    copia = base / origen
                    copia.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(real, copia)
                    words.add(copia)
        if words:
            with patch.object(gestor, "preparar", side_effect=AssertionError("Se intentó regenerar un Word recuperado")), patch("builtins.print"):
                procesamiento = gestor.procesar(base, archivos_seleccionados=sorted(words))
            assert procesamiento["ignorados"] == len(words), procesamiento
            assert procesamiento["nuevos"] == procesamiento["actualizados"] == procesamiento["errores"] == 0, procesamiento
        else:
            procesamiento = None
        assert {p.name for p in copias} == archivos
        assert {p: huella(p) for p in seleccion} == originales, "Se modificaron los JSON reales."
        assert (huella(base_real) if base_real.exists() else None) == huella_base_real, "Se modificó la base real."
        reporte = {"json_probados": 20, "base_vacia_detectada_y_recuperada": True, "respaldo_creado": resultado["respaldo"] is not None,
            "versiones_recuperadas": resultado["versiones"], "relaciones": resultado["relaciones"],
            "actuales": actuales, "wip": resultado["wip"], "omitidos": resultado["omitidos"],
            "nombres_conservados": True, "contenido_conservado_durante_reconstruccion": True,
            "archivos_reales_y_base_real_sin_cambios": True, "words_omitidos_sin_regenerar": len(words),
            "procesamiento": procesamiento, "integridad_sqlite": "ok",
            "muestra": [p.relative_to(BASE).as_posix() for p in seleccion]}
        (BASE / "tests" / "REPORTE_RECUPERACION_20_JSON.json").write_text(json.dumps(reporte, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(reporte, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
