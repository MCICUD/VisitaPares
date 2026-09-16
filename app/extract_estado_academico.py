"""Bronze -> Silver

Lee los 7 listados oficiales de estudiantes por proyecto curricular
(sistema Cóndor, exportados a Data/Bronze/Estados/) y calcula SOLO
conteos agregados por estado académico.

Estos CSV tienen nombre, documento de identidad y correo personal de cada
estudiante por fila — por eso esta capa Silver/Gold nunca guarda ni expone
esas columnas, únicamente cuenta filas. Los archivos originales se dejan
fuera del catálogo público de Data/Bronze (ver EXCLUIR_DEL_CATALOGO en
app/build_bronze_manifest.py) para no publicar esos datos personales; la
cifra siempre cita el nombre del archivo para que la coordinación (que sí
tiene acceso al original) pueda corroborarla.
"""
from __future__ import annotations

import csv
import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ESTADOS_DIR = ROOT / "Data/Bronze/Estados"
SILVER_DIR = ROOT / "Data/Silver"

# Columnas del CSV por posición (el encabezado repite "Estado" dos veces;
# csv.DictReader se quedaría solo con la segunda, así que leemos por índice).
COL_PROYECTO_CURRICULAR = 6
COL_COD_PROYECTO = 5
COL_DESCRIPCION = 8


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def leer_csv(path: Path) -> list[list[str]]:
    with path.open(encoding="utf-8-sig") as fh:
        lines = fh.readlines()
    # la primera línea es un título suelto ("Listado de estudiantes por estado"),
    # la segunda es el encabezado real.
    reader = csv.reader(lines[1:])
    next(reader)  # descartar encabezado
    return [row for row in reader if any(cell.strip() for cell in row)]


def main() -> None:
    if not ESTADOS_DIR.exists():
        raise FileNotFoundError(f"No existe {ESTADOS_DIR}")

    proyectos = []
    for path in sorted(ESTADOS_DIR.glob("Listado_de_estudiantes_por_estado_*.csv")):
        filas = leer_csv(path)
        if not filas:
            continue
        cod_proyecto = filas[0][COL_COD_PROYECTO].strip()
        nombre_proyecto = filas[0][COL_PROYECTO_CURRICULAR].strip()
        conteo_por_estado = Counter(row[COL_DESCRIPCION].strip() for row in filas)
        proyectos.append({
            "cod_proyecto": cod_proyecto,
            "proyecto_curricular": nombre_proyecto,
            "total_estudiantes": len(filas),
            "conteo_por_estado": dict(conteo_por_estado),
            "fuente": {
                "archivo": rel(path),
                "nota": (
                    "Archivo con nombre/documento/correo por estudiante; no se publica "
                    "en el catálogo de Data/Bronze. Esta cifra es un conteo de filas."
                ),
            },
        })

    data = {"proyectos": proyectos}
    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "estado_academico_agregado.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Estado académico agregado ({len(proyectos)} proyectos curriculares) -> {out_path.relative_to(ROOT)}")
    for p in proyectos:
        print(f"  - Cod.{p['cod_proyecto']} {p['proyecto_curricular']}: {p['total_estudiantes']} registros")


if __name__ == "__main__":
    main()
