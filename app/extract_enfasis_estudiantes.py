"""Bronze -> Silver

Lee el consolidado de estudiantes por énfasis (Geomática, Ingeniería de
Software, Teleinformática, Inteligencia Artificial) que la coordinación ya
mantiene actualizado.

El archivo trae 4 filas de categoría: "Grande - Investigación",
"Grande - Profundización", "Investigación" y "Profundización". "Grande" no
es una tercera modalidad: es la misma cohorte de Investigación/Profundización
vista con un grupo de tamaño distinto (asignatura "Seminario de Investigación
Grande" vs. grupo normal). Por eso aquí se suman fila a fila
("Grande - Investigación" + "Investigación", "Grande - Profundización" +
"Profundización") en vez de mostrarlas como categorías aparte — no se pierde
ningún estudiante, solo se deja de fragmentar la misma modalidad en dos
filas. La fila "Total" del archivo se conserva tal cual y sirve para
verificar la suma (Investigación + Profundización = Total en cada énfasis).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib.xlsx_reader import fuente, load_sheet  # noqa: E402

ARCHIVO = ROOT / "Data/Bronze/Maestria CIC/Énfasis estudiantes - Consolidado.xlsx"
SHEET_NAME = "Hoja1"
SILVER_DIR = ROOT / "Data/Silver"

ENFASIS_ROW = 4
FILA_GRANDE_INVESTIGACION = 5
FILA_GRANDE_PROFUNDIZACION = 6
FILA_INVESTIGACION = 7
FILA_PROFUNDIZACION = 8
FILA_TOTAL = 9
COLUMNS = ("C", "D", "E", "F")

# Cada fila fusionada suma su categoría "Grande" homónima con la normal.
FILAS_A_FUSIONAR = (
    ("MCIC. Énfasis en Investigación", FILA_INVESTIGACION, FILA_GRANDE_INVESTIGACION),
    ("MCIC. Énfasis en Profundización", FILA_PROFUNDIZACION, FILA_GRANDE_PROFUNDIZACION),
)


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def _leer_fila(ws, row: int, columns: tuple[str, ...]) -> dict:
    return {col: ws[f"{col}{row}"].value for col in columns}


def main() -> None:
    if not ARCHIVO.exists():
        raise FileNotFoundError(f"No existe {ARCHIVO}")

    ws = load_sheet(ARCHIVO, SHEET_NAME)
    archivo_rel = rel(ARCHIVO)

    enfasis = [ws[f"{col}{ENFASIS_ROW}"].value for col in COLUMNS]
    titulo = ws["B2"].value

    filas = []
    for etiqueta, fila_normal, fila_grande in FILAS_A_FUSIONAR:
        normal = _leer_fila(ws, fila_normal, COLUMNS)
        grande = _leer_fila(ws, fila_grande, COLUMNS)
        valores = {
            enfasis[i]: (normal[col] or 0) + (grande[col] or 0)
            for i, col in enumerate(COLUMNS)
        }
        filas.append({
            "etiqueta": etiqueta,
            "valores": valores,
            "fuente": fuente(archivo_rel, SHEET_NAME, f"{fila_normal}+{fila_grande} (fusionadas)"),
        })

    etiqueta_total = ws[f"B{FILA_TOTAL}"].value
    valores_total = {enfasis[i]: ws[f"{COLUMNS[i]}{FILA_TOTAL}"].value for i in range(len(COLUMNS))}
    filas.append({
        "etiqueta": etiqueta_total,
        "valores": valores_total,
        "fuente": fuente(archivo_rel, SHEET_NAME, FILA_TOTAL),
    })

    data = {
        "titulo_archivo": titulo,
        "enfasis": enfasis,
        "filas": filas,
        "fuente": fuente(archivo_rel, SHEET_NAME, f"{ENFASIS_ROW},{FILA_GRANDE_INVESTIGACION}-{FILA_TOTAL}"),
        "nota_fusion": (
            "Las categorías 'Grande - Investigación' y 'Grande - Profundización' del archivo fuente "
            "se sumaron a 'Investigación' y 'Profundización' respectivamente (misma modalidad, grupo de "
            "tamaño distinto) — sin pérdida de estudiantes."
        ),
    }

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "enfasis_estudiantes.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Consolidado de énfasis extraído -> {out_path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
