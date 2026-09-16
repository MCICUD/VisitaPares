"""Bronze -> Silver

Lee el consolidado de estudiantes por énfasis (Geomática, Ingeniería de
Software, Teleinformática, Inteligencia Artificial) que la coordinación ya
mantiene actualizado. Se transcribe tal cual: las mismas etiquetas de fila
que trae el archivo (no se reinterpretan ni renombran categorías).
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
DATA_ROWS = (5, 6, 7, 8, 9)  # Grande-Investigación, Grande-Profundización, Investigación, Profundización, Total
COLUMNS = ("C", "D", "E", "F")


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def main() -> None:
    if not ARCHIVO.exists():
        raise FileNotFoundError(f"No existe {ARCHIVO}")

    ws = load_sheet(ARCHIVO, SHEET_NAME)
    archivo_rel = rel(ARCHIVO)

    enfasis = [ws[f"{col}{ENFASIS_ROW}"].value for col in COLUMNS]
    titulo = ws["B2"].value

    filas = []
    for row in DATA_ROWS:
        etiqueta = ws[f"B{row}"].value
        valores = {enfasis[i]: ws[f"{COLUMNS[i]}{row}"].value for i in range(len(COLUMNS))}
        filas.append({
            "etiqueta": etiqueta,
            "valores": valores,
            "fuente": fuente(archivo_rel, SHEET_NAME, row),
        })

    data = {
        "titulo_archivo": titulo,
        "enfasis": enfasis,
        "filas": filas,
        "fuente": fuente(archivo_rel, SHEET_NAME, f"{ENFASIS_ROW},{DATA_ROWS[0]}-{DATA_ROWS[-1]}"),
    }

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "enfasis_estudiantes.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Consolidado de énfasis extraído -> {out_path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
