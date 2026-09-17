"""Bronze -> Silver

Lee el consolidado de estudiantes por énfasis (Geomática, Ingeniería de
Software, Teleinformática, Inteligencia Artificial) que la coordinación ya
mantiene actualizado.

El archivo trae 5 filas de datos (más "Total"), cada una con sus valores
originales del archivo (nunca sumadas ni reinterpretadas): "Grande -
Investigación" / "Grande - Profundización" se muestran como "MCIC. Énfasis
en Investigación" / "MCIC. Énfasis en Profundización", y "Investigación" /
"Profundización" se muestran como "MCIC. Investigación" / "MCIC.
Profundización" — en ambos casos es el nombre institucional completo del
mismo dato del archivo, la fila no cambia.
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

# Etiqueta a mostrar por fila: None = usar la etiqueta literal del archivo (columna B).
ETIQUETA_OVERRIDE = {
    FILA_GRANDE_INVESTIGACION: "MCIC. Énfasis en Investigación",
    FILA_GRANDE_PROFUNDIZACION: "MCIC. Énfasis en Profundización",
    FILA_INVESTIGACION: "MCIC. Investigación",
    FILA_PROFUNDIZACION: "MCIC. Profundización",
}
FILAS_DATA = (FILA_GRANDE_INVESTIGACION, FILA_GRANDE_PROFUNDIZACION, FILA_INVESTIGACION, FILA_PROFUNDIZACION, FILA_TOTAL)


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
    for row in FILAS_DATA:
        etiqueta = ETIQUETA_OVERRIDE.get(row) or ws[f"B{row}"].value
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
        "fuente": fuente(archivo_rel, SHEET_NAME, f"{ENFASIS_ROW},{FILAS_DATA[0]}-{FILAS_DATA[-1]}"),
    }

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "enfasis_estudiantes.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Consolidado de énfasis extraído -> {out_path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
