"""Bronze -> Silver

Calcula SOLO conteos agregados de egresados (por énfasis y por año de
grado) a partir de un consolidado que trae nombre, documento y correo por
egresado. Igual que con el roster de Estados, ese archivo se guarda en
Data/Bronze/PII_Interno/ (excluido del catálogo público) y aquí solo se
cuenta — nunca se publican filas individuales.

Cobertura conocida: el archivo fuente está fechado en 2023 y no incluye
graduaciones de 2024-2026 (no hay uno más reciente disponible todavía);
esa limitación se deja explícita en el resultado, no se rellena con nada.
"""
from __future__ import annotations

import datetime
import json
import sys
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SILVER_DIR = ROOT / "Data/Silver"
ARCHIVO = ROOT / "Data/Bronze/PII_Interno/Informacion egresados 2023.xlsx"

ANIO_MIN, ANIO_MAX = 2022, 2026  # alcance solicitado para lo que se presenta en el sitio

# Las hojas 195/295/395/495 son los códigos de proyecto curricular por énfasis
# (mismos códigos que Data/Bronze/Estados) — no se renombran ni se reinterpretan.
HOJAS_ENFASIS = ("195", "295", "395", "495")


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def main() -> None:
    if not ARCHIVO.exists():
        raise FileNotFoundError(f"No existe {ARCHIVO}")

    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = openpyxl.load_workbook(ARCHIVO, data_only=True, read_only=True)

    por_enfasis = []
    total_general = 0
    por_anio_general: dict[int, int] = {}

    for nombre_hoja in HOJAS_ENFASIS:
        if nombre_hoja not in wb.sheetnames:
            continue
        ws = wb[nombre_hoja]
        filas = list(ws.iter_rows(values_only=True))[1:]  # saltar encabezado
        promedios = []
        por_anio: dict[int, int] = {}
        total_hoja = 0
        for fila in filas:
            if not fila or fila[0] is None:
                continue
            fecha_grado = fila[3] if len(fila) > 3 else None
            if not isinstance(fecha_grado, datetime.datetime):
                continue
            if not (ANIO_MIN <= fecha_grado.year <= ANIO_MAX):
                continue
            total_hoja += 1
            por_anio[fecha_grado.year] = por_anio.get(fecha_grado.year, 0) + 1
            por_anio_general[fecha_grado.year] = por_anio_general.get(fecha_grado.year, 0) + 1
            promedio_raw = fila[9] if len(fila) > 9 else None
            try:
                promedios.append(float(str(promedio_raw).replace(",", ".")))
            except (TypeError, ValueError):
                pass

        total_general += total_hoja
        por_enfasis.append({
            "cod_proyecto": nombre_hoja,
            "total_graduados": total_hoja,
            "por_anio": dict(sorted(por_anio.items())),
            "promedio_academico": round(sum(promedios) / len(promedios), 2) if promedios else None,
        })

    data = {
        "rango_presentado": f"{ANIO_MIN}-{ANIO_MAX}",
        "total_graduados": total_general,
        "por_anio": dict(sorted(por_anio_general.items())),
        "por_enfasis": por_enfasis,
        "limitacion": (
            "El archivo fuente está fechado en 2023 y no incluye graduaciones "
            "posteriores (2024-2026); no existe todavía un consolidado más reciente."
        ),
        "fuente": {
            "archivo": rel(ARCHIVO),
            "nota": "Archivo con nombre/documento/correo por egresado; no se publica en el catálogo de Data/Bronze.",
        },
    }

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "egresados_agregado.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Egresados agregados ({total_general} graduados {ANIO_MIN}-{ANIO_MAX}) -> {out_path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
