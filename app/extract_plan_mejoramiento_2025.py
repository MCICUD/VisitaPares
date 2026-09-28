"""Bronze -> Silver

Lee el formato anterior "AA-FR-001 Plan de Mejoramiento" (el que se radicó
ante el MEN en 2024, un solo archivo para las dos modalidades) y produce un
JSON normalizado en Data/Silver/, con la misma trazabilidad (archivo, hoja,
fila) que el resto de la pipeline.

Este archivo es más simple que el CC-FR-001 vigente (Data/Bronze/Maestria
CIC/2026/AUTOEVALUACION/...): no tiene columnas de "tipo" (Fortaleza/
Oportunidad), objetivo, tipo de indicador, línea base, meta, actividades,
periodicidad, apoyo requerido, responsable, recursos ni seguimiento por
cortes — lo que la celda no trae, se guarda como null, igual que en
extract_plan_mejoramiento.py.

El sitio muestra este plan bajo la etiqueta "2025-2026" a pedido del equipo
de coordinación, aunque la celda C9 del archivo trae realmente "2024 - 2026"
(se conserva ese dato original en `fecha_proyeccion_plan_original` para no
perder trazabilidad).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib.xlsx_reader import as_iso_date, fuente, load_sheet  # noqa: E402

SHEET_NAME = "Plan de mejoramiento"
FACTOR_START_ROW = 12
FACTOR_END_ROW = 23  # 12 factores del modelo CNA

ARCHIVO_FUENTE = (
    ROOT / "Data/Bronze/Maestria CIC/Acreditacion/PLAN DE MEJORAMIENTO 2024/"
    "Radicación a MEN/EE1345_Formato AA-FR-001-Plan de Mejoramiento-MCIC.xlsx"
)

SILVER_DIR = ROOT / "Data/Silver"

TIPO_NOTA = (
    "Este formato (AA-FR-001) no trae una columna de clasificación "
    "Fortaleza/Oportunidad de mejora como sí trae el CC-FR-001 vigente; se "
    "muestra como 'Oportunidad de mejora' porque el documento completo es, "
    "por su propia naturaleza, un plan de mejoramiento (todas sus filas "
    "describen una acción para corregir una debilidad, ninguna es una "
    "fortaleza declarada)."
)

PERIODO_NOTA = (
    "La celda C9 del archivo fuente trae como fecha de proyección real "
    "'2024 - 2026'; en el sitio se presenta como '2025 - 2026' a pedido del "
    "equipo de coordinación de la Maestría, dejando aquí visible el dato "
    "original para no perder trazabilidad."
)


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def extract_header(ws, archivo_rel: str) -> dict:
    return {
        "codigo_formato": "AA-FR-001",
        "version_formato": None,
        "macroproceso": ws["D3"].value,
        "proceso": ws["D4"].value,
        "facultad": ws["C6"].value,
        "programa_academico": ws["C7"].value,
        "registro_calificado": ws["C8"].value,
        "registro_calificado_nota": (
            "Esta celda ya trae el registro calificado de las dos "
            "modalidades en un solo texto ('Prof. .../ Inv. ...'); a "
            "diferencia del CC-FR-001 vigente, este documento no separa "
            "el plan por modalidad, así que se muestra tal cual."
        ),
        "registro_calificado_vigencia": ws["E8"].value,
        "acreditacion_alta_calidad": ws["H8"].value,
        "acreditacion_alta_calidad_vigencia": None,
        "nivel_formacion": "POSTGRADO",
        "fecha_proyeccion_plan": "2025 - 2026",
        "fecha_proyeccion_plan_original": ws["C9"].value,
        "fecha_proyeccion_plan_nota": PERIODO_NOTA,
        "fuente": fuente(archivo_rel, SHEET_NAME, "2-9 (cabecera)"),
    }


def extract_factores(ws, archivo_rel: str) -> list[dict]:
    factores = []
    for row in range(FACTOR_START_ROW, FACTOR_END_ROW + 1):
        factores.append({
            "factor": ws[f"B{row}"].value,
            "tipo": "Oportunidad de mejora",
            "tipo_nota": TIPO_NOTA,
            "origen": ws[f"C{row}"].value,
            "descripcion": ws[f"D{row}"].value,
            "proyecto": ws[f"E{row}"].value,
            "objetivo": None,
            "articulacion_plan_institucional": None,
            "periodo_inicio": as_iso_date(ws[f"F{row}"].value),
            "periodo_fin": as_iso_date(ws[f"G{row}"].value),
            "peso_prioridad": ws[f"H{row}"].value,
            "indicador_cumplimiento": ws[f"I{row}"].value,
            "tipo_indicador": None,
            "linea_base": None,
            "meta": None,
            "actividades": None,
            "periodicidad": None,
            "tipo_actividad": None,
            "apoyo_requerido": {
                "programa_academico": False,
                "facultad": False,
                "institucion": False,
            },
            "responsable": None,
            "recursos": None,
            "seguimiento": {
                "corte_1": {
                    "fecha_seguimiento": None,
                    "descripcion_avance_cualitativo": None,
                    "pct_avance_cuantitativo": None,
                    "unidad_medida_evidencia": None,
                    "otras_observaciones": None,
                },
                "corte_2": {
                    "fecha_seguimiento": None,
                    "descripcion_avance_cualitativo": None,
                    "pct_avance_cuantitativo": None,
                    "unidad_medida_evidencia": None,
                    "otras_observaciones": None,
                },
                "acumulado": {
                    "balance_cualitativo": None,
                    "pct_avance_cuantitativo": None,
                    "unidad_medida_evidencia": None,
                    "enlace_evidencia": None,
                    "otras_observaciones": None,
                },
            },
            # La evidencia de cada factor de este plan es el mismo documento
            # radicado ante el MEN (no hay evidencia de seguimiento separada
            # por factor como en el CC-FR-001 vigente).
            "fuente": fuente(archivo_rel, SHEET_NAME, row),
        })
    return factores


def main() -> None:
    if not ARCHIVO_FUENTE.exists():
        raise FileNotFoundError(f"No existe el archivo fuente esperado: {ARCHIVO_FUENTE}")

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    archivo_rel = rel(ARCHIVO_FUENTE)
    ws = load_sheet(ARCHIVO_FUENTE, SHEET_NAME)
    data = {
        "periodo_id": "2025-2026",
        "archivo_fuente": archivo_rel,
        "cabecera": extract_header(ws, archivo_rel),
        "factores": extract_factores(ws, archivo_rel),
    }
    out_path = SILVER_DIR / "plan_mejoramiento_2025_2026.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Extracción Bronze -> Silver completada:")
    print(f"  - 2025-2026 (documento único, ambas modalidades): {len(data['factores'])} factores -> {rel(out_path)}")


if __name__ == "__main__":
    main()
