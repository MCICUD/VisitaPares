"""Bronze -> Silver

Lee el plan de mejoramiento anterior (ciclo de autoevaluación 2025, formato
AA-FR-001, un archivo por modalidad) y produce un JSON normalizado por
modalidad en Data/Silver/, con la misma trazabilidad (archivo, hoja, fila)
que el resto de la pipeline.

Este formato es más simple que el CC-FR-001 vigente (Data/Bronze/Maestria
CIC/2026/AUTOEVALUACION/...): no tiene columnas de "tipo" (Fortaleza/
Oportunidad), objetivo, tipo de indicador, línea base, meta, actividades,
periodicidad, apoyo requerido, responsable, recursos ni seguimiento por
cortes — solo se extrae lo que la celda realmente trae.

El contenido de los 12 factores es, celda por celda, idéntico entre el
archivo de Investigación y el de Profundización (se verificó directamente:
0 diferencias) — este plan se redactó en conjunto para las dos modalidades
y luego se guardó en dos archivos separados; no es un error de esta
pipeline.

El sitio muestra este plan bajo la etiqueta "2025-2026" a pedido del equipo
de coordinación, aunque la celda C9 de ambos archivos trae realmente
"2024 - 2026" (se conserva ese dato original en
`fecha_proyeccion_plan_original` para no perder trazabilidad).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from extract_plan_mejoramiento import separar_registro_calificado  # noqa: E402
from lib.xlsx_reader import as_iso_date, fuente, load_sheet  # noqa: E402

SHEET_NAME = "Plan de mejoramiento"
FACTOR_START_ROW = 12
FACTOR_END_ROW = 23  # 12 factores del modelo CNA

FUENTES = {
    "investigacion": ROOT / "Data/Bronze/Maestria CIC/Acreditacion/AUTOEVALUACION 2025/"
    "Docu Autoevaluacion RAA/ANEXOS/Anexos/Anexos especificos Investigacion/"
    "Plan de Mejoramiento MCIC Investigacion.xlsx",
    "profundizacion": ROOT / "Data/Bronze/Maestria CIC/Acreditacion/AUTOEVALUACION 2025/"
    "Docu Autoevaluacion RAA/ANEXOS/Anexos/Anexos especificos Profundizacion/"
    "Plan de Mejoramiento MCIC Profundizacion.xlsx",
}

SILVER_DIR = ROOT / "Data/Silver"

TIPO_NOTA = (
    "Este formato (ciclo de autoevaluación 2025, AA-FR-001) no trae una "
    "columna de clasificación Fortaleza/Oportunidad de mejora como sí trae "
    "el CC-FR-001 vigente; se muestra como 'Oportunidad de mejora' porque "
    "el documento completo es, por su propia naturaleza, un plan de "
    "mejoramiento (todas sus filas describen una acción para corregir una "
    "debilidad, ninguna es una fortaleza declarada)."
)

PERIODO_NOTA = (
    "La celda C9 del archivo fuente trae como fecha de proyección real "
    "'2024 - 2026'; en el sitio se presenta como '2025 - 2026' a pedido del "
    "equipo de coordinación de la Maestría, dejando aquí visible el dato "
    "original para no perder trazabilidad."
)


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def extract_header(ws, archivo_rel: str, modalidad: str) -> dict:
    registro_calificado, registro_calificado_nota = separar_registro_calificado(
        ws["C8"].value, modalidad
    )
    return {
        "facultad": ws["C6"].value,
        "programa_academico": ws["C7"].value,
        "registro_calificado": registro_calificado,
        "registro_calificado_nota": registro_calificado_nota,
        "registro_calificado_vigencia": ws["E8"].value,
        "acreditacion_alta_calidad": ws["H8"].value,
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
            "periodo_inicio": as_iso_date(ws[f"F{row}"].value),
            "periodo_fin": as_iso_date(ws[f"G{row}"].value),
            "peso_prioridad": ws[f"H{row}"].value,
            "indicador_cumplimiento": ws[f"I{row}"].value,
            "fuente": fuente(archivo_rel, SHEET_NAME, row),
        })
    return factores


def main() -> None:
    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    resumen = []
    for modalidad, path in FUENTES.items():
        if not path.exists():
            raise FileNotFoundError(f"No existe el archivo fuente esperado: {path}")
        archivo_rel = rel(path)
        ws = load_sheet(path, SHEET_NAME)
        data = {
            "periodo_id": "2025-2026",
            "modalidad": modalidad,
            "archivo_fuente": archivo_rel,
            "cabecera": extract_header(ws, archivo_rel, modalidad),
            "factores": extract_factores(ws, archivo_rel),
        }
        out_path = SILVER_DIR / f"plan_mejoramiento_2025_{modalidad}.json"
        out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        resumen.append((modalidad, len(data["factores"]), rel(out_path)))

    print("Extracción Bronze -> Silver completada (plan anterior, 2025-2026):")
    for modalidad, n_factores, out in resumen:
        print(f"  - {modalidad}: {n_factores} factores -> {out}")


if __name__ == "__main__":
    main()
