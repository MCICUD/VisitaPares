"""Bronze -> Silver

Lee el plan de mejoramiento anterior (2024-2026, ciclo de autoevaluación
2025, formato AA-FR-001, un archivo por modalidad) y produce un JSON normalizado por
modalidad en Data/Silver/, con la misma trazabilidad (archivo, hoja, fila)
que el resto de la pipeline.

Este formato es más simple que el CC-FR-001 vigente (Data/Bronze/Maestria
CIC/2026/AUTOEVALUACION/...): no tiene columnas de "tipo" (Fortaleza/
Oportunidad), objetivo, tipo de indicador, periodicidad ni seguimiento por
cortes. Sí trae línea base (col. J), meta (col. K) y descripción de las
actividades (col. L), que se extraen tal cual.

En la fila del FACTOR 2 las columnas K y L vienen intercambiadas en el
archivo fuente (K trae las actividades "a) ... b) ..." y L la meta); se
detecta y se corrige dejando constancia en `meta_nota`.

Las filas de los FACTORES 1, 7 y 8 difieren en J/K/L entre el archivo de
Investigación y el de Profundización; el resto es idéntico.

El periodo del plan es el que trae la celda C9 de ambos archivos
("2024 - 2026"); junto con el plan vigente (2026-2027) cubre 2024-2027.
"""
from __future__ import annotations

import json
import re
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

META_INTERCAMBIADA_NOTA = (
    "En el archivo fuente las columnas META (K) y DESCRIPCIÓN DE LAS "
    "ACTIVIDADES (L) de este factor vienen intercambiadas; aquí se muestran "
    "en el campo que les corresponde."
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
        "fecha_proyeccion_plan": ws["C9"].value,
        "fuente": fuente(archivo_rel, SHEET_NAME, "2-9 (cabecera)"),
    }


def _limpiar(value):
    if not isinstance(value, str):
        return value
    # Las actividades vienen como "a) ...     b) ..." separadas por espacios.
    value = re.sub(r"(?<=[.;\s])\s*(?=[b-h]\)\s?)", "\n", value)
    lineas = [" ".join(linea.split()) for linea in value.splitlines()]
    texto = "\n".join(linea for linea in lineas if linea)
    return re.sub(r" {2,}", " ", texto).strip()


def _parece_lista_actividades(value) -> bool:
    return isinstance(value, str) and value.strip().lower().startswith("a)")


def extract_factores(ws, archivo_rel: str) -> list[dict]:
    factores = []
    for row in range(FACTOR_START_ROW, FACTOR_END_ROW + 1):
        meta, actividades = ws[f"K{row}"].value, ws[f"L{row}"].value
        meta_nota = None
        if _parece_lista_actividades(meta) and not _parece_lista_actividades(actividades):
            meta, actividades = actividades, meta
            meta_nota = META_INTERCAMBIADA_NOTA
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
            "linea_base": _limpiar(ws[f"J{row}"].value),
            "meta": _limpiar(meta),
            "meta_nota": meta_nota,
            "actividades": _limpiar(actividades),
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
            "periodo_id": "2024-2026",
            "modalidad": modalidad,
            "archivo_fuente": archivo_rel,
            "cabecera": extract_header(ws, archivo_rel, modalidad),
            "factores": extract_factores(ws, archivo_rel),
        }
        out_path = SILVER_DIR / f"plan_mejoramiento_2025_{modalidad}.json"
        out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        resumen.append((modalidad, len(data["factores"]), rel(out_path)))

    print("Extracción Bronze -> Silver completada (plan anterior, 2024-2026):")
    for modalidad, n_factores, out in resumen:
        print(f"  - {modalidad}: {n_factores} factores -> {out}")


if __name__ == "__main__":
    main()
