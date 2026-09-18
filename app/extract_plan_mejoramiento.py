"""Bronze -> Silver

Lee los dos formatos CC-FR-001 "Plan de Mejoramiento - Programas" (Investigación
y Profundización) y produce un JSON normalizado por modalidad en Data/Silver/,
donde cada registro conserva su referencia exacta (archivo, hoja, fila) para
que cualquier dato mostrado en el sitio pueda verificarse contra el .xlsx
original.

No se inventa ni se completa ningún valor: lo que la celda no trae, se guarda
como null.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib.xlsx_reader import as_flag, as_iso_date, fuente, load_sheet  # noqa: E402

SHEET_NAME = "Plan de mejoramiento"
FACTOR_START_ROW = 15
FACTOR_END_ROW = 26  # 12 factores del modelo CNA

FUENTES = {
    "investigacion": ROOT / "Data/Bronze/Maestria CIC/2026/AUTOEVALUACION/MCIC- INVESTIGACIÓN/CC-FR-001 Plan de mejoramiento INV.xlsx",
    "profundizacion": ROOT / "Data/Bronze/Maestria CIC/2026/AUTOEVALUACION/MCICI- PRODUNDIZACIÓN/CC-FR-001 Plan de mejoramiento PROF.xlsx",
}

SILVER_DIR = ROOT / "Data/Silver"


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


# La celda D8 del archivo de Investigación trae, pegado en un solo texto, el
# número de registro calificado de LAS DOS modalidades ("Prof. 9925 21 JUN
# 2023/       Inv. 16163 05 SEP 2023") — así quedó diligenciada esa celda en
# el CC-FR-001 real; el archivo de Profundización, en cambio, sí trae solo su
# propio número ("Resolución 9925 21 JUN 2023"). Se verificó contra los
# Cuadros Maestro CNA de cada modalidad (Data/Bronze/ACREDITACIÓN DE ALTA
# CALIDAD/SNIES17528-.../"Resolución Registro Calificado: 16163 de septiembre
# de 2023" para Investigación; SNIES116070-.../"9925 de 21 junio de 2023"
# para Profundización) que 16163 es el número propio de Investigación y 9925
# el de Profundización — no es una cifra inventada, es la misma cifra que ya
# trae la celda, solo separada por modalidad para no mostrar en la página de
# Investigación un número que es de Profundización (y viceversa).
REGISTRO_COMBINADO_RE = re.compile(
    r"Prof\.\s*(?P<prof>.+?)\s*/\s*Inv\.\s*(?P<inv>.+)", re.IGNORECASE
)


def separar_registro_calificado(valor: object, modalidad: str) -> tuple[object, str | None]:
    """Si `valor` trae el patrón combinado "Prof. .../ Inv. ...", devuelve solo
    la parte de `modalidad` más una nota explicando el recorte (para no perder
    trazabilidad); si no calza el patrón (caso normal, incluido el archivo de
    Profundización), devuelve `valor` tal cual y sin nota."""
    if not isinstance(valor, str):
        return valor, None
    m = REGISTRO_COMBINADO_RE.search(valor)
    if not m:
        return valor, None
    numero_fecha = m.group("inv" if modalidad == "investigacion" else "prof").strip()
    # El fragmento "Inv. ..." de la celda combinada no repite la palabra
    # "Resolución" (a diferencia del archivo de Profundización, que sí la
    # trae completa); se antepone aquí solo para que las dos modalidades se
    # vean en el mismo formato — el número y la fecha son los de la celda.
    propio = numero_fecha if numero_fecha.lower().startswith("resoluci") else f"Resolución {numero_fecha}"
    nota = (
        f"La celda fuente trae el registro calificado de ambas modalidades en un solo texto "
        f"({valor!r}); aquí se muestra solo el de esta modalidad ('{propio}'), verificado contra "
        f"el Cuadro Maestro CNA correspondiente."
    )
    return propio, nota


def extract_header(ws, archivo_rel: str, modalidad: str) -> dict:
    registro_calificado, registro_calificado_nota = separar_registro_calificado(
        ws["D8"].value, modalidad
    )
    return {
        "codigo_formato": ws["F2"].value,
        "version_formato": ws["F3"].value,
        "fecha_aprobacion_formato": ws["F4"].value,
        "macroproceso": ws["C3"].value,
        "proceso": ws["C4"].value,
        "facultad": ws["D6"].value,
        "programa_academico": ws["D7"].value,
        "registro_calificado": registro_calificado,
        "registro_calificado_nota": registro_calificado_nota,
        "registro_calificado_vigencia": ws["K8"].value,
        "acreditacion_alta_calidad": ws["S8"].value,
        "acreditacion_alta_calidad_vigencia": ws["Z8"].value,
        "nivel_formacion": ws["AH8"].value,
        "fecha_proyeccion_plan": ws["D9"].value,
        "fuente": fuente(archivo_rel, SHEET_NAME, "2-9 (cabecera)"),
    }


def extract_seguimiento(ws, row: int) -> dict:
    return {
        "corte_1": {
            "fecha_seguimiento": ws[f"X{row}"].value,
            "descripcion_avance_cualitativo": ws[f"Y{row}"].value,
            "pct_avance_cuantitativo": ws[f"Z{row}"].value,
            "unidad_medida_evidencia": ws[f"AA{row}"].value,
            "otras_observaciones": ws[f"AB{row}"].value,
        },
        "corte_2": {
            "fecha_seguimiento": ws[f"AC{row}"].value,
            "descripcion_avance_cualitativo": ws[f"AD{row}"].value,
            "pct_avance_cuantitativo": ws[f"AE{row}"].value,
            "unidad_medida_evidencia": ws[f"AF{row}"].value,
            "otras_observaciones": ws[f"AG{row}"].value,
        },
        "acumulado": {
            "balance_cualitativo": ws[f"AH{row}"].value,
            "pct_avance_cuantitativo": ws[f"AI{row}"].value,
            "unidad_medida_evidencia": ws[f"AJ{row}"].value,
            "enlace_evidencia": ws[f"AK{row}"].value,
            "otras_observaciones": ws[f"AL{row}"].value,
        },
    }


def extract_factores(ws, archivo_rel: str) -> list[dict]:
    factores = []
    for row in range(FACTOR_START_ROW, FACTOR_END_ROW + 1):
        factores.append({
            "factor": ws[f"B{row}"].value,
            "tipo": (ws[f"C{row}"].value or "").strip() or None,
            "origen": ws[f"D{row}"].value,
            "descripcion": ws[f"E{row}"].value,
            "proyecto": ws[f"F{row}"].value,
            "objetivo": ws[f"G{row}"].value,
            "articulacion_plan_institucional": ws[f"H{row}"].value,
            "periodo_inicio": as_iso_date(ws[f"I{row}"].value),
            "periodo_fin": as_iso_date(ws[f"J{row}"].value),
            "peso_prioridad": ws[f"K{row}"].value,
            "indicador_cumplimiento": ws[f"L{row}"].value,
            "tipo_indicador": ws[f"M{row}"].value,
            "linea_base": ws[f"N{row}"].value,
            "meta": ws[f"O{row}"].value,
            "actividades": ws[f"P{row}"].value,
            "periodicidad": ws[f"Q{row}"].value,
            "tipo_actividad": ws[f"R{row}"].value,
            "apoyo_requerido": {
                "programa_academico": as_flag(ws[f"S{row}"].value),
                "facultad": as_flag(ws[f"T{row}"].value),
                "institucion": as_flag(ws[f"U{row}"].value),
            },
            "responsable": ws[f"V{row}"].value,
            "recursos": ws[f"W{row}"].value,
            "seguimiento": extract_seguimiento(ws, row),
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
            "modalidad": modalidad,
            "archivo_fuente": archivo_rel,
            "cabecera": extract_header(ws, archivo_rel, modalidad),
            "factores": extract_factores(ws, archivo_rel),
        }
        out_path = SILVER_DIR / f"plan_mejoramiento_{modalidad}.json"
        out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        resumen.append((modalidad, len(data["factores"]), rel(out_path)))

    print("Extracción Bronze -> Silver completada:")
    for modalidad, n_factores, out in resumen:
        print(f"  - {modalidad}: {n_factores} factores -> {out}")


if __name__ == "__main__":
    main()
