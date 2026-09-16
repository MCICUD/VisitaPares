"""Silver -> Gold

Toma los JSON normalizados de Data/Silver/ y arma el objeto GOLD_DATA que
consume el micrositio. Las únicas cifras "nuevas" que aparecen aquí son
agregados calculados por código a partir de los mismos 12 factores de cada
modalidad (conteos, promedios, distribución) — nunca datos escritos a mano.
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SILVER_DIR = ROOT / "Data/Silver"
GOLD_DIR = ROOT / "Data/Gold"

# Dato institucional conocido (dado directamente por el usuario, no extraído de Bronze):
SHAREPOINT_URL = (
    "https://udistritaleduco-my.sharepoint.com/:f:/g/personal/mcic_udistrital_edu_co/"
    "IgDD4JkqpY8YRqm7HlQGnwa7ATrRmbdIThjfnlLXwk2FPDM?e=kfKTko"
)

# Los 4 documentos directamente ligados al Plan de Mejoramiento (título +
# ruta exacta). La ruta se valida contra el catálogo completo de Bronze
# (bronze_manifest.json): si el archivo no está ahí, main() falla en vez de
# mostrar un documento que no existe.
DOCUMENTOS_PRINCIPALES_DEF = [
    {
        "titulo": "Plan de Mejoramiento CC-FR-001 — Investigación",
        "modalidad": "investigacion",
        "archivo": "Data/Bronze/Maestria CIC/2026/AUTOEVALUACION/MCIC- INVESTIGACIÓN/CC-FR-001 Plan de mejoramiento INV.xlsx",
    },
    {
        "titulo": "Autoevaluación Permanente Institucional — Investigación",
        "modalidad": "investigacion",
        "archivo": "Data/Bronze/Maestria CIC/2026/AUTOEVALUACION/MCIC- INVESTIGACIÓN/MCIC AutoevaluacionPermanenteInstitucional INV.docx",
    },
    {
        "titulo": "Plan de Mejoramiento CC-FR-001 — Profundización",
        "modalidad": "profundizacion",
        "archivo": "Data/Bronze/Maestria CIC/2026/AUTOEVALUACION/MCICI- PRODUNDIZACIÓN/CC-FR-001 Plan de mejoramiento PROF.xlsx",
    },
    {
        "titulo": "Autoevaluación Permanente Institucional — Profundización",
        "modalidad": "profundizacion",
        "archivo": "Data/Bronze/Maestria CIC/2026/AUTOEVALUACION/MCICI- PRODUNDIZACIÓN/MCIC AutoevaluacionPermanenteInstitucional PROF.docx",
    },
]


def compute_comparacion_modalidades(silver_inv: dict, silver_prof: dict) -> dict:
    """Compara campo a campo INV vs PROF para que el sitio nunca sugiera una
    diferencia que no esté realmente en los archivos (ni oculte que, en los
    12 factores, el texto es literalmente el mismo en ambos .xlsx)."""
    campos_cabecera_distintos = [
        campo for campo in silver_inv["cabecera"]
        if campo != "fuente" and silver_inv["cabecera"][campo] != silver_prof["cabecera"][campo]
    ]
    factores_con_diferencias = []
    for f_inv, f_prof in zip(silver_inv["factores"], silver_prof["factores"]):
        campos_distintos = [
            campo for campo in f_inv
            # "evidencia_seguimiento" se excluye a propósito: es normal que una
            # modalidad tenga más archivos de evidencia cargados que la otra;
            # esta comparación es solo sobre el TEXTO oficial del CC-FR-001.
            if campo not in ("fuente", "evidencia_seguimiento") and f_inv[campo] != f_prof[campo]
        ]
        if campos_distintos:
            factores_con_diferencias.append({"factor": f_inv["factor"], "campos_distintos": campos_distintos})
    return {
        "campos_cabecera_distintos": campos_cabecera_distintos,
        "factores_con_diferencias": factores_con_diferencias,
        "nota": (
            "Comparación calculada campo a campo entre los dos archivos .xlsx. "
            "Si 'factores_con_diferencias' está vacío, significa que el texto de los "
            "12 factores es idéntico en ambos archivos fuente (solo cambia el "
            "encabezado institucional); no es un error de esta pipeline."
        ),
    }


def load_bronze_manifest() -> list[dict]:
    path = SILVER_DIR / "bronze_manifest.json"
    if not path.exists():
        raise FileNotFoundError(f"No existe {path}. Corre primero app/build_bronze_manifest.py")
    return json.loads(path.read_text(encoding="utf-8"))


def build_documentos_principales(manifest: list[dict]) -> list[dict]:
    por_ruta = {r["archivo"]: r for r in manifest}
    resultado = []
    for doc in DOCUMENTOS_PRINCIPALES_DEF:
        registro = por_ruta.get(doc["archivo"])
        if registro is None:
            raise FileNotFoundError(
                f"'{doc['archivo']}' no aparece en bronze_manifest.json: "
                "revisa que el archivo siga existiendo en Data/Bronze."
            )
        resultado.append({**doc, "tamano_legible": registro["tamano_legible"], "extension": registro["extension"]})
    return resultado


EXTRACTO_CHARS = 320


def merge_texto_extraido(manifest: list[dict]) -> None:
    """Añade a cada registro del catálogo si se pudo leer su texto y un
    extracto corto (para buscar por contenido en el sitio). El texto
    completo se queda en Data/Silver/texto_bronze.json — no se manda al
    navegador para no inflar gold_data.js con documentos completos."""
    texto_path = SILVER_DIR / "texto_bronze.json"
    if not texto_path.exists():
        raise FileNotFoundError(f"No existe {texto_path}. Corre primero app/extract_texto_bronze.py")
    textos = json.loads(texto_path.read_text(encoding="utf-8"))
    for r in manifest:
        info = textos.get(r["archivo"])
        if info is None:
            r["texto_extraido"] = False
            r["extracto"] = None
            continue
        r["texto_extraido"] = bool(info["extraido"])
        texto = info["texto"] or ""
        r["extracto"] = (texto[:EXTRACTO_CHARS] + "…") if len(texto) > EXTRACTO_CHARS else (texto or None)


def manifest_stats(manifest: list[dict]) -> dict:
    por_modalidad: dict[str, int] = {}
    por_extension: dict[str, int] = {}
    por_carpeta: dict[str, int] = {}
    con_texto = 0
    for r in manifest:
        por_modalidad[r["modalidad"]] = por_modalidad.get(r["modalidad"], 0) + 1
        por_extension[r["extension"]] = por_extension.get(r["extension"], 0) + 1
        por_carpeta[r["carpeta_raiz"]] = por_carpeta.get(r["carpeta_raiz"], 0) + 1
        if r.get("texto_extraido"):
            con_texto += 1
    return {
        "total": len(manifest),
        "por_modalidad": por_modalidad,
        "por_extension": por_extension,
        "por_carpeta_raiz": por_carpeta,
        "con_texto_extraido": con_texto,
    }


def compute_stats(silver: dict) -> dict:
    factores = silver["factores"]
    tipos = Counter((f["tipo"] or "Sin especificar") for f in factores)
    tipos_actividad = Counter((f["tipo_actividad"] or "Sin especificar") for f in factores)
    pesos = [f["peso_prioridad"] for f in factores if isinstance(f["peso_prioridad"], (int, float))]
    inicios = sorted(f["periodo_inicio"] for f in factores if f["periodo_inicio"])
    fines = sorted(f["periodo_fin"] for f in factores if f["periodo_fin"])
    return {
        "total_factores": len(factores),
        "conteo_por_tipo": dict(tipos),
        "conteo_por_tipo_actividad": dict(tipos_actividad),
        "peso_prioridad_promedio": round(sum(pesos) / len(pesos), 2) if pesos else None,
        "peso_prioridad_total": sum(pesos) if pesos else None,
        "periodo_ejecucion_min": inicios[0] if inicios else None,
        "periodo_ejecucion_max": fines[-1] if fines else None,
        "derivado_de": {
            "archivo": silver["archivo_fuente"],
            "hoja": "Plan de mejoramiento",
            "filas": f"{factores[0]['fuente']['fila']}-{factores[-1]['fuente']['fila']}",
            "nota": "Conteos y promedios calculados por app/build_gold.py a partir de estas filas; no son datos manuales.",
        },
    }


def load_silver(modalidad: str) -> dict:
    path = SILVER_DIR / f"plan_mejoramiento_{modalidad}.json"
    if not path.exists():
        raise FileNotFoundError(
            f"No existe {path}. Corre primero app/extract_plan_mejoramiento.py"
        )
    return json.loads(path.read_text(encoding="utf-8"))


FACTOR_RE = re.compile(r"^FACTOR\s*0*(\d{1,2})\b", re.IGNORECASE)


def load_json(name: str) -> dict:
    path = SILVER_DIR / name
    if not path.exists():
        raise FileNotFoundError(f"No existe {path}. Corre primero el script que lo genera en app/.")
    return json.loads(path.read_text(encoding="utf-8"))


def merge_evidencia_seguimiento(factores: list[dict], evidencia: dict) -> None:
    """Añade a cada factor su evidencia real de seguimiento (Data/Bronze),
    indexada por app/extract_seguimiento_evidencia.py. Modifica in-place."""
    por_numero = evidencia["factores"]
    for f in factores:
        m = FACTOR_RE.match(f["factor"] or "")
        numero = m.group(1) if m else None
        entrada = por_numero.get(numero) if numero else None
        actividades = entrada["actividades"] if entrada else []
        total_archivos = sum(len(a["archivos"]) for a in actividades)
        f["evidencia_seguimiento"] = {
            "actividades": actividades,
            "total_archivos": total_archivos,
        }


def build_comunidad_estudiantil() -> dict:
    enfasis = load_json("enfasis_estudiantes.json")
    estado = load_json("estado_academico_agregado.json")
    egresados = load_json("egresados_agregado.json")
    return {"enfasis": enfasis, "estadoAcademico": estado, "egresados": egresados}


def main() -> None:
    GOLD_DIR.mkdir(parents=True, exist_ok=True)

    silver_inv = load_silver("investigacion")
    silver_prof = load_silver("profundizacion")
    manifest = load_bronze_manifest()
    merge_texto_extraido(manifest)

    grupos_investigacion = load_json("grupos_investigacion.json")

    evidencia_inv = load_json("seguimiento_evidencia_investigacion.json")
    evidencia_prof = load_json("seguimiento_evidencia_profundizacion.json")
    merge_evidencia_seguimiento(silver_inv["factores"], evidencia_inv)
    merge_evidencia_seguimiento(silver_prof["factores"], evidencia_prof)

    gold = {
        "meta": {
            "institucion": "Universidad Distrital Francisco José de Caldas",
            "programa": silver_inv["cabecera"]["programa_academico"],
            "facultad": silver_inv["cabecera"]["facultad"],
            "sharepointUrl": SHAREPOINT_URL,
            "modalidades": {
                "investigacion": silver_inv["cabecera"],
                "profundizacion": silver_prof["cabecera"],
            },
            "generadoPor": "app/run_pipeline.py (Bronze -> Silver -> Gold)",
        },
        "factores": {
            "investigacion": silver_inv["factores"],
            "profundizacion": silver_prof["factores"],
        },
        "stats": {
            "investigacion": compute_stats(silver_inv),
            "profundizacion": compute_stats(silver_prof),
        },
        "comparacionModalidades": compute_comparacion_modalidades(silver_inv, silver_prof),
        "documentosPrincipales": build_documentos_principales(manifest),
        "documentosBronze": manifest,
        "documentosBronzeStats": manifest_stats(manifest),
        "evidenciaProcesosRcAac": {
            "investigacion": evidencia_inv["procesos_rc_aac"],
            "profundizacion": evidencia_prof["procesos_rc_aac"],
        },
        "evidenciaDocumentosGenerales": {
            "investigacion": evidencia_inv["documentos_generales"],
            "profundizacion": evidencia_prof["documentos_generales"],
        },
        "comunidadEstudiantil": build_comunidad_estudiantil(),
        "gruposInvestigacion": grupos_investigacion,
    }

    json_path = GOLD_DIR / "gold_data.json"
    js_path = GOLD_DIR / "gold_data.js"

    json_text = json.dumps(gold, ensure_ascii=False, indent=2)
    json_path.write_text(json_text, encoding="utf-8")
    js_path.write_text(f"const GOLD_DATA = {json_text};\n", encoding="utf-8")

    print("Construcción Silver -> Gold completada:")
    print(f"  - {json_path.relative_to(ROOT)}")
    print(f"  - {js_path.relative_to(ROOT)}")
    print(f"  - Factores Investigación: {len(gold['factores']['investigacion'])}")
    print(f"  - Factores Profundización: {len(gold['factores']['profundizacion'])}")


if __name__ == "__main__":
    main()
