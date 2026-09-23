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
        "archivo": "Data/Bronze/Maestria CIC/2026/AUTOEVALUACION/MCIC- INVESTIGACIÓN/MCIC AutoevaluacionPermanenteInstitucional INV.pdf",
    },
    {
        "titulo": "Plan de Mejoramiento CC-FR-001 — Profundización",
        "modalidad": "profundizacion",
        "archivo": "Data/Bronze/Maestria CIC/2026/AUTOEVALUACION/MCICI- PRODUNDIZACIÓN/CC-FR-001 Plan de mejoramiento PROF.xlsx",
    },
    {
        "titulo": "Autoevaluación Permanente Institucional — Profundización",
        "modalidad": "profundizacion",
        "archivo": "Data/Bronze/Maestria CIC/2026/AUTOEVALUACION/MCICI- PRODUNDIZACIÓN/MCIC AutoevaluacionPermanenteInstitucional PROF.pdf",
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


ENLACE_MODULO_EGRESADOS = {
    "tipo": "enlace_externo",
    "titulo": "Módulo institucional de Hoja de Vida de Egresados (UD)",
    "url": "https://egresados.udistrital.edu.co/hoja-de-vida-impulsa-tu-perfil-profesional",
}


def merge_cuadros_maestros(factores: list[dict], cuadros: dict, resumen_grupos_factor8: dict) -> None:
    """Añade a los Factores 4 (Egresados), 8 (Aportes de la investigación) y
    9 (Bienestar) los datos reales del Cuadro Maestro CNA (Data/Silver/
    cuadros_maestros.json) — evidencia oficial que hoy no se mostraba en
    absoluto. Modifica in-place, igual que merge_evidencia_seguimiento."""
    for f in factores:
        m = FACTOR_RE.match(f["factor"] or "")
        numero = m.group(1) if m else None
        if numero == "4":
            f["datos_cuadro_maestro_cna"] = {
                "enlace_modulo_egresados": ENLACE_MODULO_EGRESADOS,
            }
            # Add data from community for Factor 4
            f["datos_comunidad_factor4"] = {
                "egresados": load_json("egresados_agregado.json"),
                "estados": load_json("estado_academico_agregado.json"),
            }
        elif numero == "8":
            f["datos_cuadro_maestro_cna"] = {
                "grupos_produccion": cuadros["grupos_produccion"],
                "resumen_grupos_investigacion": resumen_grupos_factor8,
            }
        elif numero == "9":
            f["datos_cuadro_maestro_cna"] = {
                "bienestar": cuadros["bienestar"],
            }


def merge_convenios(factores: list[dict], convenios: dict) -> None:
    """Añade al Factor 7 (Interacción con el entorno nacional e internacional)
    los convenios institucionales vigentes (Data/Silver/convenios.json,
    descargado de URELINTER: ver app/extract_convenios.py) — evidencia real
    que hoy no se mostraba en absoluto. Es un listado institucional, igual
    para las dos modalidades (no diferenciado por Investigación/
    Profundización). Modifica in-place, igual que merge_cuadros_maestros."""
    for f in factores:
        m = FACTOR_RE.match(f["factor"] or "")
        numero = m.group(1) if m else None
        if numero == "7":
            f["datos_convenios"] = convenios


def build_comunidad_estudiantil() -> dict:
    enfasis = load_json("enfasis_estudiantes.json")
    estado = load_json("estado_academico_agregado.json")
    egresados = load_json("egresados_agregado.json")
    descarga = {
        # Generado por app/build_comunidad_xlsx.py (paso final de run_pipeline.py) a partir de
        # estos mismos 3 JSON de Data/Silver/ — no se valida su existencia aquí porque ese paso
        # corre después de este (Silver -> Gold: xlsx), no antes.
        "archivo": "Data/Gold/comunidad_estudiantil.xlsx",
        "titulo": "Consolidado Comunidad Estudiantil (.xlsx)",
    }
    return {"enfasis": enfasis, "estadoAcademico": estado, "egresados": egresados, "descarga": descarga}


def normalizar_sigla(texto: str) -> str:
    """Sin tildes, mayúsculas, solo alfanuméricos — para comparar la sigla
    de un grupo (Data/Silver/grupos_investigacion.json) contra el texto
    libre de 'grupo_investigacion' que extrajo app/extract_seguimiento_tesis.py
    de las cartas reales (que a veces trae el nombre completo del grupo con
    la sigla entre paréntesis, o solo un fragmento del nombre)."""
    import unicodedata

    sin_tildes = "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    )
    return re.sub(r"[^A-Z0-9]", "", sin_tildes.upper())


def merge_proyectos_grado(grupos_investigacion: dict) -> dict:
    """Añade a cada grupo de investigación (por sigla) los proyectos de
    grado (595/695) cuya carta de radicación/viabilidad menciona ese grupo.
    Un proceso cuyo 'grupo_investigacion' no calza con ninguna sigla
    conocida se reporta aparte (proyectosGradoSinGrupo) en vez de perderse
    en silencio — puede ser un grupo nuevo no listado en el directorio, o
    una sigla mal escrita en la carta original."""
    seguimiento = load_json("seguimiento_tesis.json")
    procesos = seguimiento["procesos"]

    siglas_normalizadas = {
        normalizar_sigla(g["sigla"]): g["sigla"] for g in grupos_investigacion["grupos"]
    }
    nombres_normalizados = {
        normalizar_sigla(g["nombre"]): g["sigla"] for g in grupos_investigacion["grupos"] if g.get("nombre")
    }
    for g in grupos_investigacion["grupos"]:
        g["proyectos_grado"] = []

    por_sigla = {g["sigla"]: g for g in grupos_investigacion["grupos"]}
    sin_grupo = []

    for proceso in procesos:
        grupo_texto = proceso.get("grupo_investigacion")
        proyecto_publico = {
            "titulo_proyecto": proceso["titulo_proyecto"],
            "director": proceso["director"],
            "codirector": proceso["codirector"],
            "etapa_actual": proceso["etapa_actual"],
            "cod_proyecto": proceso["cod_proyecto"],
        }
        if not grupo_texto:
            sin_grupo.append({**proyecto_publico, "grupo_investigacion_texto": None})
            continue

        grupo_norm = normalizar_sigla(grupo_texto)
        sigla_encontrada = None
        if grupo_norm in siglas_normalizadas:
            sigla_encontrada = siglas_normalizadas[grupo_norm]
        elif grupo_norm in nombres_normalizados:
            sigla_encontrada = nombres_normalizados[grupo_norm]
        else:
            for sigla_norm, sigla_original in siglas_normalizadas.items():
                if sigla_norm and (sigla_norm == grupo_norm or sigla_norm in grupo_norm or grupo_norm in sigla_norm):
                    sigla_encontrada = sigla_original
                    break

        if sigla_encontrada:
            por_sigla[sigla_encontrada]["proyectos_grado"].append(proyecto_publico)
        else:
            sin_grupo.append({**proyecto_publico, "grupo_investigacion_texto": grupo_texto})

    resumen_etapas: dict[str, int] = {}
    for p in procesos:
        resumen_etapas[p["etapa_actual"]] = resumen_etapas.get(p["etapa_actual"], 0) + 1

    resumen_seguimiento_tesis = {
        **seguimiento["resumen"],
        "por_etapa": resumen_etapas,
        "grupos_con_proyectos_asociados": sum(1 for g in grupos_investigacion["grupos"] if g["proyectos_grado"]),
        "proyectos_sin_grupo_identificado": len(sin_grupo),
        "metodologia": seguimiento["metodologia"],
    }

    return sin_grupo, resumen_seguimiento_tesis


def main() -> None:
    GOLD_DIR.mkdir(parents=True, exist_ok=True)

    silver_inv = load_silver("investigacion")
    silver_prof = load_silver("profundizacion")
    manifest = load_bronze_manifest()
    merge_texto_extraido(manifest)

    grupos_investigacion = load_json("grupos_investigacion.json")
    proyectos_grado_sin_grupo, resumen_seguimiento_tesis = merge_proyectos_grado(grupos_investigacion)

    evidencia_inv = load_json("seguimiento_evidencia_investigacion.json")
    evidencia_prof = load_json("seguimiento_evidencia_profundizacion.json")
    merge_evidencia_seguimiento(silver_inv["factores"], evidencia_inv)
    merge_evidencia_seguimiento(silver_prof["factores"], evidencia_prof)

    cuadros_maestros = load_json("cuadros_maestros.json")
    por_mod = resumen_seguimiento_tesis.get("por_modalidad", {})
    resumen_factor8_inv = {
        "total_grupos": len(grupos_investigacion["grupos"]),
        "total_docentes_disponibles": sum(len(g["integrantes"]) for g in grupos_investigacion["grupos"]),
        "total_proyectos_grado_detectados": por_mod.get("investigacion", {}).get("total", resumen_seguimiento_tesis["total_procesos"]),
        "proyectos_grado_por_etapa": por_mod.get("investigacion", {}).get("por_etapa", resumen_seguimiento_tesis["por_etapa"]),
        "total_programa_consolidado": resumen_seguimiento_tesis["total_procesos"],
        "programa_consolidado_por_etapa": resumen_seguimiento_tesis["por_etapa"],
        "fuentes_consultadas": resumen_seguimiento_tesis.get("fuentes_consultadas", []),
    }
    resumen_factor8_prof = {
        "total_grupos": len(grupos_investigacion["grupos"]),
        "total_docentes_disponibles": sum(len(g["integrantes"]) for g in grupos_investigacion["grupos"]),
        "total_proyectos_grado_detectados": por_mod.get("profundizacion", {}).get("total", resumen_seguimiento_tesis["total_procesos"]),
        "proyectos_grado_por_etapa": por_mod.get("profundizacion", {}).get("por_etapa", resumen_seguimiento_tesis["por_etapa"]),
        "total_programa_consolidado": resumen_seguimiento_tesis["total_procesos"],
        "programa_consolidado_por_etapa": resumen_seguimiento_tesis["por_etapa"],
        "fuentes_consultadas": resumen_seguimiento_tesis.get("fuentes_consultadas", []),
    }
    merge_cuadros_maestros(silver_inv["factores"], cuadros_maestros, resumen_factor8_inv)
    merge_cuadros_maestros(silver_prof["factores"], cuadros_maestros, resumen_factor8_prof)

    convenios = load_json("convenios.json")
    merge_convenios(silver_inv["factores"], convenios)
    merge_convenios(silver_prof["factores"], convenios)

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
        "proyectosGradoSinGrupo": proyectos_grado_sin_grupo,
        "resumenSeguimientoTesis": resumen_seguimiento_tesis,
        "cuadrosMaestrosCNA": cuadros_maestros,
        "solicitudesPares": load_json("solicitudes_pares.json"),
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
