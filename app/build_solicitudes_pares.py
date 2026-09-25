"""SolicitudesPares -> Silver

Recorre SolicitudesPares/ (material entregado a los pares durante la visita)
y produce Data/Silver/solicitudes_pares.json, fuente de la sección
"Solicitudes Pares" del sitio. Estructura esperada, un directorio por día:

    SolicitudesPares/<DIA N>/APERTURA/...               -> presentaciones
    SolicitudesPares/<DIA N>/SOLICITUDES DE PARES/<área>/... -> solicitudes

Si un día nuevo aparece con esta misma estructura, se lista automáticamente.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from build_bronze_manifest import tamano_legible

ROOT = Path(__file__).resolve().parent.parent
SOLICITUDES_DIR = ROOT / "SolicitudesPares"
SILVER_DIR = ROOT / "Data/Silver"

CARPETAS_PRESENTACIONES = ("APERTURA", "PRESENTACIONES")
CARPETA_SOLICITUDES = "SOLICITUDES DE PARES"

# Archivos de presentaciones que no se publican: la lista de asistencia contiene
# nombres y firmas de los asistentes.
EXCLUIR_PRESENTACIONES = ("lista de asistencia",)

# Avisos informativos por día (ej. documentos preliminares / borrador)
AVISOS_DIAS = {
    2: "El archivo PDF cargado es un documento borrador.",
}


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def archivos(carpeta: Path) -> list[Path]:
    return sorted(
        p for p in carpeta.rglob("*")
        if p.is_file() and not p.name.startswith("~$")
    )


def registro(path: Path, borrador: bool = False) -> dict:
    size_bytes = path.stat().st_size
    item = {
        "archivo": rel(path),
        "nombre": path.name,
        "extension": path.suffix.lower().lstrip("."),
        "tamano_bytes": size_bytes,
        "tamano_legible": tamano_legible(size_bytes),
    }
    if borrador:
        item["borrador"] = True
        item["nota"] = "Documento borrador"
    return item


def recolectar_documentos_area(carpeta_area: Path, borrador: bool = False) -> list[dict]:
    """Recolecta documentos de un área de Solicitudes.
    
    Si existen versiones filtradas '(MCIC Filtrado).xlsx', empareja el documento
    original con su versión filtrada y excluye los archivos sin relación a MCIC.
    Si un documento es en origen específico para la maestría (ej. Bienestar), se
    incluye también emparejado consigo mismo.
    """
    all_files = archivos(carpeta_area)
    if not all_files:
        return []

    filtered_files = [f for f in all_files if "(MCIC Filtrado)" in f.name]
    has_filtering = (
        bool(filtered_files)
        or "bienestar" in carpeta_area.name.lower()
        or any("(MCIC Filtrado)" in f.name for f in carpeta_area.parent.rglob("*"))
    )

    if not has_filtering:
        return [registro(p, borrador=borrador) for p in all_files]

    docs = []
    processed_stems = set()

    # 1. Documentos con versión filtrada generada
    for f_filt in filtered_files:
        orig_name = f_filt.name.replace(" (MCIC Filtrado)", "")
        orig_path = f_filt.parent / orig_name
        if not orig_path.exists():
            for cand in all_files:
                if cand.name.replace(" (MCIC Filtrado)", "") == orig_name:
                    orig_path = cand
                    break

        if orig_path.exists():
            processed_stems.add(orig_path.name)
            processed_stems.add(f_filt.name)
            orig_size = orig_path.stat().st_size
            filt_size = f_filt.stat().st_size
            docs.append({
                "archivo": rel(orig_path),
                "nombre": orig_path.name,
                "extension": orig_path.suffix.lower().lstrip("."),
                "tamano_bytes": orig_size,
                "tamano_legible": tamano_legible(orig_size),
                "archivo_original": rel(orig_path),
                "tamano_original": tamano_legible(orig_size),
                "archivo_filtrado": rel(f_filt),
                "tamano_filtrado": tamano_legible(filt_size),
                "tiene_version_filtrada": True,
                "borrador": borrador,
            })

    # 2. Documentos que en origen ya son específicos de la Maestría (ej. Bienestar)
    for p in all_files:
        if p.name in processed_stems:
            continue
        if re.search(r"ciencias de la informacion|maestr[ií]a", p.name, re.IGNORECASE):
            processed_stems.add(p.name)
            size = p.stat().st_size
            docs.append({
                "archivo": rel(p),
                "nombre": p.name,
                "extension": p.suffix.lower().lstrip("."),
                "tamano_bytes": size,
                "tamano_legible": tamano_legible(size),
                "archivo_original": rel(p),
                "tamano_original": tamano_legible(size),
                "archivo_filtrado": rel(p),
                "tamano_filtrado": tamano_legible(size),
                "tiene_version_filtrada": False,
                "es_especifico_origen": True,
                "borrador": borrador,
            })

    return sorted(docs, key=lambda d: d["nombre"])


def numero_dia(carpeta: Path) -> int:
    match = re.search(r"\d+", carpeta.name)
    return int(match.group()) if match else 0


def main() -> None:
    if not SOLICITUDES_DIR.exists():
        raise FileNotFoundError(f"No existe {SOLICITUDES_DIR}")

    dias = []
    for carpeta_dia in sorted((p for p in SOLICITUDES_DIR.iterdir() if p.is_dir()), key=numero_dia):
        num_dia = numero_dia(carpeta_dia)
        es_borrador = (num_dia == 2)

        # Presentaciones: buscar en APERTURA y PRESENTACIONES
        presentaciones = []
        for nom_pres in CARPETAS_PRESENTACIONES:
            carpeta_pres = carpeta_dia / nom_pres
            if carpeta_pres.exists():
                for p in archivos(carpeta_pres):
                    if not any(ex in p.name.lower() for ex in EXCLUIR_PRESENTACIONES):
                        if not any(item["archivo"] == rel(p) for item in presentaciones):
                            presentaciones.append(registro(p, borrador=es_borrador))
        presentaciones.sort(key=lambda x: x["nombre"])

        areas = []
        carpeta_sol = carpeta_dia / CARPETA_SOLICITUDES
        if carpeta_sol.exists():
            sueltos = [
                registro(p, borrador=es_borrador)
                for p in sorted(carpeta_sol.iterdir())
                if p.is_file() and not p.name.startswith("~$")
            ]
            if sueltos:
                areas.append({"area": "General", "documentos": sueltos})
            for carpeta_area in sorted(p for p in carpeta_sol.iterdir() if p.is_dir()):
                documentos = recolectar_documentos_area(carpeta_area, borrador=es_borrador)
                if documentos:
                    areas.append({"area": carpeta_area.name, "documentos": documentos})

        # Archivos directos en la raíz de la carpeta del día (sin subdirectorios)
        directos = [
            registro(p, borrador=es_borrador)
            for p in sorted(carpeta_dia.iterdir())
            if p.is_file() and not p.name.startswith("~$")
        ]

        tiene_subsecciones = bool(presentaciones or areas)
        sin_subsecciones = not tiene_subsecciones and bool(directos)

        dias.append({
            "dia": carpeta_dia.name,
            "numero": num_dia,
            "presentaciones": presentaciones,
            "solicitudes": areas,
            "documentos": directos,
            "sin_subsecciones": sin_subsecciones,
            "aviso": AVISOS_DIAS.get(num_dia),
        })

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "solicitudes_pares.json"
    out_path.write_text(json.dumps({"dias": dias}, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Solicitudes de pares: {len(dias)} día(s) -> {rel(out_path)}")
    for d in dias:
        total_sol = sum(len(a["documentos"]) for a in d["solicitudes"])
        print(f"  - {d['dia']}: {len(d['presentaciones'])} presentación(es), {total_sol} solicitud(es) en {len(d['solicitudes'])} área(s), {len(d['documentos'])} doc(s) directo(s)")


if __name__ == "__main__":
    main()
