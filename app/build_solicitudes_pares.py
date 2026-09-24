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

CARPETA_PRESENTACIONES = "APERTURA"
CARPETA_SOLICITUDES = "SOLICITUDES DE PARES"

# Archivos de APERTURA que no se publican: la lista de asistencia contiene
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

        presentaciones = []
        carpeta_pres = carpeta_dia / CARPETA_PRESENTACIONES
        if carpeta_pres.exists():
            presentaciones = [
                registro(p, borrador=es_borrador) for p in archivos(carpeta_pres)
                if not any(ex in p.name.lower() for ex in EXCLUIR_PRESENTACIONES)
            ]

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
                documentos = [registro(p, borrador=es_borrador) for p in archivos(carpeta_area)]
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
