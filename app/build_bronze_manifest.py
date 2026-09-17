"""Bronze -> Silver (catálogo)

Recorre TODO Data/Bronze/ y produce un inventario completo de la data RAW
(un registro por archivo, sin excepción) en Data/Silver/bronze_manifest.json.
Este catálogo es la fuente de la sección "Documentos" del sitio: si un
archivo existe en Bronze, aparece aquí; si no está aquí, no se muestra en
el sitio como si existiera.

La "modalidad" de cada archivo es una heurística basada en el nombre de su
ruta (contiene "INVESTIGACION"/"PROFUNDIZACION" normalizado sin tildes);
cuando la ruta no menciona ninguna de las dos, se clasifica como "general"
(aplica a ambas modalidades o es transversal institucional).
"""
from __future__ import annotations

import json
import sys
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BRONZE_DIR = ROOT / "Data/Bronze"
SILVER_DIR = ROOT / "Data/Silver"

UNIDADES = ["B", "KB", "MB", "GB"]


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def normalizar(texto: str) -> str:
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).upper()


def clasificar_modalidad(ruta_rel: str) -> str:
    ruta_norm = normalizar(ruta_rel)
    tiene_prof = "PROFUNDIZ" in ruta_norm or "PRODUNDIZ" in ruta_norm  # "PRODUNDIZACIÓN" es un typo real en Bronze
    tiene_inv = "INVESTIGACION" in ruta_norm
    if tiene_prof and tiene_inv:
        return "ambas"
    if tiene_prof:
        return "profundizacion"
    if tiene_inv:
        return "investigacion"
    return "general"


def tamano_legible(num_bytes: int) -> str:
    tamano = float(num_bytes)
    for unidad in UNIDADES:
        if tamano < 1024 or unidad == UNIDADES[-1]:
            return f"{tamano:.0f} {unidad}" if unidad == "B" else f"{tamano:.1f} {unidad}"
        tamano /= 1024
    return f"{tamano:.1f} GB"


# Carpetas que SÍ viven bajo Data/Bronze (para que la pipeline pueda leerlas)
# pero que no se listan en el catálogo público de "Documentos": contienen
# nombre, documento de identidad y correo personal de cada estudiante por
# fila (roster oficial Cóndor). app/extract_estado_academico.py sí las lee,
# solo para calcular conteos agregados.
# "Seguimiento_Tesis" tiene el mismo problema pero por RUTA (carpetas
# "<código>/<categoría>/<archivo original>", cuyo nombre de archivo original
# suele incluir el nombre completo del estudiante) — app/extract_seguimiento_tesis.py
# sí la lee, solo para extraer título/director/etapa por proceso de tesis.
EXCLUIR_DEL_CATALOGO = ("Estados", "PII_Interno", "Seguimiento_Tesis")


def main() -> None:
    if not BRONZE_DIR.exists():
        raise FileNotFoundError(f"No existe {BRONZE_DIR}")

    registros = []
    for path in sorted(BRONZE_DIR.rglob("*")):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue  # archivo de bloqueo temporal de Office, no es un documento real
        ruta_bajo_bronze = str(path.relative_to(BRONZE_DIR)).replace("\\", "/")
        carpeta_raiz = ruta_bajo_bronze.split("/")[0]
        if carpeta_raiz in EXCLUIR_DEL_CATALOGO:
            continue
        ruta_rel = rel(path)
        size_bytes = path.stat().st_size
        registros.append({
            "archivo": ruta_rel,
            "nombre": path.name,
            "extension": path.suffix.lower().lstrip("."),
            "carpeta_raiz": carpeta_raiz,
            "carpeta_contenedora": str(path.parent.relative_to(BRONZE_DIR)).replace("\\", "/"),
            "tamano_bytes": size_bytes,
            "tamano_legible": tamano_legible(size_bytes),
            "modalidad": clasificar_modalidad(ruta_bajo_bronze),
        })

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "bronze_manifest.json"
    out_path.write_text(json.dumps(registros, ensure_ascii=False, indent=2), encoding="utf-8")

    por_modalidad: dict[str, int] = {}
    por_extension: dict[str, int] = {}
    for r in registros:
        por_modalidad[r["modalidad"]] = por_modalidad.get(r["modalidad"], 0) + 1
        por_extension[r["extension"]] = por_extension.get(r["extension"], 0) + 1

    print(f"Catálogo Bronze completo: {len(registros)} archivos -> {rel(out_path)}")
    print(f"  - por modalidad: {por_modalidad}")
    print(f"  - por extensión: {por_extension}")


if __name__ == "__main__":
    main()
