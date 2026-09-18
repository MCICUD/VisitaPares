"""Bronze -> Silver

Indexa la evidencia REAL de seguimiento del Plan de Mejoramiento que la
coordinación ya ha ido archivando por factor y por actividad (carpetas
"a.", "b.", "c." ... que corresponden 1 a 1 con las "ACTIVIDADES REQUERIDAS
PARA LOGRAR META" de cada factor en el CC-FR-001).

Esto es intencionalmente solo un INVENTARIO (nombre, ruta, tamaño) — no se
interpreta ni se resume el contenido de cada documento, para no inventar
nada. Si una actividad no tiene ningún archivo cargado todavía en Bronze,
el sitio debe decir exactamente eso: que no hay evidencia cargada aún (el
Plan 2026-2027 recién empieza a ejecutarse en agosto de 2026).
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib.xlsx_reader import fuente  # noqa: E402
from build_bronze_manifest import tamano_legible  # noqa: E402

SILVER_DIR = ROOT / "Data/Silver"

MODALIDADES = {
    "investigacion": ROOT / "Data/Bronze/MCIC.INVESTIGACION/Procesos de Renocavion y acreditación",
    "profundizacion": ROOT / "Data/Bronze/MCIC-PROFUNDIZACION/Procesos de Renocavion y acreditación",
}

FACTOR_RE = re.compile(r"^FACTOR\s*0*(\d{1,2})\b", re.IGNORECASE)


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def listar_archivos(carpeta: Path) -> list[dict]:
    archivos = []
    if not carpeta.exists():
        return archivos
    for path in sorted(carpeta.rglob("*")):
        if path.is_file():
            archivos.append({
                "nombre": path.name,
                "archivo": rel(path),
                "tamano_legible": tamano_legible(path.stat().st_size),
            })
    return archivos


def procesar_modalidad(modalidad: str, base: Path) -> dict:
    seguimiento_dir = base / "Plan de Mejoramiento"
    rc_aac_dir = base / "Procesos de RC y AAC"

    documentos_generales = []
    factores: dict[str, dict] = {str(n): {"factor_nombre": None, "actividades": []} for n in range(1, 13)}

    if seguimiento_dir.exists():
        for item in sorted(seguimiento_dir.iterdir()):
            m = FACTOR_RE.match(item.name) if item.is_dir() else None
            if item.is_file():
                documentos_generales.append({
                    "nombre": item.name,
                    "archivo": rel(item),
                    "tamano_legible": tamano_legible(item.stat().st_size),
                })
            elif m:
                numero = m.group(1)
                actividades = []
                for actividad_dir in sorted(item.iterdir()):
                    if actividad_dir.is_dir():
                        actividades.append({
                            "nombre": actividad_dir.name,
                            "archivos": listar_archivos(actividad_dir),
                        })
                    elif actividad_dir.is_file():
                        # archivo suelto directamente bajo el FACTOR, sin actividad asignada
                        actividades.append({
                            "nombre": "(sin actividad específica)",
                            "archivos": [{
                                "nombre": actividad_dir.name,
                                "archivo": rel(actividad_dir),
                                "tamano_legible": tamano_legible(actividad_dir.stat().st_size),
                            }],
                        })
                factores[numero] = {"factor_nombre": item.name, "actividades": actividades}

    return {
        "modalidad": modalidad,
        "documentos_generales": documentos_generales,
        "procesos_rc_aac": listar_archivos(rc_aac_dir),
        "factores": factores,
    }


def main() -> None:
    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    resumen = []
    for modalidad, base in MODALIDADES.items():
        data = procesar_modalidad(modalidad, base)
        out_path = SILVER_DIR / f"seguimiento_evidencia_{modalidad}.json"
        out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        n_archivos = sum(
            len(a["archivos"]) for f in data["factores"].values() for a in f["actividades"]
        ) + len(data["documentos_generales"]) + len(data["procesos_rc_aac"])
        resumen.append((modalidad, n_archivos, out_path.relative_to(ROOT)))

    print("Evidencia de seguimiento indexada:")
    for modalidad, n, out in resumen:
        print(f"  - {modalidad}: {n} archivos -> {out}")


if __name__ == "__main__":
    main()
