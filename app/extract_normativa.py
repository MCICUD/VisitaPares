"""Bronze -> Silver

Normativa que rige a la Maestría (no está ligada a un factor del plan): reglamento de
posgrados, régimen de matrículas, reglamento de trabajo de grado y estatuto estudiantil.
Cada norma indica a qué modalidad aplica y, cuando el documento tiene partes por
modalidad, la página que corresponde a cada una.
"""
from __future__ import annotations

import json
from pathlib import Path

from build_bronze_manifest import tamano_legible

ROOT = Path(__file__).resolve().parent.parent
SILVER_DIR = ROOT / "Data/Silver"
CARPETA = ROOT / "Data/Bronze/Normativa/NORMATIVA"

AMBAS = "Investigación y Profundización"
NORMAS = [
    {
        "archivo": "Acuerdo N 01 de 2025- REGLAMENTO DE POSGRADO NUEVO.pdf",
        "titulo": "Reglamento de posgrados de la Facultad de Ingeniería",
        "norma": "Acuerdo 01 de 2025", "fecha": "4 de septiembre de 2025", "expide": "Consejo de Facultad de Ingeniería",
        "estado": "Vigente (deroga el Acuerdo 01 de 2009)",
        "aplica": {"investigacion": AMBAS, "profundizacion": AMBAS},
    },
    {
        "archivo": "Acuerdo 001 de 2009 REGLAMENTO POSGRADOS ANTERIOR.pdf",
        "titulo": "Reglamento de estudiantes de posgrado de la Facultad de Ingeniería",
        "norma": "Acuerdo 001 de 2009", "fecha": "2009", "expide": "Consejo de Facultad de Ingeniería",
        "estado": "Anterior (derogado por el Acuerdo 01 de 2025)",
        "aplica": {"investigacion": AMBAS, "profundizacion": AMBAS},
    },
    {
        "archivo": "ACUERDO TRABAJOS DE GRADO MAESTRÍAS.pdf",
        "titulo": "Reglamento del trabajo de grado de las maestrías de la Facultad de Ingeniería (con la modificación de 2022)",
        "norma": "Acuerdo 01 de 2019 y Acuerdo 04 de 2022", "fecha": "5 de marzo de 2019 y 27 de abril de 2022", "expide": "Consejo Académico",
        "estado": "Vigente",
        "aplica": {"investigacion": AMBAS, "profundizacion": AMBAS},
        "detalle": {
            "investigacion": {"texto": "Le corresponde la Parte 1 (maestrías de investigación, págs. 2-9) y las modificaciones de los artículos 7, 9 y 16 (págs. 23-24).", "pagina": 2},
            "profundizacion": {"texto": "Le corresponde la Parte 2 (maestrías de profundización, págs. 10-19) y las modificaciones de los artículos 29 y 34 (págs. 25-26).", "pagina": 10},
        },
    },
    {
        "archivo": "ESTATUTO ESTUDIANTIL.pdf",
        "titulo": "Estatuto estudiantil de la Universidad Distrital",
        "norma": "Acuerdo 027 de 1993", "fecha": "23 de diciembre de 1993 (actualizado a agosto de 2019)", "expide": "Consejo Superior Universitario",
        "estado": "Vigente",
        "aplica": {"investigacion": AMBAS, "profundizacion": AMBAS},
    },
    {
        "archivo": "ACUERDO 004 DEL 2006 DESCUENTOS.pdf",
        "titulo": "Régimen de liquidación de matrículas de los estudiantes",
        "norma": "Acuerdo 004 de 2006", "fecha": "25 de enero de 2006", "expide": "Consejo Superior Universitario",
        "estado": "Vigente",
        "aplica": {"investigacion": AMBAS, "profundizacion": AMBAS},
    },
]


def main() -> None:
    normas = []
    for n in NORMAS:
        ruta = CARPETA / n["archivo"]
        if not ruta.exists():
            raise FileNotFoundError(ruta)
        normas.append({**n, "nombre": n["archivo"], "archivo": str(ruta.relative_to(ROOT)).replace("\\", "/"), "tamano_legible": tamano_legible(ruta.stat().st_size)})
    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    (SILVER_DIR / "normativa.json").write_text(json.dumps({"normas": normas}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Normativa: {len(normas)} normas -> Data/Silver/normativa.json")


if __name__ == "__main__":
    main()
