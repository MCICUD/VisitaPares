"""Modalidad (Investigación / Profundización) de cada estudiante MCIC.

Prioridad (la misma que usa el resto del proyecto):
  1. código de proyecto dentro del código estudiantil (AAAAP595NNN / AAAAP695NNN);
  2. proyecto actual en Cóndor (Data/Bronze/Estados, 595 / 695);
  3. bases de datos MCIC 2026 (hojas N-A Investigación, N-A Profundización y Pasantías);
  4. plan anterior (proyectos 95, 195, 295, 395 y 495 sin modalidad registrada): se toma como Investigación.
"""
from __future__ import annotations

import csv
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
ESTADOS_DIR = ROOT / "Data/Bronze/Estados"
BASES_DIR = ROOT / "Data/Bronze/Maestria CIC/2026/BASES  DE DATOS ESTUDIANTES"
PLAN_ANTERIOR = ("95", "195", "295", "395", "495")

INVESTIGACION = "Investigación"
PROFUNDIZACION = "Profundización"


def _proyecto_actual() -> dict[str, str]:
    out: dict[str, str] = {}
    for f in sorted(ESTADOS_DIR.glob("Listado_de_estudiantes_por_estado_*.csv")):
        lineas = f.read_text(encoding="utf-8", errors="replace").splitlines()[1:]
        for r in csv.DictReader(lineas):
            cod = (r.get("Cod. Estudiante") or "").strip()
            if cod:
                out[cod] = (r.get("Cod. Proyecto") or "").strip()
    return out


def _bases() -> dict[str, str]:
    import openpyxl

    fuentes = [
        ("MCIC - Base de datos INVESTIGACION.xlsx", ["N-A Investigación"], INVESTIGACION),
        ("MCIC - Base de datos Profundizacion.xlsx", ["N-A Profundizacion", "Pasantías"], PROFUNDIZACION),
        ("MCIC - Base de datos V2.xlsx", ["N-A Investigación", "N-A Profundizacion", "Pasantías"], None),
    ]
    out: dict[str, str] = {}
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        for archivo, hojas, fija in fuentes:
            wb = openpyxl.load_workbook(BASES_DIR / archivo, read_only=True, data_only=True)
            for h in hojas:
                filas = list(wb[h].iter_rows(values_only=True))
                enc = [str(c).strip() if c else "" for c in filas[0]]
                i_mod = enc.index("MODALIDAD") if "MODALIDAD" in enc else None
                for r in filas[1:]:
                    if not r or not r[0]:
                        continue
                    cod = str(r[0]).split(".")[0].strip()
                    if cod in out:
                        continue
                    mod = fija
                    if mod is None:
                        v = str(r[i_mod] or "") if i_mod is not None else ""
                        if "rofund" in v or h == "Pasantías":
                            mod = PROFUNDIZACION
                        elif "nvestig" in v:
                            mod = INVESTIGACION
                        else:
                            mod = INVESTIGACION if h == "N-A Investigación" else PROFUNDIZACION
                    out[cod] = mod
    return out


class ClasificadorModalidad:
    def __init__(self) -> None:
        self.proyecto = _proyecto_actual()
        self.bases = _bases()

    def de(self, codigo: str) -> str | None:
        """Modalidad del estudiante o None si no es de la Maestría (código sin proyecto reconocible)."""
        proy = codigo[5:8] if len(codigo) == 11 else ""
        if proy == "595":
            return INVESTIGACION
        if proy == "695":
            return PROFUNDIZACION
        actual = self.proyecto.get(codigo)
        if actual == "595":
            return INVESTIGACION
        if actual == "695":
            return PROFUNDIZACION
        if codigo in self.bases:
            return self.bases[codigo]
        if proy in PLAN_ANTERIOR or actual in PLAN_ANTERIOR or (len(codigo) == 7 and actual == "95"):
            return INVESTIGACION
        return None
