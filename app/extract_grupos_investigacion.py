"""Bronze -> Silver

Lee el directorio de grupos de investigación de la Facultad con docentes
que pueden dirigir trabajos de grado en la MCIC (una hoja por grupo).
Datos institucionales/profesionales de docentes (nombre, correo
institucional, CVLAC público) — no es información personal sensible.
"""
from __future__ import annotations

import json
import sys
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SILVER_DIR = ROOT / "Data/Silver"
ARCHIVO = ROOT / "Data/Bronze/Maestria CIC/Directorio Grupos de Inv MCIC.xlsx"

CAMPOS_ENCABEZADO = {
    "nombre:": "nombre",
    "líneas de investigación:": "lineas_investigacion",
    "clasificación - colciencias:": "clasificacion",
    "gruplac:": "gruplac",
    "pagina web": "pagina_web",
    "e-mail:": "email",
    "lider:": "lider",
}


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def limpiar(v) -> str | None:
    if v is None:
        return None
    v = str(v).strip()
    return v or None


def procesar_hoja(ws) -> dict:
    filas = list(ws.iter_rows(values_only=True))
    grupo = {v: None for v in CAMPOS_ENCABEZADO.values()}
    integrantes = []
    fila_header_integrantes = None

    for i, fila in enumerate(filas):
        etiqueta = limpiar(fila[1]) if len(fila) > 1 else None
        valor = limpiar(fila[2]) if len(fila) > 2 else None
        if etiqueta and etiqueta.lower() in CAMPOS_ENCABEZADO:
            grupo[CAMPOS_ENCABEZADO[etiqueta.lower()]] = valor
        if etiqueta and etiqueta.lower().startswith("nombres y apellidos"):
            fila_header_integrantes = i

    if fila_header_integrantes is not None:
        for fila in filas[fila_header_integrantes + 1:]:
            nombre = limpiar(fila[1]) if len(fila) > 1 else None
            if not nombre:
                continue
            integrantes.append({
                "nombre": nombre,
                "correo": limpiar(fila[2]) if len(fila) > 2 else None,
                "tematicas": limpiar(fila[3]) if len(fila) > 3 else None,
                "cvlac": limpiar(fila[4]) if len(fila) > 4 else None,
            })

    grupo["integrantes"] = integrantes
    return grupo


def main() -> None:
    if not ARCHIVO.exists():
        raise FileNotFoundError(f"No existe {ARCHIVO}")

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = __import__("openpyxl").load_workbook(ARCHIVO, data_only=True, read_only=True)

    # Mapeo oficial de grupos y clasificaciones
    CLASIFICACIONES_OFICIALES = {
        "GIIRA": "A1",
        "INTECSE": "A",
        "INTERNET INTELIGENTE": "A",
        "LIDER": "A",
        "LIFAE": "A",
        "NIDE": "A",
        "GICOECOL": "B",
        "GITEM++": "B",
        "GRECO": "B",
        "LASER": "B",
        "GCEM": "C",
        "GEFEM": "C",
        "GESETIC": "C",
        "GICOGE": "C",
        "GITUD": "C",
        "LAMIC": "C",
        "ITI": "Facultad Tecnológica",
        "ARMOS": "Facultad Tecnológica",
        "GIDENUTAS": "Facultad Tecnológica"
    }

    grupos = []
    grupos_encontrados = set()

    for nombre_hoja in wb.sheetnames:
        llave = nombre_hoja.upper()
        if llave not in CLASIFICACIONES_OFICIALES:
            continue  # Excluir grupos que no están en la lista oficial
            
        grupo = procesar_hoja(wb[nombre_hoja])
        grupo["sigla"] = nombre_hoja
        grupo["fuente"] = {"archivo": rel(ARCHIVO), "hoja": nombre_hoja}
        grupo["clasificacion"] = CLASIFICACIONES_OFICIALES[llave]  # Forzar la clasificación
        grupos.append(grupo)
        grupos_encontrados.add(llave)

    # Añadir grupos que no estaban en el Excel (ej. ARMOS, GIDENUTAS)
    for sigla, clasificacion in CLASIFICACIONES_OFICIALES.items():
        if sigla not in grupos_encontrados:
            grupos.append({
                "nombre": sigla,
                "sigla": sigla,
                "clasificacion": clasificacion,
                "integrantes": [],
                "proyectos_grado": [],
                "fuente": {"archivo": "Lista oficial manual", "hoja": sigla}
            })

    data = {"grupos": grupos}
    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "grupos_investigacion.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    total_integrantes = sum(len(g["integrantes"]) for g in grupos)
    print(f"Grupos de investigación: {len(grupos)} grupos, {total_integrantes} docentes -> {out_path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
