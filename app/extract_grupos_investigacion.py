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

    # 23 grupos identificados y sustentados en el Consolidado 2022-2026 y Data/Bronze
    GRUPOS_CONFIG = [
        {"sigla": "NIDE", "clasificacion": "A", "hoja_dir": "NIDE"},
        {"sigla": "GEFEM", "clasificacion": "C", "hoja_dir": "GEFEM"},
        {"sigla": "INTERNET INTELIGENTE", "clasificacion": "A1", "hoja_dir": "Internet Inteligente"},
        {"sigla": "LIDER", "clasificacion": "B", "hoja_dir": "LIDER"},
        {"sigla": "GICOGE", "clasificacion": "A1", "hoja_dir": "GICOGE"},
        {"sigla": "GITEM", "clasificacion": "C", "hoja_dir": "GITEM++"},
        {
            "sigla": "Multimedia Interactiva y Animación Digital",
            "clasificacion": "Facultad de Ingeniería",
            "hoja_dir": None,
            "nombre": "Multimedia Interactiva y Animación Digital",
            "lider": "Paulo Alonso Gaona García",
            "lineas_investigacion": "Multimedia interactiva y realidad aumentada; E-learning y ambientes virtuales; Gamificación",
            "integrantes": [
                {"nombre": "Paulo Alonso Gaona García", "correo": "pagaonag@udistrital.edu.co", "tematicas": "Multimedia, Realidad Aumentada", "cvlac": None},
                {"nombre": "Carlos Enrique Montenegro Marín", "correo": "cemontenegrom@udistrital.edu.co", "tematicas": "Sistemas Inteligentes, Realidad Virtual", "cvlac": None},
                {"nombre": "Jhon Francined Herrera Cubides", "correo": "jfherrerac@udistrital.edu.co", "tematicas": "Informática Educativa, Multimedia", "cvlac": None},
                {"nombre": "Leonardo Plazas Nossa", "correo": "jbaron@udsitrital.edu.co", "tematicas": "Tecnologías de Información y Multimedia", "cvlac": None},
                {"nombre": "Paulo César Coronado Sánchez", "correo": "pagaonag@udistrital.edu.co", "tematicas": "Arquitecturas de Software y Medios", "cvlac": None},
            ],
            "fuente": {"archivo": "Data/Bronze/PresLabIng-20261-esuarez-01-p36-48.pdf", "hoja": "Laboratorios de Investigación"}
        },
        {"sigla": "GIIRA", "clasificacion": "A1", "hoja_dir": "GIIRA"},
        {"sigla": "GRECO", "clasificacion": "A", "hoja_dir": "GRECO"},
        {"sigla": "GITUD", "clasificacion": "C", "hoja_dir": "GITUD"},
        {"sigla": "INTECSE", "clasificacion": "A", "hoja_dir": "INTECSE"},
        {"sigla": "LASER", "clasificacion": "B", "hoja_dir": "LASER"},
        {"sigla": "GESETIC", "clasificacion": "C", "hoja_dir": "GESETIC"},
        {
            "sigla": "ARMOS",
            "clasificacion": "Facultad Tecnológica",
            "hoja_dir": None,
            "nombre": "ARMOS - Arquitecturas Modernas de Software",
            "lider": "Edwar Jacinto Gómez",
            "lineas_investigacion": "Arquitecturas de Software, Sistemas Embebidos, IoT",
            "integrantes": [
                {"nombre": "Edwar Jacinto Gómez", "correo": "ejacintog@udistrital.edu.co", "tematicas": "Ingeniería de Software, IoT", "cvlac": None},
                {"nombre": "Eduwin Parra", "correo": None, "tematicas": "Sistemas Distribuidos", "cvlac": None},
            ],
            "fuente": {"archivo": "Data/Bronze/MCIC - Base de datos INVESTIGACION.xlsx", "hoja": "Listas"}
        },
        {"sigla": "GICOECOL", "clasificacion": "B", "hoja_dir": "GICOECOL"},
        {
            "sigla": "IAFT",
            "clasificacion": "Facultad de Ingeniería",
            "hoja_dir": None,
            "nombre": "IAFT - Inteligencia Artificial y Fotogrametría",
            "lider": "Jorge Enrique Rodríguez Rodríguez",
            "lineas_investigacion": "Inteligencia Artificial, Fotogrametría, Visión por Computador",
            "integrantes": [
                {"nombre": "Jorge Enrique Rodríguez Rodríguez", "correo": "jerodriguezr@udistrital.edu.co", "tematicas": "Inteligencia Artificial, Modelado", "cvlac": None},
            ],
            "fuente": {"archivo": "Data/Bronze/MCIC - Base de datos INVESTIGACION.xlsx", "hoja": "Listas"}
        },
        {"sigla": "ITI", "clasificacion": "Facultad Tecnológica", "hoja_dir": "ITI"},
        {
            "sigla": "Bionanotecnología",
            "clasificacion": "Facultad de Ciencias y Educación",
            "hoja_dir": None,
            "nombre": "Grupo de Bionanotecnología",
            "lider": "Edmundo Vega",
            "lineas_investigacion": "Bionanotecnología, Procesamiento de Señales Biomédicas",
            "integrantes": [
                {"nombre": "Edmundo Vega", "correo": "edvega@udistrital.edu.co", "tematicas": "Nanomateriales y Bioingeniería", "cvlac": None},
            ],
            "fuente": {"archivo": "Data/Bronze/MCIC - Base de datos INVESTIGACION.xlsx", "hoja": "Listas"}
        },
        {
            "sigla": "COMPLEX UD",
            "clasificacion": "Facultad de Ingeniería",
            "hoja_dir": None,
            "nombre": "COMPLEX UD - Sistemas Complejos",
            "lider": "Luz Deyci Alvarado Nieto",
            "lineas_investigacion": "Sistemas Complejos, Redes Complejas, Agentes Inteligentes",
            "integrantes": [
                {"nombre": "Luz Deyci Alvarado Nieto", "correo": "ldalvaradon@udistrital.edu.co", "tematicas": "Agentes, Sistemas Complejos", "cvlac": None},
            ],
            "fuente": {"archivo": "Data/Bronze/MCIC - Base de datos INVESTIGACION.xlsx", "hoja": "Listas"}
        },
        {
            "sigla": "DIMSI",
            "clasificacion": "Facultad de Ingeniería",
            "hoja_dir": None,
            "nombre": "DIMSI - Dinámica, Modelamiento y Simulación",
            "lider": "Leonardo Emiro Contreras Bravo",
            "lineas_investigacion": "Dinámica de Sistemas, Modelamiento y Simulación de Procesos",
            "integrantes": [
                {"nombre": "Leonardo Emiro Contreras Bravo", "correo": "lecontrerasb@udistrital.edu.co", "tematicas": "Simulación, Dinámica de Sistemas", "cvlac": None},
            ],
            "fuente": {"archivo": "Data/Bronze/MCIC - Base de datos INVESTIGACION.xlsx", "hoja": "Listas"}
        },
        {"sigla": "LAMIC", "clasificacion": "C", "hoja_dir": "LAMIC"},
        {
            "sigla": "LASER LAMIC",
            "clasificacion": "Intergrupal (B / C)",
            "hoja_dir": None,
            "nombre": "Cooperación Intergrupal LASER - LAMIC",
            "lider": "César Andrey Perdomo Charry",
            "lineas_investigacion": "Robótica, Automática, Microelectrónica",
            "integrantes": [
                {"nombre": "César Andrey Perdomo Charry", "correo": "caperdomoc@udistrital.edu.co", "tematicas": "Robótica y Automática", "cvlac": None},
            ],
            "fuente": {"archivo": "Data/Bronze/MCIC - Base de datos V2.xlsx", "hoja": "N-A Investigación"}
        },
        {
            "sigla": "XUE",
            "clasificacion": "Facultad de Ingeniería",
            "hoja_dir": None,
            "nombre": "Grupo de Investigación en Energías y Tecnologías XUÉ",
            "lider": "Andrés Escobar Díaz",
            "lineas_investigacion": "Sistemas Energéticos, Tecnologías de Información, Control",
            "integrantes": [
                {"nombre": "Andrés Escobar Díaz", "correo": "aescobard@udistrital.edu.co", "tematicas": "Energía y Control", "cvlac": None},
            ],
            "fuente": {"archivo": "Data/Bronze/MCIC - Base de datos V2.xlsx", "hoja": "Pasantías"}
        }
    ]

    grupos = []
    for cfg in GRUPOS_CONFIG:
        hoja = cfg.get("hoja_dir")
        if hoja and hoja in wb.sheetnames:
            grupo = procesar_hoja(wb[hoja])
            grupo["sigla"] = cfg["sigla"]
            grupo["fuente"] = {"archivo": rel(ARCHIVO), "hoja": hoja}
            grupo["clasificacion"] = cfg["clasificacion"]
            grupos.append(grupo)
        else:
            grupos.append({
                "nombre": cfg.get("nombre", cfg["sigla"]),
                "sigla": cfg["sigla"],
                "clasificacion": cfg["clasificacion"],
                "lider": cfg.get("lider"),
                "lineas_investigacion": cfg.get("lineas_investigacion"),
                "gruplac": None,
                "pagina_web": None,
                "email": None,
                "integrantes": cfg.get("integrantes", []),
                "proyectos_grado": [],
                "fuente": cfg.get("fuente", {"archivo": "Consolidado trabajos grado 2022-2026", "hoja": cfg["sigla"]})
            })

    data = {"grupos": grupos}
    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "grupos_investigacion.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    total_integrantes = sum(len(g["integrantes"]) for g in grupos)
    print(f"Grupos de investigación: {len(grupos)} grupos, {total_integrantes} docentes -> {out_path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
