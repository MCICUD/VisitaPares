"""Bronze -> Silver

Egresados de la MCIC, cruzando TRES fuentes crudas para ser lo más completo
y transparente posible sobre "cuántos se han graduado":

  1. Data/Bronze/Estados/Listado_de_estudiantes_por_estado_{...}.csv — el
     roster oficial Cóndor de los 7 proyectos curriculares de la MCIC. Trae
     el estado ACTUAL de cada estudiante ("E - Graduado" incluido) y su
     "Ultima Matricula" (periodo AAAA-N), pero NO trae fecha de grado.
  2. Data/Bronze/PII_Interno/Informacion egresados 2023.xlsx — acta de grado
     real (con FECHA GRADO y PROMEDIO) de los 4 proyectos con énfasis
     (195/295/395/495), pero el archivo está fechado en 2023: no existe
     una versión más reciente con fecha de grado para 2024-2026.
  3. Data/Bronze/PII_Interno/Información Egresados Maestria y Doctorados
     2025.xlsx — roster de egresados actualizado a 2025 (más completo que
     el acta de 2023: no se limita a quienes tienen acta digitalizada), sin
     fecha de grado tampoco. Se usa para verificar el total por proyecto
     contra Estados/.

Por qué no hay un desglose 2024-2026 con fecha exacta
------------------------------------------------------
Se revisó Estados/, el acta de grado y el roster 2025, además de la carpeta
cruda equivalente en el repositorio de centralización de datos del programa
(WorkStudyStudent/Bronze) — NINGUNA fuente disponible guarda la fecha de
grado de un estudiante graduado entre 2024 y 2026: Cóndor solo conserva el
estado actual ("Graduado"), no cuándo ocurrió. Por eso este script no
inventa un desglose año a año con fecha real para ese tramo.

Qué se hace en su lugar (a pedido explícito, con validación)
----------------------------------------------------------------
Se usa el año de "Ultima Matricula" de cada estudiante marcado como
Graduado en Estados/ como ESTIMACIÓN de su año de grado (normalmente el
grado ocurre pocos meses después del último periodo matriculado). Esto sí
cubre 2022-2026 porque Estados/ es el roster vivo del sistema. Para saber
qué tan buena es esta estimación, se compara contra los casos donde SÍ hay
fecha real (los graduados de 2022-2023 del acta): se mide en cuántos de
esos casos el año estimado coincide con el año real.

El desglose se deja POR PROYECTO CURRICULAR (los 7 códigos reales de
Estados/), no forzado a las 4 categorías de "énfasis" del archivo de
énfasis vigentes — 595/695/95 no son énfasis, son modalidades/código
histórico distintos, y no hay una tabla de equivalencia oficial para
mapearlos a Geomática/Ing.Software/Teleinf./Sistemas de Información sin
inventarla. Mostrar el proyecto curricular real evita esa invención.
"""
from __future__ import annotations

import csv
import datetime
import json
import re
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SILVER_DIR = ROOT / "Data/Silver"
ESTADOS_DIR = ROOT / "Data/Bronze/Estados"
ARCHIVO_ACTA = ROOT / "Data/Bronze/PII_Interno/Informacion egresados 2023.xlsx"
ARCHIVO_ROSTER_2025 = ROOT / "Data/Bronze/PII_Interno/Información Egresados Maestria y Doctorados 2025.xlsx"

ANIO_MIN, ANIO_MAX = 2022, 2026  # alcance solicitado para lo que se presenta en el sitio

# Columnas del CSV de Estados por posición (igual que app/extract_estado_academico.py).
COL_COD_ESTUDIANTE = 3
COL_COD_PROYECTO = 5
COL_PROYECTO_CURRICULAR = 6
COL_ESTADO_CODIGO = 7
COL_ULTIMA_MATRICULA = 14

HOJAS_ACTA = ("195", "295", "395", "495")  # únicos 4 proyectos con acta de grado digitalizada
PROYECTOS_TODOS = ("95", "195", "295", "395", "495", "595", "695")

ANIO_RE = re.compile(r"^(\d{4})")


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def leer_csv_estados(path: Path) -> list[list[str]]:
    with path.open(encoding="utf-8-sig") as fh:
        lines = fh.readlines()
    reader = csv.reader(lines[1:])  # saltar título suelto
    next(reader)  # descartar encabezado
    return [row for row in reader if any(cell.strip() for cell in row)]


def anio_de_ultima_matricula(texto: str) -> int | None:
    if not texto:
        return None
    m = ANIO_RE.match(texto.strip())
    return int(m.group(1)) if m else None


def cargar_graduados_estados() -> dict[str, dict]:
    """Devuelve {cod_proyecto: {"proyecto_curricular", "graduados": [{"codigo", "anio_estimado"}]}}."""
    resultado = {}
    for cod_proyecto in PROYECTOS_TODOS:
        path = ESTADOS_DIR / f"Listado_de_estudiantes_por_estado_{cod_proyecto}.csv"
        if not path.exists():
            continue
        filas = leer_csv_estados(path)
        graduados = []
        nombre_proyecto = filas[0][COL_PROYECTO_CURRICULAR].strip() if filas else cod_proyecto
        for fila in filas:
            if fila[COL_ESTADO_CODIGO].strip() != "E":
                continue
            graduados.append({
                "codigo": fila[COL_COD_ESTUDIANTE].strip(),
                "anio_estimado": anio_de_ultima_matricula(fila[COL_ULTIMA_MATRICULA]),
            })
        resultado[cod_proyecto] = {"proyecto_curricular": nombre_proyecto, "graduados": graduados}
    return resultado


def cargar_totales_roster_2025() -> dict[str, set[str]]:
    """Devuelve {proyecto_curricular_texto: {codigos}} desde el roster 2025 (solo cubre los 4 énfasis)."""
    if not ARCHIVO_ROSTER_2025.exists():
        return {}
    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = openpyxl.load_workbook(ARCHIVO_ROSTER_2025, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    filas = list(ws.iter_rows(values_only=True))
    # fila 0: título institucional suelto; fila 1: encabezado; fila 2: separador en blanco
    por_proyecto: dict[str, set[str]] = {}
    for fila in filas[3:]:
        if not fila or fila[1] is None:
            continue
        codigo = str(fila[1]).strip()
        programa = (fila[7] or "").strip() if len(fila) > 7 else ""
        if not programa:
            continue
        por_proyecto.setdefault(programa, set()).add(codigo)
    return por_proyecto


def cargar_acta_2023() -> dict[str, dict]:
    """Devuelve {cod_proyecto: {"total_graduados", "por_anio", "promedio_academico", "codigos_por_anio"}}."""
    if not ARCHIVO_ACTA.exists():
        raise FileNotFoundError(f"No existe {ARCHIVO_ACTA}")
    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = openpyxl.load_workbook(ARCHIVO_ACTA, data_only=True, read_only=True)

    resultado = {}
    for nombre_hoja in HOJAS_ACTA:
        if nombre_hoja not in wb.sheetnames:
            continue
        ws = wb[nombre_hoja]
        filas = list(ws.iter_rows(values_only=True))[1:]  # saltar encabezado
        por_anio: dict[int, int] = {}
        codigos_por_anio: dict[int, list[str]] = {}
        promedios = []
        total_hoja = 0
        for fila in filas:
            if not fila or fila[0] is None:
                continue
            fecha_grado = fila[3] if len(fila) > 3 else None
            if not isinstance(fecha_grado, datetime.datetime):
                continue
            if not (ANIO_MIN <= fecha_grado.year <= ANIO_MAX):
                continue
            total_hoja += 1
            anio = fecha_grado.year
            por_anio[anio] = por_anio.get(anio, 0) + 1
            codigos_por_anio.setdefault(anio, []).append(str(int(fila[0])))
            promedio_raw = fila[9] if len(fila) > 9 else None
            try:
                promedios.append(float(str(promedio_raw).replace(",", ".")))
            except (TypeError, ValueError):
                pass
        resultado[nombre_hoja] = {
            "total_graduados": total_hoja,
            "por_anio": dict(sorted(por_anio.items())),
            "codigos_por_anio": codigos_por_anio,
            "promedio_academico": round(sum(promedios) / len(promedios), 2) if promedios else None,
        }
    return resultado


def validar_estimacion_contra_fecha_real(graduados_estados: dict, acta: dict) -> dict:
    """Para los códigos que SÍ tienen fecha de grado real (acta 2022-2023),
    compara esa fecha contra el año estimado a partir de 'Ultima Matricula'
    en Estados/, para reportar qué tan confiable es la estimación antes de
    usarla para 2024-2026 (donde no hay fecha real con qué comparar)."""
    anio_real_por_codigo: dict[str, int] = {}
    for datos in acta.values():
        for anio, codigos in datos["codigos_por_anio"].items():
            for cod in codigos:
                anio_real_por_codigo[cod] = anio

    anio_estimado_por_codigo: dict[str, int | None] = {}
    for datos in graduados_estados.values():
        for g in datos["graduados"]:
            anio_estimado_por_codigo[g["codigo"]] = g["anio_estimado"]

    comparables = 0
    exactos = 0
    mas_menos_1 = 0
    sin_estimacion = 0
    for cod, anio_real in anio_real_por_codigo.items():
        if cod not in anio_estimado_por_codigo:
            continue  # graduado en el acta pero no encontrado como "Graduado" en Estados/ (raro, se reporta aparte)
        comparables += 1
        anio_est = anio_estimado_por_codigo[cod]
        if anio_est is None:
            sin_estimacion += 1
            continue
        diferencia = abs(anio_est - anio_real)
        if diferencia == 0:
            exactos += 1
        elif diferencia == 1:
            mas_menos_1 += 1

    graduados_acta_no_en_estados = sum(
        1 for cod in anio_real_por_codigo if cod not in anio_estimado_por_codigo
    )

    return {
        "descripcion": (
            "Para los graduados de 2022-2023 (única ventana con fecha de grado real, del acta), se compara "
            "esa fecha contra el año estimado a partir de 'Ultima Matricula' en Estados/ — la misma técnica "
            "que se usa para estimar el año de grado en 2024-2026, donde no existe fecha real con qué comparar."
        ),
        "estudiantes_comparables": comparables,
        "coinciden_anio_exacto": exactos,
        "coinciden_anio_mas_menos_1": mas_menos_1,
        "sin_anio_estimable": sin_estimacion,
        "graduados_en_acta_sin_registro_graduado_en_estados": graduados_acta_no_en_estados,
    }


def main() -> None:
    graduados_estados = cargar_graduados_estados()
    roster_2025 = cargar_totales_roster_2025()
    acta = cargar_acta_2023()

    total_historico_graduados = sum(len(d["graduados"]) for d in graduados_estados.values())

    por_proyecto_historico = []
    for cod_proyecto, datos in graduados_estados.items():
        total_estados = len(datos["graduados"])
        total_roster_2025 = len(roster_2025.get(datos["proyecto_curricular"], set()))
        por_proyecto_historico.append({
            "cod_proyecto": cod_proyecto,
            "proyecto_curricular": datos["proyecto_curricular"],
            "total_graduados_historico": total_estados,
            "total_graduados_roster_2025": total_roster_2025 if datos["proyecto_curricular"] in roster_2025 else None,
        })

    por_anio_estimado: dict[int, int] = {}
    por_proyecto_por_anio_estimado = []
    for cod_proyecto, datos in graduados_estados.items():
        por_anio_proyecto: dict[int, int] = {}
        for g in datos["graduados"]:
            anio = g["anio_estimado"]
            if anio is None or not (ANIO_MIN <= anio <= ANIO_MAX):
                continue
            por_anio_proyecto[anio] = por_anio_proyecto.get(anio, 0) + 1
            por_anio_estimado[anio] = por_anio_estimado.get(anio, 0) + 1
        por_proyecto_por_anio_estimado.append({
            "cod_proyecto": cod_proyecto,
            "proyecto_curricular": datos["proyecto_curricular"],
            "por_anio_estimado": dict(sorted(por_anio_proyecto.items())),
        })

    con_fecha_real_por_anio: dict[int, int] = {}
    con_fecha_real_total = 0
    por_enfasis_fecha_real = []
    for cod_proyecto in HOJAS_ACTA:
        datos = acta.get(cod_proyecto)
        if datos is None:
            continue
        con_fecha_real_total += datos["total_graduados"]
        for anio, n in datos["por_anio"].items():
            con_fecha_real_por_anio[anio] = con_fecha_real_por_anio.get(anio, 0) + n
        por_enfasis_fecha_real.append({
            "cod_proyecto": cod_proyecto,
            "total_graduados": datos["total_graduados"],
            "por_anio": datos["por_anio"],
            "promedio_academico": datos["promedio_academico"],
        })

    validacion = validar_estimacion_contra_fecha_real(graduados_estados, acta)

    data = {
        "rango_presentado": f"{ANIO_MIN}-{ANIO_MAX}",
        "total_historico_graduados": total_historico_graduados,
        "por_proyecto_historico": por_proyecto_historico,
        "por_anio_estimado": dict(sorted(por_anio_estimado.items())),
        "por_proyecto_por_anio_estimado": por_proyecto_por_anio_estimado,
        "con_fecha_real": {
            "rango": f"{min(con_fecha_real_por_anio)}-{max(con_fecha_real_por_anio)}" if con_fecha_real_por_anio else None,
            "total_graduados": con_fecha_real_total,
            "por_anio": dict(sorted(con_fecha_real_por_anio.items())),
            "por_enfasis": por_enfasis_fecha_real,
        },
        "validacion_estimacion": validacion,
        "metodologia": (
            "Cóndor (Data/Bronze/Estados/) no guarda la fecha de grado, solo el estado actual del estudiante. "
            "'total_historico_graduados' y 'por_proyecto_historico' cuentan TODOS los estudiantes marcados "
            "'E - Graduado' en el roster, de cualquier año. 'por_anio_estimado' aproxima el año de grado con el "
            "año de la última matrícula registrada (columna 'Ultima Matricula') de cada graduado — no es la "
            "fecha real de grado, es la mejor aproximación disponible; ver 'validacion_estimacion' para su "
            "margen de error real, medido contra los únicos casos con fecha exacta conocida (2022-2023). "
            "'con_fecha_real' es la única sección con fecha de grado exacta (acta oficial), pero el archivo "
            "fuente no tiene versión posterior a 2023."
        ),
        "limitacion": (
            "No existe, en ningún sistema disponible (Cóndor ni los archivos sueltos de Secretaría), un registro "
            "de fecha de grado exacta para 2024-2026. El desglose de esos años es una estimación explícita, no "
            "un dato de acta; no se presenta como si lo fuera."
        ),
        "fuente": {
            "estados": {
                "carpeta": "Data/Bronze/Estados/",
                "nota": "7 CSV por proyecto curricular; solo se usan las columnas Cod. Estudiante (para cruzar), Estado y Ultima Matricula.",
            },
            "acta_fecha_real": {
                "archivo": rel(ARCHIVO_ACTA),
                "nota": "Archivo con nombre/documento/correo por egresado; no se publica en el catálogo de Data/Bronze.",
            },
            "roster_2025": {
                "archivo": rel(ARCHIVO_ROSTER_2025),
                "nota": "Roster de egresados actualizado a 2025, usado solo para verificar el total por proyecto contra Estados/; no se publica en el catálogo de Data/Bronze.",
            },
        },
    }

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "egresados_agregado.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Egresados: {total_historico_graduados} graduados históricos (todos los proyectos, todos los años) -> {out_path.relative_to(ROOT)}")
    print(f"  - Con fecha de grado real ({data['con_fecha_real']['rango']}): {con_fecha_real_total}")
    print(f"  - Estimado por año (Última Matrícula), {ANIO_MIN}-{ANIO_MAX}: {dict(sorted(por_anio_estimado.items()))}")
    print(f"  - Validación estimación vs. fecha real: {validacion}")


if __name__ == "__main__":
    main()
