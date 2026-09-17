"""Silver -> Gold (artefacto descargable)

Consolida los 3 datasets de la pestaña "Comunidad Estudiantil" (estado
académico histórico por énfasis, estudiantes activos por énfasis y
egresados) en un único archivo .xlsx descargable desde el micrositio, para
que alguien pueda revisar/cruzar esta información en Excel sin tener que
leer los JSON de Data/Silver/ uno por uno.

Cada hoja incluye una columna "Archivo fuente" citando de dónde sale esa
fila en Data/Bronze/ — la misma trazabilidad que ya tiene el sitio, solo que
en formato descargable. No se agrega ni se calcula ningún dato nuevo aquí:
todo viene tal cual de los JSON que ya produce el resto de la pipeline
(Data/Silver/estado_academico_agregado.json, enfasis_estudiantes.json,
egresados_agregado.json).
"""
from __future__ import annotations

import json
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path(__file__).resolve().parent.parent
SILVER_DIR = ROOT / "Data/Silver"
GOLD_DIR = ROOT / "Data/Gold"
OUT_XLSX = GOLD_DIR / "comunidad_estudiantil.xlsx"

ENCABEZADO_FONT = Font(bold=True, color="FFFFFF")
ENCABEZADO_FILL = "1E3A5F"
TITULO_FONT = Font(bold=True, size=13)


def cargar(nombre: str) -> dict:
    path = SILVER_DIR / nombre
    if not path.exists():
        raise FileNotFoundError(f"No existe {path}. Corre primero el script que lo genera en app/.")
    return json.loads(path.read_text(encoding="utf-8"))


def _escribir_titulo(ws: Worksheet, texto: str, num_cols: int) -> None:
    ws.cell(row=1, column=1, value=texto).font = TITULO_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(num_cols, 1))


def _escribir_encabezados(ws: Worksheet, fila: int, encabezados: list[str]) -> None:
    from openpyxl.styles import PatternFill
    for col, texto in enumerate(encabezados, start=1):
        celda = ws.cell(row=fila, column=col, value=texto)
        celda.font = ENCABEZADO_FONT
        celda.fill = PatternFill("solid", fgColor=ENCABEZADO_FILL)
        celda.alignment = Alignment(horizontal="center", wrap_text=True)


def _autoancho(ws: Worksheet, anchos: list[int]) -> None:
    for i, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = ancho


def hoja_leeme(wb: openpyxl.Workbook) -> None:
    ws = wb.active
    ws.title = "Léeme"
    filas = [
        ("Hoja", "Contenido", "Fuente cruda (Data/Bronze/)"),
        (
            "Estado académico",
            "Conteo de registros del roster oficial Cóndor por énfasis y por estado académico "
            "(Matriculado, Graduado, Inactivo, etc.) — incluye TODOS los estados, no solo los que "
            "se muestran en el micrositio.",
            "Data/Bronze/Estados/Listado_de_estudiantes_por_estado_*.csv",
        ),
        (
            "Estudiantes por énfasis",
            "Consolidado de estudiantes activos por énfasis y categoría (Investigación/Profundización, "
            "incluyendo las cohortes 'Grande'), tal como lo mantiene la coordinación.",
            "Data/Bronze/Maestria CIC/Énfasis estudiantes - Consolidado.xlsx",
        ),
        (
            "Egresados - histórico",
            "Total de graduados por énfasis, de cualquier año, según el estado 'E - Graduado' del "
            "roster Cóndor vigente.",
            "Data/Bronze/Estados/Listado_de_estudiantes_por_estado_*.csv",
        ),
        (
            "Egresados - por año 2022-2026",
            "Graduados por énfasis y por año dentro de 2022-2026, ESTIMADO a partir del año de la "
            "última matrícula de cada graduado (Cóndor no registra fecha de grado). Validado contra "
            "los únicos casos con fecha real conocida: ver hoja 'Egresados - fecha real'.",
            "Data/Bronze/Estados/Listado_de_estudiantes_por_estado_*.csv",
        ),
        (
            "Egresados - fecha real 2022-2023",
            "Graduados con FECHA DE GRADO real (acta oficial), por énfasis y por año — único tramo "
            "con fecha exacta disponible; no existe una versión del acta posterior a 2023.",
            "Data/Bronze/PII_Interno/Informacion egresados 2023.xlsx",
        ),
    ]
    for r, fila in enumerate(filas, start=1):
        for c, valor in enumerate(fila, start=1):
            celda = ws.cell(row=r, column=c, value=valor)
            if r == 1:
                celda.font = ENCABEZADO_FONT
                from openpyxl.styles import PatternFill
                celda.fill = PatternFill("solid", fgColor=ENCABEZADO_FILL)
            celda.alignment = Alignment(wrap_text=True, vertical="top")
    _autoancho(ws, [26, 70, 45])
    ws.freeze_panes = "A2"


def hoja_estado_academico(wb: openpyxl.Workbook, data: dict) -> None:
    ws = wb.create_sheet("Estado académico")
    encabezados = ["Cód. Proyecto", "Énfasis / Proyecto Curricular", "Estado", "Estudiantes", "Archivo fuente"]
    _escribir_encabezados(ws, 1, encabezados)
    fila = 2
    for p in data["proyectos"]:
        for estado, conteo in p["conteo_por_estado"].items():
            ws.append([p["cod_proyecto"], p["proyecto_curricular"], estado, conteo, p["fuente"]["archivo"]])
            fila += 1
        ws.append([p["cod_proyecto"], p["proyecto_curricular"], "TOTAL (todos los estados)", p["total_estudiantes"], p["fuente"]["archivo"]])
        for c in range(1, 6):
            ws.cell(row=fila, column=c).font = Font(bold=True)
        fila += 1
    _autoancho(ws, [14, 55, 40, 14, 55])
    ws.freeze_panes = "A2"


def hoja_estudiantes_enfasis(wb: openpyxl.Workbook, data: dict) -> None:
    ws = wb.create_sheet("Estudiantes por énfasis")
    nombres_enfasis = data["enfasis"]
    encabezados = ["Categoría", *nombres_enfasis, "Archivo fuente"]
    ws.cell(row=1, column=1, value=data["titulo_archivo"]).font = Font(italic=True, size=10)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(encabezados))
    _escribir_encabezados(ws, 2, encabezados)
    for i, fila_dato in enumerate(data["filas"], start=3):
        valores = [fila_dato["etiqueta"], *[fila_dato["valores"][n] for n in nombres_enfasis], data["fuente"]["archivo"]]
        ws.append(valores)
        if fila_dato["etiqueta"] == "Total":
            for c in range(1, len(encabezados) + 1):
                ws.cell(row=i, column=c).font = Font(bold=True)
    _autoancho(ws, [32, *([16] * len(nombres_enfasis)), 55])
    ws.freeze_panes = "A3"


def hoja_egresados_historico(wb: openpyxl.Workbook, data: dict) -> None:
    ws = wb.create_sheet("Egresados - histórico")
    encabezados = ["Cód. Proyecto", "Énfasis / Proyecto Curricular", "Graduados históricos (todos los años)", "Roster egresados 2025 (verificación)", "Archivo fuente"]
    _escribir_encabezados(ws, 1, encabezados)
    for p in data["por_proyecto_historico"]:
        ws.append([
            p["cod_proyecto"], p["proyecto_curricular"], p["total_graduados_historico"],
            p["total_graduados_roster_2025"], "Data/Bronze/Estados/",
        ])
    fila_total = len(data["por_proyecto_historico"]) + 2
    ws.append(["", "TOTAL histórico MCIC", data["total_historico_graduados"], "", ""])
    for c in range(1, 6):
        ws.cell(row=fila_total, column=c).font = Font(bold=True)
    _autoancho(ws, [14, 55, 22, 24, 30])
    ws.freeze_panes = "A2"


def hoja_egresados_por_anio(wb: openpyxl.Workbook, data: dict) -> None:
    ws = wb.create_sheet("Egresados - por año 2022-2026")
    anios = sorted({a for p in data["por_proyecto_por_anio_estimado"] for a in p["por_anio_estimado"]})
    ws.cell(
        row=1, column=1,
        value="Estimado a partir del año de última matrícula de cada graduado (Cóndor no registra fecha de grado exacta para 2024-2026).",
    ).font = Font(italic=True, size=10)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(anios) + 3)
    encabezados = ["Cód. Proyecto", "Énfasis / Proyecto Curricular", *[str(a) for a in anios], "Total"]
    _escribir_encabezados(ws, 2, encabezados)
    fila = 3
    totales_por_anio = {a: 0 for a in anios}
    total_general = 0
    for p in data["por_proyecto_por_anio_estimado"]:
        valores_anio = [p["por_anio_estimado"].get(str(a), 0) for a in anios]
        total_fila = sum(valores_anio)
        ws.append([p["cod_proyecto"], p["proyecto_curricular"], *valores_anio, total_fila])
        for a, v in zip(anios, valores_anio):
            totales_por_anio[a] += v
        total_general += total_fila
        fila += 1
    ws.append(["", "TOTAL", *[totales_por_anio[a] for a in anios], total_general])
    for c in range(1, len(encabezados) + 1):
        ws.cell(row=fila, column=c).font = Font(bold=True)
    _autoancho(ws, [14, 55, *([10] * len(anios)), 12])
    ws.freeze_panes = "A3"


def hoja_egresados_fecha_real(wb: openpyxl.Workbook, data: dict) -> None:
    fecha_real = data["con_fecha_real"]
    if not fecha_real.get("total_graduados"):
        return
    ws = wb.create_sheet("Egresados - fecha real")
    anios = sorted(fecha_real["por_anio"].keys())
    encabezados = ["Cód. Proyecto", *anios, "Total", "Promedio académico", "Archivo fuente"]
    _escribir_encabezados(ws, 1, encabezados)
    for p in fecha_real["por_enfasis"]:
        valores_anio = [p["por_anio"].get(a, 0) for a in anios]
        ws.append([p["cod_proyecto"], *valores_anio, p["total_graduados"], p["promedio_academico"], data["fuente"]["acta_fecha_real"]["archivo"]])
    _autoancho(ws, [14, *([10] * len(anios)), 10, 18, 55])
    ws.freeze_panes = "A2"


def main() -> None:
    estado = cargar("estado_academico_agregado.json")
    enfasis = cargar("enfasis_estudiantes.json")
    egresados = cargar("egresados_agregado.json")

    wb = openpyxl.Workbook()
    hoja_leeme(wb)
    hoja_estado_academico(wb, estado)
    hoja_estudiantes_enfasis(wb, enfasis)
    hoja_egresados_historico(wb, egresados)
    hoja_egresados_por_anio(wb, egresados)
    hoja_egresados_fecha_real(wb, egresados)

    GOLD_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_XLSX)
    print(f"Consolidado Comunidad Estudiantil (.xlsx) -> {OUT_XLSX.relative_to(ROOT)}")
    print(f"  - Hojas: {wb.sheetnames}")


if __name__ == "__main__":
    main()
