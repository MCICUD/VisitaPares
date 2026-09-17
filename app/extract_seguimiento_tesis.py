"""Bronze -> Silver

Extrae, del CONTENIDO real de los documentos del proceso de trabajo de
grado (carta de radicación, solicitud de jurados, carta de aval y
viabilidad del grupo, acta de sustentación) de los estudiantes de
Investigación (595) y Profundización (695) con última matrícula entre 2022
y 2026, los campos: título del proyecto, director, codirector, grupo de
investigación y (si ya sustentó) nota/carácter/jurados.

Los documentos fuente viven en Data/Bronze/Seguimiento_Tesis/<código>/<categoría>/
(ver Data/Bronze/Seguimiento_Tesis/MANIFEST.json para el origen exacto de
cada archivo). Esa carpeta está excluida del catálogo público de Bronze
(build_bronze_manifest.EXCLUIR_DEL_CATALOGO) porque el nombre de carpeta y
de archivo suele incluir el nombre completo del estudiante.

Política de privacidad de este script (no negociable): el código y el
nombre del estudiante se usan SOLO en memoria, para agrupar los documentos
de un mismo proceso de tesis. Nunca se escriben en Data/Silver/seguimiento_tesis.json
— ni el código, ni el nombre, ni la ruta del archivo (esa ruta contiene el
código en el nombre de la carpeta). Lo único que identifica de dónde salió
cada dato es la lista de categorías de documento encontradas
("fuente_categorias"), igual de verificable pero sin exponer identidad.

Si un campo no se puede extraer con confianza (documento escaneado sin
texto, o el patrón no calza), se deja en None — nunca se inventa.
"""
from __future__ import annotations

import csv
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib.doc_reader import extraer_texto  # noqa: E402
from lib import tesis_extractores as ext  # noqa: E402

SILVER_DIR = ROOT / "Data/Silver"
ESTADOS_DIR = ROOT / "Data/Bronze/Estados"
SEGUIMIENTO_DIR = ROOT / "Data/Bronze/Seguimiento_Tesis"

COL_COD_ESTUDIANTE = 3
COL_ULTIMA_MATRICULA = 14
PROYECTOS = ("595", "695")
ANIO_RE = re.compile(r"^(\d{4})")

CATEGORIAS = ("carta_radicacion", "solicitud_jurados", "viabilidad", "acta_sustentacion")

EXTRACTOR_POR_CATEGORIA = {
    "carta_radicacion": ext.extraer_radicacion,
    "solicitud_jurados": ext.extraer_solicitud_jurados,
    "viabilidad": ext.extraer_viabilidad,
    "acta_sustentacion": ext.extraer_acta_sustentacion,
}


def leer_csv_estados(path: Path) -> list[list[str]]:
    with path.open(encoding="utf-8-sig") as fh:
        lines = fh.readlines()
    reader = csv.reader(lines[1:])
    next(reader)
    return [row for row in reader if any(c.strip() for c in row)]


def cargar_cod_proyecto_por_estudiante() -> dict[str, str]:
    """{codigo_estudiante: cod_proyecto} para 595 y 695 (los únicos que
    tiene sentido cruzar aquí: son las dos líneas de trabajo de grado)."""
    resultado = {}
    for cod_proyecto in PROYECTOS:
        path = ESTADOS_DIR / f"Listado_de_estudiantes_por_estado_{cod_proyecto}.csv"
        if not path.exists():
            continue
        for fila in leer_csv_estados(path):
            resultado[fila[COL_COD_ESTUDIANTE].strip()] = cod_proyecto
    return resultado


def etapa_actual(categorias_presentes: set[str], acta_campos: dict | None) -> str:
    if "acta_sustentacion" in categorias_presentes:
        caracter = (acta_campos or {}).get("caracter") or "sin carácter registrado"
        return f"Sustentado ({caracter})"
    if "solicitud_jurados" in categorias_presentes:
        return "Jurados solicitados, pendiente sustentación"
    if "carta_radicacion" in categorias_presentes or "viabilidad" in categorias_presentes:
        return "Anteproyecto radicado, pendiente asignación de jurados"
    return "Sin anteproyecto radicado"


def procesar_estudiante(codigo_dir: Path, cod_proyecto: str) -> dict:
    """Lee todos los documentos de un estudiante y devuelve su entrada de
    proceso de tesis (sin código ni nombre)."""
    campos_por_categoria: dict[str, dict] = {}
    categorias_presentes: set[str] = set()
    leidos, no_leidos = 0, 0

    for categoria in CATEGORIAS:
        cat_dir = codigo_dir / categoria
        if not cat_dir.is_dir():
            continue
        archivos = sorted(p for p in cat_dir.iterdir() if p.is_file())
        if not archivos:
            continue
        categorias_presentes.add(categoria)
        texto = None
        for archivo in archivos:
            r = extraer_texto(archivo)
            if r["extraido"]:
                texto = r["texto"]
                leidos += 1
                break
            no_leidos += 1
        if texto:
            campos_por_categoria[categoria] = EXTRACTOR_POR_CATEGORIA[categoria](texto)

    radic = campos_por_categoria.get("carta_radicacion") or {}
    viab = campos_por_categoria.get("viabilidad") or {}
    jur = campos_por_categoria.get("solicitud_jurados") or {}
    acta = campos_por_categoria.get("acta_sustentacion") or {}

    etapa = etapa_actual(categorias_presentes, acta)

    entrada = {
        "cod_proyecto": cod_proyecto,
        "etapa_actual": etapa,
        "titulo_proyecto": radic.get("titulo_proyecto") or viab.get("titulo_proyecto") or jur.get("titulo_proyecto") or acta.get("titulo_proyecto"),
        "director": radic.get("director") or viab.get("director") or jur.get("director") or acta.get("director"),
        "codirector": radic.get("codirector") or viab.get("codirector") or jur.get("codirector") or acta.get("codirector"),
        "grupo_investigacion": radic.get("grupo_investigacion") or viab.get("grupo_investigacion"),
        "numero_acta_sustentacion": acta.get("numero_acta"),
        "nota_sustentacion": acta.get("nota"),
        "caracter_sustentacion": acta.get("caracter"),
        "jurados_sustentacion": acta.get("jurados") or None,
        "fuente_categorias": sorted(categorias_presentes),
    }
    return entrada, leidos, no_leidos


def main() -> None:
    if not SEGUIMIENTO_DIR.exists():
        raise FileNotFoundError(f"No existe {SEGUIMIENTO_DIR}")

    cod_proyecto_por_estudiante = cargar_cod_proyecto_por_estudiante()

    procesos = []
    total_leidos, total_no_leidos = 0, 0
    estudiantes_sin_proyecto_conocido = 0

    for codigo_dir in sorted(p for p in SEGUIMIENTO_DIR.iterdir() if p.is_dir()):
        codigo = codigo_dir.name
        cod_proyecto = cod_proyecto_por_estudiante.get(codigo)
        if cod_proyecto is None:
            estudiantes_sin_proyecto_conocido += 1
            continue
        entrada, leidos, no_leidos = procesar_estudiante(codigo_dir, cod_proyecto)
        total_leidos += leidos
        total_no_leidos += no_leidos
        procesos.append(entrada)

    data = {
        "procesos": procesos,
        "resumen": {
            "total_procesos": len(procesos),
            "documentos_leidos": total_leidos,
            "documentos_no_leidos": total_no_leidos,
            "estudiantes_sin_proyecto_conocido": estudiantes_sin_proyecto_conocido,
        },
        "metodologia": (
            "Cada entrada corresponde a un estudiante de Investigación (595) o Profundización (695) con "
            "última matrícula entre 2022 y 2026 que tiene al menos un documento de su proceso de trabajo de "
            "grado localizado en Data/Bronze/Seguimiento_Tesis/. Los campos se extraen leyendo el CONTENIDO "
            "real de la carta de radicación, la solicitud de jurados, la carta de aval y viabilidad del grupo, "
            "y el acta de sustentación (regex sobre texto, sin IA) — ver app/lib/tesis_extractores.py. Un campo "
            "vacío significa que el documento no se pudo leer (escaneado sin texto) o que su formato no calzó "
            "con el patrón esperado; nunca se completa con un valor inventado. No se publica el código ni el "
            "nombre del estudiante en este archivo — solo el proceso de tesis en sí."
        ),
    }

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "seguimiento_tesis.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Seguimiento de tesis: {len(procesos)} procesos -> {out_path.relative_to(ROOT)}")
    print(f"  - documentos leídos: {total_leidos} | no leídos (escaneado/sin texto): {total_no_leidos}")
    if estudiantes_sin_proyecto_conocido:
        print(f"  - carpetas de estudiante sin cod_proyecto 595/695 conocido (omitidas): {estudiantes_sin_proyecto_conocido}")


if __name__ == "__main__":
    main()
