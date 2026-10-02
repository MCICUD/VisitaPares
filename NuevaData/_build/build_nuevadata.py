"""Construye NuevaData/ (propuesta de evidencias para aprobación).

    NuevaData/
      Original/<Modalidad>/...        evidencia original por FACTOR (partió de una copia de
                                      Data/Bronze/<MCIC.*>/.../Plan de Mejoramiento y se amplió
                                      con lo que cargó la coordinación). ES LA FUENTE: el script
                                      solo la lee, nunca la borra ni la modifica.
      Original/_transcripciones/      transcripción manual de planillas (datos personales)
      Presentacion/<Modalidad>/...    evidencia analizada, ordenada por FACTOR / actividad,
                                      sin datos sensibles y separada por modalidad

Uso:  .venv/bin/python NuevaData/_build/build_nuevadata.py

El script es idempotente: borra y regenera solo Presentacion/. Lee Original/ y,
además, Data/Bronze, Data/Silver, SolicitudesPares/ y el consolidado de trabajos
de grado de la raíz; no modifica ninguno de ellos.
La transcripción manual de las planillas escaneadas (nombres y códigos) vive
en NuevaData/_build/_privado/transcripciones.json (fuera de git).
"""
from __future__ import annotations

import csv
import hashlib
import json
import re
import shutil
import sys
import unicodedata
import warnings
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlsx_style as st  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
BRONZE = ROOT / "Data/Bronze"
SILVER = ROOT / "Data/Silver"
OUT = ROOT / "NuevaData"
PRIVADO = Path(__file__).resolve().parent / "_privado/transcripciones.json"

MODALIDADES = {
    "Investigacion": {
        "clave": "investigacion",
        "etiqueta": "Investigación",
        "programa": "MCIC - INVESTIGACIÓN",
        "proyecto": "595",
        "snies": "17528",
        "plan": OUT / "Original/Investigacion",
        "web": "https://facingenieria.udistrital.edu.co/mcic-investigacion/",
        "captura_web": "Actualizacion pagiona web mcic -  investigacion.png",
    },
    "Profundizacion": {
        "clave": "profundizacion",
        "etiqueta": "Profundización",
        "programa": "MCIC - PROFUNDIZACIÓN",
        "proyecto": "695",
        "snies": "116070",
        "plan": OUT / "Original/Profundizacion",
        "web": "https://facingenieria.udistrital.edu.co/mcic-profundizacion/",
        "captura_web": "Actualizacion pagiona web mcic -  profundizacion.png",
    },
}

ENFASIS_POR_PROYECTO = {"195": "Teleinformática", "295": "Sistemas de Información", "395": "Geomática",
                        "495": "Ingeniería de Software", "95": "Teleinformática (plan antiguo)"}

# Solicitudes de los pares (visita sep 2026) y responsables, tomados de CorreccionEvidencias.ods
SOLICITUDES_PARES = {
    1: ("PEP listo. Documento según la conversación del lunes, con los ajustes y socializado en el Consejo de Carrera.", "Sebastián Vanegas"),
    2: ("Asistencias a las inducciones: documento con los datos de asistencia; actualizaciones de la página web (dejar el enlace); publicidad de las maestrías por correo; Open Day con lista de asistencia; socialización de grupos de investigación. Se actualiza cada semestre por proceso de admisiones.", "Juan y Karol"),
    3: ("Solicitar a los profesores evidencia de las capacitaciones del periodo 2025-2026 adicionales a la información del cuadro.", "Karol (correo de solicitud)"),
    4: ("Solicitar a los egresados 2022-2026 el sector en el que laboran actualmente (formulario: nombre y sector).", "Karol y Juan (tarea enviada a la OATI)"),
    5: ("Syllabus actualizados.", "—"),
    6: ("Número de estudiantes informados (activos en la ventana de tiempo); tabla de estudiantes que solicitaron prórroga; datos de PAGOT y graduados.", "Karol y Juan"),
    7: ("Correo solicitando a los profesores información sobre su participación en convenios, con soporte.", "—"),
    8: ("Cargar el Excel consolidado de trabajos de grado.", "—"),
    9: ("Personas que asisten a la inducción e información de Bienestar.", "—"),
    10: ("Biblioteca y Planes TIC.", "—"),
    11: ("Cargar las 2 autoevaluaciones y las encuestas.", "—"),
    12: ("Laboratorios.", "—"),
}

RECOMENDACIONES_CNA_2022 = [
    ("Impacto social y académico", "Adelantar estudios que permitan establecer el impacto social y académico de la misión, el Proyecto Educativo Institucional y los programas académicos, así como de los productos investigativos, de proyección social, artística y cultural, sobre el desarrollo de la región y del país."),
    ("Estudiantes", "Desplegar acciones que estimulen la movilidad entrante y saliente de estudiantes, en los ámbitos nacional e internacional."),
    ("Interacción con el entorno", "Hacer una mayor difusión entre la comunidad estudiantil de los numerosos convenios activos y específicos disponibles y de los apoyos brindados por la Institución. Asimismo, profundizar las acciones que faciliten un mayor intercambio con profesores, especialmente de otras regiones."),
    ("Interacción con el entorno", "Desarrollar proyectos cooperativos de investigación y académicos con instituciones internacionales de reconocido prestigio."),
    ("Grupos de investigación", "Seguir avanzando en el despliegue de acciones que contribuyan al fortalecimiento de los grupos de investigación que apoyan al Programa."),
    ("Investigación e innovación", "Mejorar la clasificación de los profesores del Programa como investigadores y la categorización de los grupos en el SNCT&I."),
    ("Investigación e innovación", "Diseñar estudios que evalúen el impacto de la investigación sobre el desarrollo productivo y tecnocientífico en la región y el país."),
    ("Interacción con el entorno", "Continuar fortaleciendo los convenios activos existentes y, especialmente, precisar su especificidad u oportunidades en las áreas del conocimiento relacionadas con el Programa y registrar su grado de uso."),
    ("Investigación e innovación", "Mantener los esfuerzos dirigidos a la estructuración y entrada en funcionamiento del Instituto de Investigación e Innovación en Ingeniería."),
    ("Infraestructura", "Continuar el proceso dirigido a la construcción de nueva infraestructura física en Ingeniería que permita atender las necesidades y el crecimiento del Programa, incluyendo espacios para profesores."),
    ("Permanencia y graduación", "Persistir en las medidas desarrolladas dentro del Plan de Mejoramiento orientadas a actualizar la base de datos de graduados y fortalecer su seguimiento."),
    ("Egresados", "Diseñar estudios que establezcan el impacto de los graduados sobre el desarrollo productivo, científico y tecnológico. Asimismo, enfatizar acciones que promuevan la graduación, realizar evaluaciones que identifiquen los factores que la afectan y revisar las estrategias establecidas para mejorar este aspecto."),
    ("Bienestar universitario", "Implementar acciones de bienestar universitario orientadas a atender las necesidades particulares y las especificidades de los estudiantes del Programa, considerando la baja participación evidenciada en las ofertas de bienestar."),
    ("Bienestar universitario", "Desarrollar un registro sistemático sobre la participación y utilización de las ofertas de bienestar universitario por parte de los estudiantes de la Maestría."),
    ("Aseguramiento de la calidad", "Fortalecer el sistema interno de aseguramiento de la calidad, de manera que permita identificar los logros, resultados e impactos de la implementación de las recomendaciones derivadas de los procesos de autoevaluación, evaluación de pares y recomendaciones del CNA, con el fin de consolidar el mejoramiento continuo del Programa."),
]


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def norm_tokens(s: str | None) -> set[str]:
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().upper()
    return set(re.sub(r"[^A-Z ]", " ", s).split())


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def copiar(src: Path, dst_dir: Path, nombre: str | None = None) -> Path:
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / (nombre or src.name)
    shutil.copy2(src, dst)
    return dst


def copiar_arbol(src: Path, dst: Path) -> None:
    shutil.copytree(src, dst, dirs_exist_ok=True)


def factor_dirs(plan: Path) -> dict[int, Path]:
    out = {}
    for d in plan.iterdir():
        m = re.match(r"FACTOR (\d+)\.", d.name)
        if d.is_dir() and m:
            out[int(m.group(1))] = d
    return out


def fecha_es(iso: str | None) -> str:
    if not iso:
        return "—"
    y, m, d = iso.split("-")
    return f"{d}/{m}/{y}"


def frase(s: str | None) -> str | None:
    """'CURSÓ Y APROBÓ ESPECIALIZADO I Y II (2023-1)' -> 'Cursó y aprobó especializado I y II (2023-1)'."""
    if not s:
        return s
    s = " ".join(str(s).split()).lower()
    s = re.sub(r"\b(i{1,3})\b", lambda m: m.group(1).upper(), s)
    return s[:1].upper() + s[1:]


def categoria_escalafon(s: str | None) -> str | None:
    if not s:
        return None
    for cat in ("TITULAR", "ASOCIADO", "ASISTENTE", "AUXILIAR"):
        if cat in s.upper():
            return cat.capitalize()
    return frase(s)


# ---------------------------------------------------------------------------
# Fuentes de clasificación por modalidad (Cóndor + bases de datos MCIC)
# ---------------------------------------------------------------------------
def cargar_roster() -> list[dict]:
    filas = []
    for f in sorted((BRONZE / "Estados").glob("Listado_de_estudiantes_por_estado_*.csv")):
        lineas = f.read_text(encoding="utf-8", errors="replace").splitlines()[1:]
        for r in csv.DictReader(lineas):
            filas.append({
                "codigo": (r.get("Cod. Estudiante") or "").strip(),
                "nombre": r.get("Nombre Estudiante") or "",
                "proyecto": (r.get("Cod. Proyecto") or "").strip(),
                "estado": r.get("DescripciÓn") or r.get("Estado") or "",
                "ultima": r.get("Ultima Matricula") or "",
                "documento": (r.get("Documento") or "").strip(),
            })
    return filas


def cargar_bases_mcic() -> dict[str, dict]:
    """código -> {'modalidad', 'tipo', 'estado', 'enfasis', 'ingreso', 'fuente'} (prioridad: bases 2026, luego V2)."""
    base = BRONZE / "Maestria CIC/2026/BASES  DE DATOS ESTUDIANTES"
    out: dict[str, dict] = {}
    fuentes = [
        ("MCIC - Base de datos INVESTIGACION.xlsx", ["N-A Investigación"], "Investigación"),
        ("MCIC - Base de datos Profundizacion.xlsx", ["N-A Profundizacion", "Pasantías"], "Profundización"),
        ("MCIC - Base de datos V2.xlsx", ["N-A Investigación", "N-A Profundizacion", "Pasantías"], None),
    ]
    for archivo, hojas, mod_fija in fuentes:
        wb = openpyxl.load_workbook(base / archivo, read_only=True, data_only=True)
        for h in hojas:
            filas = list(wb[h].iter_rows(values_only=True))
            enc = [str(c).strip() if c else "" for c in filas[0]]
            idx = {k: enc.index(k) for k in enc if k}
            for r in filas[1:]:
                if not r or not r[0]:
                    continue
                cod = str(r[0]).split(".")[0].strip()
                if cod in out:
                    continue
                mod = mod_fija
                if mod is None:
                    v = str(r[idx["MODALIDAD"]] or "") if "MODALIDAD" in idx else ""
                    if "rofund" in v or h == "Pasantías":
                        mod = "Profundización"
                    elif "nvestig" in v:
                        mod = "Investigación"
                    else:
                        mod = "Investigación" if h == "N-A Investigación" else "Profundización"
                out[cod] = {
                    "modalidad": mod,
                    "tipo": r[idx["TIPO DE ESTUDIANTE"]] if "TIPO DE ESTUDIANTE" in idx else None,
                    "estado": r[idx["ESTADO"]] if "ESTADO" in idx else None,
                    "enfasis": r[idx["ÉNFASIS"]] if "ÉNFASIS" in idx else (r[idx["ENFASIS"]] if "ENFASIS" in idx else None),
                    "ingreso": r[idx["INGRESO"]] if "INGRESO" in idx else None,
                    "fuente": f"{archivo} / {h}",
                }
    return out


ROSTER = cargar_roster()
ROSTER_POR_CODIGO = {r["codigo"]: r for r in ROSTER}
ROSTER_POR_DOC: dict[str, list[dict]] = defaultdict(list)
for _r in ROSTER:
    if _r["documento"]:
        ROSTER_POR_DOC[_r["documento"]].append(_r)
BASES = cargar_bases_mcic()


def buscar_en_roster(nombre: str) -> dict | None:
    t = norm_tokens(nombre)
    cands = [r for r in ROSTER if t and t <= norm_tokens(r["nombre"])]
    if not cands:
        cands = [r for r in ROSTER if len(t & norm_tokens(r["nombre"])) >= max(3, len(t) - 1)]
    if not cands:
        return None
    cands.sort(key=lambda r: (r["proyecto"] not in ("595", "695"), r["estado"] != "Matriculado"))
    return cands[0]


def modalidad_de(codigo: str | None, nombre: str | None = None, declarada: str | None = None) -> tuple[str | None, str]:
    """Devuelve (modalidad, criterio)."""
    if declarada:
        return declarada, "Declarada en la solicitud"
    if codigo == "OTRO_PROGRAMA":
        return None, "Otro programa"
    if not codigo and nombre:
        r = buscar_en_roster(nombre)
        codigo = r["codigo"] if r else None
    if not codigo:
        return None, "Sin código"
    proy = codigo[5:8] if len(codigo) == 11 else ""
    if proy == "595":
        return "Investigación", "Código de proyecto 595"
    if proy == "695":
        return "Profundización", "Código de proyecto 695"
    r = ROSTER_POR_CODIGO.get(codigo)
    if r and r["proyecto"] in ("595", "695"):
        return ("Investigación" if r["proyecto"] == "595" else "Profundización"), "Cóndor (proyecto actual)"
    if codigo in BASES:
        return BASES[codigo]["modalidad"], "Base de datos MCIC"
    return None, "Sin registro de modalidad"


def enfasis_normalizado(texto: str | None) -> str | None:
    t = norm_tokens(texto)
    for clave, etiqueta in (("SOFTWARE", "Ingeniería de Software"), ("TELEINFORMATICA", "Teleinformática"),
                            ("GEOMATICA", "Geomática"), ("INTELIGENCIA", "Inteligencia Artificial")):
        if clave in t:
            return etiqueta
    return frase(texto) if texto else None


def enfasis_de(codigo: str | None) -> str | None:
    if not codigo or len(codigo) != 11:
        return None
    proy = codigo[5:8]
    if proy in ENFASIS_POR_PROYECTO:
        return ENFASIS_POR_PROYECTO[proy]
    b = BASES.get(codigo)
    return enfasis_normalizado(b["enfasis"]) if b and b.get("enfasis") else None


def contar_por_modalidad(personas: list[tuple[str, str | None]]) -> Counter:
    c = Counter()
    for nombre, codigo in personas:
        mod, _ = modalidad_de(codigo, nombre)
        if mod is None and codigo:
            mod = modalidad_con_plan_anterior(codigo)
        c[mod or "Otro programa"] += 1
    return c


# ---------------------------------------------------------------------------
# Plan de mejoramiento: metas (anterior 2024-2026 y vigente 2026-2027)
# ---------------------------------------------------------------------------
def metas(clave: str) -> dict[int, dict]:
    ant = json.loads((SILVER / f"plan_mejoramiento_2025_{clave}.json").read_text(encoding="utf-8"))["factores"]
    vig = json.loads((SILVER / f"plan_mejoramiento_{clave}.json").read_text(encoding="utf-8"))["factores"]
    out = {}
    for i, (a, v) in enumerate(zip(ant, vig), start=1):
        out[i] = {"factor": v["factor"], "meta_anterior": a.get("meta"), "linea_base_anterior": a.get("linea_base"),
                  "meta_vigente": " ".join((v.get("meta") or "").split()), "responsable": v.get("responsable")}
    return out


def encabezado_factor(ws, fila: int, n: int, m: dict, ancho: int = 8) -> int:
    fila = st.seccion(ws, fila, "Metas del plan de mejoramiento", ancho)
    filas = [["Plan anterior (2024-2026)", m[n]["meta_anterior"]],
             ["Plan vigente (2026-2027)", m[n]["meta_vigente"]],
             ["Solicitud de los pares (visita sep. 2026)", SOLICITUDES_PARES[n][0]]]
    for etiqueta, texto in filas:
        a = ws.cell(row=fila, column=1, value=etiqueta)
        a.font, a.fill, a.alignment, a.border = st.FONT_HEADER, st.FILL_HEADER, st.WRAP_TOP, st.BORDE
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=2)
        b = ws.cell(row=fila, column=3, value=texto)
        b.font, b.alignment, b.border = st.FONT_CELDA, st.WRAP_TOP, st.BORDE
        ws.merge_cells(start_row=fila, start_column=3, end_row=fila, end_column=ancho)
        ws.row_dimensions[fila].height = max(18, 14 * (1 + len(texto or "") // 110))
        fila += 1
    return fila + 1


# ---------------------------------------------------------------------------
# FACTOR 1 — PEP y jornadas con docentes (sin cambios de contenido)
# ---------------------------------------------------------------------------
def factor1(mod: dict, dest: Path) -> list[str]:
    fdir = factor_dirs(mod["plan"])[1]
    entregado = []
    # Orientaciones institucionales (Vicerrectoría Académica, 2025) para evaluar los propósitos de formación y aprendizaje:
    # se entrega con el PEP vigente mientras llega el PEP definitivo.
    pfa = "2026IE4045-GC-RemisionEvaluacionPFA-ANEXO1.pdf"
    copiar(fdir / pfa, dest / "a. PEP")
    entregado.append(f"a. PEP/{pfa}")
    for act in ("a. PEP", "b. Jornadas de trabajo con docentes"):
        for f in sorted((fdir / act).iterdir()):
            copiar(f, dest / act)
            entregado.append(f"{act}/{f.name}")
    return entregado


# ---------------------------------------------------------------------------
# FACTOR 2 — Divulgación, páginas web e inducciones
# ---------------------------------------------------------------------------
def factor2(nombre_mod: str, mod: dict, dest: Path, tr: dict, m: dict) -> list[str]:
    etiqueta = mod["etiqueta"]
    otra = "Profundización" if etiqueta == "Investigación" else "Investigación"
    src = factor_dirs(mod["plan"])[2] / "a. Divulgación de propuestas académicas"
    act = dest / "a. Divulgación de propuestas académicas"
    sop = act / "Anexos"
    imagenes = [mod["captura_web"], "Actualizacion pagiona web mcic.png", "Inducción 2025-1.jpg",
                "Divulgación grupos de investigación 2025-1.jpg", "Divulgación grupos de investigación 2025-3.jpg",
                "Encuentro estudiantes evaluacion docente.jpg", "Realización del OPEN DAY 3.0.jpeg", "Inducción 2026-3.jpeg"]
    for f in imagenes + ["Open Day.pdf"]:
        copiar(src / f, sop)
    orig = f"Original/{nombre_mod}/{src.relative_to(mod['plan'])}"
    ev = tr["eventos"]

    def cnt(personas):
        c = contar_por_modalidad(personas)
        return c.get(etiqueta, 0), c.get(otra, 0), sum(c.values())

    ind25 = cnt([(n, None) for n in ev["induccion_2025_1_correo"]["destinatarios"]])
    div25 = cnt([(n, None) for n in ev["divulgacion_grupos_2025_3_correo"]["destinatarios"]])
    ind26 = cnt([tuple(x) for x in ev["induccion_2026_1_lista"]["asistentes"]])
    gru26 = cnt([tuple(x) for x in ev["presentacion_grupos_2026_1_lista"]["asistentes"]])
    soc26 = cnt([tuple(x) for x in ev["socializacion_grupos_2026_3_lista"]["asistentes_mcic"]])

    eventos = [
        # fecha, periodo, actividad, tipo, lugar, convoca, alcance, n_mod, n_otra, total, soporte_pres, soporte_orig, nota
        ("2025-01-31", "2025-1", "Comunicación de inicio de clases y convocatoria a la inducción (7 de febrero de 2025, 6:15 p. m.) con el calendario del semestre", "Comunicación a estudiantes", "Correo institucional", "Coordinación MCIC (cuenta de Investigación)", "Investigación", ind25[0], ind25[1], ind25[2] + ev["induccion_2025_1_correo"]["destinatarios_externos"], None, "Induccion 2025-1.pdf (reservado)", "Destinatarios en copia oculta."),
        ("2025-02-07", "2025-1", "Inducción de estudiantes nuevos 2025-1 (registro fotográfico)", "Inducción", "Auditorio Sabio Caldas", "Coordinación MCIC", "Ambas", None, None, None, "Anexos/Inducción 2025-1.jpg", None, "Sin planilla de asistencia en la evidencia."),
        (None, "2025-1", "Divulgación de grupos de investigación 2025-1 (registro fotográfico)", "Divulgación de grupos de investigación", "Auditorio, Facultad de Ingeniería", "Facultad de Ingeniería", "Ambas", None, None, None, "Anexos/Divulgación grupos de investigación 2025-1.jpg", None, "Sin planilla de asistencia en la evidencia."),
        ("2025-08-22", "2025-3", "Invitación a la jornada de divulgación de grupos de investigación (asistencia obligatoria para quienes cursan Seminario de Investigación)", "Comunicación a estudiantes", "Correo institucional", "Coordinación MCIC (cuenta de Investigación)", "Investigación", div25[0], div25[1], div25[2] + ev["divulgacion_grupos_2025_3_correo"]["destinatarios_sin_nombre"], None, "Divulgacio grupos 2025-3.pdf (reservado)", "Incluye el afiche del evento."),
        ("2025-08-30", "2025-3", "Jornada de divulgación de grupos de investigación de la Facultad (8:00 a. m. a 12:00 m.). Invitan: Maestría en Ingeniería Industrial, MCIC Investigación y Profundización, Maestría en Gerencia Integral de Proyectos y Maestría en Telecomunicaciones Móviles", "Divulgación de grupos de investigación", "Auditorio Sabio Caldas", "Facultad de Ingeniería", "Ambas", None, None, None, "Anexos/Divulgación grupos de investigación 2025-3.jpg", None, "Registro fotográfico."),
        (None, "—", "Encuentro con estudiantes – evaluación docente (registro fotográfico)", "Encuentro con estudiantes", "Sala de cómputo, Facultad de Ingeniería", "Coordinación MCIC", "Ambas", None, None, None, "Anexos/Encuentro estudiantes evaluacion docente.jpg", None, "El soporte no trae fecha."),
        ("2026-02-02", "2026-1", "Inducción 2026-1 (control de asistencia GD-PR-008-FR-026)", "Inducción", "Facultad de Ingeniería", "Coordinación MCIC", "Ambas", ind26[0], ind26[1], ind26[2], None, "LISTA ASISTENCIA INDUCCIONES 2026-1.pdf (reservado)", "Planilla con firmas."),
        ("2026-02-21", "2026-1", "Presentación de grupos de investigación – Facultad de Ingeniería 2026-1 (8:00 a. m. a 12:00 m.)", "Divulgación de grupos de investigación", "Auditorio Sabio Caldas", "Facultad de Ingeniería – Posgrados", "Ambas", gru26[0], gru26[1], gru26[2], None, "Presentación de Grupos de Investigación MCIC2026-1.pdf (reservado)", f"Además asistieron {tr['eventos']['presentacion_grupos_2026_1_lista']['docentes_y_grupos']} docentes y representantes de grupos de investigación."),
        ("2026-04-30", "2026-1", "Difusión del Open Day 3.0 de posgrados UD (evento del 15 de mayo de 2026, 5:00 p. m.)", "Divulgación de la oferta académica", "Correo institucional", "Decanatura Facultad de Ingeniería – Eventos", "Ambas", None, None, None, "Anexos/Open Day.pdf", None, "Material publicitario para canales de cada programa de posgrado; incluye enlace de inscripción."),
        ("2026-05-15", "2026-1", "Realización del Open Day 3.0 (registro fotográfico)", "Divulgación de la oferta académica", "Auditorio Sabio Caldas y Muro de Escalar", "Facultad de Ingeniería", "Ambas", None, None, None, "Anexos/Realización del OPEN DAY 3.0.jpeg", None, "Sin planilla de asistencia en la evidencia."),
        (None, "2026-3", "Inducción 2026-3 – presentación de grupos de investigación (registro fotográfico, grupo LIDER)", "Inducción", "Auditorio, Facultad de Ingeniería", "Coordinación MCIC", "Ambas", None, None, None, "Anexos/Inducción 2026-3.jpeg", None, "Registro fotográfico."),
        ("2026-08-22", "2026-3", "Socialización de grupos de investigación (8:00 a. m. a 12:00 m.)", "Divulgación de grupos de investigación", "Auditorio Sabio Caldas", "Facultad de Ingeniería", "Ambas", soc26[0], soc26[1], soc26[2], None, "Asistencia 22-08-22 2026-3.pdf (reservado)", f"Planilla compartida con otros programas ({ev['socializacion_grupos_2026_3_lista']['total_registros_planilla']} registros en total)."),
    ]
    eventos = [e for e in eventos if e[6] in ("Ambas", etiqueta)]

    wb = st.nuevo_libro()
    # --- Resumen
    ws = st.hoja(wb, "Resumen", f"Factor 2 · Estudiantes — {mod['programa']}",
                 "Divulgación de la oferta académica, páginas web, inducciones y socialización de grupos de investigación. "
                 "Consolidado a partir de capturas, fotografías, correos y planillas de asistencia; los registros con datos personales se resguardan y no se publican.")
    fila = encabezado_factor(ws, 4, 2, m)
    con_lista = [e for e in eventos if e[7] is not None]
    fila = st.kpis(ws, fila, [
        ("Página web de la modalidad", "Actualizada"),
        ("Actividades documentadas", len(eventos)),
        (f"Registros de estudiantes de {etiqueta}", sum(e[7] for e in con_lista)),
    ])
    fila = st.nota(ws, fila + 1, "Los conteos por modalidad se obtienen cruzando cada nombre/código de las planillas con Cóndor (proyectos 595 = Investigación, 695 = Profundización) y con las bases de datos MCIC. "
                   "Un mismo estudiante puede aparecer en varias actividades; los registros no se suman como personas únicas.")
    for col, w in zip("ABCDEFGH", (16, 16, 16, 16, 16, 16, 16, 16)):
        ws.column_dimensions[col].width = w

    # --- Páginas web
    ws = st.hoja(wb, "Páginas web", "Páginas web de la Maestría", "Enlaces vigentes y capturas de la actualización.", 5)
    webs = [[etiqueta, mod["web"], mod["web"], "Sitio propio de la modalidad", "Captura de la actualización", f"Anexos/{mod['captura_web']}"],
            ["Común (ambas modalidades)", "Sitio MCIC con menú para aspirantes de Investigación y Profundización", None, "Banner de inscripciones abiertas y acreditación de alta calidad (Res. 024858 de 2022)", "Captura del sitio común", "Anexos/Actualizacion pagiona web mcic.png"]]
    st.tabla(ws, 4, ["Modalidad", "Enlace", "_url", "Contenido", "Soporte", "_sop"], webs, [26, 60, 50, 40], links={1: 2, 4: 5})

    # --- Actividades
    ws = st.hoja(wb, "Actividades", "Actividades de divulgación, inducción y socialización",
                 f"Columna «Estudiantes de {etiqueta}»: registros de la planilla o destinatarios del correo que pertenecen a esta modalidad.", 12)
    filas = []
    for e in eventos:
        filas.append([fecha_es(e[0]) if e[0] else "—", e[1], e[2], e[3], e[4], e[5],
                      e[6], e[7] if e[7] is not None else "—", e[8] if e[8] is not None else "—",
                      e[9] if e[9] is not None else "—",
                      "Ver soporte" if e[10] else "—", e[10],
                      e[11] or "—", e[12]])
    st.tabla(ws, 4, ["Fecha", "Periodo", "Actividad", "Tipo", "Lugar / medio", "Convoca", "Alcance",
                     f"Estudiantes de {etiqueta}", f"Estudiantes de {otra}", "Total registros", "Soporte (Presentación)", "_s",
                     "Documento original (reservado: contiene datos personales)", "Observación"],
             filas, [11, 9, 48, 20, 20, 22, 12, 13, 13, 11, 14, 42, 34], links={10: 11})

    # --- Galería
    ws = st.hoja(wb, "Galería", "Registro fotográfico y capturas", "Miniaturas de los anexos copiados en la carpeta «Anexos».", 6)
    fila = 4
    for i, f in enumerate([x for x in imagenes]):
        col = "A" if i % 2 == 0 else "E"
        if i % 2 == 0 and i:
            fila += 18
        ws[f"{col}{fila}"] = f.rsplit(".", 1)[0]
        ws[f"{col}{fila}"].font = st.FONT_SECCION
        ws[f"{col}{fila}"].hyperlink = f"Anexos/{f}"
        st.miniatura(ws, f"{col}{fila + 1}", sop / f, 300)
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 12

    nombre = f"F2_Divulgacion_y_Estudiantes_{nombre_mod}.xlsx"
    wb.save(act / nombre)
    return [f"a. Divulgación de propuestas académicas/{nombre}"] + [f"a. Divulgación de propuestas académicas/Anexos/{x}" for x in imagenes + ["Open Day.pdf"]]


# ---------------------------------------------------------------------------
# FACTOR 3 — Participación docente en capacitación (planta compartida)
# ---------------------------------------------------------------------------
# Certificados recientes de los profesores (transcritos de los PDF de Original/<Mod>/FACTOR 3/…/Certificados).
# (profesor, capacitación, entidad, fecha, horas, año)
CERTIFICADOS_F3 = [
    ("Ortiz Davila Alvaro Enrique", "Introduction to GeoWebApp Building with open-source tools (curso en línea)", "Universidad de Twente (ITC) y Geoversity", "13/05/2024", 40, 2024),
    ("Ortiz Davila Alvaro Enrique", "Curso Internacional de Competencias Docentes y Diseño de Materiales Didácticos para la Educación en Línea (título propio, calificación 9,60)", "Universidad Internacional de La Rioja (UNIR)", "04/05/2026 al 05/07/2026", 125, 2026),
    ("Gomez Vargas Ernesto", "Taller de formación EUR-ACE «Aseguramiento de calidad y acreditación internacional EUR-ACE de programas educativos en ingeniería»", "ENAEE, ENTEA y LACCEI (Girardot, Colombia)", "08/11/2022 al 10/11/2022", None, 2022),
    ("Plazas Nossa Leonardo", "Data Science and Machine Learning: Making Data-Driven Decisions", "MIT Schwarzman College of Computing y MIT IDSS", "agosto de 2023", None, 2023),
    ("Vanegas Ayala Sebastian Camilo", "Taller «Instalación local de modelos de IA»", "Vicerrectoría Académica y Comité PlanEsTIC-UD", "16/07/2025", 1, 2025),
    ("Vanegas Ayala Sebastian Camilo", "Taller «Automatización de tareas de ofimática con IA a través del uso de protocolo MCP»", "Vicerrectoría Académica y Comité PlanEsTIC-UD", "15/07/2025", 1, 2025),
    ("Vanegas Ayala Sebastian Camilo", "Taller «Da vida a tus ideas con presentaciones creativas e impactantes en Gamma AI»", "Vicerrectoría Académica y Comité PlanEsTIC-UD", "20/08/2025", 1, 2025),
]


def factor3(nombre_mod: str, mod: dict, dest: Path, tr: dict, m: dict) -> list[str]:
    src = factor_dirs(mod["plan"])[3] / "a. Informes de participación de los docentes/Participcion docente.xlsx"
    wb_src = openpyxl.load_workbook(src, data_only=True)
    fuera = {frozenset(norm_tokens(n)) for n in tr["docentes_excluidos_f3"]}

    def incluido(nombre):
        return frozenset(norm_tokens(nombre)) not in fuera

    cm = openpyxl.load_workbook(BRONZE / "ACREDITACIÓN DE ALTA CALIDAD/SNIES17528-MCIC-Investigación/SNIES17528-Cuadros maestros.xlsx",
                                read_only=True, data_only=True)["Profesores Listado_Detallad "]
    cm_info = {}
    for r in cm.iter_rows(min_row=9, values_only=True):
        if r[3]:
            cm_info[frozenset(norm_tokens(r[3]))] = {"grupo": r[10], "cat": r[11], "formacion": r[4]}

    docentes = []
    for r in wb_src["info 1"].iter_rows(min_row=5, values_only=True):
        nombre = r[2]
        if not nombre or not incluido(nombre):
            continue
        info = cm_info.get(frozenset(norm_tokens(nombre)), {})
        cat_inv = info.get("cat")
        cat_inv = re.sub(r"^\d+\.\s*", "", str(cat_inv)) if cat_inv and str(cat_inv).strip() not in ("N/A", "None") else None
        grupo = info.get("grupo")
        grupo = frase(grupo) if grupo and grupo.upper().startswith("SIN ") else grupo
        docentes.append([None, " ".join(w.capitalize() for w in nombre.split()), categoria_escalafon(r[3]),
                         frase(r[4]), r[5], frase(r[6]), r[7],
                         grupo, frase(cat_inv) if cat_inv and cat_inv.isupper() else cat_inv, None])
    for prof, titulo, entidad, fecha, horas, _anio in CERTIFICADOS_F3:
        clave = frozenset(norm_tokens(prof))
        fila_d = next((d for d in docentes if frozenset(norm_tokens(d[1])) == clave), None)
        if fila_d is None:  # profesor que no figura en el listado de Decanatura
            info = cm_info.get(clave, {})
            cat_inv = re.sub(r"^\d+\.\s*", "", str(info.get("cat"))) if info.get("cat") and str(info["cat"]).strip() not in ("N/A", "None") else None
            grupo = info.get("grupo")
            grupo = frase(grupo) if grupo and grupo.upper().startswith("SIN ") else grupo
            fila_d = [None, " ".join(w.capitalize() for w in prof.split()), None, None, None, None, None, grupo, frase(cat_inv) if cat_inv and cat_inv.isupper() else cat_inv, None]
            docentes.append(fila_d)
        linea = f"{titulo} — {entidad} ({fecha}{f', {horas} h' if horas else ''})"
        fila_d[9] = f"{fila_d[9]}\n{linea}" if fila_d[9] else linea
    docentes.sort(key=lambda d: d[1])
    cert_dir = src.parent / "Certificados"
    cert_pdfs = sorted(cert_dir.glob("*.pdf")) if cert_dir.exists() else []
    for i, d in enumerate(docentes, 1):
        d[0] = i
        toks = set(norm_tokens(d[1]))
        propios_pdf = [f for f in cert_pdfs if len(toks & set(norm_tokens(f.stem))) >= 2]
        d.append("\n".join(f.name for f in propios_pdf) or None)
        d.append(f"Anexos/Certificados/{propios_pdf[0].name if len(propios_pdf) == 1 else ''}" if propios_pdf else None)
    con_cap = sum(1 for d in docentes if d[3] or d[5] or d[9])

    movilidad = []
    for r in wb_src["DECANATURA  movilidad"].iter_rows(min_row=4, values_only=True):
        nombre = r[3]
        if not nombre or not incluido(nombre):
            continue
        tipo = ("Entrante" if (r[9] or r[10]) else "Saliente") + " – " + ("Nacional" if (r[9] or r[11]) else "Internacional") if any(r[9:13]) else "No especificada"
        ini, fin = r[13], r[14]
        dias = (fin - ini).days + 1 if isinstance(ini, datetime) and isinstance(fin, datetime) else None
        movilidad.append([" ".join(w.capitalize() for w in str(nombre).split()), tipo, " ".join(str(r[4] or "").split()), r[5],
                          " ".join(str(r[6] or "").split()), str(r[8] or "").replace("1. ", "").replace("2. ", ""),
                          ini.date() if isinstance(ini, datetime) else ini, fin.date() if isinstance(fin, datetime) else fin, dias])
    movilidad.sort(key=lambda x: (x[6] or date.min), reverse=True)

    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 3 · Profesores — {mod['programa']}",
                 "Participación de los profesores de la Maestría en capacitación académica y administrativa, y movilidad académica. "
                 "Incluye las capacitaciones adicionales con certificado reportadas por los profesores (columna J de la hoja «Capacitación docente»); los certificados se adjuntan como soporte en la carpeta «Anexos/Certificados» (columna K).")
    fila = encabezado_factor(ws, 4, 3, m)
    pct = con_cap / len(docentes) * 100 if docentes else 0
    fila = st.kpis(ws, fila, [("Profesores de la Maestría", len(docentes)), ("Con capacitación registrada", con_cap),
                              ("Porcentaje (meta: 15 %)", f"{pct:.0f} %"), ("Movilidades registradas", len(movilidad))])
    st.nota(ws, fila + 1, "La planta docente es común a las dos modalidades: el Cuadro Maestro CNA No. 05 de SNIES 17528 (Investigación) y el de SNIES 116070 (Profundización) relacionan los mismos profesores. "
            "Por eso este libro es el mismo en las carpetas de Investigación y de Profundización. Grupo y categoría de investigador: Cuadro Maestro CNA No. 05.")
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 16

    ws = st.hoja(wb, "Capacitación docente", "Capacitación docente (SCRUM e ILUD) y perfil investigativo", "Fuentes: Decanatura (capacitaciones) y Cuadro Maestro CNA No. 05 (grupo y categoría).", 10)
    st.tabla(ws, 4, ["N°", "Profesor", "Categoría escalafón", "Capacitación SCRUM", "Año", "Segunda lengua (ILUD)", "Año",
                     "Grupo de investigación", "Categoría investigador MinCiencias", "Otras capacitaciones con certificado",
                     "Soporte (certificados en Anexos)", "_c"],
             docentes, [5, 34, 13, 18, 7, 40, 7, 18, 16, 70, 44], links={10: 11})

    ws = st.hoja(wb, "Movilidad", "Movilidad académica de profesores (entrante y saliente)", "Fuente: Decanatura – movilidad. Ordenada de la más reciente a la más antigua.", 9)
    st.tabla(ws, 4, ["Profesor / invitado", "Tipo", "Entidad o evento", "Lugar", "Objeto", "Modalidad", "Inicio", "Fin", "Días"],
             movilidad, [30, 24, 40, 22, 60, 12, 11, 11, 6])
    for row in ws.iter_rows(min_row=5, min_col=7, max_col=8):
        for c in row:
            c.number_format = "DD/MM/YYYY"

    act = dest / "a. Informes de participación de los docentes"
    act.mkdir(parents=True, exist_ok=True)
    nombre = "F3_Participacion_Docente_Capacitacion.xlsx"
    wb.save(act / nombre)
    anexos = []
    for f in cert_pdfs:
        copiar(f, act / "Anexos" / "Certificados")
        anexos.append(f"a. Informes de participación de los docentes/Anexos/Certificados/{f.name}")
    return [f"a. Informes de participación de los docentes/{nombre}"] + anexos


# ---------------------------------------------------------------------------
# FACTOR 4 — Egresados: servicios (se retiran Experiencias UD e infografía) e
# informe de caracterización e impacto con la Hoja de Vida de Egresados (OATI)
# ---------------------------------------------------------------------------
GRUPOS_EGRESADOS = ["Investigación", "Profundización"]  # el plan anterior se toma como Investigación
PLAN_ANTERIOR = ("95", "195", "295", "395", "495")

SECTORES = [  # (sector, palabras clave sin tildes, en minúscula); el primero que coincide gana
    ("Defensa y seguridad", ["defensa", "policia", "ejercito", "armada", "fuerza aerea", "fuerzas militares", "inteligencia"]),
    ("Educación", ["universidad", "colegio", "educacion", "sena", "servicio nacional de aprendizaje", "escuela", "instituto tecnico",
                   "politecnico", "corporacion universitaria", "fundacion universitaria", "uniminuto", "unad", "ecci", "fuac", "docente", "institucion educativa"]),
    ("Gobierno y entidades públicas", ["alcaldia", "gobernacion", "ministerio", "secretaria", "superintendencia", "igac", "instituto geografico",
                                       "catastro", "unidad administrativa", "departamento administrativo", "contraloria", "procuraduria",
                                       "fiscalia", "judicatura", "registraduria", "dane", "dian", "ideam", "instituto distrital", "agencia", "servicio geologico", "empresa de"]),
    ("Energía y servicios públicos", ["energia", "vanti", "gas", "acueducto", "codensa", "enel", "esp", "electrica", "petrol", "ecopetrol"]),
    ("Financiero y seguros", ["banco", "seguros", "financier", "fiduciaria", "bancolombia"]),
    ("Salud", ["salud", "hospital", "clinica", "eps"]),
    ("Telecomunicaciones y TIC", ["telefonica", "claro", "movistar", "tigo", "etb", "software", "tecnolog", "technolog", "sistemas", "informatic",
                                  "telecom", "digital", "datos", "systems", "ztech", "sofka", "epam", "redes", "cyber", "cloud", "system", "it"]),
    ("Geoinformación, consultoría e ingeniería", ["geomat", "geograf", "geoespac", "geodes", "geoinform", "sig", "mapper", "consult", "ingenier", "topograf", "cartograf", "swissphoto", "constructora", "proyectos", "asesor"]),
]


def _sin_tildes(s: str | None) -> str:
    return unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()


def sector_de(nombre: str | None) -> str:
    n = " " + re.sub(r"[^a-z0-9 ]", " ", _sin_tildes(nombre)) + " "
    for sector, claves in SECTORES:
        if any((" " + c.strip() + " " in n) if len(c.strip()) <= 3 else (c.strip() in n) for c in claves):
            return sector
    return "Sin clasificar"


ALIAS_INSTITUCIONES = [("distrital", "Universidad Distrital Francisco José de Caldas"), ("codazzi", "Instituto Geográfico Agustín Codazzi (IGAC)"),
                       ("igac", "Instituto Geográfico Agustín Codazzi (IGAC)"), ("servicio nacional de aprendizaje", "SENA"), ("sena-", "SENA"),
                       ("ecci", "Universidad ECCI"), ("nacional abierta", "UNAD"), ("unad", "UNAD"), ("uptc", "UPTC"),
                       ("pedagogica y tecnologica", "UPTC"), ("minuto de dios", "UNIMINUTO"), ("autonoma de colombia", "FUAC"),
                       ("udca", "UDCA"), ("ciencias aplicadas y ambientales", "UDCA")]
MINUSCULAS = {"de", "la", "del", "los", "las", "el", "y", "en", "para", "a"}


def institucion_canonica(nombre: str | None) -> str:
    n = _sin_tildes(nombre)
    for clave, canon in ALIAS_INSTITUCIONES:
        if clave in n:
            return canon
    palabras = re.sub(r"\s+", " ", (nombre or "").strip(" .-")).split()
    return " ".join(w.lower() if w.lower() in MINUSCULAS else (w.capitalize() if len(w) > 3 else w.upper()) for w in palabras)


def modalidad_con_plan_anterior(codigo: str) -> str | None:
    """Modalidad del estudiante; los del plan anterior (proyectos 95-495 sin modalidad registrada) se toman como Investigación."""
    mod, _ = modalidad_de(codigo)
    if mod:
        return mod
    r = ROSTER_POR_CODIGO.get(codigo)
    proy = codigo[5:8] if len(codigo) == 11 else ""
    return "Investigación" if proy in PLAN_ANTERIOR or (r and r["proyecto"] in PLAN_ANTERIOR) else None


def grupo_egresado(documento: str) -> tuple[str, dict | None]:
    """Modalidad de un egresado según Cóndor (por documento): primero su registro de graduación."""
    regs = sorted(ROSTER_POR_DOC.get(documento, []), key=lambda r: r["estado"] != "Graduado")
    if not regs:
        return "Sin registro en Cóndor", None
    for r in regs:
        mod = modalidad_con_plan_anterior(r["codigo"])
        if mod:
            return mod, regs[0]
    return "Sin registro en Cóndor", regs[0]


def analizar_egresados(path: Path) -> tuple[list[dict], dict]:
    import pandas as pd
    hojas = pd.read_excel(path, sheet_name=None, dtype=str)
    for d in hojas.values():
        if "identificacion" in d.columns:
            d["identificacion"] = d["identificacion"].astype(str).str.strip()
    por_persona = {h: d.groupby("identificacion") for h, d in hojas.items() if "identificacion" in d.columns and h != "datos basicos"}
    personas = []
    for doc in hojas["datos basicos"]["identificacion"].dropna().astype(str).str.strip().unique():
        grupo, reg = grupo_egresado(doc)
        get = lambda h: por_persona[h].get_group(doc) if doc in por_persona[h].groups else None  # noqa: E731
        prof, docn, inv, prod, idi = get("experiencia profesional"), get("experiencia docente"), get("grupos investigacion"), get("productos investigacion"), get("segunda lengua")
        p = {"grupo": grupo, "estado": reg["estado"] if reg else "Sin registro", "codigo": reg["codigo"] if reg else None,
             "proyecto_condor": reg["proyecto"] if reg else None}
        p["anio_grado"] = int(reg["ultima"][:4]) if reg and reg["estado"] == "Graduado" and re.match(r"\d{4}", reg["ultima"] or "") else None
        p["enfasis"] = enfasis_de(reg["codigo"]) if reg else None
        p["n_prof"] = 0 if prof is None else len(prof)
        act = prof[prof["actual"] == "S"] if prof is not None else None
        p["trabaja_actual"] = act is not None and len(act) > 0
        p["nivel_actual"] = sorted(set(act["nivel_institucion"].dropna())) if p["trabaja_actual"] else []
        p["sector_actual"] = sorted({sector_de(x) for x in act["nombre_institucion"]}) if p["trabaja_actual"] else []
        p["sectores_hist"] = sorted({sector_de(x) for x in prof["nombre_institucion"]}) if prof is not None else []
        p["inst_prof"] = sorted({institucion_canonica(x) for x in prof["nombre_institucion"].dropna()}) if prof is not None else []
        p["inst_prof_actual"] = sorted({institucion_canonica(x) for x in act["nombre_institucion"].dropna()}) if p["trabaja_actual"] else []
        p["n_doc"] = 0 if docn is None else len(docn)
        p["docente_actual"] = docn is not None and (docn["actual"] == "S").any()
        p["niveles_doc"] = sorted(set(docn["nivel_docencia"].dropna())) if docn is not None else []
        p["docente_posgrado"] = any(n in ("Maestria", "Especialización", "Doctorado") for n in p["niveles_doc"])
        p["inst_doc"] = sorted({institucion_canonica(x) for x in docn["nombre_institucion"].dropna()}) if docn is not None else []
        p["docente_ud"] = any("Distrital" in i for i in p["inst_doc"])
        p["n_grupos"] = 0 if inv is None else int(inv["grupo_investigacion"].notna().sum())
        p["n_proyectos"] = 0 if inv is None else int(inv["titulo_investigacion"].fillna("").str.strip().ne("").sum())
        p["cat_grupos"] = sorted({str(c).strip().upper() for c in inv["categoria_grupo"].dropna()}) if inv is not None else []
        p["en_investigacion"] = inv is not None and len(inv) > 0
        p["prod"] = dict(prod["tipo_producto"].dropna().value_counts()) if prod is not None else {}
        p["idiomas"] = [] if idi is None else [(r["idioma"], r["nivel_certificado"]) for _, r in idi.iterrows()]
        niv = [n for _, n in p["idiomas"] if n]
        orden = ["A1", "A2", "B1", "B2", "C1", "C2"]
        p["mejor_nivel"] = max(niv, key=lambda x: orden.index(x) if x in orden else -1) if niv else None
        p["secciones"] = sum([p["n_prof"] > 0, p["n_doc"] > 0, p["en_investigacion"] or bool(p["prod"]), bool(p["idiomas"])])
        personas.append(p)
    # cobertura: graduados distintos en Cóndor por grupo
    grad = {}
    for r in ROSTER:
        if r["estado"] == "Graduado" and r["documento"]:
            g, _ = grupo_egresado(r["documento"])
            grad[r["documento"]] = (g, int(r["ultima"][:4]) if re.match(r"\d{4}", r["ultima"] or "") else None)
    return personas, {"graduados_condor": grad, "hojas": {h: len(d) for h, d in hojas.items()}}


def _pct(n, d):
    return f"{n / d * 100:.0f} %" if d else "—"


ARCHIVO_ENCUESTA_EGRESADOS = "Caracterización e impacto de Egresados- MCIC (1-22).xlsx"
NIVELES_APORTE = ["Muy alto", "Alto", "Medio", "Bajo", "Ningún aporte"]
DIMENSIONES_APORTE = [21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]  # columnas de la encuesta (posición)
ITEMS_VALORACION = [37, 38, 39, 40, 41, 42, 43]


def analizar_encuesta_egresados(path: Path, etiqueta: str) -> dict | None:
    """Agregados de la encuesta de caracterización e impacto para la modalidad (sin nombres, documentos, organizaciones ni texto libre)."""
    import pandas as pd
    if not path.exists():
        return None
    d = pd.read_excel(path, dtype=str)
    col = lambda i: d.columns[i]  # noqa: E731
    mod = d[col(9)].fillna("").str.strip()
    d = d[mod == etiqueta].reset_index(drop=True)
    n = len(d)

    def conteo(i, multiple=False):
        c = Counter()
        for v in d[col(i)].dropna():
            partes = [x.strip() for x in str(v).split(";")] if multiple else [str(v).strip()]
            for x in partes:
                if x:
                    c[" ".join(x.split())] += 1
        return c.most_common()

    anios = Counter(int(str(v)[:4]) for v in d[col(10)].dropna())
    aporte = []
    for i in DIMENSIONES_APORTE:
        cn = Counter(d[col(i)].fillna("").str.strip())
        aporte.append([" ".join(col(i).split())] + [cn.get(x, 0) for x in NIVELES_APORTE] + [_pct(cn.get("Muy alto", 0) + cn.get("Alto", 0), n)])
    valoracion = []
    for i in ITEMS_VALORACION:
        vals = [int(x) for x in d[col(i)].dropna() if str(x).strip().isdigit()]
        valoracion.append([" ".join(col(i).split()).rstrip("."), round(sum(vals) / len(vals), 2) if vals else None, len(vals)])
    return {"n": n, "anios": sorted(anios.items()), "enfasis": conteo(11), "grupos": conteo(13), "situacion": conteo(14),
            "sector": conteo(15), "nivel": conteo(19), "relacion": conteo(20), "aporte": aporte, "actividades": conteo(32, True),
            "formacion": conteo(34, True), "grupo_inv": conteo(35), "publicaciones": conteo(36), "valoracion": valoracion,
            "recomienda": Counter(d[col(44)].fillna("").str.strip()), "impacto": conteo(45)}


def factor4(nombre_mod: str, mod: dict, dest: Path, m: dict) -> list[str]:
    etiqueta = mod["etiqueta"]
    fdir = factor_dirs(mod["plan"])[4]
    out = []
    for f in ("Portafolio de Servicios Grupo Funcional Egresados.pdf", "PROYECTO DE ACUERDO POLÍTICA DE EGRESADOS.pdf"):
        copiar(fdir / "a. Servicios" / f, dest / "a. Servicios")
        out.append(f"a. Servicios/{f}")

    personas, extra = analizar_egresados(fdir / "HojaVidaEgresados.xlsx")
    personas = [p for p in personas if p["grupo"] in GRUPOS_EGRESADOS]
    graduados = extra["graduados_condor"]
    grupos = {g: [p for p in personas if p["grupo"] == g] for g in GRUPOS_EGRESADOS}
    propios = grupos[etiqueta]

    def ind(ps: list[dict]) -> list:
        n = len(ps)
        c = lambda f: sum(1 for p in ps if f(p))  # noqa: E731
        return [n, c(lambda p: p["estado"] == "Graduado"), c(lambda p: p["estado"] != "Graduado"),
                c(lambda p: p["n_prof"]), c(lambda p: p["trabaja_actual"]),
                c(lambda p: "Pública" in p["nivel_actual"]), c(lambda p: "Privada" in p["nivel_actual"]), c(lambda p: "Mixta" in p["nivel_actual"]),
                c(lambda p: p["n_doc"]), c(lambda p: p["docente_actual"]), c(lambda p: p["docente_posgrado"]), c(lambda p: p["docente_ud"]),
                c(lambda p: p["en_investigacion"]), c(lambda p: p["n_grupos"]), c(lambda p: p["n_proyectos"]),
                c(lambda p: p["prod"].get("Ponencia internacional")), c(lambda p: p["prod"].get("Ponencia nacional")), c(lambda p: p["prod"].get("Patente")),
                c(lambda p: p["idiomas"]), c(lambda p: p["mejor_nivel"] in ("B1", "B2", "C1", "C2")),
                c(lambda p: p["secciones"] >= 3)]

    etiquetas_ind = ["Egresados con Hoja de Vida (módulo OATI)", "  de ellos, graduados según Cóndor", "  de ellos, otros estados (matriculados, inactivos, suspendidos)",
                     "Con experiencia profesional registrada", "Con vinculación laboral actual", "  vinculación actual en entidad pública", "  vinculación actual en entidad privada",
                     "  vinculación actual en entidad mixta", "Con experiencia docente registrada", "  docente actual", "  ha dictado docencia en posgrado",
                     "  ha sido docente de la Universidad Distrital", "Con grupo o proyecto de investigación registrado", "  con pertenencia a grupo de investigación",
                     "  con proyectos de investigación registrados", "Con ponencia internacional", "Con ponencia nacional", "Con patente",
                     "Con segunda lengua declarada", "  con nivel B1 o superior certificado", "Hoja de Vida con 3 o más secciones diligenciadas (de 4)"]
    cols = {g: ind(grupos[g]) for g in GRUPOS_EGRESADOS}
    total_ind = ind(personas)
    propios_ind = cols[etiqueta]
    filas_ind = [[e] + [cols[g][i] for g in GRUPOS_EGRESADOS] + [total_ind[i], "—" if i == 0 else _pct(propios_ind[i], len(propios))] for i, e in enumerate(etiquetas_ind)]

    # Cobertura frente a los graduados de Cóndor
    grad_por_grupo = Counter(g for g, _ in graduados.values())
    oati_grad_por_grupo = Counter(p["grupo"] for p in personas if p["estado"] == "Graduado")
    cob = [[g, grad_por_grupo.get(g, 0), oati_grad_por_grupo.get(g, 0), _pct(oati_grad_por_grupo.get(g, 0), grad_por_grupo.get(g, 0))] for g in GRUPOS_EGRESADOS]
    tg, to = sum(grad_por_grupo.get(g, 0) for g in GRUPOS_EGRESADOS), sum(oati_grad_por_grupo.get(g, 0) for g in GRUPOS_EGRESADOS)
    cob.append(["Total", tg, to, _pct(to, tg)])
    anios = ["2022", "2023", "2024", "2025", "2026"]
    cob_anio = []
    for a in anios:
        nc = sum(1 for gg, y in graduados.values() if gg == etiqueta and y == int(a))
        no = sum(1 for p in propios if p["estado"] == "Graduado" and p["anio_grado"] == int(a))
        cob_anio.append([etiqueta, a, nc, no, _pct(no, nc)])

    # Perfil por persona (anónimo y grueso: sin nombre, documento, empleador ni cargo)
    perfil = []
    for i, p in enumerate(sorted(propios, key=lambda x: (x["anio_grado"] or 0, x["enfasis"] or "")), 1):
        perfil.append([f"E-{i:03d}", p["enfasis"] or "—", p["estado"], p["anio_grado"] or "—",
                       "Sí" if p["trabaja_actual"] else "No", ", ".join(p["nivel_actual"]) or "—", ", ".join(p["sector_actual"]) or "—",
                       "Sí" if p["docente_actual"] else "No", ", ".join(x for x in p["niveles_doc"]) or "—",
                       p["n_grupos"], p["n_proyectos"], p["prod"].get("Ponencia internacional", 0) + p["prod"].get("Ponencia nacional", 0),
                       p["prod"].get("Patente", 0), p["mejor_nivel"] or "—", f"{p['secciones']}/4"])

    def tabla_sector(ps, clave):
        c = Counter()
        for p in ps:
            for s in p[clave]:
                c[s] += 1
        return [[s, n] for s, n in c.most_common()]

    inst, inst_doc = Counter(), Counter()
    for p in propios:
        for i in p["inst_prof"]:
            inst[i] += 1
        for i in p["inst_doc"]:
            inst_doc[i] += 1

    enc = analizar_encuesta_egresados(next(fdir.glob("Caracteriz*(1-22).xlsx"), fdir / ARCHIVO_ENCUESTA_EGRESADOS), etiqueta)

    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 4 · Egresados — caracterización e impacto — {mod['programa']}",
                 f"Hoja de Vida de Egresados del módulo institucional de la OATI ({len(propios)} egresados de {etiqueta}) y encuesta de caracterización e impacto.")
    fila = encabezado_factor(ws, 4, 4, m)
    kpis = [(f"Egresados de {etiqueta} (OATI)", len(propios)),
            ("Con experiencia profesional", _pct(sum(1 for p in propios if p["n_prof"]), len(propios))),
            ("Docentes actuales", sum(1 for p in propios if p["docente_actual"])),
            ("Con ponencias o patentes", sum(1 for p in propios if any(p["prod"].get(t) for t in ("Ponencia internacional", "Ponencia nacional", "Patente")))),
            ("Cobertura OATI 2022-2026", f"{sum(x[3] for x in cob_anio)}/{sum(x[2] for x in cob_anio)}")]
    if enc and enc["n"]:
        recom = sum(v for x, v in enc["recomienda"].items() if x in ("Definitivamente sí", "Probablemente sí"))
        alto = next((r for r in enc["aporte"] if r[0].lower().startswith("desarrollo profesional")), None)
        kpis += [("Respuestas a la encuesta de impacto", enc["n"]), ("Recomendarían la MCIC", _pct(recom, enc["n"])),
                 ("Aporte alto o muy alto al desarrollo profesional", alto[-1] if alto else "—")]
    fila = st.kpis(ws, fila, kpis)
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 18

    ws = st.hoja(wb, "Indicadores", "Indicadores de impacto por modalidad (conteo de personas)",
                 f"Investigación incluye a los egresados del plan anterior. Columna «% {etiqueta}»: porcentaje sobre los egresados de {etiqueta}.", 5)
    st.tabla(ws, 4, ["Indicador"] + GRUPOS_EGRESADOS + ["Total OATI", f"% {etiqueta}"], filas_ind, [58, 16, 16, 14, 14], filtro=False)

    if enc and enc["n"]:
        n = enc["n"]
        ws = st.hoja(wb, "Encuesta de impacto", f"Encuesta de caracterización e impacto de egresados — {etiqueta}",
                     f"{n} respuestas de egresados de {etiqueta}. Sin nombres, documentos ni organizaciones; el texto libre no se publica.", 8)
        f2 = 4

        def bloque(titulo, filas_):
            nonlocal f2
            f2 = st.seccion(ws, f2, titulo, 8)
            f2 = st.tabla(ws, f2, ["Respuesta", "Egresados", "%"], [[k, v, _pct(v, n)] for k, v in filas_], filtro=False, congelar=False)
        bloque("Año de graduación", [(str(a), v) for a, v in enc["anios"]])
        bloque("Énfasis", enc["enfasis"])
        bloque("Grupo de investigación del trabajo de grado", enc["grupos"])
        bloque("Situación principal actual", enc["situacion"])
        bloque("Sector en el que se desempeña", enc["sector"])
        bloque("Nivel de responsabilidad del cargo", enc["nivel"])
        bloque("Relación de la actividad actual con la formación recibida", enc["relacion"])
        f2 = st.seccion(ws, f2, "Aporte de la MCIC a su desarrollo (número de egresados por nivel)", 8)
        f2 = st.tabla(ws, f2, ["Dimensión"] + NIVELES_APORTE + ["% Alto o Muy alto"], enc["aporte"], filtro=False, congelar=False)
        bloque("Actividades después de graduarse (selección múltiple)", enc["actividades"])
        bloque("Formación continuada después de graduarse (selección múltiple)", enc["formacion"])
        bloque("Pertenece o ha pertenecido a un grupo de investigación", enc["grupo_inv"])
        bloque("Ha publicado productos académicos o científicos después de graduarse", enc["publicaciones"])
        f2 = st.seccion(ws, f2, "Valoración de la formación (escala 1 a 5)", 8)
        f2 = st.tabla(ws, f2, ["Aspecto", "Promedio", "Respuestas"], enc["valoracion"], filtro=False, congelar=False)
        bloque("¿Recomendaría la MCIC a otros profesionales?", enc["recomienda"].most_common())
        bloque("¿Ha generado algún impacto en su entorno a partir de su formación?", enc["impacto"])
        ws.column_dimensions["A"].width = 64
        for c_ in "BCDEFGH":
            ws.column_dimensions[c_].width = 14

    ws = st.hoja(wb, "Perfil egresados", f"Perfil de los egresados de {etiqueta} (anónimo)",
                 "Un registro por persona con identificador interno (E-001…). No incluye nombre, documento, empleador ni cargo; la relación persona–identificador no se publica.", 15)
    st.tabla(ws, 4, ["ID", "Énfasis (plan de ingreso)", "Estado en Cóndor", "Año de grado (est.)", "Vinculación actual", "Tipo de entidad", "Sector estimado",
                     "Docente actual", "Niveles de docencia", "Grupos de investigación", "Proyectos de investigación", "Ponencias", "Patentes", "Mejor nivel de inglés/2ª lengua", "Secciones diligenciadas"],
             perfil, [8, 22, 16, 11, 11, 18, 30, 10, 28, 11, 11, 10, 9, 14, 12])

    ws = st.hoja(wb, "Sector y vinculación", "Dónde trabajan los egresados (sector estimado por nombre de la entidad)",
                 "El módulo no tiene un campo de sector económico: se estimó por palabras clave en el nombre de la entidad. Quedan «Sin clasificar» los nombres que no permiten inferirlo.", 6)
    f2 = st.seccion(ws, 4, f"Vinculación actual — {etiqueta}", 4)
    f2 = st.tabla(ws, f2, ["Sector estimado", "Egresados"], tabla_sector(propios, "sector_actual") or [["Sin vinculación actual registrada", 0]], [46, 14], filtro=False, congelar=False)
    f2 = st.seccion(ws, f2, f"Trayectoria profesional (todas las experiencias) — {etiqueta}", 4)
    st.tabla(ws, f2, ["Sector estimado", "Egresados con experiencia"], tabla_sector(propios, "sectores_hist") or [["—", 0]], [46, 14], filtro=False, congelar=False)

    ws = st.hoja(wb, "Instituciones", f"Instituciones donde han trabajado o dictado docencia los egresados de {etiqueta}",
                 "Conteo de egresados distintos por institución (las 25 principales con 2 o más egresados); no se asocia a personas.", 4)
    f2 = st.tabla(ws, 4, ["Experiencia profesional — institución", "Egresados"], [[i, n_] for i, n_ in inst.most_common(25) if n_ >= 2], [60, 12], filtro=False, congelar=False)
    st.tabla(ws, f2, ["Docencia — institución", "Egresados"], [[i, n_] for i, n_ in inst_doc.most_common(25) if n_ >= 2], [60, 12], filtro=False, congelar=False)

    ws = st.hoja(wb, "Cobertura", "Cobertura de la Hoja de Vida de Egresados frente a los graduados de Cóndor",
                 "Meta 2026-2027: al menos el 50 % de los egresados con información actualizada en el módulo de Hoja de Vida (OATI). Graduados de Cóndor = documentos distintos con estado Graduado.", 5)
    f2 = st.tabla(ws, 4, ["Grupo", "Graduados (Cóndor)", "Con Hoja de Vida (OATI)", "Cobertura"], cob, [34, 20, 24, 12], filtro=False, congelar=False)
    f2 = st.seccion(ws, f2, f"Graduados por año de grado estimado (última matrícula) — {etiqueta}", 5)
    st.tabla(ws, f2, ["Grupo", "Año", "Graduados (Cóndor)", "Con Hoja de Vida (OATI)", "Cobertura"], cob_anio, [34, 8, 20, 24, 12], filtro=False, congelar=False)

    ws = st.hoja(wb, "Privacidad", "Privacidad", None, 2)
    st.tabla(ws, 3, ["Tema", "Detalle"], [["Privacidad", "Este informe no incluye nombres, documentos, fechas de nacimiento, correos ni teléfonos; las instituciones se agregan sin asociarlas a personas."]], [28, 130], filtro=False, congelar=False)

    act = dest / "e. Caracterización e impacto de los egresados"
    act.mkdir(parents=True, exist_ok=True)
    nombre = f"F4_Egresados_Impacto_{nombre_mod}.xlsx"
    wb.save(act / nombre)
    out.append(f"{act.name}/{nombre}")
    return out


# ---------------------------------------------------------------------------
# FACTOR 5 — Syllabus actualizados y verificación Investigación vs Profundización
# ---------------------------------------------------------------------------
PLAN_RES016 = {
    "Investigación": [("I", "Seminario de investigación", "OB", 4), ("I", "Herramientas matemáticas para el manejo de la información", "OB", 4),
                      ("I", "Informática", "OB", 4), ("I", "Énfasis I (1 espacio)", "OB", 4),
                      ("II", "Trabajo de grado I", "OB", 4), ("II", "Electiva", "EI", 4), ("II", "Énfasis II (2 espacios)", "OB", 8),
                      ("III", "Trabajo de grado II", "OB", 8), ("III", "Electiva (espacios de énfasis del periodo III u otros)", "EI", 4)],
    "Profundización": [("I", "Seminario de investigación", "OB", 4), ("I", "Herramientas matemáticas para el manejo de la información", "OB", 4),
                       ("I", "Informática", "OB", 4), ("I", "Énfasis I (1 espacio)", "OB", 4),
                       ("II", "Electiva", "EI", 4), ("II", "Énfasis II (3 espacios)", "OB", 12),
                       ("III", "Trabajo de grado", "OB", 4), ("III", "Electiva", "EI", 4), ("III", "Énfasis (1 espacio)", "OB", 4)],
}
HORAS_RES016 = {"Investigación": (384, 176, 1552), "Profundización": (480, 160, 1472)}


def leer_codigo_plan(path: Path) -> str | None:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for ws in wb:
        for r in ws.iter_rows(values_only=True):
            for i, c in enumerate(r):
                if isinstance(c, str) and "CÓDIGO PLAN DE ESTUDIOS" in c.upper():
                    resto = c.upper().split("CÓDIGO PLAN DE ESTUDIOS", 1)[1].strip(" :")
                    if resto:
                        return resto
                    sig = [x for x in r[i + 1:] if x not in (None, "")]
                    return str(sig[0]) if sig else None
        break
    return None


def factor5(nombre_mod: str, mod: dict, dest: Path, m: dict) -> list[str]:
    fdir = factor_dirs(mod["plan"])[5]
    out = []
    act_a = dest / "a. Definición y socialización de ejes de formación y competencias"
    res = next(fdir.glob("*Res 016 2025-CA.pdf"))
    copiar(res, act_a, "Res 016 de 2025 Consejo Académico - Plan de estudios MCIC.pdf")
    out.append(f"{act_a.name}/Res 016 de 2025 Consejo Académico - Plan de estudios MCIC.pdf")

    act_b = dest / "b. Resultados de Aprendizaje a nivel microcurricular"
    syl = json.loads((SILVER / "syllabi.json").read_text(encoding="utf-8"))
    filas = []
    for e in syl["espacios_academicos"]:
        src = ROOT / e["archivo"]
        copiar(src, act_b / "Syllabus AA-FR-003" / e["area"])
        out.append(f"{act_b.name}/Syllabus AA-FR-003/{e['area']}/{src.name}")
        filas.append([e["codigo"], e["nombre"].capitalize(), e["area"], e["nivel"], e["creditos"],
                      f"{e['htd']}/{e['htc']}/{e['hta']}", src.name, md5(src)[:10], "Sí (mismo archivo)",
                      leer_codigo_plan(src) or "Vacío", "No"])

    # comparación de las carpetas SNIES de acreditación
    a = BRONZE / "ACREDITACIÓN DE ALTA CALIDAD/SNIES17528-MCIC-Investigación/SNIES17528-Syllabus"
    b = BRONZE / "ACREDITACIÓN DE ALTA CALIDAD/SNIES116070-MCIC-Profundización/SNIES116070-Syllabus"
    ha = {md5(p): p.relative_to(a) for p in a.rglob("*") if p.is_file()}
    hb = {md5(p): p.relative_to(b) for p in b.rglob("*") if p.is_file()}
    snies = [[str(ha[h]), str(hb.get(h, "—")), "Idéntico" if h in hb else "Diferente"] for h in sorted(ha, key=lambda k: str(ha[k]))]
    iguales = sum(1 for x in snies if x[2] == "Idéntico")

    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", "Factor 5 · Verificación de syllabus: Investigación vs Profundización",
                 "¿Se diferencian los syllabus entre la modalidad de Investigación (SNIES 17528) y la de Profundización (SNIES 116070)?")
    fila = encabezado_factor(ws, 4, 5, m)
    fila = st.kpis(ws, fila, [("Syllabus AA-FR-003", len(filas)), ("Créditos por modalidad", 44),
                              ("Archivos SNIES idénticos", f"{iguales}/{len(snies)}"), ("Syllabus con código de plan", sum(1 for f in filas if f[9] != "Vacío"))])
    con_codigo = [f for f in filas if f[9] != "Vacío"]
    identificacion = (f"{len(filas) - len(con_codigo)} de {len(filas)} syllabus dejan vacío el campo «Código plan de estudios»"
                      + (f"; los {len(con_codigo)} que lo diligencian ({', '.join(f[1] for f in con_codigo)}) registran «"
                         + ", ".join(sorted({re.sub(r'[^0-9 ]', '', f[9]).split()[0] for f in con_codigo})) + "», que no corresponde a 595 ni a 695"
                         if con_codigo else "")
                      + ". Ningún syllabus indica la modalidad: el campo «Proyecto curricular» dice solo «Maestría en Ciencias de la Información y las Comunicaciones».")
    # La coordinación retiró tres hallazgos (malla de espacios académicos, espacio sin correspondencia en la Res. 016 y denominaciones);
    # el detalle sigue disponible en CONTEXTO_NuevaData.md §4.
    hallazgos = [
        ["Plan de estudios", "Sí se diferencia", "La Res. 016 de 2025 del Consejo Académico aprueba un plan para cada modalidad (44 créditos cada uno). Investigación: Trabajo de grado I y II (12 créditos), 3 espacios de énfasis (12 créditos) y 2 electivas. Profundización: Trabajo de grado (4 créditos), 5 espacios de énfasis (20 créditos) y 2 electivas."],
        ["Horas de trabajo académico", "Sí se diferencia", "Investigación: 384 HTD / 176 HTC / 1.552 HTA. Profundización: 480 HTD / 160 HTC / 1.472 HTA (2.112 horas en ambos casos)."],
        ["Espacios académicos", "Compartidos", "Los espacios de fundamentación, seminario y énfasis son los mismos en las dos modalidades; cambia cuántos espacios de énfasis cursa el estudiante y si los del periodo III son obligatorios (Profundización) o electivos (Investigación)."],
        ["Archivos de syllabus", "No se diferencian", f"Los {len(snies)} syllabus de la carpeta SNIES 17528 y los de SNIES 116070 son byte a byte idénticos ({iguales}/{len(snies)}). Los {len(filas)} syllabus actualizados (AA-FR-003) son un único archivo por espacio académico, usado por las dos modalidades."],
        ["Identificación de la modalidad en el syllabus", "No se diferencia", identificacion],
    ]
    fila = st.seccion(ws, fila + 1, "Hallazgos")
    st.tabla(ws, fila, ["Aspecto", "Resultado", "Detalle"], hallazgos, [30, 18, 110], filtro=False, congelar=False)
    ws.column_dimensions["C"].width = 110

    ws = st.hoja(wb, "Plan de estudios Res 016", "Estructura del plan de estudios por modalidad (Res. 016 de 2025, Consejo Académico)", "OB = obligatorio · EI = electivo intrínseco.", 9)
    filas_plan = []
    for mm in ("Investigación", "Profundización"):
        for p, esp, cl, cr in PLAN_RES016[mm]:
            filas_plan.append([mm, p, esp, cl, cr])
        h = HORAS_RES016[mm]
        filas_plan.append([mm, "Total", "44 créditos · obligatorios 36 (81,82 %) · electivos 8 (18,18 %)", "", 44])
        filas_plan.append([mm, "Horas", f"HTD {h[0]} · HTC {h[1]} · HTA {h[2]} · total 2.112", "", None])
    st.tabla(ws, 4, ["Modalidad", "Periodo", "Espacio académico", "Clasificación", "Créditos"], filas_plan, [16, 9, 70, 14, 10])

    ws = st.hoja(wb, "Syllabus AA-FR-003", "Syllabus actualizados (formato AA-FR-003)", "Un archivo por espacio académico; el mismo archivo aplica a las dos modalidades.", 11)
    st.tabla(ws, 4, ["Código", "Espacio académico", "Área / énfasis", "Nivel", "Créditos", "HTD/HTC/HTA", "Archivo", "Huella (MD5)",
                     "¿Aplica a ambas modalidades?", "Campo «Código plan de estudios»", "¿Indica modalidad?"],
             filas, [11, 42, 28, 6, 8, 12, 46, 12, 16, 16, 12])

    ws = st.hoja(wb, "Comparación SNIES", "Comparación de las carpetas de syllabus de acreditación", "SNIES 17528 (Investigación) frente a SNIES 116070 (Profundización), por huella MD5.", 3)
    st.tabla(ws, 4, ["Archivo en SNIES 17528", "Archivo en SNIES 116070", "Resultado"], snies, [70, 70, 12])

    for orig in sorted((fdir / "SNIES17528-Syllabus").rglob("*")):
        if orig.is_file():
            partes = [re.sub(r"(.)╠ü", lambda mm: unicodedata.normalize("NFC", mm.group(1) + "\u0301"), x) for x in orig.relative_to(fdir / "SNIES17528-Syllabus").parts]
            copiar(orig, act_b / "Anexos" / "Syllabus SNIES (acreditación)" / Path(*partes[:-1]), partes[-1])
            out.append(f"{act_b.name}/Anexos/Syllabus SNIES (acreditación)/{'/'.join(partes)}")
    nombre = "F5_Verificacion_Syllabus_Investigacion_vs_Profundizacion.xlsx"
    act_b.mkdir(parents=True, exist_ok=True)
    wb.save(act_b / nombre)
    out.append(f"{act_b.name}/{nombre}")
    return out


# ---------------------------------------------------------------------------
# FACTOR 6 — Permanencia y graduación (prórrogas 2026-3, PAGOT y graduados)
# ---------------------------------------------------------------------------
def modalidad_declarada(txt: str | None) -> str | None:
    t = _sin_tildes(txt)
    if "prof" in t:
        return "Profundización"
    if "inv" in t:
        return "Investigación"
    return None


def fecha_correo(pdf: Path) -> str | None:
    import subprocess
    txt = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True).stdout
    m = re.search(r"Fecha\s+\w{3}\s+(\d{1,2})/(\d{2})/(\d{4})", txt)
    return f"{m.group(3)}-{m.group(2)}-{int(m.group(1)):02d}" if m else None


def leer_prorrogas_2026_3(base: Path) -> list[dict]:
    """Filas de los tres «Estudiantes Prorroga 2026-3» (1.ª, 2.ª y 3.ª solicitud) con la fecha del correo de cada estudiante."""
    filas = []
    for ronda, carpeta in enumerate(("PRIMERA SOLICITUD", "SEGUNDA SOLICITUD", "TERCERA SOLICITUD"), 1):
        d = base / carpeta
        xlsx = next(p for p in d.glob("*.xlsx") if "Prorroga" in p.name or "Prórroga" in p.name)
        correos = {p: norm_tokens(p.stem) for p in (d / "CORREOS").glob("*.pdf")}
        ws = openpyxl.load_workbook(xlsx, data_only=True).active
        for r in ws.iter_rows(min_row=5, values_only=True):
            if not r[2]:
                continue
            codigo = str(r[2]).strip().split(".")[0]
            nombre = " ".join(str(r[3]).split())
            toks = norm_tokens(nombre)
            pdf = next((p for p, t in correos.items() if t == toks or t <= toks or toks <= t), None)
            filas.append({"ronda": ronda, "num": int(r[1]), "codigo": codigo, "nombre": nombre,
                          "tipo": "PAGOT" if str(r[4]).strip().upper() == "PAGOT" else "Regular",
                          "modalidad": modalidad_declarada(r[5]), "viabilidad": str(r[6] or "").strip().upper(),
                          "fecha_aval": r[7].date() if isinstance(r[7], datetime) else None,
                          "componente1": " ".join(str(r[8] or "—").split()), "tesis_radicada": str(r[9] or "").strip().upper(),
                          "creditos": r[10], "trabajo_grado": " ".join(str(r[11] or "—").split()),
                          "fecha_correo": fecha_correo(pdf) if pdf else None})
    return filas


def leer_pagot_2026_3(path: Path) -> list[dict]:
    ws = openpyxl.load_workbook(path, data_only=True)["SEGUMIENTO PAGOT"]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[4]:
            continue
        resp = _sin_tildes(r[9])
        if "no respondio" in resp:
            cat = "No respondió"
        elif "no pago" in resp:
            cat = "Pago incompleto"
        elif "acta" in resp and "si acepto" not in resp:
            cat = "Acta pendiente o no enviada"
        elif "si acepto" in resp.replace("  ", " "):
            cat = "Aceptó"
        elif "pendiente" in resp or not resp:
            cat = "Pendiente"
        else:
            cat = "Otra respuesta"
        plan = str(r[5] or "")
        out.append({"codigo": str(r[1]).strip(), "ultimo": " ".join(str(r[2] or "—").split()), "plan_est": r[3],
                    "modalidad": "Investigación" if plan.startswith("595") else "Profundización" if plan.startswith("695") else None,
                    "creditos": r[6], "periodos": r[7], "matriculas": r[8], "respuesta": cat,
                    "pago": "Sí" if re.match(r"\s*2026-3\s*A", str(r[15] or "")) else "No",
                    "p2026_3": r[11], "p2027_1": r[12], "p2027_3": r[13],
                    "estado": r[14] if re.search(r"matriculad|renovaci|p[eé]rdida|graduad|abandono|inactiv", _sin_tildes(r[14]) or "") else None,
                    "recibo": r[15] if re.match(r"\s*\d{4}-\d", str(r[15] or "")) else None})
    return out


def leer_inscritos_pagot(path: Path) -> list[dict]:
    ws = openpyxl.load_workbook(path, data_only=True).active
    out, per = [], None
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[0]:
            per = str(r[0]).strip()
        if r[1]:
            out.append({"periodo": per, "codigo": str(r[1]).strip(),
                        "modalidad": "Profundización" if "rofund" in str(r[4]) else "Investigación"})
    return out


def factor6(nombre_mod: str, mod: dict, dest: Path, tr: dict, m: dict) -> list[str]:
    etiqueta = mod["etiqueta"]
    otra = "Profundización" if etiqueta == "Investigación" else "Investigación"
    f6 = factor_dirs(mod["plan"])[6]
    pr = leer_prorrogas_2026_3(f6 / "PRORROGA 2026-3")
    pagot26 = leer_pagot_2026_3(f6 / "PAGOT 20263.xlsx")
    inscritos = leer_inscritos_pagot(f6 / "Aspirantes Inscritos 2023-3 al 2025-3 MCIComunicaciones.xlsx")

    # --- Solicitudes de prórroga 2026-3 (modalidad declarada en el cuadro de la coordinación; se contrasta con Cóndor/bases)
    filas = []
    for p in pr:
        mmod = p["modalidad"] or modalidad_con_plan_anterior(p["codigo"])
        base = BASES.get(p["codigo"], {})
        cond = ROSTER_POR_CODIGO.get(p["codigo"], {})
        id_ = f"P{p['ronda']}-{p['num']:02d}"
        fila = [id_, f"{p['ronda']}.ª", fecha_es(p["fecha_correo"]), mmod, enfasis_de(p["codigo"]) or frase(base.get("enfasis")) or "—",
                p["tipo"], p["viabilidad"] or "—", fecha_es(p["fecha_aval"].isoformat() if p["fecha_aval"] else None), p["componente1"], p["tesis_radicada"] or "—",
                p["creditos"], p["trabajo_grado"], f"{cond.get('estado', '—')} (última matrícula {cond.get('ultima', '—')})" if cond else "—"]
        if mmod == etiqueta:
            filas.append(fila)
    codigos_pr = {p["codigo"] for p in pr if (p["modalidad"] or modalidad_con_plan_anterior(p["codigo"])) == etiqueta}
    unicos = len(codigos_pr)
    n_pagot = sum(1 for f in filas if f[5] == "PAGOT")
    por_ronda = Counter(f[1] for f in filas)

    # --- PAGOT: cruce de las cuatro fuentes
    arch = "MCIC - Base de datos INVESTIGACION.xlsx" if etiqueta == "Investigación" else "MCIC - Base de datos Profundizacion.xlsx"
    hoja = "N-A Investigación" if etiqueta == "Investigación" else "N-A Profundizacion"
    wbb = openpyxl.load_workbook(BRONZE / "Maestria CIC/2026/BASES  DE DATOS ESTUDIANTES" / arch, read_only=True, data_only=True)
    rows = list(wbb[hoja].iter_rows(values_only=True))
    enc = [str(c).strip() if c else "" for c in rows[0]]
    it, ig = enc.index("TIPO DE ESTUDIANTE"), enc.index("INGRESO")
    base_pagot = {str(r[0]).split(".")[0].strip(): r for r in rows[1:] if r and r[0] and r[it] == "PAGOT"}
    regulares = [r for r in rows[1:] if r and r[0] and r[it] == "REGULAR"]
    seg = [x for x in pagot26 if (x["modalidad"] or modalidad_con_plan_anterior(x["codigo"])) == etiqueta]
    seg_otra = [x for x in pagot26 if (x["modalidad"] or modalidad_con_plan_anterior(x["codigo"])) == otra]
    seg_cod = {x["codigo"] for x in seg}
    pagaron = [x for x in seg if x["pago"] == "Sí"]
    insc = [x for x in inscritos if x["modalidad"] == etiqueta]
    insc_cod = {x["codigo"] for x in insc}
    pr_pagot = {p["codigo"] for p in pr if p["tipo"] == "PAGOT" and (p["modalidad"] or modalidad_con_plan_anterior(p["codigo"])) == etiqueta}
    cruce = [
        ["Base MCIC (columna «Tipo de estudiante» = PAGOT)", len(base_pagot), len(set(base_pagot) & seg_cod), len(set(base_pagot) & pr_pagot)],
        ["Inscritos oficiales a PAGOT 2024-1 a 2025-3 (listado de aspirantes)", len(insc_cod), len(insc_cod & seg_cod), len(insc_cod & pr_pagot)],
        ["Seguimiento PAGOT 2026-3 (Plan vigente " + ("595" if etiqueta == "Investigación" else "695") + ")", len(seg_cod), len(seg_cod), len(seg_cod & pr_pagot)],
        ["    de ellos, pagaron la matrícula 2026-3 (recibo activo)", len(pagaron), len(pagaron), len({x['codigo'] for x in pagaron} & pr_pagot)],
        ["Solicitudes de prórroga 2026-3 de estudiantes PAGOT", len(pr_pagot), len(pr_pagot & seg_cod), len(pr_pagot)],
    ]
    ing = Counter(x["periodo"] for x in insc)
    resp = Counter(x["respuesta"] for x in seg)
    est = Counter(" ".join(str(x["estado"] or "Sin dato de estado").split())[:48] for x in seg)
    pagot_filas = []
    for i, x in enumerate(sorted(seg, key=lambda z: z["ultimo"]), 1):
        pagot_filas.append([f"G-{i:02d}", enfasis_de(x["codigo"]) or "—", x["ultimo"], x["creditos"], x["periodos"], x["matriculas"], x["respuesta"], x["pago"],
                            x["p2026_3"] or "—", x["p2027_1"] or "—", x["p2027_3"] or "—", " ".join(str(x["estado"] or "—").split()), x["recibo"] or "—",
                            "Sí" if x["codigo"] in codigos_pr else "No", "Sí" if x["codigo"] in base_pagot else "No", "Sí" if x["codigo"] in insc_cod else "No"])
    # Reunión de acompañamiento (convocatoria del 06/04/2026)
    conv = tr["eventos"]["reunion_acompanamiento_2026_04_06"]["convocados_codigos"]
    conv_mod = Counter(modalidad_con_plan_anterior(c) for c in conv if modalidad_con_plan_anterior(c))

    # Graduados por año (estimado por última matrícula) y modalidad
    grad = Counter()
    for r in ROSTER:
        if r["estado"] == "Graduado" and r["ultima"][:4].isdigit() and 2022 <= int(r["ultima"][:4]) <= 2026:
            mm = modalidad_con_plan_anterior(r["codigo"])
            if mm:
                grad[(mm, r["ultima"][:4])] += 1
    anios = ["2022", "2023", "2024", "2025", "2026"]
    grad_filas = []
    vals = [grad.get((etiqueta, a), 0) for a in anios]
    grad_filas.append([etiqueta] + vals + [sum(vals)])

    norma = openpyxl.load_workbook(f6 / "Normativa_PAGOT_UD.xlsx", data_only=True)["Normativa PAGOT"]
    normas = []
    for r in norma.iter_rows(min_row=3, values_only=True):
        if r[0]:
            normas.append([r[0], r[1], r[2], r[3], "Ver norma" if r[4] else "—", r[4]])

    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 6 · Permanencia y graduación — {mod['programa']}",
                 "Estudiantes con pendiente de trabajo de grado y riesgo de pérdida de calidad: solicitudes de prórroga 2026-3, estudiantes PAGOT, acompañamiento y graduados. "
                 "Los estudiantes se identifican con un ID; los nombres, códigos y correos originales se resguardan y no se publican.")
    fila = encabezado_factor(ws, 4, 6, m)
    fila = st.kpis(ws, fila, [("Solicitudes de prórroga radicadas 2026-3", len(filas)), ("Estudiantes que solicitaron prórroga", unicos),
                              ("Solicitantes que son PAGOT", n_pagot), ("PAGOT en seguimiento 2026-3", len(seg)),
                              ("PAGOT que pagaron la matrícula 2026-3", len(pagaron)), ("Inscritos oficiales PAGOT", len(insc_cod)),
                              ("Convocados a acompañamiento", conv_mod.get(etiqueta, 0)), (f"Graduados 2022-2026", grad_filas[0][-1])])
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 17

    enc_p = ["ID", "Ronda", "Fecha del correo", "Modalidad", "Énfasis", "Tipo de estudiante", "Aval anteproyecto", "Fecha del aval", "Componente 1 (evento o revista)",
             "Tesis radicada", "Créditos aprobados", "Trabajo de grado", "Estado en Cóndor"]
    anch = [8, 7, 12, 14, 20, 11, 11, 11, 38, 9, 10, 26, 28]
    ws = st.hoja(wb, "Solicitudes de prórroga", f"Solicitudes de prórroga de permanencia 2026-3 — {etiqueta}",
                 "Tres rondas de radicación ante la coordinación (mayo-junio de 2026) para estudio del Consejo de Facultad.", 14)
    st.tabla(ws, 4, enc_p, filas, anch)

    ws = st.hoja(wb, "PAGOT", f"Estudiantes PAGOT — {etiqueta}", "Cruce de las cuatro fuentes disponibles y seguimiento individual 2026-3 (anónimo).", 16)
    f2 = st.seccion(ws, 4, "Cruce de fuentes", 8)
    f2 = st.tabla(ws, f2, ["Fuente", "Estudiantes", "También en seguimiento 2026-3", "También solicitó prórroga 2026-3"], cruce, [64, 14, 22, 22], filtro=False, congelar=False)
    f2 = st.seccion(ws, f2, "Respuesta a la convocatoria de PAGOT (seguimiento 2026-3)", 8)
    f2 = st.tabla(ws, f2, ["Respuesta", "Estudiantes"], [[k, v] for k, v in resp.most_common()], filtro=False, congelar=False)
    f2 = st.seccion(ws, f2, "Estado en Cóndor (seguimiento 2026-3)", 8)
    f2 = st.tabla(ws, f2, ["Estado", "Estudiantes"], [[k, v] for k, v in est.most_common()], filtro=False, congelar=False)
    f2 = st.seccion(ws, f2, "Inscritos oficiales a PAGOT por periodo", 8)
    f2 = st.tabla(ws, f2, ["Periodo", "Estudiantes"], [[k, v] for k, v in sorted(ing.items())], filtro=False, congelar=False)
    ws = st.hoja(wb, "Seguimiento PAGOT 2026-3", f"Seguimiento individual PAGOT 2026-3 — {etiqueta}",
                 "Un registro por estudiante con ID interno (G-01…); sin nombre, documento ni código. El plan de trabajo son los espacios que cursará cada estudiante en cada periodo.", 16)
    st.tabla(ws, 4, ["ID", "Énfasis (plan de ingreso)", "Último periodo activo", "Créditos aprobados", "Periodos", "Matrículas", "Respuesta a la convocatoria", "Pago 2026-3",
                     "Plan 2026-3", "Plan 2027-1", "Plan 2027-3", "Estado en Cóndor", "Último recibo", "Solicitó prórroga 2026-3", "En base MCIC como PAGOT", "Inscrito oficial PAGOT"],
             pagot_filas, [7, 22, 12, 10, 9, 10, 22, 9, 26, 22, 22, 30, 12, 12, 13, 12])
    ws = st.hoja(wb, "Normativa PAGOT", "Normativa institucional del Plan de Graduación Oportuna (PAGOT)", "Fuente: Normativa_PAGOT_UD.xlsx.", 5)
    st.tabla(ws, 4, ["Tipo", "Número", "Año", "Expedida por", "Enlace", "_u"], normas, [14, 9, 8, 32, 14], links={4: 5})

    ws = st.hoja(wb, "Acompañamiento", "Reunión de acompañamiento académico (06/04/2026, 7:00 p. m., virtual)",
                 "Convocatoria de la coordinación a estudiantes cuyo tiempo de permanencia vence en 2026-1, para orientar rutas de culminación del trabajo de grado.", 3)
    st.tabla(ws, 4, ["Modalidad", "Estudiantes convocados", "Fuente"],
             [[k, v, "Correo de convocatoria de la coordinación (6 de abril de 2026)"] for k, v in sorted(conv_mod.items())], [24, 22, 70], filtro=False)

    ws = st.hoja(wb, "Graduados", f"Graduados 2022-2026 — {etiqueta}", "Año estimado a partir de la última matrícula registrada en Cóndor.", 7)
    st.tabla(ws, 4, ["Modalidad"] + anios + ["Total"], grad_filas, [28, 9, 9, 9, 9, 9, 10], filtro=False)

    act = dest / "d. Seguimiento al avance de trabajos de grado y tiempos de permanencia"
    act.mkdir(parents=True, exist_ok=True)
    nombre = f"F6_Permanencia_y_Graduacion_{nombre_mod}.xlsx"
    wb.save(act / nombre)
    copiar(f6 / "Normativa_PAGOT_UD.xlsx", act)
    return [f"{act.name}/{nombre}", f"{act.name}/Normativa_PAGOT_UD.xlsx"]



# ---------------------------------------------------------------------------
# Insumos de los factores 7, 8 y 12 (viven en Original/<Modalidad>/FACTOR N; ver CONTEXTO §12)
# ---------------------------------------------------------------------------
def compartido(mod: dict, n: int, relativo: str, patron: str) -> list[Path]:
    """Archivos que coinciden en la carpeta de la modalidad; si no están, en la de Investigación (documentos pesados compartidos por las dos modalidades)."""
    for m in (mod, MODALIDADES["Investigacion"]):
        base = factor_dirs(m["plan"])[n] / relativo
        encontrados = sorted(base.glob(patron)) if base.exists() else []
        if encontrados:
            return encontrados
    return []


def carpeta_factor(mod: dict, n: int) -> Path:
    return factor_dirs(mod["plan"])[n]


CEDULAS_CONVENIO_IGAC = ["79.571.941", "51.975.477", "79.367.935", "60261576"]  # documentos de identidad que aparecen en los soportes


def redactar_documentos_identidad(src: Path, dst: Path) -> None:
    """Copia el PDF tapando los números de cédula de las personas que firman o son designadas."""
    import pymupdf
    patron = re.compile(r"(?:c[eé]dula[^0-9]{0,90}?|\bC\.?\s?C\.?\s*(?:No\.?)?\s*)(\d{1,3}(?:[.,]\d{3}){2}|\d{7,10})", re.I)
    doc = pymupdf.open(src)
    for page in doc:
        numeros = set(CEDULAS_CONVENIO_IGAC) | {m.group(1) for m in patron.finditer(" ".join(page.get_text("text").split()))}
        for n in numeros:
            for r in page.search_for(n):
                page.add_redact_annot(r, fill=(0, 0, 0))
        page.apply_redactions()
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst, garbage=3, deflate=True)


SOPORTES_CONVENIO_IGAC = {  # nombre original (por prefijo) -> nombre publicado. «0-Apertura Financiera» no se publica (datos bancarios).
    "CONVENIO IGAC": "Convenio específico de cooperación 5570 de 2025 (IGAC - UDFJC).pdf",
    "SECOP II": "SECOP II - Convenio específico 5570 de 2025.pdf",
    "ACTA REUNI": "Acta de la mesa técnica del 3 de junio de 2026 (apoyo a la supervisión).pdf",
    "Designación comité técnico IGAC": "Designación del comité técnico por el IGAC (septiembre de 2026).pdf",
    "Designación comité técnico UD": "Designación del representante de la UD ante el comité técnico (agosto de 2026).pdf",
    "Designación de apoyo": "Designación de apoyo a la supervisión (marzo de 2026).pdf",
}
CARPETA_SOPORTES_IGAC = "Convenio específico IGAC 5570 de 2025"


def aligerar_pptx(src: Path, dst: Path) -> None:
    """Copia una presentación dejando los videos incrustados como imagen fija (el archivo original pesa más de 190 MB)."""
    from lxml import etree
    from pptx import Presentation
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main", "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
          "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main"}
    xp = lambda el, e: etree._Element.xpath(el, e, namespaces=ns)  # noqa: E731
    prs = Presentation(src)
    for sl in prs.slides:
        for vf in xp(sl._element, ".//a:videoFile"):
            vf.getparent().remove(vf)
        for ext in xp(sl._element, ".//p:ext[p14:media]"):
            ext.getparent().remove(ext)
        for rid, rel in list(sl.part.rels.items()):
            if rel.reltype.endswith(("/video", "/media", "/audio")):
                sl.part.drop_rel(rid)
    dst.parent.mkdir(parents=True, exist_ok=True)
    prs.save(dst)


def informe_grupos_corregido(src: Path, dst: Path) -> None:
    """Informe de impacto social por grupo; «LASER LAMIC» no es un grupo (confusión de LASER): su sección se integra a LASER."""
    import docx
    d = docx.Document(src)
    ps = list(d.paragraphs)
    idx = {p.text.strip(): i for i, p in enumerate(ps) if p.style.name.startswith("Heading")}
    h_laser, h_mix = idx.get("Grupo de investigación: LASER"), idx.get("Grupo de investigación: LASER LAMIC")
    if h_laser is not None and h_mix is not None:
        fin_mix = next((i for i in range(h_mix + 1, len(ps)) if ps[i].style.name.startswith("Heading")), len(ps))
        cuerpo_mix = ps[h_mix + 1:fin_mix]
        fin_laser = next(i for i in range(h_laser + 1, len(ps)) if ps[i].style.name.startswith("Heading"))
        primero = ps[h_laser + 1]
        for r in primero.runs:
            r.text = r.text.replace("seis trabajos de grado, cinco sustentados", "siete trabajos de grado, seis sustentados")
        if cuerpo_mix and cuerpo_mix[0].runs:
            r0 = cuerpo_mix[0].runs[0]
            r0.text = "Adicionalmente, " + r0.text[:1].lower() + r0.text[1:]
        ancla = ps[fin_laser - 1]._p
        for par in cuerpo_mix:
            ancla.addnext(par._p)
            ancla = par._p
        ps[h_mix]._p.getparent().remove(ps[h_mix]._p)
    dst.parent.mkdir(parents=True, exist_ok=True)
    d.save(dst)


def _toks(s: str) -> set[str]:
    return {t for t in re.sub(r"[^a-z0-9 ]", " ", _sin_tildes(s)).split() if len(t) > 3}


def tesis_por_modalidad(todos: list[list], carpeta: Path) -> dict[str, list[tuple[Path, str, list]]]:
    """Cada PDF de la carpeta de tesis se asocia, por su título, a una fila del consolidado: {modalidad: [(pdf, nombre_sin_codigos, fila)]}."""
    import subprocess
    out: dict[str, list] = {"Investigación": [], "Profundización": []}
    usados: Counter = Counter()
    for pdf in sorted(carpeta.glob("*.pdf")):
        tt = _toks(subprocess.run(["pdftotext", "-l", "4", str(pdf), "-"], capture_output=True, text=True).stdout)
        mejor = max((f for f in todos if f[2]), key=lambda f: len(_toks(str(f[2])) & tt) / max(1, len(_toks(str(f[2])))))
        if len(_toks(str(mejor[2])) & tt) / max(1, len(_toks(str(mejor[2])))) < 0.6 or mejor[4] not in out:
            continue
        titulo = re.sub(r"[\\/:*?\"<>|\n]", " ", " ".join(str(mejor[2]).split()))[:90].strip()
        usados[mejor[0]] += 1
        nombre = f"Tesis - {titulo}{' (2)' if usados[mejor[0]] > 1 else ''}.pdf"
        out[mejor[4]].append((pdf, nombre, mejor))
    return out


# ---------------------------------------------------------------------------
# FACTOR 8 — Trabajos de grado vinculados a grupos (consolidado por modalidad)
# ---------------------------------------------------------------------------
def modalidad_por_nombre(estudiantes: str | None) -> tuple[str | None, str]:
    """Modalidad inferida por el nombre de los estudiantes (Cóndor/bases). Solo si todos coinciden y el nombre identifica a una única persona."""
    nombres = [n.strip() for n in re.split(r";|\n| y |,", str(estudiantes or "")) if n.strip()]
    if not nombres:
        return None, "Sin estudiantes"
    mods = set()
    for n in nombres:
        t = norm_tokens(n)
        if len(t) < 3:
            return None, "Nombre incompleto"
        cands = [r for r in ROSTER if t <= norm_tokens(r["nombre"])]
        if not cands or len({r["documento"] or r["codigo"] for r in cands}) != 1:
            return None, "Sin coincidencia única en Cóndor"
        mm = {modalidad_con_plan_anterior(r["codigo"]) for r in cands} - {None}
        if len(mm) != 1:
            return None, "Cóndor/bases sin modalidad"
        mods |= mm
    return (mods.pop(), "Inferida por nombre (Cóndor/bases)") if len(mods) == 1 else (None, "Estudiantes de modalidades distintas")


def clasificar_consolidado(fuente: Path) -> tuple[list[list], list[str]]:
    """Cada fila: [ID, grupo, título, año, modalidad|None, directores, estudiantes, fecha, RIUD, estado, tipo, nota de modalidad, observación]."""
    src = openpyxl.load_workbook(fuente, data_only=True)
    filas = []
    for r in src["Consolidado"].iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        cod = str(r[1]).split(".")[0] if r[1] else None
        modal, obs_mod, tipo = r[5], None, None
        grupo_ = "LASER" if re.sub(r"[^A-Z]", "", str(r[2] or "").upper()) == "LASERLAMIC" else r[2]  # «LASER LAMIC» es una confusión de LASER
        if modal == "Pasantía":
            modal, obs_mod, tipo = "Profundización", "Pasantía (opción de grado de Profundización)", "Pasantía"
        if modal not in ("Investigación", "Profundización"):
            mm, criterio = modalidad_de(cod) if cod else (None, "Sin código")
            if mm:
                modal, obs_mod = mm, f"Modalidad completada por cruce: {criterio.lower()}"
            else:
                mm, criterio = modalidad_por_nombre(r[7])
                if mm:
                    modal, obs_mod = mm, f"{criterio}; verificar"
                else:  # sin registro en Cóndor: corresponde al plan anterior, que se toma como Investigación
                    modal, obs_mod = "Investigación", "Plan anterior (asignada como Investigación); verificar"
        tipo = tipo or ("Trabajo de investigación" if modal == "Investigación" else "Trabajo de profundización" if modal == "Profundización" else None)
        # Se omiten el código estudiantil y la ruta del archivo fuente (contienen códigos); siguen en el consolidado de Original/
        filas.append([r[0], grupo_, r[3], r[4], modal, r[6], r[7], r[8], r[9], r[11], tipo, obs_mod or "", re.sub(r"\b\d{11}\b", "[código]", str(r[12] or ""))])
    return filas, [c.value for c in src["Consolidado"][1]]


def factor8(nombre_mod: str, mod: dict, dest: Path, m: dict) -> list[str]:
    etiqueta = mod["etiqueta"]
    otra = "Profundización" if etiqueta == "Investigación" else "Investigación"
    fdir = factor_dirs(mod["plan"])[8]
    out = []
    act_a = dest / "a. Presentación a nuevos estudiantes de procesos de investigación"
    copiar(fdir / "Directorio Grupos de Inv MCIC.xlsx", act_a)
    out.append(f"{act_a.name}/Directorio Grupos de Inv MCIC.xlsx")
    act_b = dest / "b. Socialización y vinculación de actividades de investigación"
    if etiqueta == "Investigación":
        for f in sorted((fdir / "ANEXOS PONENCIAS").iterdir()):
            copiar(f, act_b / "Anexos" / "Ponencias")
            out.append(f"{act_b.name}/Anexos/Ponencias/{f.name}")

    fuente = fdir / "Consolidado_trabajos_grado_MCIC_2022_2026 (7).xlsx"
    todos, _ = clasificar_consolidado(fuente)
    por_mod = {k: [f for f in todos if f[4] == k] for k in ("Investigación", "Profundización")}
    propios = por_mod[etiqueta]
    sustentado = lambda f: str(f[9]).startswith("Sustentado")  # noqa: E731
    sust = [f for f in propios if sustentado(f)]
    pend = [f for f in propios if not sustentado(f)]
    por_anio = Counter((f[3], sustentado(f)) for f in propios)
    grupos = Counter(f[1] for f in propios)

    def fila_comp(nombre, lista):
        s = sum(1 for f in lista if sustentado(f))
        return [nombre, len(lista), s, len(lista) - s, len({f[1] for f in lista}), sum(1 for f in lista if f[10] == "Pasantía")]
    comparativo = [fila_comp("Investigación", por_mod["Investigación"]), fila_comp("Profundización", por_mod["Profundización"]),
                   ["Total consolidado", len(todos), sum(1 for f in todos if sustentado(f)), sum(1 for f in todos if not sustentado(f)), len({f[1] for f in todos}), sum(1 for f in todos if f[10] == "Pasantía")]]
    todos_grupos = sorted({f[1] for f in todos if f[1]}, key=lambda g: -sum(1 for f in todos if f[1] == g))
    por_grupo = [[g, sum(1 for f in por_mod["Investigación"] if f[1] == g), sum(1 for f in por_mod["Profundización"] if f[1] == g)] for g in todos_grupos]

    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 8 · Trabajos de grado vinculados a grupos de investigación — {etiqueta}",
                 "Versión por modalidad del «Consolidado de trabajos de grado MCIC 2022-2026» (el libro completo del consolidado es el documento de origen). "
                 "Se omiten el código estudiantil y la ruta del archivo fuente.")
    fila = encabezado_factor(ws, 4, 8, m)
    fila = st.kpis(ws, fila, [(f"Trabajos de {etiqueta}", len(propios)), ("Sustentados", len(sust)), ("Pendientes por sustentar", len(pend)), ("Grupos de investigación", len(grupos))])
    fila += 1
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 17
    fila = st.seccion(ws, fila + 2, "Proyectos de investigación y de profundización (comparativo)")
    fila = st.tabla(ws, fila, ["Modalidad", "Trabajos", "Sustentados", "Pendientes por sustentar", "Grupos distintos", "Pasantías"], comparativo, filtro=False, congelar=False)
    fila = st.seccion(ws, fila, f"Por año — {etiqueta}")
    anios = sorted({f[3] for f in propios if f[3]})
    fila = st.tabla(ws, fila, ["Año", "Sustentados", "Pendientes por sustentar", "Total"],
                    [[a, por_anio.get((a, True), 0), por_anio.get((a, False), 0), por_anio.get((a, True), 0) + por_anio.get((a, False), 0)] for a in anios],
                    filtro=False, congelar=False)
    fila = st.seccion(ws, fila, f"Por grupo de investigación — {etiqueta}")
    st.tabla(ws, fila, ["Grupo", "Trabajos de grado"], [[g, n] for g, n in grupos.most_common()], filtro=False, congelar=False)

    ws = st.hoja(wb, "Comparativo por grupo", "Trabajos de grado por grupo de investigación y modalidad",
                 "Permite ver qué grupos acompañan proyectos de investigación y cuáles de profundización.", 4)
    st.tabla(ws, 4, ["Grupo de investigación", "Investigación", "Profundización"], por_grupo, [50, 16, 16], filtro=False)

    encab = ["ID", "Grupo de investigación", "Título del trabajo de grado", "Año", "Modalidad", "Director(es)", "Estudiante(s)", "Fecha / soporte de finalización",
             "Enlace RIUD", "Estado de sustentación", "Tipo de trabajo", "Nota de modalidad", "Observación de cruce"]
    anch = [6, 22, 60, 7, 14, 28, 28, 18, 30, 22, 22, 34, 50]
    ws = st.hoja(wb, "Consolidado", f"Trabajos de grado 2022-2026 — {etiqueta}", None, 13)
    st.tabla(ws, 3, encab, propios, anch)
    ws = st.hoja(wb, "Pendientes por sustentar", f"Casos programados, radicados o avalados sin sustentación — {etiqueta}", None, 13)
    st.tabla(ws, 3, encab, pend, anch)
    act_b = dest / "b. Socialización y vinculación de actividades de investigación"
    act_b.mkdir(parents=True, exist_ok=True)
    nombre = f"F8_Trabajos_de_grado_por_grupo_{nombre_mod}.xlsx"
    wb.save(act_b / nombre)
    out.append(f"{act_b.name}/{nombre}")

    # Insumos de los grupos (presentaciones, informe de impacto, proyectos ejecutados y tesis), tomados de Original/
    rel_a, rel_b = "a. Presentación a nuevos estudiantes de procesos de investigación", "b. Socialización y vinculación de actividades de investigación"
    for g in compartido(mod, 8, rel_a, "*.pptx"):
        nombre_g = g.name.replace("_2026-3", " 2026-3")
        aligerar_pptx(g, act_a / nombre_g)
        out.append(f"{act_a.name}/{nombre_g}")
    for informe in compartido(mod, 8, rel_b, "Informe Grupos*.docx")[:1]:
        informe_grupos_corregido(informe, act_b / "Informe de impacto social de los grupos de investigación 2022-2026.docx")
        out.append(f"{act_b.name}/Informe de impacto social de los grupos de investigación 2022-2026.docx")
    for f in compartido(mod, 8, rel_b + "/Proyectos Ejecutados", "*.pdf"):
        copiar(f, act_b / "Anexos" / "Proyectos ejecutados")
        out.append(f"{act_b.name}/Anexos/Proyectos ejecutados/{f.name}")
    carpeta_tesis = mod["plan"] / factor_dirs(mod["plan"])[8].name / rel_b / "Tesis"
    if carpeta_tesis.exists():
        filas_t = []
        for pdf, nombre_t, fila_t in tesis_por_modalidad(todos, carpeta_tesis)[etiqueta]:
            copiar(pdf, act_b / "Anexos" / "Tesis", nombre_t)
            out.append(f"{act_b.name}/Anexos/Tesis/{nombre_t}")
            filas_t.append([fila_t[1], fila_t[2], fila_t[6], fila_t[5], fila_t[3], fila_t[9], nombre_t])
        if filas_t:
            wb2 = openpyxl.load_workbook(act_b / f"F8_Trabajos_de_grado_por_grupo_{nombre_mod}.xlsx")
            ws = st.hoja(wb2, "Tesis con documento", f"Trabajos de grado con documento completo — {etiqueta}", "Documentos disponibles en la carpeta «Anexos / Tesis».", 7)
            st.tabla(ws, 4, ["Grupo de investigación", "Título", "Estudiante(s)", "Director(es)", "Año", "Estado de sustentación", "Documento (anexo)"], filas_t, [26, 70, 30, 30, 7, 20, 60])
            wb2.save(act_b / f"F8_Trabajos_de_grado_por_grupo_{nombre_mod}.xlsx")
    return out


# ---------------------------------------------------------------------------
# FACTOR 9 — Bienestar
# ---------------------------------------------------------------------------
def bienestar_pptx() -> list[list]:
    p = ROOT / "SolicitudesPares/DIA 2/PRESENTACIONES/Bienestar MCIC.pptx"
    x = zipfile.ZipFile(p).read("ppt/slides/slide7.xml").decode("utf8")
    toks = [t.strip() for t in re.findall(r"<a:t>([^<]*)</a:t>", x) if t.strip()]
    filas, i = [], 0
    servicios = ["MEDICINA", "ENFERMERIA", "ODONTOLOGÍA", "PSICOLOGÍA", "FISIOTERAPIA"]
    for s in servicios:
        i = toks.index(s, i)
        j = toks.index("N° ESTUDIANTES", i)
        vals = [int(v) for v in toks[j + 1:j + 8]]
        filas.append([s.capitalize()] + vals + [sum(vals)])
        i = j
    return filas


def factor9(nombre_mod: str, mod: dict, dest: Path, tr: dict, m: dict) -> list[str]:
    etiqueta = mod["etiqueta"]
    fdir = factor_dirs(mod["plan"])[9]
    out = []
    act_a = dest / "a. Solicitud a Bienestar de las actividades ofrecidas"
    act_b = dest / "b. Divulgación de los servicios ofrecidos por Bienestar"
    copiar(ROOT / "SolicitudesPares/DIA 2/PRESENTACIONES/Bienestar MCIC.pptx", act_a)
    out.append(f"{act_a.name}/Bienestar MCIC.pptx")
    copiar(fdir / "acuerdo_02_2019_beca_ecaes.pdf", act_b)
    out.append(f"{act_b.name}/acuerdo_02_2019_beca_ecaes.pdf")
    if etiqueta == "Investigación":
        copiar(fdir / "res_2025-143 para becas de la ODI.pdf", act_b)
        out.append(f"{act_b.name}/res_2025-143 para becas de la ODI.pdf")

    ev = tr["eventos"]
    ind26 = contar_por_modalidad([tuple(x) for x in ev["induccion_2026_1_lista"]["asistentes"]])
    admit = sum(1 for r in ROSTER if r["codigo"].startswith("20261" + mod["proyecto"]))
    matric = sum(1 for r in ROSTER if r["estado"] == "Matriculado" and modalidad_con_plan_anterior(r["codigo"]) == etiqueta)

    divulgacion = [
        ["31/01/2025", "Comunicación de inicio de clases e inducción 2025-1", "Correo a estudiantes", "Investigación", "Correo de inicio de clases (reservado)"],
        ["02/02/2026", f"Inducción 2026-1: {ind26.get(etiqueta, 0)} estudiantes de {etiqueta} asistieron (cohorte 2026-1 admitida: {admit})", "Inducción", "Ambas", "Planilla de asistencia (reservada)"],
        ["—", "Encuentro con estudiantes – evaluación docente (registro fotográfico)", "Encuentro", "Ambas", "Presentacion/…/FACTOR 2/…/Anexos/Encuentro estudiantes evaluacion docente.jpg"],
        ["23-24/09/2026", "Presentación «Bienestar Universitario y Buen Vivir» a los pares (portafolio de servicios y uso por estudiantes MCIC)", "Presentación institucional", "Ambas", "a. Solicitud a Bienestar…/Bienestar MCIC.pptx"],
    ]
    divulgacion = [d for d in divulgacion if d[3] in ("Ambas", etiqueta)]
    estimulos = [["Acuerdo 02 de 2019 (Consejo Académico)", "Incentivo de matrícula en posgrado para egresados de pregrado de la UD con los mejores resultados en Saber Pro.", "Ambas modalidades", "acuerdo_02_2019_beca_ecaes.pdf"]]
    if etiqueta == "Investigación":
        estimulos.append(["Resolución 143 de 2025 (Rectoría)", "Procedimiento para asignar estímulos económicos de sostenimiento mensual del Programa de Excelencia Académica (Acuerdo 03 de 2024), dirigido a maestrías de investigación y doctorados.", "Solo Investigación", "res_2025-143 para becas de la ODI.pdf"])
    servicios = bienestar_pptx()

    # El Cuadro Maestro No. 10 es idéntico en SNIES 17528 y 116070 (Silver guarda una sola copia)
    bien = json.loads((SILVER / "cuadros_maestros.json").read_text(encoding="utf-8")).get("bienestar")
    cm_filas = []
    if bien:
        for s in bien["servicios"]:
            vals = [(s["por_semestre"][p] or {}).get("estudiantes_atendidos") for p in bien["semestres"]]
            cm_filas.append([frase(s["servicio"])] + [v if v is not None else "—" for v in vals])

    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 9 · Bienestar de la comunidad académica — {etiqueta}",
                 "Divulgación de servicios y estímulos de Bienestar a los estudiantes y uso de los servicios por estudiantes MCIC.")
    fila = encabezado_factor(ws, 4, 9, m)
    fila = st.kpis(ws, fila, [(f"Matriculados {etiqueta} (Cóndor)", matric), ("Asistieron a la inducción 2026-1", ind26.get(etiqueta, 0)),
                              ("Estímulos divulgados", len(estimulos)), ("Servicios de salud con registro", len(servicios))])
    st.nota(ws, fila + 1, "Las estadísticas de Bienestar (Cuadro Maestro CNA No. 10 y presentación de Bienestar) se reportan para la Maestría en conjunto: son idénticas en SNIES 17528 y SNIES 116070 y no permiten separar por modalidad.")
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 17

    ws = st.hoja(wb, "Divulgación", "Divulgación de servicios de Bienestar a estudiantes", None, 5)
    st.tabla(ws, 3, ["Fecha", "Actividad", "Tipo", "Alcance", "Soporte"], divulgacion, [14, 70, 22, 14, 60])
    ws = st.hoja(wb, "Estímulos y becas", "Estímulos y becas divulgados", None, 4)
    st.tabla(ws, 3, ["Norma", "Contenido", "Aplica a", "Soporte (carpeta b.)"], estimulos, [36, 80, 18, 36])
    ws = st.hoja(wb, "Uso de servicios", "Estudiantes MCIC atendidos por servicio de Bienestar (2020 – 2026-1)", "Fuente: presentación «Bienestar MCIC» entregada a los pares (Día 2).", 9)
    st.tabla(ws, 4, ["Servicio", "2020", "2021", "2022", "2023", "2024", "2025", "2026-1", "Total"], servicios, [18, 8, 8, 8, 8, 8, 8, 8, 8], filtro=False)
    if cm_filas:
        ws = st.hoja(wb, "Cuadro Maestro No. 10", "Estudiantes atendidos por semestre (Cuadro Maestro CNA No. 10)", f"Fuente: Cuadros maestros SNIES {mod['snies']}.", 11)
        st.tabla(ws, 4, ["Servicio"] + bien["semestres"], cm_filas, [36] + [8] * len(bien["semestres"]), filtro=False)

    act_b.mkdir(parents=True, exist_ok=True)
    nombre = f"F9_Bienestar_{nombre_mod}.xlsx"
    wb.save(act_b / nombre)
    out.append(f"{act_b.name}/{nombre}")
    return out


# ---------------------------------------------------------------------------
# FACTOR 7 — normas institucionales + enlace a URELINTER (convenios vigentes)
# ---------------------------------------------------------------------------
URL_URELINTER = "https://urelinter.udistrital.edu.co/convenios/cooperacion-redes-asociaciones"


def factor7(nombre_mod: str, mod: dict, dest: Path, m: dict) -> list[str]:
    out = sin_cambios(7, mod, dest)
    conv = json.loads((SILVER / "convenios.json").read_text(encoding="utf-8"))
    wb = st.nuevo_libro()
    ws = st.hoja(wb, "URELINTER", f"Factor 7 · Convenios vigentes — {mod['programa']}",
                 "Los convenios los administra la Unidad de Relaciones Interinstitucionales (URELINTER). Se deja el enlace oficial mientras se recibe la información de los profesores sobre su participación en convenios.", 6)
    fila = encabezado_factor(ws, 4, 7, m)
    fila = st.seccion(ws, fila, "Enlace oficial", 6)
    ws.cell(row=fila, column=1, value="Convenios vigentes de la Universidad Distrital (URELINTER)").font = st.FONT_CELDA
    c = ws.cell(row=fila, column=3, value=URL_URELINTER)
    c.hyperlink, c.font = URL_URELINTER, st.FONT_LINK
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=2)
    ws.merge_cells(start_row=fila, start_column=3, end_row=fila, end_column=6)
    fila += 2
    fila = st.kpis(ws, fila, [("Convenios institucionales vigentes", conv["convenios_totales"]), ("Internacionales", conv["por_nivel"]["Internacional"]),
                              ("Nacionales", conv["por_nivel"]["Nacional"]), ("Aplican a la MCIC", conv["total_convenios_aplicables_mcic"]), ("Complementan a la MCIC", conv.get("total_convenios_complementarios_mcic", 0))], por_fila=4)
    fila = st.nota(ws, fila + 1, "El listado de URELINTER es institucional (toda la Universidad) y no se puede filtrar por programa. "
                   f"Fuente: {conv['fuente']['url']} · descargado el {fecha_es(conv['fuente']['fecha_descarga'])}. {conv['fuente']['nota']}", 6)
    for col in "ABCDEF":
        ws.column_dimensions[col].width = 22
    ws = st.hoja(wb, "Por tipo", "Convenios vigentes por tipo", f"Total institucional: {conv['convenios_totales']}.", 2)
    st.tabla(ws, 4, ["Tipo de convenio", "Convenios"], [[k, v] for k, v in conv["por_tipo"].items()], [34, 14], filtro=False)
    if conv.get("convenios"):
        ws = st.hoja(wb, "Convenios vigentes", "Convenios que aplican o complementan a la MCIC", f"De los {conv['convenios_totales']} convenios vigentes de URELINTER (descargado el {fecha_es(conv['fuente']['fecha_descarga'])}).", 9)
        etiqueta_rel = {"aplica_posgrado": "Aplica a la MCIC", "complementa": "Complementa a la MCIC", "ingenieria": "Relacionado con Ingeniería", None: ""}
        st.tabla(ws, 4, ["Relación", "Código", "Nivel", "País / categoría", "Institución", "Tipo", "Denominación", "Vigencia hasta", "Objeto"],
                 [[etiqueta_rel[c["relacion"]], c["codigo"], c["nivel"], c["pais_categoria"], c["institucion"], c["tipo"], c["denominacion"],
                   f"{c['fecha_fin']} ({c['estado']})", " ".join(str(c["objeto"]).split())] for c in conv["convenios"]],
                 [22, 12, 14, 18, 46, 18, 20, 26, 90])
    act = dest / "a. Diagnostico de convenios vigentes"
    carpeta_igac = next((d for d in (carpeta_factor(mod, 7) / "a. Diagnostico de convenios vigentes").glob("CONVENIO IGAC*") if d.is_dir()), None) if (carpeta_factor(mod, 7) / "a. Diagnostico de convenios vigentes").exists() else None
    soportes_igac = []
    if carpeta_igac:
        for f in sorted(carpeta_igac.glob("*.pdf")):
            nombre_pub = next((v for k, v in SOPORTES_CONVENIO_IGAC.items() if unicodedata.normalize("NFC", f.name).startswith(unicodedata.normalize("NFC", k))), None)
            if nombre_pub:
                redactar_documentos_identidad(f, act / "Anexos" / CARPETA_SOPORTES_IGAC / nombre_pub)
                out.append(f"{act.name}/Anexos/{CARPETA_SOPORTES_IGAC}/{nombre_pub}")
                soportes_igac.append(nombre_pub)
    if soportes_igac:
        ws = st.hoja(wb, "Convenio IGAC", "Convenio específico de cooperación N.° 5570 de 2025 — Instituto Geográfico Agustín Codazzi (IGAC) y Universidad Distrital",
                     "Aunar esfuerzos académicos y administrativos para un programa de fortalecimiento en geografía, geomática, catastro, inteligencia artificial y ciencia de datos, mediante proyectos de maestría y doctorado. "
                     "Nació de las conversaciones del IGAC con la coordinación de la MCIC; supervisa por la UD la Oficina de Investigaciones con apoyo de docentes de la Maestría.", 3)
        st.tabla(ws, 4, ["Dato", "Detalle", ""], [["Vigencia", "24 meses desde el acta de inicio del 11 de noviembre de 2025 (hasta el 7 de noviembre de 2027); estado: en ejecución (SECOP II)"],
                                                   ["Aportes", "IGAC $50.000.000 en dinero; Universidad Distrital $36.480.000 en especie (total $86.480.000)"],
                                                   ["Marco", "Convenio marco IGAC-UDFJC C-2023-1 (cooperación en investigación y programas de formación)"]], [18, 120, 4], filtro=False, congelar=False)
        st.tabla(ws, 11, ["Soporte (carpeta Anexos)", "", ""], [[x] for x in soportes_igac], [90, 4, 4], filtro=False, congelar=False)

    act.mkdir(parents=True, exist_ok=True)
    nombre = f"F7_Convenios_URELINTER_{nombre_mod}.xlsx"
    wb.save(act / nombre)
    return out + [f"{act.name}/{nombre}"]


# ---------------------------------------------------------------------------
# FACTOR 10 — Biblioteca, Planes TIC y ambientes de aprendizaje
# ---------------------------------------------------------------------------
BIBLIOTECA_CIFRAS = [  # (indicador, Universidad (todas las bibliotecas), Biblioteca Facultad de Ingeniería)
    ("Área (m²)", "5.489", "261"),
    ("Puestos de lectura", "1.426", "72"),
    ("Colecciones", "375.000+ (210.000 libros · 91 mil documentos en RIUD · 45 mil audiovisuales · 30 mil revistas · 47 bases de datos suscritas)", "24.000+ (14 mil libros · 500 audiovisuales · 10 mil trabajos de grado)"),
    ("Tecnología", "2.370 equipos (497 portátiles · 1.675 tablets · 62 de funcionarios · 29 catálogos · 9 carteleras digitales · 14 TV · 84 audífonos)", "321 equipos (43 portátiles · 272 tablets · 1 catálogo · 1 cartelera digital · 2 TV · 2 audífonos)"),
    ("Espacios", "88 (16 salas de lectura · 6 auditorios · 1 teatro · 9 salas grupales · 5 expositivas · 1 multifuncional · 2 de informática · 4 multimedia · 2 maker spaces · 21 mediatecas · 2 de capacitación)", "2 (1 sala de lectura · 1 sala grupal)"),
]


def datos_biblioteca(xlsx: Path) -> tuple[list[str], list[list]]:
    ws = openpyxl.load_workbook(xlsx, data_only=True)["A. Uso_Servicios CRAI+desglose"]
    filas, anios = [], []
    for r in ws.iter_rows(values_only=True):
        r = list(r[1:])  # la tabla empieza en la columna B
        if r and r[0] == "SERVICIOS CRAI+":
            anios = [str(x) for x in r[1:10]]
        elif anios and r[0] and (str(r[0]).startswith("Servicios CRAI+") or r[0] == "TOTAL"):
            filas.append([" ".join(str(r[0]).split())] + [int(v or 0) for v in r[1:10]] + [int(r[10] or 0)])
            if r[0] == "TOTAL":
                break
    return anios, filas


def factor10(nombre_mod: str, mod: dict, dest: Path, m: dict) -> list[str]:
    fdir = factor_dirs(mod["plan"])[10]
    inv_dir = factor_dirs(MODALIDADES["Investigacion"]["plan"])[10]
    out = []
    a = dest / "a. Diagnóstico sobre ambientes de aprendizaje"
    b = dest / "b. Consolidación de propuesta de medios educativos"
    copiar(fdir / "Ambientes de Aprendizaje Facultad Ing.png", a)
    out.append(f"{a.name}/Ambientes de Aprendizaje Facultad Ing.png")
    for f in ("02 Biblioteca - Versión acreditación 2026 - Posgrado (1).pptx", "Presentación PlanesTIC MCIC.pptx"):
        copiar(fdir / f, b)
        out.append(f"{b.name}/{f}")
    anexo = "Anexo 1 Biblioteca Estadísticas Ingenieria_Maestria Ciencias.xlsx"
    copiar(inv_dir / anexo, b)  # estadística institucional: se entrega en ambas modalidades
    out.append(f"{b.name}/{anexo}")
    anios, uso = datos_biblioteca(inv_dir / anexo)

    from pptx import Presentation
    tic = Presentation(fdir / "Presentación PlanesTIC MCIC.pptx")
    hitos = sorted((sh.left, sh.text_frame.text) for sh in tic.slides[14].shapes if sh.has_text_frame and sh.name.startswith("CuadroTexto 2") and "TRAYECTORIA" not in sh.text_frame.text)
    anios_h = ["2017", "2018", "2020", "2021", "2022", "2024", "2025", "2026"]
    trayectoria = [[a_, "; ".join(x.strip(" .") for x in t.split("\n") if x.strip())] for a_, (_, t) in zip(anios_h, hitos)]

    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 10 · Medios educativos y ambientes de aprendizaje — {mod['programa']}",
                 "Propuesta de medios educativos articulada con la Unidad de Biblioteca y el Comité de Planes TIC (PlanEsTIC). Información institucional, común a las dos modalidades.")
    fila = encabezado_factor(ws, 4, 10, m)
    fila = st.kpis(ws, fila, [("Colección de la Biblioteca de Ingeniería", "24.000+"), ("Puestos de lectura (Ingeniería)", 72), ("Bases de datos suscritas (UD)", 47),
                              ("Casos de soporte virtual en posgrados", 608), ("…de ellos, de la Maestría desde 2024", 77), ("Servicios CRAI+ usados 2025", f"{uso[-1][8]:,}".replace(",", "."))])
    st.nota(ws, fila + 1, "Las cifras de Biblioteca y PlanEsTIC son institucionales: no se pueden separar por modalidad de la Maestría (el anexo estadístico es de la Facultad de Ingeniería y la Maestría).")
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 18

    ws = st.hoja(wb, "Biblioteca - cifras", "Unidad de Biblioteca: infraestructura, colecciones y tecnología", "Fuente: presentación «02 Biblioteca – Versión acreditación 2026 – Posgrado» (Unidad de Biblioteca, 30 de agosto de 2024).", 3)
    st.tabla(ws, 4, ["Indicador", "Sistema de Bibliotecas UD", "Biblioteca Vytautas Gabriunas – Facultad de Ingeniería"], [list(x) for x in BIBLIOTECA_CIFRAS], [26, 70, 60], filtro=False)
    ws = st.hoja(wb, "Biblioteca - uso CRAI+", "Uso de los servicios CRAI+ por año (Sistema de Bibliotecas)", "Fuente: «Anexo 1 Biblioteca – Estadísticas» (datos del Sistema de Bibliotecas al 30 de junio de 2026). 2026 es parcial.", 12)
    st.tabla(ws, 4, ["Servicio CRAI+"] + anios + ["Total"], uso, [64] + [11] * 9 + [12], filtro=False)
    for row in ws.iter_rows(min_row=5, min_col=2, max_col=11):
        for c in row:
            c.number_format = "#,##0"
    ws = st.hoja(wb, "PlanEsTIC", "Maestría en articulación con PlanEsTIC (educación virtual)", "Fuente: presentación «Maestría en Ciencias de la Información y las Comunicaciones en articulación con PlanEsTIC».", 2)
    f2 = st.tabla(ws, 4, ["Componente", "Contenido"], [
        ["Estructura virtual", "Equipo virtual con roles: decano / coordinador de programa y de PlanEsTIC, autor de contenido y asesor pedagógico, profesional de servicio y soporte virtual."],
        ["Diseño pedagógico", "Modelo ADDIE para el diseño instruccional de los cursos de la modalidad virtual (calidad de los recursos y mejora continua)."],
        ["Gestión académica", "Gestión académica con facultades y coordinaciones; prevención de la deserción; atención a docentes y estudiantes; calidad de los cursos; seguimiento de la acción tutorial."],
        ["Gestión tecnológica", "Plataforma de aprendizaje (Moodle institucional), herramientas web, software y hardware para producción de contenidos; encuentros sincrónicos por Teams institucional."],
        ["Soporte técnico virtual", "608 casos atendidos en posgrados; 77 casos de la Maestría desde 2024, además de la atención por chat y correo."],
        ["Ambiente virtual de aprendizaje", "Aulas virtuales creadas por solicitud, alistadas desde un esqueleto de curso, con matrícula de usuarios, capacitación y lanzamiento."],
    ], [30, 120], filtro=False, congelar=False)
    f2 = st.seccion(ws, f2, "Trayectoria del programa (modalidad presencial con apoyo de la virtualidad) — espacios académicos por hito", 2)
    st.tabla(ws, f2, ["Año del hito", "Espacios académicos"], trayectoria, [30, 120], filtro=False, congelar=False)

    b.mkdir(parents=True, exist_ok=True)
    nombre = f"F10_Medios_Educativos_{nombre_mod}.xlsx"
    wb.save(b / nombre)
    out.append(f"{b.name}/{nombre}")
    return out


# ---------------------------------------------------------------------------
# FACTOR 11 — Autoevaluación 2025 (anterior) y 2026 diferenciadas
# ---------------------------------------------------------------------------
def resultados_autoevaluacion_2026(path: Path) -> list[list]:
    """Por instrumento y factor: respuestas máximas y promedio (1-5) de las preguntas de escala."""
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for hoja, etiqueta in (("ESTUDIANTES", "Estudiantes"), ("DOCENTES", "Docentes"), ("DIRECTIVOS", "Directivos")):
        factor, n_max, acum = None, 0, defaultdict(lambda: [0, 0.0, 0])  # factor -> [n_pesos, suma, n_preguntas]
        n_por_factor = defaultdict(int)
        filas = [r for r in wb[hoja].iter_rows(values_only=True)]
        i = 0
        while i < len(filas):
            celdas = [c for c in filas[i] if c is not None]
            txt = str(celdas[0]) if celdas else ""
            mf = re.match(r"(FACTOR \d+)\.", txt)
            if mf:
                factor = mf.group(1)
            if txt.startswith("Personas que han contestado") and factor:
                n = int(celdas[1])
                j, vals = i + 3, []
                while j < len(filas):
                    c2 = [c for c in filas[j] if c is not None]
                    if len(c2) >= 2 and re.fullmatch(r"[1-5]", str(c2[0]).strip()):
                        vals.append((int(str(c2[0]).strip()), int(c2[1])))
                        j += 1
                    else:
                        break
                if vals and sum(v for _, v in vals) == n:
                    acum[factor][0] += n
                    acum[factor][1] += sum(k * v for k, v in vals)
                    acum[factor][2] += 1
                    n_por_factor[factor] = max(n_por_factor[factor], n)
                i = j
                continue
            i += 1
        for f, (nn, suma, nq) in sorted(acum.items(), key=lambda kv: int(kv[0].split()[1])):
            out.append([etiqueta, f, n_por_factor[f], nq, round(suma / nn, 2)])
    return out


def contar_respuestas_2025(carpeta: Path) -> list[list]:
    cuentas = []
    for etiqueta, patron in (("Estudiantes", "Proceso de Autoevaluación 2025 - Estudiantes.xlsx"), ("Docentes", "Proceso de Autoevaluación 2025 - Docentes.xlsx"),
                             ("Egresados", "Proceso de Autoevaluación 2025 - Egresados.xlsx")):
        f = next((p for p in carpeta.rglob("*.xlsx") if unicodedata.normalize("NFC", p.name).lower() == unicodedata.normalize("NFC", patron).lower()), None)
        if f:
            ws = openpyxl.load_workbook(f, read_only=True, data_only=True).worksheets[0]
            cuentas.append([etiqueta, sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0])])
    return cuentas


def factor11(nombre_mod: str, mod: dict, dest: Path, m: dict) -> list[str]:
    etiqueta = mod["etiqueta"]
    sigla = "INV" if etiqueta == "Investigación" else "PROF"
    fdir = factor_dirs(mod["plan"])[11]
    ant = next(p for p in fdir.iterdir() if p.is_dir() and unicodedata.normalize("NFC", p.name).lower().startswith("autoevaluación anterior"))
    out = []
    c26 = dest / "c. Implementación de los lineamientos del SIAC" / "Autoevaluación 2026 (vigente)"
    c25 = dest / "c. Implementación de los lineamientos del SIAC" / "Autoevaluación 2025 (anterior)"
    d26 = dest / "d. Construcción de reportes" / "Autoevaluación 2026 (vigente)"
    d25 = dest / "d. Construcción de reportes" / "Autoevaluación 2025 (anterior)"
    rel = lambda p: str(p.relative_to(dest))  # noqa: E731
    for f in (f"CC-FR-001 Plan de mejoramiento {sigla}.xlsx", f"MCIC AutoevaluacionPermanenteInstitucional {sigla}.pdf"):
        copiar(fdir / f, c26)
        out.append(rel(c26 / f))
    for f in ("ANÁLISIS GENERAL DE RESULTADOS APLICACIÓN INSTRUMENTOS MCIC.pdf",
              "Resultados Proceso Autoevaluación 2026-1 Maestría en Ciencias de la Información y las Comunicaciones.xlsx"):
        copiar(fdir / f, d26)
        out.append(rel(d26 / f))
    informe = next(ant.glob("Informe autoev*.pdf"))
    copiar(informe, c25)
    out.append(rel(c25 / informe.name))
    enc = ant / "Encuestas"
    for grupo in ("Estudiantes", "Docentes", "Egresados"):
        for png in sorted((enc / grupo).rglob("*.png")):
            copiar(png, d25 / "Anexos" / "Gráficos de las encuestas" / grupo)
            out.append(rel(d25 / "Anexos" / "Gráficos de las encuestas" / grupo / png.name))
    for pdf in sorted((enc / "Soportes").glob("*.pdf")):
        copiar(pdf, d25 / "Anexos" / "Invitaciones a las encuestas")
        out.append(rel(d25 / "Anexos" / "Invitaciones a las encuestas" / pdf.name))

    res26 = resultados_autoevaluacion_2026(fdir / "Resultados Proceso Autoevaluación 2026-1 Maestría en Ciencias de la Información y las Comunicaciones.xlsx")
    n26 = {}
    for inst, _f, n, _q, _p in res26:
        n26[inst] = max(n26.get(inst, 0), n)
    n25 = dict(contar_respuestas_2025(enc))
    if not n25:  # las encuestas 2025 se aplicaron a toda la Maestría; las bases de respuestas están en la carpeta de Investigación
        inv_f11 = factor_dirs(MODALIDADES["Investigacion"]["plan"])[11]
        n25 = dict(contar_respuestas_2025(next(p for p in inv_f11.iterdir() if p.is_dir() and unicodedata.normalize("NFC", p.name).lower().startswith("autoevaluación anterior")) / "Encuestas"))
    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 11 · Autoevaluación 2025 (anterior) y 2026 (vigente) — {mod['programa']}",
                 "Los dos procesos están en carpetas separadas, con el año en el nombre: «Autoevaluación 2025 (anterior)» y «Autoevaluación 2026 (vigente)».")
    fila = encabezado_factor(ws, 4, 11, m)
    fila = st.kpis(ws, fila, [("Estudiantes que respondieron 2025", n25.get("Estudiantes", "—")), ("Estudiantes que respondieron 2026-1", n26.get("Estudiantes", "—")),
                              ("Docentes que respondieron 2025", n25.get("Docentes", "—")), ("Docentes que respondieron 2026-1", n26.get("Docentes", "—")),
                              ("Egresados que respondieron 2025", n25.get("Egresados", "—")), ("Directivos que respondieron 2026-1", n26.get("Directivos", "—"))])
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 18
    ws = st.hoja(wb, "2025 vs 2026", "Diferencias entre la autoevaluación 2025 y la de 2026", None, 4)
    st.tabla(ws, 3, ["Aspecto", "Autoevaluación 2025 (anterior)", "Autoevaluación 2026 (vigente)"], [
        ["Propósito", "Informe de autoevaluación con fines de renovación de la acreditación en alta calidad (octubre de 2025).", "Proceso de autoevaluación permanente (Sistema Interno de Aseguramiento de la Calidad), 2026."],
        ["Documento principal", f"{informe.name}", f"MCIC AutoevaluacionPermanenteInstitucional {sigla}.pdf (29 páginas)"],
        ["Instrumentos", "Encuestas a estudiantes, docentes, egresados, directivos y empleadores aplicadas en marzo de 2025.", "Instrumentos de apreciación institucionales (Modelo CNA 2020) a estudiantes, docentes y directivos; periodo 2026-1."],
        ["Resultados", "Gráficos por pregunta de cada encuesta (carpeta «Anexos»). Las bases de respuestas contienen datos personales y se resguardan.", "Libro «Resultados Proceso Autoevaluación 2026-1» y análisis general por estamento (elaborado con apoyo de IA, con recomendación de análisis propio)."],
        ["Plan de mejoramiento asociado", "Plan 2024-2026 (formato AA-FR-001), ver Factor 11 de este plan.", f"Plan 2026-2027 (formato CC-FR-001) — «CC-FR-001 Plan de mejoramiento {sigla}.xlsx»."],
    ], [28, 80, 80], filtro=False, congelar=False)
    ws = st.hoja(wb, "Resultados 2026-1", "Promedio por factor de las preguntas de escala 1-5 (2026-1)",
                 "Solo preguntas de selección única con escala 1 a 5; el detalle por pregunta está en el libro de resultados. No hay un cálculo equivalente para 2025 porque sus respuestas individuales contienen datos personales.", 5)
    st.tabla(ws, 4, ["Instrumento", "Factor", "Respuestas (máx.)", "Preguntas de escala", "Promedio (1-5)"], res26, [16, 14, 18, 18, 14], filtro=False)
    ws = st.hoja(wb, "Archivos", "Archivos por autoevaluación", "Rutas relativas a esta carpeta de factor.", 3)
    fil = []
    for r in out:
        fil.append(["2025 (anterior)" if "2025" in r else "2026 (vigente)", r.split("/")[0], "/".join(r.split("/")[1:])])
    st.tabla(ws, 3, ["Autoevaluación", "Actividad", "Archivo"], fil, [18, 52, 100])
    d26.mkdir(parents=True, exist_ok=True)
    nombre = f"F11_Autoevaluacion_2025_vs_2026_{nombre_mod}.xlsx"
    wb.save(d26.parent / nombre)
    out.append(rel(d26.parent / nombre))
    return out


# ---------------------------------------------------------------------------
# FACTOR 12 — Recursos físicos e infraestructura: laboratorios (salas de informática) y software
# ---------------------------------------------------------------------------
def quitar_ultima_diapositiva(src: Path, destino: Path) -> None:
    """Copia la presentación sin su última diapositiva («Estudiantes impactados», que no corresponde a recursos físicos)."""
    from pptx import Presentation
    prs = Presentation(src)
    lst = prs.slides._sldIdLst
    ultimo = list(lst)[-1]
    prs.part.drop_rel(ultimo.rId)
    lst.remove(ultimo)
    destino.parent.mkdir(parents=True, exist_ok=True)
    prs.save(destino)


def datos_laboratorios(src: Path) -> tuple[list[list], list[list]]:
    """(salas, solicitudes) leídos de las tablas de la presentación de laboratorios."""
    from pptx import Presentation
    prs = Presentation(src)
    solicitudes, vistos, salas = [], set(), []
    for s in prs.slides:
        titulo = next((sh.text_frame.text.strip() for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip().lower().startswith("sala de inform")), None)
        tablas = [sh.table for sh in s.shapes if sh.has_table]
        if titulo and tablas:
            cuerpo = next((sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip().startswith("Docente")), "")
            docente = re.search(r"Docentes?:\s*(.*?)\s*(?:\n|$)", cuerpo)
            asig = re.search(r"Asignaturas?:\s*(.*?)\s*Software:", cuerpo.replace("\n", " "))
            soft = [c.text.strip() for r in tablas[0].rows for c in r.cells if c.text.strip()]
            salas.append([titulo.replace("Sala de informática", "").strip(), " ".join(docente.group(1).split()).title() if docente else "—",
                          " ".join(asig.group(1).split()).capitalize() if asig else "—", len(soft), "; ".join(soft)])
        elif tablas:
            for r in tablas[0].rows:
                celdas = [c.text.strip() for c in r.cells]
                if len(celdas) == 3 and re.match(r"20\d\d-I+$", celdas[0]) and tuple(celdas) not in vistos:
                    vistos.add(tuple(celdas))
                    solicitudes.append([celdas[0].replace("-III", "-3").replace("-I", "-1"), " ".join(celdas[1].split()).title(), "; ".join(x.strip() for x in celdas[2].split("\x0b") if x.strip())])
    return salas, solicitudes


def factor12(nombre_mod: str, mod: dict, dest: Path, m: dict) -> list[str]:
    fdir = factor_dirs(mod["plan"])[12]
    out = []
    a = dest / "a. Solicitud de informe de avance de la obra"
    c = dest / "c. Socialización con la comunidad académica de acciones y avances"
    copiar(fdir / "G312-3 Requerimientos equipos tecnológicos y de software.pdf", a)
    out.append(f"{a.name}/G312-3 Requerimientos equipos tecnológicos y de software.pdf")
    planos = next(iter(compartido(mod, 12, ".", "Planos*.pdf")), None)
    if planos:
        copiar(planos, a, "Planos laboratorios nuevo edificio - diseño arquitectónico (agosto 2020).pdf")
        out.append(f"{a.name}/Planos laboratorios nuevo edificio - diseño arquitectónico (agosto 2020).pdf")
    nombre_pptx = "Laboratorios Maestría - MIC.pptx"
    quitar_ultima_diapositiva(fdir / nombre_pptx, c / nombre_pptx)
    out.append(f"{c.name}/{nombre_pptx}")

    salas, solicitudes = datos_laboratorios(fdir / nombre_pptx)
    n_soft = len({x.strip().lower() for s_ in salas for x in s_[4].split(";") if x.strip()})
    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 12 · Recursos físicos y tecnológicos — {mod['programa']}",
                 "Salas de informática de la Facultad de Ingeniería usadas por la Maestría y software requerido por los docentes (Aulas de Software, desde 2025-1). Información institucional, común a las dos modalidades.")
    fila = encabezado_factor(ws, 4, 12, m)
    fila = st.kpis(ws, fila, [("Salas de informática", len(salas)), ("Programas de software distintos", n_soft),
                              ("Solicitudes de software", len(solicitudes)), ("Periodos con solicitudes", len({x[0] for x in solicitudes}))])
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 18
    ws = st.hoja(wb, "Salas de informática", "Salas de informática y software instalado por asignatura", "Fuente: presentación «Laboratorios Facultad de Ingeniería – Software Maestría».", 5)
    st.tabla(ws, 4, ["Sala", "Docente(s)", "Asignatura(s)", "Programas", "Software"], salas, [8, 44, 50, 11, 110], filtro=False)
    ws = st.hoja(wb, "Solicitudes de software", "Software solicitado por los docentes de la Maestría", "Registros de Aulas de Software desde el periodo 2025-1, con denominación estandarizada.", 3)
    st.tabla(ws, 4, ["Periodo", "Docente", "Software o programas solicitados"], solicitudes, [11, 40, 100])
    wb.save(c / f"F12_Laboratorios_y_Software_{nombre_mod}.xlsx")
    out.append(f"{c.name}/F12_Laboratorios_y_Software_{nombre_mod}.xlsx")
    return out


# ---------------------------------------------------------------------------
# FACTOR 7 (normas) — sin cambios de contenido: se ubican en su actividad
# ---------------------------------------------------------------------------
UBICACION_SIN_CAMBIOS = {
    7: lambda f: "d. Definición de acción para la formalización de convenios",
}


def sin_cambios(n: int, mod: dict, dest: Path) -> list[str]:
    fdir = factor_dirs(mod["plan"])[n]
    out = []
    for f in sorted(p for p in fdir.iterdir() if p.is_file()):
        act = UBICACION_SIN_CAMBIOS[n](f)
        copiar(f, dest / act)
        out.append(f"{act}/{f.name}")
    return out


# ---------------------------------------------------------------------------
# Índice por modalidad
# ---------------------------------------------------------------------------
ESTADO = {n: "Disponible" for n in range(1, 13)}
ESTADO.update({1: "Disponible (PEP definitivo en elaboración)", 3: "Disponible (en recolección)", 5: "Disponible (syllabus en PDF en preparación)"})
OBS = {
    1: "PEP de la modalidad, actas de las jornadas con docentes y orientaciones PFA de la Vicerrectoría Académica.",
    2: "Libro con páginas web, actividades y galería; anexos con el registro fotográfico y las capturas.",
    3: "Participación en capacitación y movilidad docente. En recolección: soportes de capacitaciones 2025-2026 reportados por los profesores.",
    4: "Portafolio de servicios, proyecto de acuerdo de política de egresados e informe de caracterización e impacto con la Hoja de Vida de Egresados (OATI), separado por modalidad.",
    5: "Plan de estudios por modalidad (Res. 016 de 2025), syllabus AA-FR-003 y verificación entre modalidades. En preparación: syllabus en PDF.",
    6: "Prórrogas 2026-3, PAGOT, acompañamiento y graduados, con identificadores en lugar de datos personales.",
    7: "Normativa institucional y enlace oficial de convenios vigentes de URELINTER.",
    8: "Consolidado de trabajos de grado por modalidad (proyectos de investigación y de profundización diferenciados) y ponencias (solo Investigación) como anexo.",
    9: "Estímulos (la Res. 143 de 2025 solo aplica a Investigación) y estadísticas de uso de Bienestar.",
    10: "Presentaciones de Biblioteca y Planes TIC, anexo estadístico de Biblioteca y libro resumen (información institucional, común a las dos modalidades).",
    11: "Autoevaluación 2025 y autoevaluación 2026 en carpetas separadas, con libro comparativo.",
    12: "Presentación de laboratorios con los estudiantes activos por énfasis 2022-2026 de la modalidad.",
}


def indice(nombre_mod: str, mod: dict, base: Path, entregado: dict[int, list[str]], m: dict) -> None:
    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Índice", f"Evidencias del Plan de Mejoramiento — {mod['programa']} (SNIES {mod['snies']})",
                 "Evidencias del Plan de Mejoramiento. Periodo 2024-2027: plan anterior 2024-2026 y plan vigente 2026-2027. "
                 "Cada factor se entrega en su carpeta, con los archivos dentro de la actividad a la que corresponden.", 8)
    filas = []
    for n in range(1, 13):
        filas.append([n, m[n]["factor"].split(". ", 1)[1].rstrip("."), m[n]["meta_anterior"], m[n]["meta_vigente"],
                      SOLICITUDES_PARES[n][0], SOLICITUDES_PARES[n][1], ESTADO[n], len(entregado[n]), OBS[n]])
    st.tabla(ws, 4, ["N°", "Factor", "Meta plan anterior (2024-2026)", "Meta plan vigente (2026-2027)", "Solicitud de los pares (sep. 2026)",
                     "Responsable", "Estado", "Archivos", "Observaciones"], filas, [5, 30, 45, 45, 45, 20, 22, 9, 60])
    ws = st.hoja(wb, "Archivos", "Archivos entregados por factor", "Rutas relativas a esta carpeta.", 3)
    rutas = []
    carpetas = {n: d.name for n, d in factor_dirs(mod["plan"]).items()}
    for n in range(1, 13):
        for r in entregado[n]:
            rutas.append([n, carpetas[n], r])
    st.tabla(ws, 3, ["N°", "Carpeta del factor", "Archivo (actividad / nombre)"], rutas, [5, 60, 100])
    ws = st.hoja(wb, "Recomendaciones CNA 2022", "Recomendaciones de los pares académicos / CNA (visita 2022)", None, 3)
    st.tabla(ws, 3, ["N°", "Tema", "Recomendación"], [[i, t, r] for i, (t, r) in enumerate(RECOMENDACIONES_CNA_2022, 1)], [5, 28, 120])
    wb.save(base / f"00_Indice_Evidencias_{nombre_mod}.xlsx")


# ---------------------------------------------------------------------------
def main() -> None:
    tr = json.loads(PRIVADO.read_text(encoding="utf-8"))
    shutil.rmtree(OUT / "Presentacion", ignore_errors=True)  # Original/ es la fuente: no se toca
    copiar(PRIVADO, OUT / "Original/_transcripciones")

    for nombre_mod, mod in MODALIDADES.items():
        base = OUT / "Presentacion" / nombre_mod
        m = metas(mod["clave"])
        carpetas = factor_dirs(mod["plan"])
        d = {n: base / carpetas[n].name for n in range(1, 13)}
        entregado = {
            1: factor1(mod, d[1]),
            2: factor2(nombre_mod, mod, d[2], tr, m),
            3: factor3(nombre_mod, mod, d[3], tr, m),
            4: factor4(nombre_mod, mod, d[4], m),
            5: factor5(nombre_mod, mod, d[5], m),
            6: factor6(nombre_mod, mod, d[6], tr, m),
            7: factor7(nombre_mod, mod, d[7], m),
            8: factor8(nombre_mod, mod, d[8], m),
            9: factor9(nombre_mod, mod, d[9], tr, m),
            10: factor10(nombre_mod, mod, d[10], m),
            11: factor11(nombre_mod, mod, d[11], m),
            12: factor12(nombre_mod, mod, d[12], m),
        }
        indice(nombre_mod, mod, base, entregado, m)
        total = sum(len(v) for v in entregado.values())
        print(f"Presentacion/{nombre_mod}: {total} archivos en 12 factores")


if __name__ == "__main__":
    main()
