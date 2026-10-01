"""Construye NuevaData/ (propuesta de evidencias para aprobación).

    NuevaData/
      Original/<Modalidad>/...        copia fiel de Data/Bronze/<MCIC.*>/.../Plan de Mejoramiento
      Original/_transcripciones/      transcripción manual de planillas (datos personales)
      Presentacion/<Modalidad>/...    evidencia analizada, ordenada por FACTOR / actividad,
                                      sin datos sensibles y separada por modalidad

Uso:  .venv/bin/python NuevaData/_build/build_nuevadata.py

El script es idempotente: borra y regenera Original/ y Presentacion/. Lee
únicamente archivos de Data/Bronze, Data/Silver, SolicitudesPares/ y el
consolidado de trabajos de grado de la raíz; no modifica ninguno de ellos.
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
        "plan": BRONZE / "MCIC.INVESTIGACION/Procesos de Renocavion y acreditación/Plan de Mejoramiento",
        "web": "https://facingenieria.udistrital.edu.co/mcic-investigacion/",
        "captura_web": "Actualizacion pagiona web mcic -  investigacion.png",
    },
    "Profundizacion": {
        "clave": "profundizacion",
        "etiqueta": "Profundización",
        "programa": "MCIC - PROFUNDIZACIÓN",
        "proyecto": "695",
        "snies": "116070",
        "plan": BRONZE / "MCIC-PROFUNDIZACION/Procesos de Renocavion y acreditación/Plan de Mejoramiento",
        "web": "https://facingenieria.udistrital.edu.co/mcic-profundizacion/",
        "captura_web": "Actualizacion pagiona web mcic -  profundizacion.png",
    },
}

ENFASIS_POR_PROYECTO = {"195": "Teleinformática", "295": "Sistemas de Información", "395": "Geomática",
                        "495": "Ingeniería de Software", "95": "Teleinformática (plan antiguo)"}

# Solicitudes de los pares (visita sep 2026) y responsables, tomados de CorreccionEvidencias.ods
SOLICITUDES_PARES = {
    1: ("PEP listo. Documento según la conversación del lunes, con los ajustes y socializado en el Consejo de Carrera.", "Sebastián Vanegas"),
    2: ("Asistencias a las inducciones: generar documento con los datos (quitar ruido); actualizaciones de la página web (dejar el enlace); publicidad de las maestrías por correo; Open Day con lista de asistencia; socialización de grupos de investigación. Se actualiza cada semestre por proceso de admisiones.", "Juan y Karol"),
    3: ("Solicitar a los profesores evidencia de las capacitaciones del periodo 2025-2026 adicionales a la información del cuadro.", "Karol (correo de solicitud)"),
    4: ("Retirar el documento de experiencias y la infografía; solicitar a los egresados 2022-2026 el sector en el que laboran actualmente (formulario: nombre y sector).", "Karol y Juan (tarea enviada a la OATI)"),
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
        c[mod or "Sin modalidad / otro programa"] += 1
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
    sop = act / "soportes"
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
        ("2025-01-31", "2025-1", "Comunicación de inicio de clases y convocatoria a la inducción (7 de febrero de 2025, 6:15 p. m.) con el calendario del semestre", "Comunicación a estudiantes", "Correo institucional", "Coordinación MCIC (cuenta de Investigación)", "Investigación", ind25[0], ind25[1], ind25[2] + ev["induccion_2025_1_correo"]["destinatarios_externos"], None, f"{orig}/Induccion 2025-1.pdf", "Destinatarios en copia oculta."),
        ("2025-02-07", "2025-1", "Inducción de estudiantes nuevos 2025-1 (registro fotográfico)", "Inducción", "Auditorio Sabio Caldas", "Coordinación MCIC", "Ambas", None, None, None, "soportes/Inducción 2025-1.jpg", None, "Sin planilla de asistencia en la evidencia."),
        (None, "2025-1", "Divulgación de grupos de investigación 2025-1 (registro fotográfico)", "Divulgación de grupos de investigación", "Auditorio, Facultad de Ingeniería", "Facultad de Ingeniería", "Ambas", None, None, None, "soportes/Divulgación grupos de investigación 2025-1.jpg", None, "Sin planilla de asistencia en la evidencia."),
        ("2025-08-22", "2025-3", "Invitación a la jornada de divulgación de grupos de investigación (asistencia obligatoria para quienes cursan Seminario de Investigación)", "Comunicación a estudiantes", "Correo institucional", "Coordinación MCIC (cuenta de Investigación)", "Investigación", div25[0], div25[1], div25[2] + ev["divulgacion_grupos_2025_3_correo"]["destinatarios_sin_nombre"], None, f"{orig}/Divulgacio grupos 2025-3.pdf", "Incluye el afiche del evento."),
        ("2025-08-30", "2025-3", "Jornada de divulgación de grupos de investigación de la Facultad (8:00 a. m. a 12:00 m.). Invitan: Maestría en Ingeniería Industrial, MCIC Investigación y Profundización, Maestría en Gerencia Integral de Proyectos y Maestría en Telecomunicaciones Móviles", "Divulgación de grupos de investigación", "Auditorio Sabio Caldas", "Facultad de Ingeniería", "Ambas", None, None, None, "soportes/Divulgación grupos de investigación 2025-3.jpg", None, "Registro fotográfico."),
        (None, "—", "Encuentro con estudiantes – evaluación docente (registro fotográfico)", "Encuentro con estudiantes", "Sala de cómputo, Facultad de Ingeniería", "Coordinación MCIC", "Ambas", None, None, None, "soportes/Encuentro estudiantes evaluacion docente.jpg", None, "El soporte no trae fecha."),
        ("2026-02-02", "2026-1", "Inducción 2026-1 (control de asistencia GD-PR-008-FR-026)", "Inducción", "Facultad de Ingeniería", "Coordinación MCIC", "Ambas", ind26[0], ind26[1], ind26[2], None, f"{orig}/LISTA ASISTENCIA INDUCCIONES 2026-1.pdf", "Planilla con firmas."),
        ("2026-02-21", "2026-1", "Presentación de grupos de investigación – Facultad de Ingeniería 2026-1 (8:00 a. m. a 12:00 m.)", "Divulgación de grupos de investigación", "Auditorio Sabio Caldas", "Facultad de Ingeniería – Posgrados", "Ambas", gru26[0], gru26[1], gru26[2], None, f"{orig}/Presentación de Grupos de Investigación MCIC2026-1.pdf", f"Además firmaron {tr['eventos']['presentacion_grupos_2026_1_lista']['docentes_y_grupos']} docentes y representantes de grupos de investigación."),
        ("2026-04-30", "2026-1", "Difusión del Open Day 3.0 de posgrados UD (evento del 15 de mayo de 2026, 5:00 p. m.)", "Divulgación de la oferta académica", "Correo institucional", "Decanatura Facultad de Ingeniería – Eventos", "Ambas", None, None, None, "soportes/Open Day.pdf", None, "Material publicitario para canales de cada programa de posgrado; incluye enlace de inscripción."),
        ("2026-05-15", "2026-1", "Realización del Open Day 3.0 (registro fotográfico)", "Divulgación de la oferta académica", "Auditorio Sabio Caldas y Muro de Escalar", "Facultad de Ingeniería", "Ambas", None, None, None, "soportes/Realización del OPEN DAY 3.0.jpeg", None, "Sin planilla de asistencia en la evidencia."),
        (None, "2026-3", "Inducción 2026-3 – presentación de grupos de investigación (registro fotográfico, grupo LIDER)", "Inducción", "Auditorio, Facultad de Ingeniería", "Coordinación MCIC", "Ambas", None, None, None, "soportes/Inducción 2026-3.jpeg", None, "Registro fotográfico."),
        ("2026-08-22", "2026-3", "Socialización de grupos de investigación (8:00 a. m. a 12:00 m.)", "Divulgación de grupos de investigación", "Auditorio Sabio Caldas", "Facultad de Ingeniería", "Ambas", soc26[0], soc26[1], soc26[2], None, f"{orig}/Asistencia 22-08-22 2026-3.pdf", f"Planilla compartida con otros programas ({ev['socializacion_grupos_2026_3_lista']['total_registros_planilla']} registros en total)."),
    ]
    eventos = [e for e in eventos if e[6] in ("Ambas", etiqueta)]

    # Cobertura de cohortes nuevas (admitidos según Cóndor)
    def admitidos(prefijo):
        return sum(1 for r in ROSTER if r["codigo"].startswith(prefijo + mod["proyecto"]))

    def asistentes_cohorte(personas, prefijo):
        n = 0
        for nombre, codigo in personas:
            if not codigo:
                r = buscar_en_roster(nombre)
                codigo = r["codigo"] if r else None
            if codigo and codigo.startswith(prefijo + mod["proyecto"]):
                n += 1
        return n

    ind_l = [tuple(x) for x in ev["induccion_2026_1_lista"]["asistentes"]]
    gru_l = [tuple(x) for x in ev["presentacion_grupos_2026_1_lista"]["asistentes"]]
    soc_l = [tuple(x) for x in ev["socializacion_grupos_2026_3_lista"]["asistentes_mcic"]]
    a261, a263 = admitidos("20261"), admitidos("20262")
    cobertura = [
        ["2026-1", a261, "Inducción 2026-1 (02/02/2026)", asistentes_cohorte(ind_l, "20261")],
        ["2026-1", a261, "Presentación de grupos de investigación (21/02/2026)", asistentes_cohorte(gru_l, "20261")],
        ["2026-3", a263, "Socialización de grupos de investigación (22/08/2026)", asistentes_cohorte(soc_l, "20262")],
    ]
    for c in cobertura:
        c.append(f"{(c[3] / c[1] * 100):.0f} %" if c[1] else "—")

    wb = st.nuevo_libro()
    # --- Resumen
    ws = st.hoja(wb, "Resumen", f"Factor 2 · Estudiantes — {mod['programa']}",
                 "Divulgación de la oferta académica, páginas web, inducciones y socialización de grupos de investigación. "
                 "Consolidado a partir de capturas, fotografías, correos y planillas de asistencia; los soportes con datos personales se conservan en Original/.")
    fila = encabezado_factor(ws, 4, 2, m)
    con_lista = [e for e in eventos if e[7] is not None]
    fila = st.kpis(ws, fila, [
        ("Página web de la modalidad", "Actualizada"),
        ("Actividades documentadas", len(eventos)),
        (f"Registros de estudiantes de {etiqueta}", sum(e[7] for e in con_lista)),
        ("Cobertura inducción 2026-1", cobertura[0][4]),
    ])
    fila = st.nota(ws, fila + 1, "Los conteos por modalidad se obtienen cruzando cada nombre/código de las planillas con Cóndor (proyectos 595 = Investigación, 695 = Profundización) y con las bases de datos MCIC. "
                   "Un mismo estudiante puede aparecer en varias actividades; los registros no se suman como personas únicas.")
    for col, w in zip("ABCDEFGH", (16, 16, 16, 16, 16, 16, 16, 16)):
        ws.column_dimensions[col].width = w

    # --- Páginas web
    ws = st.hoja(wb, "Páginas web", "Páginas web de la Maestría", "Enlaces vigentes y capturas de la actualización.", 5)
    webs = [[etiqueta, mod["web"], mod["web"], "Sitio propio de la modalidad", "Captura de la actualización", f"soportes/{mod['captura_web']}"],
            ["Común (ambas modalidades)", "Sitio MCIC con menú para aspirantes de Investigación y Profundización", None, "Banner de inscripciones abiertas y acreditación de alta calidad (Res. 024858 de 2022)", "Captura del sitio común", "soportes/Actualizacion pagiona web mcic.png"]]
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
                     "Soporte original (contiene datos personales)", "Observación"],
             filas, [11, 9, 48, 20, 20, 22, 12, 13, 13, 11, 14, 42, 34], links={10: 11})

    # --- Cobertura
    ws = st.hoja(wb, "Cobertura cohortes", f"Cobertura de las actividades sobre las cohortes nuevas de {etiqueta}",
                 f"Admitidos: estudiantes con código de ingreso del periodo en el proyecto {mod['proyecto']} según Cóndor.", 5)
    st.tabla(ws, 4, ["Cohorte", "Admitidos", "Actividad", "Asistentes de la cohorte", "Cobertura"], cobertura, [12, 12, 52, 22, 12])

    # --- Galería
    ws = st.hoja(wb, "Galería", "Registro fotográfico y capturas", "Miniaturas de los soportes copiados en la carpeta «soportes».", 6)
    fila = 4
    for i, f in enumerate([x for x in imagenes]):
        col = "A" if i % 2 == 0 else "E"
        if i % 2 == 0 and i:
            fila += 18
        ws[f"{col}{fila}"] = f.rsplit(".", 1)[0]
        ws[f"{col}{fila}"].font = st.FONT_SECCION
        ws[f"{col}{fila}"].hyperlink = f"soportes/{f}"
        st.miniatura(ws, f"{col}{fila + 1}", sop / f, 300)
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 12

    nombre = f"F2_Divulgacion_y_Estudiantes_{nombre_mod}.xlsx"
    wb.save(act / nombre)
    return [f"a. Divulgación de propuestas académicas/{nombre}"] + [f"a. Divulgación de propuestas académicas/soportes/{x}" for x in imagenes + ["Open Day.pdf"]]


# ---------------------------------------------------------------------------
# FACTOR 3 — Participación docente en capacitación (planta compartida)
# ---------------------------------------------------------------------------
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
    docentes.sort(key=lambda d: d[1])
    for i, d in enumerate(docentes, 1):
        d[0] = i
    con_cap = sum(1 for d in docentes if d[3] or d[5])

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
                 "Información en actualización: se solicitó a los profesores reportar con soporte las capacitaciones adicionales del periodo 2025-2026 (columna J de la hoja «Capacitación docente»).")
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
                     "Grupo de investigación", "Categoría investigador MinCiencias", "Capacitaciones 2025-2026 reportadas por el profesor (por diligenciar)"],
             docentes, [5, 34, 13, 18, 7, 40, 7, 18, 16, 38])

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
    return [f"a. Informes de participación de los docentes/{nombre}"]


# ---------------------------------------------------------------------------
# FACTOR 4 — Egresados (se retiran Experiencias UD e infografía)
# ---------------------------------------------------------------------------
def factor4(mod: dict, dest: Path) -> list[str]:
    src = factor_dirs(mod["plan"])[4] / "a. Servicios"
    out = []
    for f in ("Portafolio de Servicios Grupo Funcional Egresados.pdf", "PROYECTO DE ACUERDO POLÍTICA DE EGRESADOS.pdf"):
        copiar(src / f, dest / "a. Servicios")
        out.append(f"a. Servicios/{f}")
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
    hallazgos = [
        ["Plan de estudios", "Sí se diferencia", "La Res. 016 de 2025 del Consejo Académico aprueba un plan para cada modalidad (44 créditos cada uno). Investigación: Trabajo de grado I y II (12 créditos), 3 espacios de énfasis (12 créditos) y 2 electivas. Profundización: Trabajo de grado (4 créditos), 5 espacios de énfasis (20 créditos) y 2 electivas."],
        ["Horas de trabajo académico", "Sí se diferencia", "Investigación: 384 HTD / 176 HTC / 1.552 HTA. Profundización: 480 HTD / 160 HTC / 1.472 HTA (2.112 horas en ambos casos)."],
        ["Espacios académicos", "Compartidos", "Los espacios de fundamentación, seminario y énfasis son los mismos en las dos modalidades; cambia cuántos espacios de énfasis cursa el estudiante y si los del periodo III son obligatorios (Profundización) o electivos (Investigación)."],
        ["Archivos de syllabus", "No se diferencian", f"Los {len(snies)} syllabus de la carpeta SNIES 17528 y los de SNIES 116070 son byte a byte idénticos ({iguales}/{len(snies)}). Los {len(filas)} syllabus actualizados (AA-FR-003) son un único archivo por espacio académico, usado por las dos modalidades."],
        ["Identificación de la modalidad en el syllabus", "No se diferencia", identificacion],
        ["Malla «Información Espacios Académicos.xlsx»", "Parcial", "Tiene una hoja por modalidad y diferencia el trabajo de grado (Investigación: TG I y TG II; Profundización: TG). Marca todos los espacios de énfasis como obligatorios en ambas hojas y la hoja de Profundización registra el código de proyecto 595 (el de Profundización es 695)."],
        ["Espacio sin correspondencia en la Res. 016", "Revisar", "«Matemática avanzada y geoprocesamiento» tiene syllabus AA-FR-003, pero no figura en el plan de estudios de la Res. 016 (podría ser electiva)."],
        ["Denominaciones", "Revisar", "La Res. 016 usa «Modelo de simulación y redes», «Bases de datos especiales» y «Procesamiento digital de imágenes y elaboración de productos cartográficos»; los syllabus usan «Modelado y simulación de redes», «Bases de datos espaciales» y «… productos cartográficos o fotogramétricos»."],
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

    nombre = "F5_Verificacion_Syllabus_Investigacion_vs_Profundizacion.xlsx"
    act_b.mkdir(parents=True, exist_ok=True)
    wb.save(act_b / nombre)
    out.append(f"{act_b.name}/{nombre}")
    return out


# ---------------------------------------------------------------------------
# FACTOR 6 — Permanencia y graduación
# ---------------------------------------------------------------------------
def factor6(nombre_mod: str, mod: dict, dest: Path, tr: dict, m: dict) -> list[str]:
    etiqueta = mod["etiqueta"]
    filas, por_confirmar = [], []
    for p in tr["prorrogas"]:
        mmod, criterio = modalidad_de(p["codigo"], declarada=p.get("modalidad_declarada"))
        base = BASES.get(p["codigo"], {})
        tipo = "PAGOT" if (p["pagot"] or base.get("tipo") == "PAGOT") else "No registrado como PAGOT"
        cond = ROSTER_POR_CODIGO.get(p["codigo"], {})
        fila = [p["id"], fecha_es(p["fecha"]), mmod or "Por confirmar", enfasis_de(p["codigo"]) or frase(base.get("enfasis")) or "—", tipo,
                p["solicitud"], p["c1"], p["producto_c1"], p["c2"], p["adjuntos"],
                f"{cond.get('estado', '—')} (última matrícula {cond.get('ultima', '—')})" if cond else "—",
                p.get("nota_modalidad") or criterio]
        if mmod == etiqueta:
            filas.append(fila)
        elif mmod is None:
            por_confirmar.append(fila)
    filas.sort(key=lambda f: f[0])
    c1_ok = sum(1 for f in filas if re.search(r"aceptad|aprobad", f[6], re.I))
    c2_ok = sum(1 for f in filas if re.search(r"radicad|entregad", f[8], re.I))

    # Reunión de acompañamiento (convocatoria del 06/04/2026)
    conv = tr["eventos"]["reunion_acompanamiento_2026_04_06"]["convocados_codigos"]
    conv_mod = Counter(modalidad_de(c)[0] or "Sin registro" for c in conv)

    # PAGOT de la modalidad
    archivo = "MCIC - Base de datos INVESTIGACION.xlsx" if etiqueta == "Investigación" else "MCIC - Base de datos Profundizacion.xlsx"
    hoja = "N-A Investigación" if etiqueta == "Investigación" else "N-A Profundizacion"
    wbb = openpyxl.load_workbook(BRONZE / "Maestria CIC/2026/BASES  DE DATOS ESTUDIANTES" / archivo, read_only=True, data_only=True)
    rows = list(wbb[hoja].iter_rows(values_only=True))
    enc = [str(c).strip() if c else "" for c in rows[0]]
    it, ie, ig = enc.index("TIPO DE ESTUDIANTE"), enc.index("ESTADO"), enc.index("INGRESO")
    ien = next((enc.index(k) for k in ("ÉNFASIS", "ENFASIS") if k in enc), None)
    estados = {}
    for r in openpyxl.load_workbook(BRONZE / "Maestria CIC/2026/BASES  DE DATOS ESTUDIANTES" / archivo, read_only=True, data_only=True)["Listas"].iter_rows(values_only=True):
        if r and len(r) > 2 and r[1] and r[2] and len(str(r[1])) == 1:
            estados[r[1]] = frase(str(r[2]).replace("A?O", "AÑO"))
    pagot = [r for r in rows[1:] if r and r[0] and r[it] == "PAGOT"]
    regulares = [r for r in rows[1:] if r and r[0] and r[it] == "REGULAR"]
    pag_ing = Counter(r[ig] for r in pagot)
    pag_est = Counter(estados.get(r[ie], r[ie]) for r in pagot)
    pag_enf = Counter(enfasis_normalizado(r[ien]) if ien is not None and r[ien] else "Sin dato" for r in pagot)

    # Graduados por año (estimado por última matrícula, mismo criterio del sitio) y modalidad
    grad = Counter()
    for r in ROSTER:
        if r["estado"] == "Graduado" and r["ultima"][:4].isdigit() and 2022 <= int(r["ultima"][:4]) <= 2026:
            mm, _ = modalidad_de(r["codigo"])
            grad[(mm or "Sin modalidad registrada", r["ultima"][:4])] += 1
    anios = ["2022", "2023", "2024", "2025", "2026"]
    grad_filas = []
    for mm in (etiqueta, "Sin modalidad registrada"):
        vals = [grad.get((mm, a), 0) for a in anios]
        grad_filas.append([mm] + vals + [sum(vals)])

    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 6 · Permanencia y graduación — {mod['programa']}",
                 "Estudiantes con pendiente de trabajo de grado y riesgo de pérdida de calidad: solicitudes de prórroga (mayo de 2026), acompañamiento, estudiantes PAGOT y graduados. "
                 "Los estudiantes se identifican con un ID (S-01…); la relación ID–estudiante y los correos originales se conservan en Original/.")
    fila = encabezado_factor(ws, 4, 6, m)
    fila = st.kpis(ws, fila, [("Solicitudes de prórroga", len(filas)), ("Componente 1 aceptado/aprobado", c1_ok),
                              ("Componente 2 radicado/entregado", c2_ok), ("Estudiantes PAGOT (base MCIC)", len(pagot)),
                              ("Convocados a acompañamiento", conv_mod.get(etiqueta, 0)), ("Estudiantes regulares (base MCIC)", len(regulares)),
                              ("Graduados 2022-2026", grad_filas[0][-1]), ("Solicitudes con modalidad por confirmar", len(por_confirmar))])
    st.nota(ws, fila + 1, "Modalidad de cada solicitud: declarada por el estudiante, o según las bases de datos MCIC (hojas N-A Investigación / N-A Profundización) y Cóndor (595/695). "
            "Graduados: año estimado por la última matrícula en Cóndor, el mismo criterio del micrositio; «Sin modalidad registrada» agrupa graduados del plan anterior (proyectos 195–495) que no aparecen en las bases por modalidad.")
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 17

    enc_p = ["ID", "Fecha", "Modalidad", "Énfasis", "Tipo de estudiante", "Solicitud", "Componente 1", "Producto / evento del componente 1",
             "Componente 2 (trabajo de grado)", "Adjuntos", "Estado en Cóndor 2026-3", "Criterio de modalidad"]
    anch = [7, 11, 14, 18, 16, 34, 36, 34, 40, 9, 26, 26]
    ws = st.hoja(wb, "Solicitudes de prórroga", f"Solicitudes de prórroga de permanencia — {etiqueta}", "Radicadas ante la coordinación entre el 19 y el 26 de mayo de 2026 para estudio del Consejo de Facultad.", 12)
    st.tabla(ws, 4, enc_p, filas, anch)
    if por_confirmar:
        ws = st.hoja(wb, "Modalidad por confirmar", "Solicitudes sin modalidad registrada", "No se encontró la modalidad en las fuentes; confirmar antes de publicar.", 12)
        st.tabla(ws, 4, enc_p, por_confirmar, anch)

    ws = st.hoja(wb, "Acompañamiento", "Reunión de acompañamiento académico (06/04/2026, 7:00 p. m., virtual)",
                 "Convocatoria de la coordinación a estudiantes cuyo tiempo de permanencia vence en 2026-1, para orientar rutas de culminación del trabajo de grado.", 3)
    st.tabla(ws, 4, ["Modalidad", "Estudiantes convocados", "Fuente"],
             [[k, v, "Correo de convocatoria (Original/…/FACTOR 6/Reunión de acompañamiento.pdf)"] for k, v in sorted(conv_mod.items())], [24, 22, 70], filtro=False)

    ws = st.hoja(wb, "PAGOT", f"Estudiantes PAGOT — {etiqueta}", f"Fuente: {archivo} (hoja {hoja}), columna «Tipo de estudiante».", 4)
    f2 = st.tabla(ws, 4, ["Periodo de ingreso a PAGOT", "Estudiantes"], [[k, v] for k, v in sorted(pag_ing.items(), key=lambda x: str(x[0]))], [30, 14], filtro=False, congelar=False)
    f2 = st.tabla(ws, f2, ["Estado académico", "Estudiantes"], [[k, v] for k, v in pag_est.most_common()], [30, 14], filtro=False, congelar=False)
    st.tabla(ws, f2, ["Énfasis", "Estudiantes"], [[k, v] for k, v in pag_enf.most_common()], [30, 14], filtro=False, congelar=False)

    ws = st.hoja(wb, "Graduados", f"Graduados 2022-2026 — {etiqueta}", "Año estimado a partir de la última matrícula registrada en Cóndor.", 7)
    st.tabla(ws, 4, ["Modalidad"] + anios + ["Total"], grad_filas, [28, 9, 9, 9, 9, 9, 10], filtro=False)

    act = dest / "d. Seguimiento al avance de trabajos de grado y tiempos de permanencia"
    act.mkdir(parents=True, exist_ok=True)
    nombre = f"F6_Permanencia_y_Graduacion_{nombre_mod}.xlsx"
    wb.save(act / nombre)
    return [f"{act.name}/{nombre}"]


# ---------------------------------------------------------------------------
# FACTOR 8 — Trabajos de grado vinculados a grupos (consolidado por modalidad)
# ---------------------------------------------------------------------------
CONSOLIDADO = ROOT / "Consolidado_trabajos_grado_MCIC_2022_2026 (7).xlsx"


def factor8(nombre_mod: str, mod: dict, dest: Path, m: dict) -> list[str]:
    etiqueta = mod["etiqueta"]
    fdir = factor_dirs(mod["plan"])[8]
    out = []
    act_a = dest / "a. Presentación a nuevos estudiantes de procesos de investigación"
    copiar(fdir / "Directorio Grupos de Inv MCIC.xlsx", act_a)
    out.append(f"{act_a.name}/Directorio Grupos de Inv MCIC.xlsx")
    if etiqueta == "Investigación":
        for f in sorted((fdir / "ANEXOS PONENCIAS").iterdir()):
            copiar(f, dest / "ANEXOS PONENCIAS")
            out.append(f"ANEXOS PONENCIAS/{f.name}")

    src = openpyxl.load_workbook(CONSOLIDADO, data_only=True)
    enc = [c.value for c in src["Consolidado"][1]]
    filas_mod, por_confirmar = [], []
    for r in src["Consolidado"].iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        r = list(r)
        cod = str(r[1]).split(".")[0] if r[1] else None
        modal = r[5]
        obs_mod = None
        if modal == "Pasantía":
            modal, obs_mod = "Profundización", "Pasantía (opción de grado de Profundización)"
        if modal not in ("Investigación", "Profundización"):
            mm, criterio = modalidad_de(cod) if cod else (None, "Sin código")
            if mm:
                modal, obs_mod = mm, f"Modalidad completada por cruce: {criterio.lower()}"
            else:
                modal = None
        # Se omiten el código estudiantil y la ruta del archivo fuente (contiene códigos); ambos siguen en el consolidado de Original/
        fila = [r[0], r[2], r[3], r[4], modal or "Por confirmar", r[6], r[7], r[8], r[9], r[11], obs_mod or "", re.sub(r"\b\d{11}\b", "[código]", str(r[12] or ""))]
        if modal == etiqueta:
            filas_mod.append(fila)
        elif modal is None:
            por_confirmar.append(fila)

    sust = [f for f in filas_mod if str(f[9]).startswith("Sustentado")]
    pend = [f for f in filas_mod if not str(f[9]).startswith("Sustentado")]
    por_anio = Counter((f[3], str(f[9]).startswith("Sustentado")) for f in filas_mod)
    grupos = Counter(f[1] for f in filas_mod)

    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Resumen", f"Factor 8 · Trabajos de grado vinculados a grupos de investigación — {etiqueta}",
                 "Versión por modalidad del «Consolidado de trabajos de grado MCIC 2022-2026» (el libro completo se conserva sin cambios en Original/). "
                 "Se omiten el código estudiantil y la ruta del archivo fuente.")
    fila = encabezado_factor(ws, 4, 8, m)
    fila = st.kpis(ws, fila, [("Trabajos de grado", len(filas_mod)), ("Sustentados", len(sust)),
                              ("Pendientes / sin verificación", len(pend)), ("Grupos de investigación", len(grupos))])
    st.nota(ws, fila + 1, f"Casos del consolidado sin modalidad en la fuente y sin código para cruzar: {len(por_confirmar)} (hoja «Modalidad por confirmar», la misma en ambas modalidades).")
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 17
    fila = st.seccion(ws, fila + 3, "Por año")
    anios = sorted({f[3] for f in filas_mod if f[3]})
    fila = st.tabla(ws, fila, ["Año", "Sustentados", "No sustentados", "Total"],
                    [[a, por_anio.get((a, True), 0), por_anio.get((a, False), 0), por_anio.get((a, True), 0) + por_anio.get((a, False), 0)] for a in anios],
                    filtro=False, congelar=False)
    fila = st.seccion(ws, fila, "Por grupo de investigación")
    st.tabla(ws, fila, ["Grupo", "Trabajos de grado"], [[g, n] for g, n in grupos.most_common()], filtro=False, congelar=False)

    encab = ["ID", "Grupo de investigación", "Título del trabajo de grado", "Año", "Modalidad", "Director(es)", "Estudiante(s)",
             "Fecha / soporte de finalización", "Enlace RIUD", "Estado de sustentación", "Nota de modalidad", "Observación de cruce"]
    anch = [6, 22, 60, 7, 14, 28, 28, 18, 30, 22, 30, 50]
    ws = st.hoja(wb, "Consolidado", f"Trabajos de grado 2022-2026 — {etiqueta}", None, 12)
    st.tabla(ws, 3, encab, filas_mod, anch)
    ws = st.hoja(wb, "Pendientes sustentación", f"Casos programados, radicados o avalados sin sustentación — {etiqueta}", None, 12)
    st.tabla(ws, 3, encab, pend, anch)
    ws = st.hoja(wb, "Modalidad por confirmar", "Casos sin modalidad registrada en las fuentes", "Confirmar la modalidad antes de publicar.", 12)
    st.tabla(ws, 4, encab, por_confirmar, anch)
    ws = st.hoja(wb, "Fuentes y calidad", "Fuentes, alcance y control de calidad (tomado del consolidado original)", None, 4)
    fuentes = [[c for c in r] for r in src["Fuentes y calidad"].iter_rows(min_row=4, values_only=True) if any(r)]
    st.tabla(ws, 3, ["Fuente", "Registros incorporados", "Uso en el consolidado", "Observación"], fuentes[1:], [50, 20, 30, 80], filtro=False)

    act_b = dest / "b. Socialización y vinculación de actividades de investigación"
    act_b.mkdir(parents=True, exist_ok=True)
    nombre = f"F8_Trabajos_de_grado_por_grupo_{nombre_mod}.xlsx"
    wb.save(act_b / nombre)
    out.append(f"{act_b.name}/{nombre}")
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
    matric = sum(1 for r in ROSTER if r["proyecto"] == mod["proyecto"] and r["estado"] == "Matriculado")

    divulgacion = [
        ["31/01/2025", "Comunicación de inicio de clases e inducción 2025-1", "Correo a estudiantes", "Investigación", "Original/…/FACTOR 9/Induccion 2025-1.pdf"],
        ["02/02/2026", f"Inducción 2026-1: {ind26.get(etiqueta, 0)} estudiantes de {etiqueta} firmaron la planilla (cohorte 2026-1 admitida: {admit})", "Inducción", "Ambas", "Original/…/FACTOR 9/LISTA ASISTENCIA INDUCCIONES 2026-1.pdf"],
        ["—", "Encuentro con estudiantes – evaluación docente (registro fotográfico)", "Encuentro", "Ambas", "Presentacion/…/FACTOR 2/…/soportes/Encuentro estudiantes evaluacion docente.jpg"],
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
    fila = st.kpis(ws, fila, [(f"Matriculados {etiqueta} (Cóndor)", matric), ("Firmaron inducción 2026-1", ind26.get(etiqueta, 0)),
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
# FACTORES 7, 10, 11, 12 — sin cambios de contenido (se ubican en su actividad)
# ---------------------------------------------------------------------------
UBICACION_SIN_CAMBIOS = {
    7: lambda f: "d. Definición de acción para la formalización de convenios",
    10: lambda f: "a. Diagnóstico sobre ambientes de aprendizaje",
    11: lambda f: "c. Implementación de los lineamientos del SIAC" if f.name.startswith(("CC-FR-001", "MCIC Autoevaluacion")) else "d. Construcción de reportes",
    12: lambda f: "a. Solicitud de informe de avance de la obra",
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
ESTADO = {1: "Sin cambios (aprobado)", 2: "Actualizado", 3: "Actualizado (en recolección)", 4: "Actualizado",
          5: "Verificado", 6: "Actualizado", 7: "Sin cambios", 8: "Actualizado (para revisión)", 9: "Actualizado",
          10: "Sin cambios", 11: "Sin cambios", 12: "Sin cambios"}
OBS = {
    1: "PEP y actas de las jornadas con docentes de Ingeniería de Software y Geomática.",
    2: "Libro único con páginas web, actividades, cobertura de cohortes y galería; los soportes con datos personales quedan en Original/.",
    3: "Pendiente: soportes de capacitaciones 2025-2026 que reporten los profesores (columna por diligenciar).",
    4: "Se retiran «Experiencias UD» e «Infografía Esquema Normativo». Pendiente: formulario a egresados 2022-2026 (sector en que laboran) enviado a la OATI.",
    5: "Los syllabus no se diferencian por modalidad; el plan de estudios (Res. 016 de 2025) sí. Ver hallazgos en el libro de verificación.",
    6: "Las solicitudes de prórroga se presentan con ID, sin datos personales.",
    7: "Normativa institucional de internacionalización y cooperación. Pendiente: correo a profesores sobre participación en convenios.",
    8: "Consolidado de trabajos de grado separado por modalidad; ponencias (ANEXOS PONENCIAS) solo en Investigación, de donde proviene el soporte.",
    9: "Incluye estímulos (Res. 143 de 2025 solo aplica a Investigación) y estadísticas de uso de Bienestar.",
    10: "Pendiente: información de Biblioteca y Planes TIC.",
    11: "Plan de mejoramiento y autoevaluación propios de la modalidad, más resultados e instrumentos de autoevaluación.",
    12: "Pendiente: información de laboratorios.",
}


def indice(nombre_mod: str, mod: dict, base: Path, entregado: dict[int, list[str]], m: dict) -> None:
    wb = st.nuevo_libro()
    ws = st.hoja(wb, "Índice", f"Evidencias del Plan de Mejoramiento — {mod['programa']} (SNIES {mod['snies']})",
                 "Propuesta de evidencias para aprobación. Periodo 2024-2027: plan anterior 2024-2026 y plan vigente 2026-2027. "
                 "Cada factor se entrega en su carpeta, con los archivos dentro de la actividad a la que corresponden.", 8)
    filas = []
    for n in range(1, 13):
        filas.append([n, m[n]["factor"].split(". ", 1)[1].rstrip("."), m[n]["meta_anterior"], m[n]["meta_vigente"],
                      SOLICITUDES_PARES[n][0], SOLICITUDES_PARES[n][1], ESTADO[n], len(entregado[n]), OBS[n]])
    st.tabla(ws, 4, ["N°", "Factor", "Meta plan anterior (2024-2026)", "Meta plan vigente (2026-2027)", "Solicitud de los pares (sep. 2026)",
                     "Responsable", "Estado", "Archivos", "Observaciones"], filas, [5, 30, 45, 45, 45, 20, 18, 9, 50])
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
    for sub in ("Original", "Presentacion"):
        shutil.rmtree(OUT / sub, ignore_errors=True)

    # Original: copia fiel de cada modalidad + insumos adicionales usados
    for nombre_mod, mod in MODALIDADES.items():
        o = OUT / "Original" / nombre_mod
        copiar_arbol(mod["plan"], o)
        f8 = o / factor_dirs(mod["plan"])[8].name
        copiar(CONSOLIDADO, f8)
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
            4: factor4(mod, d[4]),
            5: factor5(nombre_mod, mod, d[5], m),
            6: factor6(nombre_mod, mod, d[6], tr, m),
            7: sin_cambios(7, mod, d[7]),
            8: factor8(nombre_mod, mod, d[8], m),
            9: factor9(nombre_mod, mod, d[9], tr, m),
            10: sin_cambios(10, mod, d[10]),
            11: sin_cambios(11, mod, d[11]),
            12: sin_cambios(12, mod, d[12]),
        }
        indice(nombre_mod, mod, base, entregado, m)
        total = sum(len(v) for v in entregado.values())
        print(f"Presentacion/{nombre_mod}: {total} archivos en 12 factores")
    print("Original/: copia fiel de Investigación y Profundización + consolidado de trabajos de grado + transcripciones")


if __name__ == "__main__":
    main()
