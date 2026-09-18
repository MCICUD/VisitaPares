"""Bronze -> Silver

Convenios institucionales, para el Factor 7 (Interacción con el entorno
nacional e internacional) del Plan de Mejoramiento.

A diferencia de todos los demás extractores de esta pipeline, la fuente NO
es un archivo que la coordinación ya tuviera en Data/Bronze: es el listado
oficial de convenios vigentes que publica URELINTER (Unidad de Relaciones
Internacionales e Interinstitucionales de la Universidad Distrital) en
https://urelinter.udistrital.edu.co/convenios/cooperacion-redes-asociaciones
(archivo descargable ahí: "Convenios vigentes URELINTER.xlsx", descrito por
esa página como "el listado de los convenios vigentes al último trimestre
2025"). Se descargó ese mismo archivo, sin modificarlo, a
Data/Bronze/Convenios/Convenios vigentes URELINTER.xlsx el 2026-09-18, para
que a partir de ahí siga el mismo tratamiento Bronze -> Silver -> Gold que
cualquier otro documento (trazable, versionable, sin depender de que el
sitio externo siga disponible).

Qué hace este script y qué NO hace
----------------------------------
Es un listado INSTITUCIONAL (los 388 convenios vigentes de TODA la
Universidad Distrital), no uno filtrado por programa: URELINTER no ofrece un
recorte por programa/facultad (se confirmó revisando el sitio). Por eso:

  1. Se reportan los totales institucionales tal cual (por nivel, por tipo),
     como contexto.
  2. Se filtran, por texto (institución + denominación + objeto), los
     convenios que mencionan explícitamente a la Facultad de Ingeniería
     (la facultad a la que pertenece la MCIC) o a alguna de sus áreas/
     énfasis afines. La lista de frases usada para ese filtro es la
     constante FRASES_FACULTAD_INGENIERIA de abajo — se deja explícita y
     completa en el JSON de salida para que se pueda auditar (qué frase
     hizo calzar cada convenio), en vez de aplicar un filtro opaco.
  3. NINGÚN convenio de los 388 menciona por su nombre completo a la
     "Maestría en Ciencias de la Información y las Comunicaciones" ni a
     "MCIC": se deja esto explícito en 'limitacion' en vez de mostrar los
     convenios de Facultad de Ingeniería como si fueran de la maestría
     específicamente.
  4. De los 7 convenios que mencionan a la Facultad de Ingeniería, se leyó a
     mano el objeto de cada uno (ver APLICA_A_POSGRADO_MCIC) para confirmar
     si de verdad aplican a un programa de posgrado como la MCIC: solo 3 lo
     hacen (uno no restringe nivel, uno dice "todos los estudiantes de la
     UDFJC", y uno menciona expresamente "pregrado y postgrado"); los otros
     4 restringen su objeto a un programa de PREGRADO (Ingeniería de
     Sistemas o Ingeniería de Software) y no benefician a la MCIC. Se
     conservan los 7 en la salida (con su clasificación) en vez de borrar
     los que no aplican, para que quede auditable por qué se descartan.
"""
from __future__ import annotations

import json
import re
import sys
import unicodedata
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SILVER_DIR = ROOT / "Data/Silver"
ARCHIVO = ROOT / "Data/Bronze/Convenios/Convenios vigentes URELINTER.xlsx"
SHEET_NAME = "Convenios URELINTER"

FUENTE_URL = "https://urelinter.udistrital.edu.co/convenios/cooperacion-redes-asociaciones"
FECHA_DESCARGA = "2026-09-18"

COLUMNAS = (
    "codigo", "nivel", "pais_categoria", "institucion", "fecha_inicio", "fecha_fin",
    "vigencia_anios", "estado", "tipo", "denominacion", "objeto",
)

# Frases (sin tildes, mayúsculas) que identifican un convenio relacionado con
# la Facultad de Ingeniería de la UD (la facultad de la MCIC) o con las
# áreas/énfasis de la maestría (Sistemas, Telemática, Teleinformática,
# Geomática, Telecomunicaciones). Se buscan sobre institución+denominación+
# objeto, sin distinguir mayúsculas ni tildes.
FRASES_FACULTAD_INGENIERIA = (
    "FACULTAD DE INGENIERIA",
    "INGENIERIA DE SISTEMAS",
    "TELEMATICA",
    "TELEINFORMATICA",
    "GEOMATICA",
    "TELECOMUNICACIONES",
    "CIENCIAS DE LA INFORMACION Y LAS COMUNICACIONES",
    "MAESTRIA EN CIENCIAS DE LA INFORMACION",
    "MCIC",
)

# Para cada uno de los convenios que calzan con FRASES_FACULTAD_INGENIERIA se
# revisó a mano el "objeto" completo (texto real del Excel, citado abajo) para
# determinar si su alcance incluye posgrado (y por tanto a la MCIC) o si el
# propio texto lo restringe a pregrado / a un proyecto curricular de pregrado
# distinto de la MCIC. No se puede automatizar con una palabra clave: p. ej.
# el convenio con NEWNET S.A. restringe a pregrado sin usar la palabra
# "pregrado" (dice "Proyecto Curricular de Ingeniería de Sistemas", que es el
# nombre del programa de pregrado). Se deja la cita exacta que sustenta cada
# decisión para que se pueda verificar contra el Excel fuente.
APLICA_A_POSGRADO_MCIC: dict[str, tuple[bool, str]] = {
    "C-2012-12": (
        True,
        "Convenio Marco de cooperación general (seminarios, proyectos conjuntos, capacitación); "
        "el objeto no restringe nivel de formación ni proyecto curricular.",
    ),
    "C-2012-36": (
        False,
        "El objeto dice explícitamente 'movilidad académica de estudiantes de pregrado de Ingeniería "
        "de Sistemas'; no incluye posgrado ni a la MCIC.",
    ),
    "C-2016-10": (
        False,
        "El objeto restringe a estudiantes 'del Proyecto Curricular de Ingeniería de Sistemas' "
        "(el programa de pregrado); no menciona posgrado ni a la MCIC.",
    ),
    "C-2019-28": (
        True,
        "El objeto dice 'podrán participar todos los estudiantes de la UDFJC'; no restringe por "
        "nivel de formación ni por proyecto curricular del lado UDFJC.",
    ),
    "C-2022-29": (
        True,
        "El objeto dice explícitamente 'movilidad académica de estudiantes (pregrado y postgrado) de "
        "la facultad de ingeniería'; incluye posgrado, y por tanto a la MCIC.",
    ),
    "C-2023-19": (
        False,
        "El objeto es sobre crear un programa académico conjunto de 'Pregrado en Ingeniería de "
        "Software'; no es un programa de posgrado, no aplica a la MCIC.",
    ),
    "C-2025-46": (
        False,
        "El objeto da acceso a 'estudiantes de pregrado de la Facultad de Ingeniería' de la UDFJC a "
        "programas de posgrado del IMT Atlantique; el lado UDFJC del convenio es solo de pregrado, "
        "no da ningún beneficio a estudiantes de la MCIC.",
    ),
}


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def normalizar(texto: str) -> str:
    sin_tildes = "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    )
    return sin_tildes.upper()


def limpiar_tipo(valor: str | None) -> str | None:
    if valor is None:
        return None
    valor = re.sub(r"\s+", " ", valor).strip()
    return valor or None


def fecha_iso(valor) -> str | None:
    import datetime

    if isinstance(valor, (datetime.datetime, datetime.date)):
        return valor.date().isoformat() if isinstance(valor, datetime.datetime) else valor.isoformat()
    if valor is None:
        return None
    return str(valor)


def main() -> None:
    if not ARCHIVO.exists():
        raise FileNotFoundError(
            f"No existe {ARCHIVO}. Descargar de {FUENTE_URL} y guardar con ese nombre exacto."
        )

    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = openpyxl.load_workbook(ARCHIVO, data_only=True, read_only=True)
    ws = wb[SHEET_NAME]
    filas = list(ws.iter_rows(values_only=True))[1:]  # fila 1 = encabezado

    convenios = []
    for i, fila in enumerate(filas, start=2):
        if not fila or fila[3] is None:  # sin institución, fila vacía de relleno
            continue
        registro = dict(zip(COLUMNAS, fila))
        registro["fecha_inicio"] = fecha_iso(registro["fecha_inicio"])
        registro["fecha_fin"] = fecha_iso(registro["fecha_fin"])
        registro["tipo"] = limpiar_tipo(registro["tipo"])
        registro["_fila"] = i
        convenios.append(registro)

    por_nivel: dict[str, int] = {}
    por_tipo: dict[str, int] = {}
    for c in convenios:
        por_nivel[c["nivel"] or "Sin especificar"] = por_nivel.get(c["nivel"] or "Sin especificar", 0) + 1
        por_tipo[c["tipo"] or "Sin especificar"] = por_tipo.get(c["tipo"] or "Sin especificar", 0) + 1

    relevantes = []
    for c in convenios:
        texto = normalizar(
            " ".join(str(c.get(k) or "") for k in ("institucion", "denominacion", "objeto"))
        )
        frases_encontradas = [f for f in FRASES_FACULTAD_INGENIERIA if f in texto]
        if frases_encontradas:
            aplica, justificacion = APLICA_A_POSGRADO_MCIC.get(
                c["codigo"], (None, "No evaluado a mano todavía; revisar el objeto contra el Excel fuente.")
            )
            relevantes.append({
                **c,
                "frases_encontradas": frases_encontradas,
                "aplica_a_posgrado_mcic": aplica,
                "justificacion_aplicabilidad": justificacion,
            })

    aplicables_mcic = [c for c in relevantes if c["aplica_a_posgrado_mcic"] is True]

    data = {
        "convenios_totales": len(convenios),
        "por_nivel": por_nivel,
        "por_tipo": por_tipo,
        "convenios_relacionados_facultad_ingenieria": relevantes,
        "total_convenios_relacionados": len(relevantes),
        "total_convenios_aplicables_mcic": len(aplicables_mcic),
        "criterio_filtro": (
            "Se buscó, sobre institución + denominación + objeto de cada convenio (sin tildes ni "
            "mayúsculas), alguna de estas frases: " + ", ".join(FRASES_FACULTAD_INGENIERIA) + ". "
            "Cada convenio relevante trae 'frases_encontradas' con las que hizo calzar, para poder "
            "verificar el filtro contra el Excel fuente. Además, cada uno trae "
            "'aplica_a_posgrado_mcic' (true/false) y 'justificacion_aplicabilidad': una lectura manual "
            "del objeto de ESE convenio (citada literalmente) que dice si su alcance incluye posgrado "
            "(y por tanto a la MCIC) o si el propio texto lo restringe a un programa de pregrado "
            "distinto — ver APLICA_A_POSGRADO_MCIC en este script."
        ),
        "limitacion": (
            "Este es el listado INSTITUCIONAL de convenios vigentes de toda la Universidad Distrital "
            "(no uno filtrado por programa: URELINTER no ofrece un recorte por programa/facultad). "
            "Ningún convenio de los 388 menciona por su nombre a la Maestría en Ciencias de la "
            "Información y las Comunicaciones (MCIC). Se encontraron por texto "
            f"{len(relevantes)} que mencionan a la Facultad de Ingeniería (la facultad a la que "
            f"pertenece la MCIC) o a alguna de sus áreas afines; de esos, solo {len(aplicables_mcic)} "
            "aplican realmente a un programa de posgrado como la MCIC según su propio objeto (los "
            "otros 4 quedaron descartados de la lista que se muestra porque su objeto restringe "
            "explícitamente el convenio a un programa de PREGRADO — Ingeniería de Sistemas o "
            "Ingeniería de Software — y no benefician a estudiantes de la maestría). El detalle de "
            "los 7 encontrados y por qué se descartaron los 4 queda en "
            "'convenios_relacionados_facultad_ingenieria' (campo 'aplica_a_posgrado_mcic') de este "
            "mismo archivo, para que se pueda auditar."
        ),
        "fuente": {
            "url": FUENTE_URL,
            "archivo_descargado": rel(ARCHIVO),
            "fecha_descarga": FECHA_DESCARGA,
            "nota": (
                "Descrito por URELINTER como 'el listado de los convenios vigentes al último trimestre "
                "2025'. Para copia de un convenio o información adicional: convenios-ceri@udistrital.edu.co."
            ),
        },
    }

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "convenios.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    print(
        f"Convenios: {len(convenios)} institucionales, {len(relevantes)} relacionados con Fac. Ingeniería, "
        f"{len(aplicables_mcic)} aplican realmente a posgrado/MCIC -> {out_path.relative_to(ROOT)}"
    )


if __name__ == "__main__":
    main()
