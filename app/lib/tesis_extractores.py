"""Extractores de campos estructurados desde el TEXTO de los 4 tipos de
documento del proceso de trabajo de grado 595/695 (Investigación/
Profundización): carta de radicación, solicitud de jurados, carta de aval y
viabilidad del grupo, y acta de sustentación.

Se inspeccionaron ~15 documentos reales (copiados a Data/Bronze/Seguimiento_Tesis/)
antes de escribir estos patrones. Son cartas institucionales de formato
bastante fijo, pero se encontraron DOS plantillas distintas de "carta de
radicación" que coexisten en la maestría:

  - "RADICACIÓN ANTEPROYECTO" (modalidad tesis): "...radicar el anteproyecto
    de trabajo de grado titulado <TÍTULO> de la autoría del <NOMBRE>
    identificado con el código estudiantil <CÓDIGO> cuyo director es el
    profesor <DIRECTOR>. Para que sea remitido al grupo de investigación
    <GRUPO>."
  - "RADICACIÓN PASANTÍA" (modalidad pasantía, la mayoría de los casos
    encontrados en Profundización): "...presentar al docente de planta
    <DIRECTOR> como director [y al señor <CODIRECTOR> como codirector] de la
    pasantía <TÍTULO> para la correspondiente aprobación...".

Si un campo no se puede extraer con confianza, se deja en None — nunca se
inventa un valor. Si ninguna de las dos plantillas de radicación calza, el
documento simplemente no aporta título/director (pero puede seguir aportando
"grupo_investigacion" si esa línea sí aparece, o simplemente contarse como
evidencia de que existe una carta de radicación para ese estudiante).
"""
from __future__ import annotations

import re


def _limpiar(s: str | None) -> str | None:
    if not s:
        return None
    s = re.sub(r"\s+", " ", s).strip(" \n\t.,:;\"“”")
    return s or None


# Textos de plantilla sin diligenciar (ej. "NOMBRE DEL CO-DIRECTOR" cuando el
# campo quedó vacío en el formato) que un extractor puede confundir con un
# valor real si solo mira "la siguiente línea no vacía". Se filtran para no
# publicar la etiqueta del formulario como si fuera un dato real.
_MARCADORES_PLANTILLA = (
    "nombre del director", "nombre del co-director", "nombre del codirector",
    "nombre del estudiante", "nombre proyecto", "codigo", "código", "correo",
    "facultad a la que", "firma del", "firma director",
)


def _es_marcador_plantilla(s: str) -> bool:
    s_norm = s.strip().lower()
    return any(s_norm.startswith(m) for m in _MARCADORES_PLANTILLA)


def _limpiar_nombre(s: str | None) -> str | None:
    """Como _limpiar, pero además descarta el resultado si no parece un
    nombre de persona (ver _parece_nombre_de_persona) — para no publicar un
    párrafo de plantilla como si fuera el nombre del director/codirector."""
    limpio = _limpiar(s)
    if limpio and not _parece_nombre_de_persona(limpio):
        return None
    return limpio


def _parece_nombre_de_persona(s: str) -> bool:
    """Descarta líneas que claramente no son un nombre propio: párrafos de
    nota/disclaimer que a veces caen justo después de una etiqueta 'Vo. Bo.'
    cuando el campo real quedó en blanco (ej. 'Nota: Esta solicitud se
    tendrá en cuenta solo si...'). Un nombre real es corto y no trae
    conectores de oración en minúscula."""
    palabras = s.split()
    if len(s) > 60 or len(palabras) > 6:
        return False
    conectores = (" se ", " que ", " del ", " para ", " tendrá ", " esta ", " este ")
    if any(c in f" {s.lower()} " for c in conectores):
        return False
    return True


RADICACION_ANTEPROYECTO_RE = re.compile(
    r"titulado\s+\(?(.*?)\)?\s*de la autor[ií]a del(?:\s+estudiante)?\s+.*?"
    r"c[oó]digo(?:\s+estudiantil)?\s*\(?\s*\d{6,11}\s*\)?\s*,?\s*cuyo director es el\s*"
    r"(?:profesor|profesora)?\s*\(?(.*?)\)?(?:\.\s|,|\s+para que sea remitido|\s+el proyecto en referencia)",
    re.IGNORECASE | re.DOTALL,
)

RADICACION_PASANTIA_RE = re.compile(
    r"presentar (?:al|a la)\s+(?:docente de planta\s+)?(.*?)\s+como\s+director"
    r"(?:\s+y al (?:señor|la señora)\s+[\"“]?(.*?)[\"”]?\s+como\s+co-?director)?"
    r"\s+de la pasant[ií]a\s+[\"“]?(.*?)[\"”]?\s+para la correspondiente",
    re.IGNORECASE | re.DOTALL,
)

# Insensible a mayúsculas solo en las palabras "grupo"/"de investigación";
# el nombre del grupo (sigla) sí debe quedar en mayúsculas para no capturar
# palabras sueltas comunes ("grupo de trabajo", "grupo curricular", etc.).
GRUPO_INVESTIGACION_RE = re.compile(
    r"(?:grupo(?:\s+de\s+investigaci[oó]n)?)\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ0-9+\-]{1,25})\b"
)
GRUPO_FALSOS_POSITIVOS = {"DE", "CURRICULAR", "INVESTIGACION", "INVESTIGACIÓN", "CONSEJO", "MCIC"}


def _extraer_grupo(plano: str) -> str | None:
    m = GRUPO_INVESTIGACION_RE.search(plano)
    if not m:
        return None
    candidato = m.group(1).strip()
    if candidato.upper() in GRUPO_FALSOS_POSITIVOS:
        return None
    return candidato


def extraer_radicacion(texto: str) -> dict:
    """Carta de radicación (anteproyecto de tesis O pasantía)."""
    plano = re.sub(r"\s+", " ", texto)
    out = {
        "titulo_proyecto": None, "director": None, "codirector": None,
        "grupo_investigacion": None, "es_pasantia": False,
    }

    m = RADICACION_ANTEPROYECTO_RE.search(plano)
    if m:
        out["titulo_proyecto"] = _limpiar(m.group(1))
        out["director"] = _limpiar_nombre(m.group(2))
    else:
        m = RADICACION_PASANTIA_RE.search(plano)
        if m:
            out["es_pasantia"] = True
            out["director"] = _limpiar_nombre(m.group(1))
            out["codirector"] = _limpiar_nombre(m.group(2))
            out["titulo_proyecto"] = _limpiar(m.group(3))

    out["grupo_investigacion"] = _extraer_grupo(plano)
    return out


def _nombre_despues_de_etiqueta(lineas: list[str], etiquetas: list[str]) -> str | None:
    """Busca una línea que EMPIECE con una de las 'etiquetas' (ej. 'VO. BO.
    DIRECTOR', que a veces aparece sola y a veces junto al nombre en la
    misma línea) y devuelve el nombre asociado: el resto de la misma línea
    si no está vacío, o si no, la siguiente línea no vacía."""
    norm_etiquetas = {e.upper().replace(".", "").replace(" ", "") for e in etiquetas}
    for i, linea in enumerate(lineas):
        norm = linea.upper().replace(".", "").replace(" ", "")
        etiqueta_match = next((e for e in norm_etiquetas if norm.startswith(e)), None)
        if not etiqueta_match:
            continue
        resto = linea[len(linea) - (len(norm) - len(etiqueta_match)):].strip() if len(norm) > len(etiqueta_match) else ""
        if resto and not resto.isupper() and not _es_marcador_plantilla(resto) and _parece_nombre_de_persona(resto):
            return _limpiar(resto)
        for j in range(i + 1, min(i + 5, len(lineas))):
            candidato = lineas[j].strip()
            if not candidato or _es_marcador_plantilla(candidato) or not _parece_nombre_de_persona(candidato):
                continue
            return _limpiar(candidato)
    return None


def extraer_solicitud_jurados(texto: str) -> dict:
    """'SOLICITUD DE JURADOS'. Título tras 'NOMBRE PROYECTO:'; director tras
    la etiqueta 'Vo. Bo. DIRECTOR' (normalmente en la línea siguiente, junto
    al nombre completo del estudiante que firma la solicitud)."""
    lineas = texto.split("\n")
    out = {"titulo_proyecto": None, "director": None, "codirector": None}

    plano = re.sub(r"\s+", " ", texto)
    m = re.search(
        r"NOMBRE PROYECTO:\s*[“\"]?(.*?)[”\"]?\s*(?:_{5,}|De igual forma)",
        plano, re.IGNORECASE,
    )
    if m:
        out["titulo_proyecto"] = _limpiar(m.group(1))

    out["director"] = _nombre_despues_de_etiqueta(lineas, ["Vo. Bo. DIRECTOR", "VO.BO. DIRECTOR"])
    out["codirector"] = _nombre_despues_de_etiqueta(
        lineas, ["Vo.Bo. CO-DIRECTOR", "Vo. Bo. CO-DIRECTOR", "VO.BO. CODIRECTOR"]
    )
    return out


def extraer_viabilidad(texto: str) -> dict:
    """'CARTA DE AVAL Y VIABILIDAD DEL GRUPO DE INVESTIGACIÓN'. Campos en
    formato 'Etiqueta: valor' bastante estable."""
    out = {
        "titulo_proyecto": None, "director": None, "codirector": None,
        "grupo_investigacion": None,
    }
    plano = re.sub(r"[ \t]+", " ", texto)

    m = re.search(r"El grupo de investigaci[oó]n\s+(.*?)\s*\n\s*avala", plano, re.IGNORECASE)
    if m:
        out["grupo_investigacion"] = _limpiar(m.group(1))
    if not out["grupo_investigacion"]:
        out["grupo_investigacion"] = _extraer_grupo(re.sub(r"\s+", " ", texto))

    m = re.search(r"T[ií]tulo del proyecto:\s*(.*?)\n\s*Nombre", plano, re.IGNORECASE | re.DOTALL)
    if m:
        out["titulo_proyecto"] = _limpiar(m.group(1))

    # [ \t]* (no \s*) tras los dos puntos: un campo vacío ("co-director (si
    # aplica): " sin nada) no debe arrastrar el siguiente párrafo real del
    # documento como si fuera el nombre.
    m = re.search(r"director del proyecto:[ \t]*([^\n]*)", plano, re.IGNORECASE)
    if m:
        out["director"] = _limpiar_nombre(m.group(1))

    m = re.search(r"co-?director del proyecto[^:\n]*:[ \t]*([^\n]*)", plano, re.IGNORECASE)
    if m:
        out["codirector"] = _limpiar_nombre(m.group(1))

    return out


def extraer_acta_sustentacion(texto: str) -> dict:
    """'ACTA DE SUSTENTACIÓN No. X – AÑO'. Etiquetas 'TITULO DEL PROYECTO:',
    'NOTA:', 'CARÁCTER:' en su propia línea; 'DIRECTOR INTERNO'/'DIRECTOR
    EXTERNO' aparecen en la línea SIGUIENTE al nombre del director."""
    lineas = [l.strip() for l in texto.split("\n")]
    out = {
        "titulo_proyecto": None, "numero_acta": None, "nota": None, "caracter": None,
        "director": None, "codirector": None, "jurados": [],
    }

    plano = re.sub(r"\s+", " ", texto)
    m = re.search(r"ACTA DE SUSTENTACI[OÓ]N\s*No\.?\s*([\w\-]+)", plano, re.IGNORECASE)
    if m:
        out["numero_acta"] = _limpiar(m.group(1))
    m = re.search(r"T[IÍ]TULO DEL PROYECTO:\s*(.*?)\s*(?:ESTUDIANTE:|CÓDIGO:)", plano, re.IGNORECASE)
    if m:
        out["titulo_proyecto"] = _limpiar(m.group(1))
    # OJO: [ \t]* (no \s*) y sobre el texto SIN aplanar — "NOTA:" y "CARÁCTER:"
    # suelen venir vacíos en el formato ("CARÁCTER: \n"), y si se aplana el
    # texto antes (\s+ -> " ") o se usa \s* tras los dos puntos, el patrón se
    # come el salto de línea y captura la primera palabra del párrafo
    # siguiente (ej. "ING." de la firma del jurado) como si fuera el valor.
    m = re.search(r"NOTA:[ \t]*([\d.,]+)", texto, re.IGNORECASE)
    if m:
        out["nota"] = _limpiar(m.group(1))
    m = re.search(r"CAR[AÁ]CTER:[ \t]*([^\n]*)", texto, re.IGNORECASE)
    if m and m.group(1).strip():
        out["caracter"] = _limpiar(m.group(1))

    ROLES = {
        "DIRECTOR": "director", "DIRECTOR INTERNO": "director", "DIRECTOR EXTERNO": "director",
        "CO-DIRECTOR": "codirector", "CODIRECTOR": "codirector", "JURADO": "jurados",
    }
    PREFIJO_RE = re.compile(r"^(Ing\.?|Dr\.?|Dra\.?|Mg\.?|MSc\.?)\s*", re.IGNORECASE)
    SUFIJO_PHD_RE = re.compile(r"\s*PHD\.?\s*$", re.IGNORECASE)
    for i, linea in enumerate(lineas):
        rol = ROLES.get(linea.upper())
        if not rol or i == 0:
            continue
        nombre = PREFIJO_RE.sub("", lineas[i - 1].strip())
        nombre = SUFIJO_PHD_RE.sub("", nombre).strip()
        if not nombre or _es_marcador_plantilla(nombre) or not _parece_nombre_de_persona(nombre):
            continue
        if rol == "jurados":
            out["jurados"].append(_limpiar(nombre))
        elif not out[rol]:
            out[rol] = _limpiar(nombre)
    return out
