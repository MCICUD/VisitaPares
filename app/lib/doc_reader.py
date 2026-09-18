"""Extracción de texto plano de cualquier archivo en Data/Bronze, sin
interpretar su contenido — solo lo convierte a texto para poder buscarlo y,
en un paso posterior (extract_indicios_seguimiento.py), detectar por regex
fechas/porcentajes que ya estén escritos literalmente en el documento.

Nunca resume, traduce ni "adivina" contenido: si un formato no se puede leer
(p. ej. .doc binario antiguo sin herramienta disponible, o una imagen sin
texto) se marca `error` y no se inventa nada en su lugar.
"""
from __future__ import annotations

import warnings
import zipfile
from pathlib import Path
from typing import Any

MAX_CHARS = 20_000  # límite razonable por archivo para no disparar el tamaño de Silver/Gold
MAX_BYTES_PDF = 30_000_000  # PDFs más pesados que esto no se procesan (escaneos enormes)


def _leer_docx(path: Path) -> str:
    import docx

    d = docx.Document(str(path))
    partes = [p.text for p in d.paragraphs if p.text.strip()]
    for tabla in d.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if celda.text.strip():
                    partes.append(celda.text.strip())
    return "\n".join(partes)


def _leer_pdf(path: Path) -> str:
    import pymupdf

    if path.stat().st_size > MAX_BYTES_PDF:
        raise ValueError(f"PDF > {MAX_BYTES_PDF / 1e6:.0f} MB, omitido por tamaño")
    partes = []
    with pymupdf.open(str(path)) as doc:
        for pagina in doc:
            texto = pagina.get_text()
            if texto.strip():
                partes.append(texto)
    return "\n".join(partes)


def _leer_pptx(path: Path) -> str:
    from pptx import Presentation

    prs = Presentation(str(path))
    partes = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                partes.append(shape.text_frame.text)
    return "\n".join(partes)


def _leer_xlsx(path: Path) -> str:
    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    partes = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for val in row:
                if val is not None and str(val).strip():
                    partes.append(str(val).strip())
    return "\n".join(partes)


def _leer_xls(path: Path) -> str:
    import xlrd

    try:
        wb = xlrd.open_workbook(str(path))
    except xlrd.XLRDError:
        # Algunos archivos ".xls" son en realidad .xlsx renombrados (formato
        # moderno con extensión antigua) — xlrd no los abre, openpyxl sí.
        return _leer_xlsx(path)
    partes = []
    for sheet in wb.sheets():
        for r in range(sheet.nrows):
            for val in sheet.row_values(r):
                if val not in (None, ""):
                    partes.append(str(val).strip())
    return "\n".join(partes)


def _leer_csv_txt(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def _leer_imagen(path: Path) -> str:
    import pytesseract
    from PIL import Image

    # Limitar el tamaño de la imagen por rendimiento si es necesario, 
    # pero pytesseract suele procesar imágenes razonables sin problema.
    img = Image.open(str(path))
    texto = pytesseract.image_to_string(img, lang='spa+eng')
    return texto.strip()


LECTORES = {
    "docx": _leer_docx,
    "pdf": _leer_pdf,
    "pptx": _leer_pptx,
    "xlsx": _leer_xlsx,
    "xlsm": _leer_xlsx,
    "xls": _leer_xls,
    "csv": _leer_csv_txt,
    "txt": _leer_csv_txt,
    "png": _leer_imagen,
    "jpg": _leer_imagen,
    "jpeg": _leer_imagen,
}

NO_SOPORTADOS = {"doc", "zip", "rar", "lnk"}


def extraer_texto(path: Path) -> dict[str, Any]:
    """Devuelve {texto, extraido, error}. `texto` viene recortado a MAX_CHARS."""
    ext = path.suffix.lower().lstrip(".")
    lector = LECTORES.get(ext)
    if lector is None:
        motivo = "formato sin lector implementado" if ext in NO_SOPORTADOS else f"extensión .{ext} no soportada"
        return {"texto": "", "extraido": False, "error": motivo}
    try:
        texto = lector(path)
    except Exception as exc:  # noqa: BLE001 - reportamos cualquier fallo de lectura, no lo ocultamos
        return {"texto": "", "extraido": False, "error": f"{type(exc).__name__}: {exc}"}
    texto = texto.strip()
    truncado = len(texto) > MAX_CHARS
    return {
        "texto": texto[:MAX_CHARS],
        "extraido": bool(texto),
        "error": None if texto else "sin texto extraíble (posible imagen/escaneo sin capa de texto)",
        "truncado": truncado,
    }
