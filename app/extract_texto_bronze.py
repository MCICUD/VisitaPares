"""Bronze -> Silver

Recorre el catálogo completo de Data/Bronze (Data/Silver/bronze_manifest.json)
y extrae el texto plano de cada archivo soportado (.docx, .pdf, .pptx,
.xlsx/.xlsm, .xls, .csv/.txt) usando app/lib/doc_reader.py.

Esto es lo que permite: (a) buscar por CONTENIDO en el catálogo de
Documentos del sitio, y (b) que extract_indicios_seguimiento.py encuentre
fechas/porcentajes ya escritos en la evidencia real, en vez de tener que
abrir cada archivo a mano.

No resume ni interpreta nada — solo indexa el texto tal como está en el
archivo. Los que no se pudieron leer (imagen sin capa de texto, .doc
binario antiguo, etc.) quedan marcados con su motivo, no con texto inventado.
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib.doc_reader import extraer_texto  # noqa: E402

SILVER_DIR = ROOT / "Data/Silver"
MANIFEST_PATH = SILVER_DIR / "bronze_manifest.json"


def main() -> None:
    if not MANIFEST_PATH.exists():
        raise FileNotFoundError(f"No existe {MANIFEST_PATH}. Corre primero app/build_bronze_manifest.py")

    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))

    resultados = {}
    inicio = time.time()
    n_ok, n_falla = 0, 0
    for i, registro in enumerate(manifest, start=1):
        path = ROOT / registro["archivo"]
        r = extraer_texto(path)
        resultados[registro["archivo"]] = r
        if r["extraido"]:
            n_ok += 1
        else:
            n_falla += 1
        if i % 100 == 0:
            print(f"  ... {i}/{len(manifest)} archivos procesados ({time.time() - inicio:.0f}s)")

    SILVER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = SILVER_DIR / "texto_bronze.json"
    out_path.write_text(json.dumps(resultados, ensure_ascii=False), encoding="utf-8")

    print(f"Texto extraído de {len(manifest)} archivos en {time.time() - inicio:.0f}s -> {out_path.relative_to(ROOT)}")
    print(f"  - con texto extraído: {n_ok}")
    print(f"  - sin texto (imagen/formato no soportado/error): {n_falla}")


if __name__ == "__main__":
    main()
