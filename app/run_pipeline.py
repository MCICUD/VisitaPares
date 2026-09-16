"""Orquesta la pipeline completa Bronze -> Silver -> Gold."""
from __future__ import annotations

import build_bronze_manifest
import build_gold
import extract_egresados_agregado
import extract_enfasis_estudiantes
import extract_estado_academico
import extract_grupos_investigacion
import extract_plan_mejoramiento
import extract_seguimiento_evidencia
import extract_texto_bronze


def main() -> None:
    print("== 1/9 Bronze -> Silver: catálogo completo ==")
    build_bronze_manifest.main()
    print()
    print("== 2/9 Bronze -> Silver: texto de todos los documentos ==")
    extract_texto_bronze.main()
    print()
    print("== 3/9 Bronze -> Silver: plan de mejoramiento ==")
    extract_plan_mejoramiento.main()
    print()
    print("== 4/9 Bronze -> Silver: evidencia de seguimiento ==")
    extract_seguimiento_evidencia.main()
    print()
    print("== 5/9 Bronze -> Silver: consolidado de énfasis ==")
    extract_enfasis_estudiantes.main()
    print()
    print("== 6/9 Bronze -> Silver: estado académico agregado ==")
    extract_estado_academico.main()
    print()
    print("== 7/9 Bronze -> Silver: egresados agregado ==")
    extract_egresados_agregado.main()
    print()
    print("== 8/9 Bronze -> Silver: grupos de investigación ==")
    extract_grupos_investigacion.main()
    print()
    print("== 9/9 Silver -> Gold ==")
    build_gold.main()


if __name__ == "__main__":
    main()
