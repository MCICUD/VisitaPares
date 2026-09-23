"""Orquesta la pipeline completa Bronze -> Silver -> Gold."""
from __future__ import annotations

import build_bronze_manifest
import build_comunidad_xlsx
import build_gold
import build_solicitudes_pares
import extract_convenios
import extract_cuadros_maestros
import extract_egresados_agregado
import extract_enfasis_estudiantes
import extract_estado_academico
import extract_grupos_investigacion
import extract_plan_mejoramiento
import extract_seguimiento_evidencia
import extract_seguimiento_tesis
import extract_texto_bronze


def main() -> None:
    print("== 1/14 Bronze -> Silver: catálogo completo ==")
    build_bronze_manifest.main()
    print()
    print("== 2/14 Bronze -> Silver: texto de todos los documentos ==")
    extract_texto_bronze.main()
    print()
    print("== 3/14 Bronze -> Silver: plan de mejoramiento ==")
    extract_plan_mejoramiento.main()
    print()
    print("== 4/14 Bronze -> Silver: evidencia de seguimiento ==")
    extract_seguimiento_evidencia.main()
    print()
    print("== 5/14 Bronze -> Silver: consolidado de énfasis ==")
    extract_enfasis_estudiantes.main()
    print()
    print("== 6/14 Bronze -> Silver: estado académico agregado ==")
    extract_estado_academico.main()
    print()
    print("== 7/14 Bronze -> Silver: egresados agregado ==")
    extract_egresados_agregado.main()
    print()
    print("== 8/14 Bronze -> Silver: grupos de investigación ==")
    extract_grupos_investigacion.main()
    print()
    print("== 9/14 Bronze -> Silver: seguimiento de trabajo de grado (595/695) ==")
    extract_seguimiento_tesis.main()
    print()
    print("== 10/14 Bronze -> Silver: cuadros maestros CNA (graduación/bienestar/grupos) ==")
    extract_cuadros_maestros.main()
    print()
    print("== 11/14 Bronze -> Silver: convenios (URELINTER) ==")
    extract_convenios.main()
    print()
    print("== 12/14 SolicitudesPares -> Silver: material entregado a los pares ==")
    build_solicitudes_pares.main()
    print()
    print("== 13/14 Silver -> Gold ==")
    build_gold.main()
    print()
    print("== 14/14 Silver -> Gold: consolidado .xlsx de Comunidad Estudiantil ==")
    build_comunidad_xlsx.main()


if __name__ == "__main__":
    main()
