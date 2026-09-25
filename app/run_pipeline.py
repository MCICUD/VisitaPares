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
import extract_syllabi
import extract_texto_bronze
import filter_solicitudes_mcic


def main() -> None:
    print("== 1/15 Bronze -> Silver: catálogo completo ==")
    build_bronze_manifest.main()
    print()
    print("== 2/15 Bronze -> Silver: texto de todos los documentos ==")
    extract_texto_bronze.main()
    print()
    print("== 3/15 Bronze -> Silver: plan de mejoramiento ==")
    extract_plan_mejoramiento.main()
    print()
    print("== 4/15 Bronze -> Silver: evidencia de seguimiento ==")
    extract_seguimiento_evidencia.main()
    print()
    print("== 5/15 Bronze -> Silver: consolidado de énfasis ==")
    extract_enfasis_estudiantes.main()
    print()
    print("== 6/15 Bronze -> Silver: estado académico agregado ==")
    extract_estado_academico.main()
    print()
    print("== 7/15 Bronze -> Silver: egresados agregado ==")
    extract_egresados_agregado.main()
    print()
    print("== 8/15 Bronze -> Silver: grupos de investigación ==")
    extract_grupos_investigacion.main()
    print()
    print("== 9/15 Bronze -> Silver: seguimiento de trabajo de grado (595/695) ==")
    extract_seguimiento_tesis.main()
    print()
    print("== 10/15 Bronze -> Silver: cuadros maestros CNA (graduación/bienestar/grupos) ==")
    extract_cuadros_maestros.main()
    print()
    print("== 11/15 Bronze -> Silver: convenios (URELINTER) ==")
    extract_convenios.main()
    print()
    print("== 12/15 Bronze -> Silver: microcurrículos (Syllabus AA-FR-003) ==")
    extract_syllabi.main()
    print()
    print("== 13/15 SolicitudesPares -> Silver: material entregado a los pares ==")
    filter_solicitudes_mcic.main()
    build_solicitudes_pares.main()
    print()
    print("== 14/15 Silver -> Gold ==")
    build_gold.main()
    print()
    print("== 15/15 Silver -> Gold: consolidado .xlsx de Comunidad Estudiantil ==")
    build_comunidad_xlsx.main()


if __name__ == "__main__":
    main()
