# NuevaData — propuesta de evidencias del Plan de Mejoramiento (pendiente de aprobación)

Este documento deja todo el contexto para que, en otra sesión, se pueda pedir:
**«integra NuevaData a Bronze, Silver y Gold»** una vez la coordinación apruebe las correcciones.

- Fecha de preparación: 30 de septiembre de 2026.
- Origen del pedido: correcciones solicitadas por los pares en la visita de septiembre de 2026
  (`CorreccionEvidencias.ods`, raíz del repo). A los pares les gustó que se conservara la data
  original, pero no que estuviera «esparcida». Por eso se entrega ya analizada, sin perder lo original.
- **Estado: PENDIENTE DE APROBACIÓN.** Nada de `NuevaData/` está en `Data/` todavía. El micrositio
  sigue mostrando las evidencias anteriores, salvo los dos cambios de código descritos en §5.

---

## 1. Estructura

```
NuevaData/
├── CONTEXTO_NuevaData.md            ← este documento (versionado en git)
├── _build/
│   ├── build_nuevadata.py           ← genera Original/ y Presentacion/ (idempotente)
│   ├── xlsx_style.py                ← estilo común de los .xlsx
│   └── _privado/transcripciones.json← transcripción manual de planillas (DATOS PERSONALES, fuera de git)
├── Original/                        ← fuera de git (contiene datos personales)
│   ├── Investigacion/               ← copia fiel de Data/Bronze/MCIC.INVESTIGACION/…/Plan de Mejoramiento
│   ├── Profundizacion/              ← copia fiel de Data/Bronze/MCIC-PROFUNDIZACION/…/Plan de Mejoramiento
│   │   (en ambas, FACTOR 8 incluye además el «Consolidado_trabajos_grado_MCIC_2022_2026 (7).xlsx» sin cambios)
│   └── _transcripciones/transcripciones.json
└── Presentacion/                    ← fuera de git (se sube a SharePoint para la aprobación)
    ├── Investigacion/
    │   ├── 00_Indice_Evidencias_Investigacion.xlsx
    │   └── FACTOR 1 … FACTOR 12/<actividad>/<archivos>
    └── Profundizacion/
        ├── 00_Indice_Evidencias_Profundizacion.xlsx
        └── FACTOR 1 … FACTOR 12/<actividad>/<archivos>
```

- Las carpetas `FACTOR N. …` usan **exactamente** los mismos nombres que en Bronze, así que
  `app/extract_seguimiento_evidencia.py` las reconoce sin cambios.
- En `Presentacion/` **ningún archivo queda suelto**: todos están dentro de una carpeta de actividad
  (`a. …`, `b. …`). El extractor solo marca «(sin actividad específica)» los archivos que quedan en la
  raíz del FACTOR, y en Presentacion no hay ninguno.
- Regenerar: `.venv/bin/python NuevaData/_build/build_nuevadata.py` (unos 10 s). El script borra y vuelve a crear
  `Original/` y `Presentacion/`. Solo lee `Data/Bronze`, `Data/Silver`, `SolicitudesPares/` y el consolidado
  de la raíz, y no modifica ninguno de ellos.
- `.gitignore` excluye `NuevaData/Original/`, `NuevaData/Presentacion/` y `NuevaData/_build/_privado/`.

## 2. Reglas aplicadas en todo Presentacion/

1. **Periodo 2024-2027**: plan anterior 2024-2026 (formato AA-FR-001) y plan vigente 2026-2027 (CC-FR-001).
2. **Meta del plan anterior visible en cada factor.** Cada libro nuevo abre con el bloque «Metas del plan
   de mejoramiento»: meta 2024-2026, meta 2026-2027 y la solicitud de los pares. El índice
   `00_Indice_Evidencias_<Mod>.xlsx` las reúne para los 12 factores.
3. **Sin datos sensibles.** Presentacion no incluye cédulas, códigos estudiantiles, teléfonos, correos
   personales, firmas ni nombres de estudiantes en planillas o solicitudes.
   - Las planillas de asistencia y los correos con destinatarios quedan **solo en Original/**. Los libros
     citan su ruta como «Soporte original (contiene datos personales)».
   - Las solicitudes de prórroga se identifican con un ID (S-01…S-25). La relación ID → archivo
     está en `_privado/transcripciones.json` y en `Original/_transcripciones/`.
   - Se verificó con un escaneo automático que no quedan códigos de 11 dígitos ni correos en los libros
     generados. Los únicos correos en archivos copiados son institucionales: decanaturas en
     `Open Day.pdf`, Egresados y los docentes de `Directorio Grupos de Inv MCIC.xlsx`, que ya se
     publica en el micrositio.
   - Los nombres de autores y directores de trabajos de grado (FACTOR 8) se conservan porque son
     información pública (RIUD).
4. **Sin archivos sueltos ni duplicados.** Las copias de evidencias del FACTOR 2 que estaban sueltas en los
   FACTORES 8 y 9 no se repiten en Presentacion; su información quedó en los libros. Los duplicados
   «(1)» del FACTOR 2 tampoco se copian.
5. **Separación Investigación / Profundización en todo lo que se pudo.** Cada estudiante se clasifica con
   esta prioridad (función `modalidad_de` del script):
   1. modalidad declarada en la solicitud (p. ej. un cambio de modalidad aprobado);
   2. código de proyecto en el código estudiantil (`AAAAP595NNN` = Investigación, `AAAAP695NNN` = Profundización);
   3. proyecto actual en Cóndor (`Data/Bronze/Estados/Listado_de_estudiantes_por_estado_*.csv`, 595/695);
   4. bases MCIC 2026 (`Maestria CIC/2026/BASES  DE DATOS ESTUDIANTES/`: hojas *N-A Investigación*,
      *N-A Profundizacion*, *Pasantías*) y después `MCIC - Base de datos V2.xlsx`.

   Lo que la fuente no permite separar se marca como compartido y se explica en el propio libro:
   planta docente, syllabus y estadísticas de Bienestar.

## 3. Qué hay por factor

Los estados coinciden con la columna «Estado» del índice.

| N° | Presentacion/<Mod>/FACTOR N | Estado | Diferencia INV / PROF |
|---|---|---|---|
| 1 | `a. PEP/` (PEP propio de la modalidad) y `b. Jornadas de trabajo con docentes/` (actas de Ing. Software y Geomática) | Sin cambios (aprobado por los pares) | El PEP es distinto en cada modalidad |
| 2 | `a. Divulgación…/F2_Divulgacion_y_Estudiantes_<Mod>.xlsx` + `soportes/` (capturas web, fotos, Open Day.pdf) | Actualizado | Página web propia de cada modalidad, conteos por modalidad y cobertura por cohorte |
| 3 | `a. Informes…/F3_Participacion_Docente_Capacitacion.xlsx` | Actualizado (en recolección) | Mismo libro en ambas: la planta docente es común (Cuadro Maestro No. 05 idéntico en los dos SNIES) |
| 4 | `a. Servicios/` Portafolio de Servicios + Proyecto de Acuerdo Política de Egresados | Actualizado (se retiraron *Experiencias UD* e *Infografía Esquema Normativo*) | Compartido |
| 5 | `a. …/Res 016 de 2025…pdf`, `b. …/Syllabus AA-FR-003/<área>/*.xlsx` (29) y `F5_Verificacion_Syllabus_Investigacion_vs_Profundizacion.xlsx` | Verificado | Ver §4 |
| 6 | `d. Seguimiento…/F6_Permanencia_y_Graduacion_<Mod>.xlsx` | Actualizado | INV: 19 prórrogas, 32 PAGOT, 72 graduados 2022-2026. PROF: 6 prórrogas, 19 PAGOT, 56 graduados. En ambos se agregan 25 graduados del plan anterior sin modalidad registrada |
| 7 | `d. Definición de acción para la formalización de convenios/` (8 normas institucionales) | Sin cambios (solo se ubicaron en una actividad) | Compartido |
| 8 | `a. …/Directorio Grupos de Inv MCIC.xlsx`, `b. …/F8_Trabajos_de_grado_por_grupo_<Mod>.xlsx` y, solo en INV, `ANEXOS PONENCIAS/` (26) | Actualizado (para revisión) | Consolidado separado: INV 66 casos, PROF 52 (incluye 3 pasantías), 30 «por confirmar» en ambos. Las ponencias vienen del paquete SNIES 17528 (Investigación); 18 autores confirmados como de Investigación y ninguno de Profundización |
| 9 | `a. …/Bienestar MCIC.pptx`, `b. …/acuerdo_02_2019_beca_ecaes.pdf`, `res_2025-143…pdf` (solo INV) y `F9_Bienestar_<Mod>.xlsx` | Actualizado | La Res. 143 de 2025 (Programa de Excelencia Académica) aplica solo a maestrías de investigación. Las estadísticas de Bienestar son de la Maestría en conjunto |
| 10 | `a. Diagnóstico…/Ambientes de Aprendizaje Facultad Ing.png` | Sin cambios | Compartido |
| 11 | `c. Implementación…/` CC-FR-001 y Autoevaluación de la modalidad; `d. Construcción de reportes/` análisis de instrumentos y resultados 2026-1 | Sin cambios (solo se ubicaron en actividades) | Plan y autoevaluación propios de cada modalidad |
| 12 | `a. Solicitud de informe de avance de la obra/G312-3 Requerimientos…pdf` | Sin cambios | Compartido |

Detalle de los libros nuevos:

- **F2**: hojas *Resumen*, *Páginas web* (enlaces vigentes `facingenieria.udistrital.edu.co/mcic-investigacion/`
  y `…/mcic-profundizacion/`), *Actividades* (12 en INV, 10 en PROF; dos correos salieron desde la cuenta de
  Investigación), *Cobertura cohortes* y *Galería* (miniaturas con enlace a `soportes/`).
  - Inducción 2026-1: 15 asistentes (5 de INV, 10 de PROF).
  - Presentación de grupos del 21/02/2026: 15 estudiantes MCIC (5 INV, 10 PROF), 2 de otros programas y
    9 docentes o representantes de grupos.
  - Socialización del 22/08/2026: 13 estudiantes MCIC (4 INV, 9 PROF) en una planilla de 55 registros.
  - Cobertura de la cohorte 2026-1: INV 5/6 (83 %), PROF 10/14 (71 %).
- **F3**: hojas *Resumen* (30 profesores, 13 con capacitación registrada, 43 % frente a la meta del 15 %),
  *Capacitación docente* (con grupo y categoría MinCiencias tomados del Cuadro Maestro No. 05, y una
  columna vacía «Capacitaciones 2025-2026 reportadas por el profesor» para lo que llegue por correo) y
  *Movilidad*. Sin cédulas.
- **F5**: hojas *Resumen* (hallazgos), *Plan de estudios Res 016*, *Syllabus AA-FR-003* y *Comparación SNIES*.
- **F6**: hojas *Resumen*, *Solicitudes de prórroga* (ID, fecha, énfasis, PAGOT, componente 1 / evento,
  componente 2, estado en Cóndor 2026-3), *Acompañamiento* (reunión del 06/04/2026: 18 INV, 2 PROF),
  *PAGOT* (por ingreso, estado y énfasis) y *Graduados*. El año de grado se estima con la última
  matrícula, igual que en el micrositio; la suma 72 + 56 + 25 = 153 cuadra con Gold.
- **F8**: hojas *Resumen* (recalculado por modalidad), *Consolidado*, *Pendientes sustentación*,
  *Modalidad por confirmar* y *Fuentes y calidad* (copiada del original). Se omiten el código estudiantil y
  la ruta del archivo fuente, y en las observaciones los códigos se reemplazan por «[código]».
- **F9**: hojas *Resumen*, *Divulgación*, *Estímulos y becas*, *Uso de servicios* (Bienestar MCIC.pptx,
  2020 – 2026-1) y *Cuadro Maestro No. 10*.

## 4. Hallazgos que necesitan decisión de la coordinación

1. **FACTOR 5: los syllabus NO se diferencian por modalidad.**
   - Los 28 archivos de `SNIES17528-Syllabus` y de `SNIES116070-Syllabus` son byte a byte idénticos.
   - La carpeta FACTOR 5 de Profundización contiene la carpeta de Investigación (`SNIES17528-Syllabus`).
   - 27 de los 29 syllabus AA-FR-003 dejan vacío el «Código plan de estudios» y los otros 2 (Redes y
     Política sectorial…) registran «919». Ninguno indica la modalidad.
   - En cambio, **el plan de estudios sí se diferencia** (Res. 016 de 2025): INV tiene TG I + TG II
     (12 créditos), 3 énfasis y los espacios del periodo III como electivos; PROF tiene TG (4 créditos) y
     5 énfasis. Las horas son 384/176/1552 frente a 480/160/1472.
   - Otros detalles: la malla `Información Espacios Académicos.xlsx` registra el proyecto 595 en la hoja
     de Profundización (debería ser 695). «Matemática avanzada y geoprocesamiento» tiene syllabus pero no
     aparece en la Res. 016. Hay denominaciones distintas entre la Res. 016 y los syllabus.
   - Decisión: ¿se diligencia el código de plan y la modalidad en los syllabus, o se aclara ante los pares
     que el espacio académico es común y lo que cambia es la estructura del plan?
2. **FACTOR 8**: 30 trabajos del consolidado no traen modalidad ni código para cruzar. Están en la hoja
   «Modalidad por confirmar» (la misma en ambos libros). También hay que confirmar si se acepta
   Pasantía = Profundización.
3. **Pendientes de insumo (no se inventó nada):**
   - F3: correo de Karol a los profesores pidiendo capacitaciones 2025-2026.
   - F4: formulario a egresados 2022-2026 (tarea enviada a la OATI).
   - F7: correo a profesores sobre convenios.
   - F10: Biblioteca y Planes TIC.
   - F11: «las 2 autoevaluaciones y encuestas»; ya están las que había en Bronze.
   - F12: laboratorios.
4. **Ubicación de archivos sueltos (F7, F10, F11, F12).** Se asignaron así; ajustar si la coordinación
   prefiere otra actividad:
   - F7: las 8 normas → `d. Definición de acción para la formalización de convenios`.
   - F11: CC-FR-001 y Autoevaluación → `c.`; análisis y resultados → `d.`.
   - F12: G312-3 → `a.`.
5. **Numeración del pedido.** El punto «7» del pedido («montaremos el consolidado de trabajos de grado»)
   se interpretó como **FACTOR 8**, porque así lo indica `CorreccionEvidencias.ods` («Cargar excel de
   consolidado»). El FACTOR 7 quedó sin cambios de contenido.

## 5. Cambios que YA se hicieron en el micrositio (fuera de NuevaData)

- `app/extract_plan_mejoramiento_2025.py` ahora extrae del plan anterior la **línea base (col. J), la meta
  (col. K) y las actividades (col. L)**, y respeta el periodo real de la celda C9 («2024 - 2026»). Antes se
  forzaba «2025 - 2026».
  - En la fila del FACTOR 2 las columnas K y L vienen intercambiadas en el Excel fuente. Se detecta y
    queda explicado en `meta_nota`.
  - Los FACTORES 1, 7 y 8 difieren entre los archivos INV y PROF; el resto es idéntico.
- `app.js` (`renderPlanAnteriorSeccionHtml`) muestra «Qué se planteó en 2024-2026 (plan anterior)» con
  **Línea base** y **Meta 2024-2026**.
- `index.html`: el parámetro de caché pasó a `?v=20260930-v1`.
- Se regeneraron `Data/Silver/plan_mejoramiento_2025_*.json` y `Data/Gold/gold_data.{json,js}` (solo con
  `extract_plan_mejoramiento_2025.py` y `build_gold.py`). La comparación con el Gold anterior mostró que lo
  único que cambió fue el agregado de `linea_base`, `meta`, `meta_nota` y `actividades` en `plan_anterior`.
- Nada de esto está commiteado.

## 6. Cómo integrar a Bronze → Silver → Gold cuando se apruebe

Hacerlo **por modalidad y por factor**, respetando los ajustes que pida la coordinación.

1. **Respaldar** `Data/Bronze/MCIC.INVESTIGACION/Procesos de Renocavion y acreditación/Plan de Mejoramiento`
   y `Data/Bronze/MCIC-PROFUNDIZACION/…/Plan de Mejoramiento` (por ejemplo, copiarlos al scratchpad).
2. **Originales con datos personales → `Data/Bronze/PII_Interno/`**, que `build_bronze_manifest.py`
   excluye del catálogo:
   `NuevaData/Original/<Mod>/` → `Data/Bronze/PII_Interno/Evidencias_Originales_Plan_Mejoramiento/<Mod>/`
   (y `_transcripciones/`). Así se conserva la data original sin publicarla en el sitio.
3. **Reemplazar** en Bronze el contenido de cada `FACTOR N. …` por el de
   `NuevaData/Presentacion/<Mod>/FACTOR N. …`:
   - `Investigacion` → `MCIC.INVESTIGACION/…/Plan de Mejoramiento/`
   - `Profundizacion` → `MCIC-PROFUNDIZACION/…/Plan de Mejoramiento/`

   Hay que borrar lo anterior de cada FACTOR para que no queden sueltos ni duplicados. El
   `00_Indice_Evidencias_<Mod>.xlsx` va en la raíz de `Plan de Mejoramiento/`: el extractor de evidencias
   lo ignora y el catálogo lo lista.
4. Quitar `Entregables Plan de Mejoramiento.xlsx` (documento interno de trabajo, solo en INV) de la raíz
   de Bronze; ya está en Original.
5. **Correr la pipeline completa**: `cd app && ../.venv/bin/python run_pipeline.py`.
6. **Revisar**:
   - `Data/Silver/seguimiento_evidencia_*.json`: ningún FACTOR debe tener «(sin actividad específica)».
   - Abrir `index.html`: modal de cada factor → «Evidencia de seguimiento cargada en Data/Bronze».
   - Catálogo de Bronze: no debe aparecer nada de `PII_Interno`.
7. **Opcional (mejoras al micrositio):**
   - FACTOR 8: renderizar en el modal el resumen del consolidado por modalidad (sustentados, pendientes,
     por grupo) leyendo `F8_Trabajos_de_grado_por_grupo_<Mod>.xlsx` desde un extractor nuevo
     (`app/extract_trabajos_grado_consolidado.py`), en vez de mostrarlo solo como archivo.
   - FACTOR 2 / 6 / 9: mostrar los KPIs de la hoja *Resumen* de cada libro.
   - `extract_syllabi.py` hoy lee `Data/Bronze/Syllabus/`. Si la coordinación diligencia el código de
     plan / modalidad en los syllabus, volver a correr y mostrar la diferencia.
8. Subir el número de versión `?v=` en `index.html` y commitear (el código, no `Data/`).

## 7. Otros datos útiles

- Correspondencias de códigos:
  - Proyectos: 595 = MCIC Investigación (SNIES 17528); 695 = MCIC Profundización (SNIES 116070).
  - Énfasis del plan anterior: 195 Teleinformática, 295 Sistemas de Información, 395 Geomática,
    495 Ingeniería de Software.
  - Código estudiantil = `AAAA` + periodo (`1`/`2`) + proyecto (3 dígitos) + consecutivo. Los periodos de
    ingreso se nombran 2026-1 y 2026-3.
- Cohortes nuevas según Cóndor:
  - 2026-1: 6 INV, 14 PROF.
  - 2026-3: 6 INV, 16 PROF.
  - Matriculados actuales: 49 INV, 58 PROF.
- Cuando se hizo este trabajo (30/09/2026) ya no estaban en el árbol de trabajo las carpetas `Data2/` y
  `Presentacion/` (LaTeX) de la raíz. Siguen en git (`git checkout -- Data2 Presentacion` las recupera).
  NuevaData no depende de ellas.
