# NuevaData — evidencias del Plan de Mejoramiento (aprobadas con ajustes, listas para revisión final)

Este documento deja todo el contexto para que, en otra sesión, se pueda pedir:
**«integra NuevaData a Bronze, Silver y Gold»** una vez el usuario revise `Presentacion/` y lo confirme.

- Primera versión: 30 de septiembre de 2026 (correcciones de los pares, `CorreccionEvidencias.ods`, raíz del repo).
- **Segunda versión: 1 de octubre de 2026.** La coordinación aprobó NuevaData con 13 comentarios (§4) y cargó
  insumos nuevos directamente en `Original/`. Esta versión los incorpora.
- **Estado: INTEGRADO (1-oct-2026).** `Presentacion/` reemplazó los FACTOR de `Data/Bronze/.../Plan de Mejoramiento` (pipeline completa corrida, Gold regenerado).
  Las subcarpetas `Anexos/` se listan en el sitio como «Anexos» plegables (extractor + `app.js`). Los originales con datos personales
  NO se publicaron: siguen solo en `NuevaData/Original/` (gitignored), no en `Data/Bronze/PII_Interno/` (que está versionado).
- Los pares valoraron conservar la data original, pero no que estuviera «esparcida»: por eso se entrega la
  original (`Original/`) y la analizada (`Presentacion/`), siempre separadas por modalidad.

---

## 1. Estructura

```
NuevaData/
├── CONTEXTO_NuevaData.md            ← este documento (versionado en git)
├── _build/
│   ├── build_nuevadata.py           ← genera Presentacion/ a partir de Original/ (idempotente)
│   ├── xlsx_style.py                ← estilo común de los .xlsx
│   └── _privado/transcripciones.json← transcripción manual de planillas (DATOS PERSONALES, fuera de git)
├── Original/                        ← FUENTE. Fuera de git (datos personales). El script NO la borra ni la modifica
│   ├── Investigacion/FACTOR 1…12/   ← copia de Bronze (plan de mejoramiento) + lo que cargó la coordinación
│   ├── Profundizacion/FACTOR 1…12/
│   └── _transcripciones/transcripciones.json
└── Presentacion/                    ← fuera de git (se sube a SharePoint para la aprobación)
    ├── Investigacion/ 00_Indice_Evidencias_Investigacion.xlsx + FACTOR 1 … FACTOR 12/<actividad>/<archivos>
    └── Profundizacion/ 00_Indice_Evidencias_Profundizacion.xlsx + FACTOR 1 … FACTOR 12/<actividad>/<archivos>
```

- Las carpetas `FACTOR N. …` usan **exactamente** los mismos nombres que en Bronze, así que
  `app/extract_seguimiento_evidencia.py` las reconoce sin cambios.
- En `Presentacion/` **ningún archivo queda suelto**: todos están dentro de una carpeta de actividad.
- Regenerar: `.venv/bin/python NuevaData/_build/build_nuevadata.py` (unos 15 s). Solo borra y recrea
  `Presentacion/`. Lee `Original/`, `Data/Bronze` (Cóndor y bases MCIC), `Data/Silver`, `SolicitudesPares/` y no modifica nada de eso.
  **Si se agrega o cambia un insumo, se coloca en `Original/<Modalidad>/FACTOR N…` (en ambas modalidades) y se vuelve a correr.**
- `.gitignore` excluye `NuevaData/Original/`, `NuevaData/Presentacion/` y `NuevaData/_build/_privado/`.
- Nota técnica: los nombres de archivo de `Original/` pueden venir con tildes en distinta normalización Unicode
  (p. ej. `Investigacio╠ün`); el script no depende de ellas salvo por los nombres exactos listados en el código.
- Se dejó un respaldo de la `Presentacion/` anterior (con carpetas duplicadas `FACTOR 01…09` creadas por la sincronización)
  en el directorio temporal de la sesión; `Presentacion/` se regeneró completa.

## 2. Reglas aplicadas en todo Presentacion/

1. **Periodo 2024-2027**: plan anterior 2024-2026 (formato AA-FR-001) y plan vigente 2026-2027 (CC-FR-001).
2. **Meta del plan anterior visible en cada factor** (bloque «Metas del plan de mejoramiento» al abrir cada libro) y
   en el índice `00_Indice_Evidencias_<Mod>.xlsx`, que ahora incluye una columna con el **comentario de la coordinación (oct. 2026)**.
3. **Sin datos sensibles.** Presentacion no incluye cédulas, códigos estudiantiles, teléfonos, correos personales, firmas
   ni nombres de estudiantes o egresados.
   - Prórrogas: ID `P<ronda>-<n>` (P1-01…). PAGOT: ID `G-01…`. Egresados: ID `E-001…`. La relación ID → persona no se publica.
   - Los libros de egresados no traen empleador ni cargo por persona; las instituciones solo se cuentan agregadas.
   - Los correos, planillas, respuestas de encuestas 2025 (`DB mensajes`, `Información empleadores`, bases de respuestas)
     y soportes de las prórrogas quedan solo en `Original/`.
   - Los nombres de autores y directores de trabajos de grado (FACTOR 8) se conservan: son públicos (RIUD).
   - Se verificó con un escaneo automático que los libros generados no contienen correos personales, códigos, documentos
     ni nombres de los estudiantes/egresados (el único correo es institucional: `convenios-ceri@udistrital.edu.co`).
4. **Sin archivos sueltos ni duplicados.** `CC-FR-001 P.xlsx` (Autoevaluación anterior, INV) es idéntico byte a byte al
   plan vigente `CC-FR-001 Plan de mejoramiento INV.xlsx`, por eso se entrega una sola vez.
5. **Separación Investigación / Profundización en todo lo que se pudo.** Cada estudiante se clasifica con
   esta prioridad (función `modalidad_de` del script):
   1. modalidad declarada (cuadro de prórrogas, archivo PAGOT o consolidado);
   2. código de proyecto en el código estudiantil (`AAAAP595NNN` = Investigación, `AAAAP695NNN` = Profundización);
   3. proyecto actual en Cóndor (`Data/Bronze/Estados/Listado_de_estudiantes_por_estado_*.csv`, 595/695);
   4. bases MCIC 2026 (`Maestria CIC/2026/BASES  DE DATOS ESTUDIANTES/`) y después `MCIC - Base de datos V2.xlsx`;
   5. (solo FACTOR 8) nombre del estudiante contra Cóndor, si identifica a una única persona; queda marcado «Inferida por nombre».

   Lo que la fuente no permite separar se marca como compartido y se explica: planta docente, syllabus, Bienestar,
   Biblioteca/PlanEsTIC, convenios URELINTER y los egresados del plan anterior (antes de que existieran las modalidades).

## 3. Qué hay por factor (versión 2)

| N° | Qué se entrega en Presentacion/<Mod>/FACTOR N | Estado | Diferencia INV / PROF |
|---|---|---|---|
| 1 | `a. PEP/` (PEP propio de cada modalidad **+ PDF «Orientaciones PFA» 2026IE4045…**) y `b. Jornadas…` | Actualizado; **falta el PEP definitivo** | El PEP es distinto en cada modalidad; el PDF PFA es institucional |
| 2 | `F2_Divulgacion_y_Estudiantes_<Mod>.xlsx` + `soportes/` | Actualizado (aprobado) | Páginas web, conteos y cohortes por modalidad |
| 3 | `F3_Participacion_Docente_Capacitacion.xlsx` | Se sube así; **mañana se ajusta** con lo que cargan los profesores | Planta docente común |
| 4 | `a. Servicios/` (2 PDF) + **`e. Caracterización e impacto de los egresados/F4_Egresados_Impacto_<Mod>.xlsx`** | Actualizado | Ver §3.1 |
| 5 | Res. 016, `F5_Verificacion_Syllabus…xlsx` (5 hallazgos), syllabus por área | Actualizado; **faltan los syllabus en PDF** | Ver §4 |
| 6 | `F6_Permanencia_y_Graduacion_<Mod>.xlsx` + `Normativa_PAGOT_UD.xlsx` | Actualizado | Ver §3.2 |
| 7 | **`a. Diagnostico de convenios vigentes/F7_Convenios_URELINTER_<Mod>.xlsx`** (enlace oficial) + 8 normas en `d.` | Actualizado | Compartido |
| 8 | `F8_Trabajos_de_grado_por_grupo_<Mod>.xlsx` (con comparativo INV vs PROF), Directorio de grupos, ponencias (solo INV) | Actualizado | Ver §3.3 |
| 9 | `Bienestar MCIC.pptx`, acuerdo 02/2019, `res_2025-143` (solo INV), `F9_Bienestar_<Mod>.xlsx` | Actualizado (sin cambios de fondo) | Estadísticas de Bienestar compartidas |
| 10 | `b.` Biblioteca (pptx + anexo estadístico xlsx), Planes TIC (pptx), `F10_Medios_Educativos_<Mod>.xlsx`; `a.` ambientes | Actualizado | Compartido (institucional) |
| 11 | `c.`/`d.` con subcarpetas **«Autoevaluación 2025 (anterior)»** y **«Autoevaluación 2026 (vigente)»** + `F11_Autoevaluacion_2025_vs_2026_<Mod>.xlsx` | Actualizado | Informe y plan propios de cada modalidad |
| 12 | `c.` `Laboratorios Maestría - MIC.pptx` con **última diapositiva actualizada** + `F12_Estudiantes_activos_por_enfasis_<Mod>.xlsx`; `a.` G312-3 | Actualizado | Tabla por modalidad |

### 3.1 Factor 4 — informe de impacto de egresados
- Fuente: `HojaVidaEgresados.xlsx` (módulo Hoja de Vida de la OATI): **146 registros, 145 egresados distintos**, 7 hojas
  (datos básicos, experiencia docente/profesional, grupos y productos de investigación, obras, segunda lengua).
  Es el insumo de la meta 2026-2027 del factor (≥ 50 % de egresados con hoja de vida actualizada). Los archivos de INV y PROF
  tienen el mismo contenido.
- Separación por modalidad (cruce por documento con Cóndor + bases MCIC): **Investigación 18, Profundización 6,
  plan anterior sin modalidad 119, sin registro en Cóndor 2**. 135 son graduados; 10 son personas que hoy figuran
  matriculadas, inactivas o suspendidas (los «profes y demás»; columna «Estado» del perfil).
- Libro: *Resumen*, *Indicadores* (las 4 categorías lado a lado), *Perfil egresados* (anónimo), *Sector y vinculación*
  (sector estimado por palabras clave del nombre de la entidad), *Instituciones* (agregadas), *Cobertura* (frente a graduados
  de Cóndor: 23 % en total; INV 14/76, PROF 4/75) y *Alcance y límites*.
- Límites: el módulo no trae sector económico ni fechas de actualización; los egresados 2022-2026 están poco cubiertos
  (el formulario a la OATI sigue pendiente). Con 6 egresados de Profundización los porcentajes son indicativos.

### 3.2 Factor 6 — prórrogas 2026-3 y PAGOT
- Fuente de las prórrogas: carpeta `PRORROGA 2026-3/` (1.ª, 2.ª y 3.ª solicitud; mismo contenido en ambas modalidades) y sus tres
  cuadros «Estudiantes Prórroga…xlsx», que traen la **modalidad declarada**. La fecha sale del correo de cada estudiante.
  **38 solicitudes = 37 estudiantes** (una estudiante radicó en la 2.ª y la 3.ª ronda): **INV 27 filas / 26 estudiantes
  (16 PAGOT); PROF 11 (4 PAGOT)**. Reemplaza las 25 solicitudes transcritas de la versión 1.
- PAGOT: `PAGOT 20263.xlsx` (seguimiento 2026-3: 35 estudiantes; modalidad = «Plan vigente» 595/695) → **INV 22 (17 con pago), PROF 13 (9 con pago)**;
  `Aspirantes Inscritos 2023-3 al 2025-3…xlsx` (inscritos oficiales a PAGOT: INV 30, PROF 14 inscripciones = 13 estudiantes); `Normativa_PAGOT_UD.xlsx` (9 normas, con enlaces).
  La hoja *PAGOT* cruza cuatro fuentes: base MCIC (INV 32 / PROF 19), inscritos oficiales, seguimiento 2026-3 y prórrogas.
- Hallazgo del cruce: **el seguimiento PAGOT 2026-3 y las prórrogas de PAGOT no comparten ningún estudiante** (son cohortes distintas:
  quienes piden prórroga están en PAGOT desde 2025-1; el seguimiento 2026-3 reúne a quienes se invitó a reingresar).
- Alertas (hoja *Cruces y alertas*): 1 solicitud con aval «NO» y fecha de aval de 2015 (probable error de digitación), 1 con modalidad del cuadro
  distinta de las bases MCIC, y **7 de los 13 PAGOT con plan vigente 695 figuran como Investigación en las bases MCIC** (se usó la del archivo PAGOT).
- Se conservan *Acompañamiento* (reunión del 06/04/2026) y *Graduados* (INV 72, PROF 56, 25 sin modalidad registrada).

### 3.3 Factor 8 — proyectos de investigación vs profundización
- El nuevo `Consolidado_trabajos_grado_MCIC_2022_2026 (7).xlsx` de `b. Socialización…` tiene el **mismo contenido** que el de la raíz
  (solo cambian los metadatos). Se agregó una hoja *Comparativo por grupo* y una tabla comparativa en *Resumen*, y una columna «Tipo de trabajo».
- Resultado: INV 74 trabajos (67 sustentados), PROF 55 (47 sustentados, 3 pasantías), **19 por confirmar** (antes 30). 11 casos se
  completaron por el nombre del estudiante contra Cóndor/bases (8 INV, 3 PROF) y quedan marcados «Inferida por nombre; verificar».

### 3.4 Factor 12 — última diapositiva de Laboratorios
- `Laboratorios Maestría - MIC.pptx` (mismo archivo en ambas modalidades). La diapositiva 13 («Estudiantes Impactados») estaba vacía:
  se le agregó una tabla nativa de **estudiantes activos por énfasis 2022-2026** de la modalidad, más una fila aparte para el plan anterior
  sin modalidad y una nota de método. El original sigue en `Original/`.
- Definición: «activo en el año Y» = Y entre el año de ingreso (código) y el de la última matrícula en Cóndor. **Es una aproximación**: no se
  conocen periodos intermedios sin matrícula. El énfasis es el del plan de ingreso (195 Teleinformática, 295 Sistemas de Información,
  395 Geomática, 495 Ingeniería de Software; para 595/695, de las bases MCIC). 2026: INV 94, PROF 92.
- Verificado solo por contenido (python-pptx): **no hay LibreOffice en el equipo para renderizar**; conviene abrirlo en PowerPoint y revisar el aspecto.

## 4. Comentarios de la coordinación (oct. 2026) y cómo se resolvieron

1. **F1**: aún falta el PEP; se entrega el PDF `2026IE4045-GC-RemisionEvaluacionPFA-ANEXO1.pdf` (Orientaciones para la evaluación de los
   propósitos de formación y aprendizaje, Vicerrectoría Académica, 54 págs.) en `a. PEP` junto al PEP de cada modalidad.
2. **F2**: se actualiza con lo ya elaborado.
3. **F3**: se sube igual; hay que rehacerlo con las capacitaciones que reporten los profesores (columna J de «Capacitación docente»).
4. **F4**: nuevo informe de impacto (§3.1).
5. **F5**: se eliminaron 3 hallazgos negativos. **Supuesto mío (confirmar)**: se retiraron los de *Malla «Información Espacios Académicos.xlsx»*,
   *Espacio sin correspondencia en la Res. 016* y *Denominaciones*, y se conservan *Plan de estudios*, *Horas*, *Espacios académicos*,
   *Archivos de syllabus* e *Identificación de la modalidad*. Si eran otros, se cambia la lista `hallazgos` en `factor5()`. Pendiente: pasar los syllabus a PDF.
   - Lo que sigue pendiente de decisión (de la v1): los syllabus no se diferencian por modalidad (los 28 de SNIES 17528 y 116070 son idénticos;
     27 de 29 syllabus AA-FR-003 dejan vacío el código de plan); el plan de estudios (Res. 016/2025) sí se diferencia.
6. **F6**: §3.2.
7. **F7**: se deja el enlace `https://urelinter.udistrital.edu.co/convenios/cooperacion-redes-asociaciones` y un resumen institucional
   (388 convenios vigentes: 221 internacionales, 167 nacionales; 7 mencionan a la Facultad de Ingeniería). Pendiente: correo a profesores.
8. **F8**: §3.3.
9. **F9**: sin cambios de fondo.
10. **F10**: se agregan Biblioteca y Planes TIC. El anexo estadístico de Biblioteca solo estaba en la carpeta de Investigación; se entrega en ambas por ser institucional.
11. **F11**: autoevaluación 2025 (informe de octubre de 2025 con fines de renovación, encuestas de marzo de 2025) y 2026 (proceso permanente, resultados 2026-1)
    en carpetas separadas. En la parte 2025 solo van el informe, los gráficos de resultados (PNG) y las invitaciones; las bases de respuestas tienen datos personales.
    El libro comparativo trae respondientes de cada año y el promedio por factor de 2026 (no hay equivalente 2025 sin abrir datos personales).
12. **F12**: §3.4.
13. **F8** (diferenciar proyectos): §3.3.

Pendientes de insumo (no se inventó nada): F3 (capacitaciones 2025-2026), F4 (formulario de sector a egresados 2022-2026, enviado a la OATI),
F7 (correo a profesores sobre convenios), F5 (syllabus en PDF), F1 (PEP definitivo).

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
   (y `_transcripciones/`). Así se conserva la data original sin publicarla en el sitio. `Original/` ya incluye lo que cargó la
   coordinación (≈770 archivos: prórrogas 2026-3, PAGOT, egresados, encuestas 2025, presentaciones); pasa completo.
3. **Reemplazar** en Bronze el contenido de cada `FACTOR N. …` por el de
   `NuevaData/Presentacion/<Mod>/FACTOR N. …`:
   - `Investigacion` → `MCIC.INVESTIGACION/…/Plan de Mejoramiento/`
   - `Profundizacion` → `MCIC-PROFUNDIZACION/…/Plan de Mejoramiento/`

   Hay que borrar lo anterior de cada FACTOR para que no queden sueltos ni duplicados. El
   `00_Indice_Evidencias_<Mod>.xlsx` va en la raíz de `Plan de Mejoramiento/`: el extractor de evidencias
   lo ignora y el catálogo lo lista.
   - Algunos factores traen subcarpetas dentro de la actividad (F5 `Syllabus AA-FR-003/<área>`, F11 `Autoevaluación 2025 (anterior)` /
     `Autoevaluación 2026 (vigente)`, F2 `soportes`): el extractor las recorre con `rglob`, no requiere cambios.
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


## 8. Ajustes de octubre (Factor 4)

- Encuesta de caracterización e impacto de egresados (18 respuestas: INV 11, PROF 7) en `Original/<Mod>/FACTOR 4…/`; hoja «Encuesta de impacto» en cada F4 (agregados, sin texto libre).
- **Regla: el plan anterior (proyectos 95-495 sin modalidad) se toma como Investigación** (F4, graduados de F6, activos de F12; `app/lib/modalidad.py` para Silver/Gold). Cóndor: graduados 577 = INV 502 + PROF 75; 2022-2026: 153 = INV 97 + PROF 56.
- OATI: 146 registros / 145 personas / 143 clasificadas (INV 137, PROF 6; 2 sin registro en Cóndor quedan fuera).
- Gold: el bloque de graduados del Factor 4 ahora es por modalidad y la sección Comunidad estudiantil suma una tabla «Graduados por modalidad» (los totales por proyecto/año no cambian).
- Discrepancia abierta: la tabla de énfasis de Comunidad (archivo «Énfasis estudiantes - Consolidado», 2026-1) da INV 79 / PROF 77 matriculados; Cóndor+bases (2026-3) da INV 83 (con 2 sin base) / PROF 73. El bloque «Matriculados» del Factor 4 sigue mostrando el total MCIC (156).

## 9. Ajustes posteriores (2-oct-2026)

- F6: pagaron = recibo «2026-3 A» (20: INV 13, PROF 7); se quitaron nota del Resumen, «Cruces y alertas» y «Modalidad por confirmar». Plan anterior = Investigación en todo (F2, F4, F6, F8, F9, F12); no debe quedar «sin modalidad».
- F8: «Pendientes por sustentar»; «LASER LAMIC» → LASER (también en Silver: `grupos_investigacion`, `seguimiento_tesis`); sin hoja «Fuentes y calidad». 9 trabajos sin registro en Cóndor se asignan a Investigación (plan anterior), marcados «verificar» en «Nota de modalidad».
- F9: «Asistieron a la inducción»; matriculados por modalidad con Cóndor + regla de plan anterior.
- F11: `app/extract_autoevaluaciones.py` → `Data/Silver/autoevaluaciones.json` → bloque «Autoevaluación 2025 (anterior) y 2026 (vigente)» en el detalle del factor. Las respuestas 2025 (53/17/36) están solo en la carpeta de Investigación y se usan para ambas modalidades.
- F12: solo recursos físicos: PPT de laboratorios sin la diapositiva «Estudiantes impactados», G312-3 y libro de salas de informática y software.
- F7: `NuevaData/Convenios/Convenios vigentes URELINTER (3).xlsx` es idéntico (388 filas) al de `Data/Bronze/Convenios/`; la tabla del sitio solo lista los convenios que aplican a posgrado (3 de 7 relacionados con la Facultad de Ingeniería).

## 10. Insumos de F7, F8 y F12 (2-oct-2026)

- F7: `NuevaData/Convenios/…(3).xlsx` (descarga 2-oct) tiene los mismos 388 convenios que el anterior; ahora la tabla del modal lista los 388 con filtros (primero los que aplican a la MCIC y los relacionados con Ingeniería). Silver `convenios.json` incluye `convenios`.
- F8: «Informe Grupos Investigación.docx» (impacto social por grupo) va en `b.` (compartido); la sección «LASER LAMIC» se integró a LASER (7 trabajos, 6 sustentados, igual que el consolidado). Presentaciones de GIIRA y Multimedia en `a.` (los videos incrustados quedan como imagen fija: el pptx de GIIRA pesaba 192 MB, sobre el límite de GitHub). Proyectos ejecutados y tesis en `b./Anexos`; las tesis se asignan por título al consolidado y por modalidad (INV 8, PROF 4). Los 11 trabajos del xlsx «Tesis Maestria Grupo Multimedia GIIRA» coinciden con el consolidado.
- F12: planos de laboratorios del nuevo edificio (81 MB) en `a.`.
- Los originales pesados están en `NuevaData/Info para 12 y 8/` (en `.gitignore`).

## 11. Convenio IGAC 5570 de 2025 (2-oct-2026)

- `NuevaData/Convenios/CONVENIO IGAC/` (7 PDF): convenio específico IGAC-UDFJC N.° 5570 de 2025 (maestrías y doctorados en geografía, geomática, catastro, IA y ciencia de datos; nació de la coordinación de la MCIC). No figura aún en el listado de URELINTER; se agrega a la tabla del Factor 7 como «Aplica a la MCIC» (aplican: 4, complementan: 2) con sus soportes en `a. Diagnostico de convenios vigentes/Anexos/Convenio específico IGAC 5570 de 2025/` (ambas modalidades).
- Los soportes se publican con los números de cédula tapados (`redactar_documentos_identidad`). `0-Apertura Financiera IGAC.pdf` NO se publica: trae datos bancarios y su nombre de proyecto no corresponde (dice «Pueblo Rrom»).
- El convenio marco C-2023-1 (IGAC, cooperación en investigación y formación) es el marco del que se deriva.

## 12. Organización final de NuevaData/ (2-oct-2026)

Solo quedan `Original/`, `Presentacion/`, `_build/` y este documento; las carpetas sueltas (`Certificados`, `Convenios`, `Encuesta egresados`, `gruposInvestigacion`, `Info para 12 y 8`) se movieron a `Original/<Modalidad>/FACTOR N…` y el generador lee de ahí:

| Insumo | Ubicación en `Original/<Modalidad>/` |
|---|---|
| Certificados de profesores | `FACTOR 3…/a. Informes de participación de los docentes/Certificados/` |
| Encuesta de egresados | `FACTOR 4…/Caracterización e impacto de Egresados- MCIC (1-18).xlsx` |
| Listado URELINTER y soportes IGAC | `FACTOR 7…/a. Diagnostico de convenios vigentes/` (`Convenios vigentes URELINTER (3).xlsx`, `CONVENIO IGAC/`) |
| Informe de grupos, proyectos ejecutados, xlsx de tesis | `FACTOR 8…/b. Socialización y vinculación…/` (ambas modalidades) |
| Tesis (PDF) | `FACTOR 8…/b. …/Tesis/` (cada una en la carpeta de su modalidad, según el consolidado) |
| Presentaciones GIIRA y Multimedia | `FACTOR 8…/a. Presentación a nuevos estudiantes…/` **solo en Investigación** |
| Planos de laboratorios del nuevo edificio | `FACTOR 12…/` **solo en Investigación** |

Los documentos pesados compartidos (presentaciones y planos) están una sola vez, en Investigación; el generador busca primero en la carpeta de la modalidad y si no están usa la de Investigación (`compartido()` en `build_nuevadata.py`). El pptx de GIIRA (192 MB) está en `.gitignore` por el límite de GitHub; sin ese archivo el generador simplemente no produce su versión sin videos.
