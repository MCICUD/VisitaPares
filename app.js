/**
 * MICROSITIO PLAN DE MEJORAMIENTO — VISITA DE PARES MCIC (Renovación AAC)
 * Universidad Distrital Francisco José de Caldas
 *
 * Todo el contenido se renderiza a partir de GOLD_DATA (Data/Gold/gold_data.js),
 * generado por app/run_pipeline.py directamente desde los archivos en Data/Bronze.
 * No hay ningún dato escrito a mano en este archivo.
 */

const MODALIDAD_LABEL = {
  investigacion: 'Investigación',
  profundizacion: 'Profundización',
  ambas: 'Ambas modalidades',
  general: 'General / Institucional',
};

const MODALIDAD_TAG_CLASS = {
  investigacion: 'tag-inv',
  profundizacion: 'tag-prof',
  ambas: 'tag-ambas',
  general: 'tag-general',
};

const MODALIDAD_ICON = {
  investigacion: '🔬',
  profundizacion: '💼',
  ambas: '🔗',
  general: '🏛️',
};

let state = {
  modalidad: 'investigacion',
  factorFiltro: 'all',
  tipoFiltro: 'all',
  busqueda: '',
};

let bronzeState = {
  modalidad: 'all',
  carpeta: 'all',
  extension: 'all',
  busqueda: '',
};

function initApp() {
  initNavigation();
  renderHeaderMeta();
  renderHeroStats();
  initPlanMejoramiento();
  initComunidad();
  initDocumentos();
  initBronzeCatalog();
  initSolicitudesPares();
  initModal();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}

// ==========================================================================
// UTILIDADES
// ==========================================================================
function escapeHtml(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function fileHref(relativePath) {
  return encodeURI(relativePath);
}

function fileLabel(relativePath) {
  return relativePath.split('/').pop();
}

function notaHtml(nota) {
  if (!nota) return '';
  return ` <span title="${escapeHtml(nota)}" style="cursor:help; opacity:0.75; font-size:11px;">ⓘ</span>`;
}

function fuenteHtml(fuente, label) {
  if (!fuente) return '';
  const texto = label || `Ver archivo fuente (hoja "${escapeHtml(fuente.hoja)}", fila ${escapeHtml(fuente.fila)})`;
  return `<a class="btn btn-outline" style="font-size:12px; padding:6px 12px;" href="${fileHref(fuente.archivo)}" target="_blank" rel="noopener noreferrer" title="${escapeHtml(fuente.archivo)}">🔎 ${texto}</a>`;
}

function modalidadTagHtml(modalidad) {
  const cls = MODALIDAD_TAG_CLASS[modalidad] || 'tag-general';
  const label = MODALIDAD_LABEL[modalidad] || modalidad;
  return `<span class="plan-modality-tag ${cls}">${MODALIDAD_ICON[modalidad] || ''} ${escapeHtml(label)}</span>`;
}

function factorNumeroYNombre(factorTexto) {
  const match = /^(FACTOR\s*\d+)\.?\s*(.*)$/i.exec(factorTexto || '');
  if (match) {
    return { numero: match[1].toUpperCase(), nombre: match[2].replace(/\.$/, '').trim() };
  }
  return { numero: '', nombre: factorTexto || '' };
}

function formatFecha(iso) {
  if (!iso) return '—';
  const [y, m] = iso.split('-');
  const meses = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  return `${meses[parseInt(m, 10) - 1]} ${y}`;
}

// ==========================================================================
// NAVEGACIÓN Y TABS
// ==========================================================================
function initNavigation() {
  const tabButtons = document.querySelectorAll('.nav-tab-btn');
  const tabSections = document.querySelectorAll('.tab-section');

  function switchTab(tabKey) {
    tabButtons.forEach((btn) => btn.classList.toggle('active', btn.dataset.tab === tabKey));
    tabSections.forEach((sec) => sec.classList.toggle('active', sec.id === 'section-' + tabKey));
    if (window.history && window.history.replaceState) {
      window.history.replaceState(null, null, '#' + tabKey);
    }
    window.scrollTo({ top: 0, behavior: 'instant' });
  }

  tabButtons.forEach((btn) => btn.addEventListener('click', () => switchTab(btn.dataset.tab)));

  const currentHash = window.location.hash.replace('#', '');
  if (currentHash) switchTab(currentHash);
}

// ==========================================================================
// ENCABEZADO: PASTILLAS SNIES / REGISTRO / ACREDITACIÓN
// ==========================================================================
function renderHeaderMeta() {
  const container = document.getElementById('header-snies-pills');
  const mods = GOLD_DATA.meta.modalidades;
  container.innerHTML = Object.entries(mods).map(([key, cab]) => `
    <div class="snies-pill">
      <strong>${MODALIDAD_LABEL[key]}:</strong> ${escapeHtml(cab.registro_calificado)}${notaHtml(cab.registro_calificado_nota)}<br>
      <span style="font-size: 10.5px; opacity: 0.85;">
        Acreditación AAC: ${escapeHtml(cab.acreditacion_alta_calidad)} · Vigencia ${escapeHtml(cab.acreditacion_alta_calidad_vigencia)}
      </span>
    </div>
  `).join('');
}

// ==========================================================================
// HERO: ESTADÍSTICAS DERIVADAS Y LISTA DE FUENTES
// ==========================================================================
function renderHeroStats() {
  const statsInv = GOLD_DATA.stats.investigacion;
  const statsProf = GOLD_DATA.stats.profundizacion;
  const totalFactores = statsInv.total_factores + statsProf.total_factores;
  const totalFortalezas = (statsInv.conteo_por_tipo['Fortaleza'] || 0) + (statsProf.conteo_por_tipo['Fortaleza'] || 0);
  const totalOportunidades = (statsInv.conteo_por_tipo['Oportunidad de mejora'] || 0) + (statsProf.conteo_por_tipo['Oportunidad de mejora'] || 0);

  const statsRow = document.getElementById('hero-stats-row');
  statsRow.innerHTML = `
    <div class="stat-box">
      <div class="stat-val">${totalFactores}</div>
      <div class="stat-lbl">Factores documentados (INV + PROF)</div>
    </div>
    <div class="stat-box">
      <div class="stat-val">${totalOportunidades}</div>
      <div class="stat-lbl">Oportunidades de mejora</div>
    </div>
    <div class="stat-box">
      <div class="stat-val">${totalFortalezas}</div>
      <div class="stat-lbl">Fortalezas identificadas</div>
    </div>
  `;

  const fuentesList = document.getElementById('hero-fuentes-list');
  fuentesList.innerHTML = GOLD_DATA.documentosPrincipales.map((doc) => `
    <div class="peer-item">
      <div class="peer-avatar">${doc.modalidad === 'investigacion' ? 'INV' : 'PRO'}</div>
      <div>
        <div class="peer-name">${escapeHtml(doc.titulo)}</div>
        <a class="peer-tag" href="${fileHref(doc.archivo)}" target="_blank" rel="noopener noreferrer">${escapeHtml(fileLabel(doc.archivo))}</a>
      </div>
    </div>
  `).join('');
}

// ==========================================================================
// PLAN DE MEJORAMIENTO
// ==========================================================================
function initPlanMejoramiento() {
  document.querySelectorAll('.modality-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.modalidad = btn.dataset.modality;
      document.querySelectorAll('.modality-btn').forEach((b) => b.classList.toggle('active', b === btn));
      populateFactorFilter();
      renderPlanBanner();
      renderPlanGrid();
    });
  });

  document.getElementById('plan-factor-filter').addEventListener('change', (e) => {
    state.factorFiltro = e.target.value;
    renderPlanGrid();
  });
  document.getElementById('plan-type-filter').addEventListener('change', (e) => {
    state.tipoFiltro = e.target.value;
    renderPlanGrid();
  });
  document.getElementById('plan-search').addEventListener('input', (e) => {
    state.busqueda = e.target.value.trim().toLowerCase();
    renderPlanGrid();
  });

  populateFactorFilter();
  renderPlanBanner();
  renderPlanGrid();
}

function populateFactorFilter() {
  const select = document.getElementById('plan-factor-filter');
  const factores = GOLD_DATA.factores[state.modalidad];
  select.innerHTML = '<option value="all">Todos los Factores</option>' +
    factores.map((f) => {
      const { numero } = factorNumeroYNombre(f.factor);
      return `<option value="${escapeHtml(f.factor)}">${escapeHtml(numero)}</option>`;
    }).join('');
  state.factorFiltro = 'all';
}

function renderPlanBanner() {
  const cab = GOLD_DATA.meta.modalidades[state.modalidad];
  const stats = GOLD_DATA.stats[state.modalidad];
  const banner = document.getElementById('plan-modality-banner');
  banner.classList.toggle('theme-profundizacion', state.modalidad === 'profundizacion');
  banner.classList.toggle('theme-investigacion', state.modalidad === 'investigacion');

  banner.innerHTML = `
    <div class="modality-banner-header">
      <div class="modality-banner-title">
        ${MODALIDAD_ICON[state.modalidad]} Estás viendo: Modalidad ${MODALIDAD_LABEL[state.modalidad]}
      </div>
      ${fuenteHtml(cab.fuente, 'Ver cabecera en el archivo fuente')}
    </div>
    <div class="modality-banner-grid">
      <div class="modality-banner-item"><span>Archivo Fuente</span><strong>${escapeHtml(fileLabel(cab.fuente.archivo))}</strong></div>
      <div class="modality-banner-item"><span>Registro Calificado</span><strong>${escapeHtml(cab.registro_calificado)}${notaHtml(cab.registro_calificado_nota)}</strong></div>
      <div class="modality-banner-item"><span>Vigencia Registro</span><strong>${escapeHtml(cab.registro_calificado_vigencia)}</strong></div>
      <div class="modality-banner-item"><span>Acreditación Alta Calidad</span><strong>${escapeHtml(cab.acreditacion_alta_calidad)}</strong></div>
      <div class="modality-banner-item"><span>Vigencia Acreditación</span><strong>${escapeHtml(cab.acreditacion_alta_calidad_vigencia)}</strong></div>
      <div class="modality-banner-item"><span>Fecha Proyección del Plan</span><strong>${escapeHtml(cab.fecha_proyeccion_plan)}</strong></div>
      <div class="modality-banner-item"><span>Peso-Prioridad Promedio (12 factores)</span><strong>${stats.peso_prioridad_promedio}</strong></div>
    </div>
  `;

  document.getElementById('plan-header-subtitle').innerHTML =
    `${escapeHtml(cab.macroproceso)} • ${escapeHtml(cab.proceso)} — ` +
    `<strong>Modalidad activa: ${MODALIDAD_LABEL[state.modalidad]}</strong> ` +
    `(archivo <code>${escapeHtml(fileLabel(cab.fuente.archivo))}</code>)`;
}

function renderPlanGrid() {
  const grid = document.getElementById('plan-grid');
  const factores = GOLD_DATA.factores[state.modalidad];

  const filtrados = factores.filter((f) => {
    if (state.factorFiltro !== 'all' && f.factor !== state.factorFiltro) return false;
    if (state.tipoFiltro !== 'all') {
      const esOportunidad = (f.tipo || '').toLowerCase().includes('oportunidad');
      if (state.tipoFiltro === 'Oportunidad' && !esOportunidad) return false;
      if (state.tipoFiltro === 'Fortaleza' && esOportunidad) return false;
    }
    if (state.busqueda) {
      const haystack = [f.proyecto, f.indicador_cumplimiento, f.meta, f.responsable, f.factor]
        .join(' ').toLowerCase();
      if (!haystack.includes(state.busqueda)) return false;
    }
    return true;
  });

  if (filtrados.length === 0) {
    grid.innerHTML = '<p style="color: var(--text-soft); padding: 20px;">No hay factores que coincidan con el filtro/búsqueda actual.</p>';
    return;
  }

  grid.innerHTML = filtrados.map((f, idx) => renderPlanCard(f, factores.indexOf(f))).join('');

  grid.querySelectorAll('[data-open-modal]').forEach((el) => {
    el.addEventListener('click', () => openFactorModal(state.modalidad, parseInt(el.dataset.openModal, 10)));
  });
}

function renderPlanCard(f, indexEnModalidad) {
  const { numero, nombre } = factorNumeroYNombre(f.factor);
  const esFortaleza = (f.tipo || '').toLowerCase().includes('fortaleza');
  return `
    <div class="plan-card ${esFortaleza ? 'fortaleza' : ''}">
      <div>
        <div class="plan-card-header">
          <span class="factor-tag">${escapeHtml(numero)}</span>
          <div style="display:flex; gap:6px; flex-wrap:wrap; justify-content:flex-end;">
            ${modalidadTagHtml(state.modalidad)}
            <span class="type-tag ${esFortaleza ? 'fortaleza' : 'oportunidad'}">${escapeHtml(f.tipo || 'Sin especificar')}</span>
          </div>
        </div>
        <h3>${escapeHtml(f.proyecto || nombre)}</h3>
        <p class="plan-card-desc">${escapeHtml(f.descripcion)}</p>
        <div class="plan-meta-box">
          <div class="plan-meta-row"><span>Peso-Prioridad</span><span>${escapeHtml(f.peso_prioridad)} / 10</span></div>
          <div class="plan-meta-row"><span>Periodo Ejecución</span><span>${formatFecha(f.periodo_inicio)} – ${formatFecha(f.periodo_fin)}</span></div>
          <div class="plan-meta-row"><span>Responsable</span><span>${escapeHtml(f.responsable)}</span></div>
        </div>
      </div>
      <div class="plan-card-actions">
        <button class="btn btn-primary" style="font-size:12px; padding:6px 12px;" data-open-modal="${indexEnModalidad}">Ver detalle completo</button>
        ${fuenteHtml(f.fuente, 'Ver fuente')}
      </div>
    </div>
  `;
}

function renderEvidenciaSeguimientoHtml(evidencia) {
  if (!evidencia || !evidencia.total_archivos) {
    return '';
  }
  const actividadesConArchivos = (evidencia.actividades || []).filter((a) => a.archivos && a.archivos.length > 0);
  if (actividadesConArchivos.length === 0) {
    return '';
  }
  return `
    <div style="margin-bottom:16px;">
      <p style="margin-bottom:8px;"><strong>📎 Evidencia de seguimiento cargada en Data/Bronze:</strong></p>
      ${actividadesConArchivos
        .map((a) => `
          <div style="margin-bottom:10px;">
            <div style="font-size:12.5px; font-weight:700; color: var(--text-main); margin-bottom:4px;">${escapeHtml(a.nombre)}</div>
            <div style="display:flex; flex-wrap:wrap; gap:6px;">
              ${a.archivos.map((arch) => `
                <a class="btn btn-outline" style="font-size:11.5px; padding:4px 10px;" href="${fileHref(arch.archivo)}" target="_blank" rel="noopener noreferrer" title="${escapeHtml(arch.archivo)}">📄 ${escapeHtml(arch.nombre)} (${escapeHtml(arch.tamano_legible)})</a>
              `).join('')}
            </div>
          </div>
        `).join('')}
    </div>
  `;
}

// ==========================================================================
// COMUNIDAD ESTUDIANTIL (énfasis y estado académico)
// ==========================================================================
function initComunidad() {
  const enfasis = GOLD_DATA.comunidadEstudiantil.enfasis;
  const estado = GOLD_DATA.comunidadEstudiantil.estadoAcademico;

  document.getElementById('comunidad-subtitle').innerHTML =
    `${escapeHtml(enfasis.titulo_archivo)} ${fuenteHtml(enfasis.fuente, 'Ver fuente')}`;

  const descarga = GOLD_DATA.comunidadEstudiantil.descarga;
  document.getElementById('comunidad-descarga-container').innerHTML = descarga ? `
    <a href="${fileHref(descarga.archivo)}" download class="btn btn-gold" style="text-decoration:none; font-size:13px;">
      <span>📥</span> ${escapeHtml(descarga.titulo)}
    </a>
  ` : '';

  const enfasisNombres = enfasis.enfasis;
  const enfasisTable = `
    <table style="width:100%; border-collapse: collapse; font-size: 13px;">
      <thead>
        <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
          <th style="padding:8px 12px;">Categoría</th>
          ${enfasisNombres.map((n) => `<th style="padding:8px 12px; text-align:right;">${escapeHtml(n)}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
        ${enfasis.filas.map((fila) => `
          <tr style="border-bottom:1px solid var(--border-color); ${fila.etiqueta === 'Total' ? 'font-weight:700;' : ''}">
            <td style="padding:7px 12px;">${escapeHtml(fila.etiqueta)}</td>
            ${enfasisNombres.map((n) => `<td style="padding:7px 12px; text-align:right;">${escapeHtml(fila.valores[n])}</td>`).join('')}
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
  document.getElementById('enfasis-table-container').innerHTML = enfasisTable;

  const ESTADOS_OCULTOS = new Set([
    'INACTIVO',
    'Abandono',
    'Pérdida de calidad de estudiante - Suspendido',
    'Pérdida de calidad de estudiante - Expulsado',
  ]);
  const todosEstados = Array.from(new Set(estado.proyectos.flatMap((p) => Object.keys(p.conteo_por_estado))))
    .filter((e) => !ESTADOS_OCULTOS.has(e))
    .sort();
  const estadoTable = `
    <div style="overflow-x:auto;">
    <table style="width:100%; border-collapse: collapse; font-size: 12.5px; min-width: 720px;">
      <thead>
        <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
          <th style="padding:8px 10px;">Énfasis (Cód.)</th>
          <th style="padding:8px 10px; text-align:right;">Total histórico<br><span style="font-weight:400; font-size:10.5px;">(todos los estados)</span></th>
          ${todosEstados.map((e) => `<th style="padding:8px 10px; text-align:right;">${escapeHtml(e)}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
        ${estado.proyectos.map((p) => `
          <tr style="border-bottom:1px solid var(--border-color);">
            <td style="padding:7px 10px;">${escapeHtml(p.proyecto_curricular)} <span class="bronze-ext-badge">Cód. ${escapeHtml(p.cod_proyecto)}</span></td>
            <td style="padding:7px 10px; text-align:right; font-weight:700;">${p.total_estudiantes}</td>
            ${todosEstados.map((e) => `<td style="padding:7px 10px; text-align:right;">${p.conteo_por_estado[e] ?? '—'}</td>`).join('')}
          </tr>
        `).join('')}
      </tbody>
    </table>
    </div>
    <p style="font-size:11px; color: var(--text-soft); margin-top:10px;">
      Fuente: ${estado.proyectos.map((p) => escapeHtml(fileLabel(p.fuente.archivo))).join(', ')} (roster Cóndor, Data/Bronze/Estados — no publicado por contener datos personales).
    </p>
  `;
  document.getElementById('estado-academico-table-container').innerHTML = estadoTable;

  renderEgresados(GOLD_DATA.comunidadEstudiantil.egresados);
  renderGruposInvestigacion(GOLD_DATA.gruposInvestigacion);
}

function renderEgresados(egresados) {
  const porProyecto2022a2026 = egresados.por_proyecto_por_anio_estimado.map((p) => ({
    ...p,
    total: Object.values(p.por_anio_estimado).reduce((acc, n) => acc + n, 0),
  }));
  const totalGeneral2022a2026 = porProyecto2022a2026.reduce((acc, p) => acc + p.total, 0);

  const totalHistoricoTable = `
    <table style="width:100%; border-collapse: collapse; font-size: 12.5px; margin-bottom:10px;">
      <thead>
        <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
          <th style="padding:8px 10px;">Énfasis (Cód.)</th>
          <th style="padding:8px 10px; text-align:right;">Graduados</th>
        </tr>
      </thead>
      <tbody>
        ${porProyecto2022a2026.map((p) => `
          <tr style="border-bottom:1px solid var(--border-color);">
            <td style="padding:7px 10px;">${escapeHtml(p.proyecto_curricular)} <span class="bronze-ext-badge">Cód. ${escapeHtml(p.cod_proyecto)}</span></td>
            <td style="padding:7px 10px; text-align:right; font-weight:700;">${p.total}</td>
          </tr>
        `).join('')}
        <tr style="font-weight:800; border-top:2px solid var(--border-color);">
          <td style="padding:7px 10px;">Total MCIC (${escapeHtml(egresados.rango_presentado)})</td>
          <td style="padding:7px 10px; text-align:right;">${totalGeneral2022a2026}</td>
        </tr>
      </tbody>
    </table>
  `;

  const anios = Object.keys(egresados.por_anio_estimado).sort();
  const estimadoTable = `
    <table style="width:100%; border-collapse: collapse; font-size: 13px; margin-bottom:18px;">
      <thead>
        <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
          <th style="padding:8px 12px;">Graduados por año</th>
          ${anios.map((a) => `<th style="padding:8px 12px; text-align:right;">${escapeHtml(a)}</th>`).join('')}
          <th style="padding:8px 12px; text-align:right; font-weight:800;">Total (${escapeHtml(egresados.rango_presentado)})</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td style="padding:7px 12px;">Total</td>
          ${anios.map((a) => `<td style="padding:7px 12px; text-align:right;">${egresados.por_anio_estimado[a]}</td>`).join('')}
          <td style="padding:7px 12px; text-align:right; font-weight:700;">${anios.reduce((acc, a) => acc + egresados.por_anio_estimado[a], 0)}</td>
        </tr>
      </tbody>
    </table>
  `;

  document.getElementById('egresados-container').innerHTML = `
    <h4 style="font-size:13.5px; margin:0 0 8px;">Graduados por énfasis (${escapeHtml(egresados.rango_presentado)})</h4>
    ${totalHistoricoTable}
    <h4 style="font-size:13.5px; margin:0 0 8px;">Por año (${escapeHtml(egresados.rango_presentado)})</h4>
    ${estimadoTable}
  `;
}

let gruposInvestigacionState = { clasificacion: 'all' };

const ETAPA_TESIS_BADGE = {
  'Sustentado': '✅',
  'Jurado': '⏳',
  'Jurados': '⏳',
  'Programado': '⏳',
  'Avalado': '📋',
  'Anteproyecto': '📝',
  'Sin anteproyecto': '⚪',
};

function etapaBadge(etapa) {
  const clave = Object.keys(ETAPA_TESIS_BADGE).find((k) => etapa.startsWith(k));
  return ETAPA_TESIS_BADGE[clave] || '•';
}

function renderGruposInvestigacion(gruposData) {
  const todosLosGrupos = gruposData.grupos.filter((g) => g.integrantes.length > 0 || g.nombre);

  const clasificaciones = Array.from(new Set(todosLosGrupos.map((g) => g.clasificacion || 'Sin clasificar'))).sort();
  const filterEl = document.getElementById('grupos-clasificacion-filter');
  filterEl.innerHTML = `<option value="all">Todas las clasificaciones MinCiencias</option>` +
    clasificaciones.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  filterEl.value = gruposInvestigacionState.clasificacion;
  filterEl.onchange = () => {
    gruposInvestigacionState.clasificacion = filterEl.value;
    pintarGruposInvestigacion(gruposData, todosLosGrupos);
  };

  pintarGruposInvestigacion(gruposData, todosLosGrupos);
}

function pintarGruposInvestigacion(gruposData, todosLosGrupos) {
  const filtro = gruposInvestigacionState.clasificacion;
  const grupos = filtro === 'all'
    ? todosLosGrupos
    : todosLosGrupos.filter((g) => (g.clasificacion || 'Sin clasificar') === filtro);

  const totalProyectos = todosLosGrupos.reduce((acc, g) => acc + (g.proyectos_grado?.length || 0), 0);
  document.getElementById('grupos-subtitle').innerHTML =
    `${todosLosGrupos.length} grupos registrados, ${todosLosGrupos.reduce((acc, g) => acc + g.integrantes.length, 0)} docentes con disponibilidad de dirección, ` +
    `${totalProyectos} proyectos de grado (595/695) asociados por carta de radicación/viabilidad ` +
    fuenteHtml(todosLosGrupos[0]?.fuente, 'Ver archivo fuente');

  document.getElementById('grupos-investigacion-container').innerHTML = `
    <div style="display:grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap:14px;">
      ${grupos.map((g) => `
        <div class="plan-meta-box" style="margin-bottom:0;">
          <div style="font-weight:700; color: var(--ud-blue); margin-bottom:6px;">${escapeHtml(g.nombre || g.sigla)} <span class="bronze-ext-badge">${escapeHtml(g.sigla)}</span></div>
          ${g.clasificacion ? `<div class="plan-meta-row"><span>Clasificación MinCiencias</span><span>${escapeHtml(g.clasificacion)}</span></div>` : ''}
          ${g.lider ? `<div class="plan-meta-row"><span>Líder</span><span>${escapeHtml(g.lider)}</span></div>` : ''}
          <div class="plan-meta-row"><span>Docentes disponibles</span><span>${g.integrantes.length}</span></div>
          ${g.integrantes.length ? `<div style="margin-top:8px; font-size:11.5px; color: var(--text-muted);">${g.integrantes.map((i) => escapeHtml(i.nombre)).join(', ')}</div>` : ''}
          ${(g.proyectos_grado && g.proyectos_grado.length) ? `
            <div style="margin-top:10px; padding-top:8px; border-top:1px dashed var(--border-color);">
              <div style="font-size:11.5px; font-weight:700; margin-bottom:6px;">🎓 Proyectos de grado asociados (${g.proyectos_grado.length})</div>
              <ul style="margin:0; padding-left:16px; font-size:11px; color: var(--text-muted);">
                ${g.proyectos_grado.map((p) => `
                  <li style="margin-bottom:5px;">
                    ${etapaBadge(p.etapa_actual)} ${escapeHtml(p.titulo_proyecto || 'Título no extraído del documento')}
                    ${p.director ? `<br><span style="opacity:0.85;">Director: ${escapeHtml(p.director)}${p.codirector ? ` · Codirector: ${escapeHtml(p.codirector)}` : ''}</span>` : ''}
                    <br><span style="opacity:0.7;">${escapeHtml(p.etapa_actual)}</span>
                  </li>
                `).join('')}
              </ul>
            </div>
          ` : ''}
        </div>
      `).join('')}
    </div>
    ${grupos.length === 0 ? '<p style="font-size:12.5px; color: var(--text-muted);">Ningún grupo tiene esa clasificación.</p>' : ''}
  `;
}

// ==========================================================================
// MODAL DE DETALLE DE FACTOR
// ==========================================================================
function initModal() {
  const overlay = document.getElementById('modal-overlay');
  document.getElementById('modal-close-btn').addEventListener('click', () => overlay.classList.remove('open'));
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) overlay.classList.remove('open');
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') overlay.classList.remove('open');
  });
}

function openFactorModal(modalidad, index) {
  const f = GOLD_DATA.factores[modalidad][index];
  const { numero, nombre } = factorNumeroYNombre(f.factor);
  document.getElementById('modal-title').textContent = `${numero} — ${nombre}`;

  document.getElementById('modal-body').innerHTML = `
    <div style="display:flex; gap:8px; margin-bottom:14px; flex-wrap:wrap;">
      ${modalidadTagHtml(modalidad)}
      <span class="type-tag ${(f.tipo || '').toLowerCase().includes('fortaleza') ? 'fortaleza' : 'oportunidad'}">${escapeHtml(f.tipo)}</span>
      ${fuenteHtml(f.fuente)}
    </div>
    <p style="margin-bottom:12px;"><strong>Proyecto / Acción global:</strong><br>${escapeHtml(f.proyecto)}</p>
    <p style="margin-bottom:12px;"><strong>Origen:</strong> ${escapeHtml(f.origen)}</p>
    <p style="margin-bottom:12px;"><strong>Descripción:</strong><br>${escapeHtml(f.descripcion)}</p>
    <p style="margin-bottom:12px;"><strong>Objetivo:</strong><br>${escapeHtml(f.objetivo)}</p>
    <p style="margin-bottom:12px;"><strong>Articulación con el Plan Institucional:</strong><br>${escapeHtml(f.articulacion_plan_institucional)}</p>

    <div class="plan-meta-box" style="margin-bottom:14px;">
      <div class="plan-meta-row"><span>Periodo de Ejecución</span><span>${formatFecha(f.periodo_inicio)} – ${formatFecha(f.periodo_fin)}</span></div>
      <div class="plan-meta-row"><span>Peso-Prioridad</span><span>${escapeHtml(f.peso_prioridad)} / 10</span></div>
      <div class="plan-meta-row"><span>Indicador de Cumplimiento</span><span>${escapeHtml(f.indicador_cumplimiento)}</span></div>
      <div class="plan-meta-row"><span>Tipo de Indicador</span><span>${escapeHtml(f.tipo_indicador)}</span></div>
      <div class="plan-meta-row"><span>Línea Base</span><span>${escapeHtml(f.linea_base)}</span></div>
      <div class="plan-meta-row"><span>Meta</span><span>${escapeHtml(f.meta)}</span></div>
      <div class="plan-meta-row"><span>Periodicidad</span><span>${escapeHtml(f.periodicidad)}</span></div>
      <div class="plan-meta-row"><span>Tipo de Actividad</span><span>${escapeHtml(f.tipo_actividad)}</span></div>
      <div class="plan-meta-row"><span>Apoyo Requerido</span><span>${['programa_academico', 'facultad', 'institucion'].filter((k) => f.apoyo_requerido[k]).map((k) => ({ programa_academico: 'Programa', facultad: 'Facultad', institucion: 'Institución' }[k])).join(', ') || '—'}</span></div>
      <div class="plan-meta-row"><span>Responsable</span><span>${escapeHtml(f.responsable)}</span></div>
      <div class="plan-meta-row"><span>Recursos</span><span>${escapeHtml(f.recursos)}</span></div>
    </div>

    <p style="margin-bottom:8px;"><strong>Actividades requeridas para lograr la meta:</strong></p>
    <p style="margin-bottom:14px; white-space: pre-line; font-size: 13px; color: var(--text-muted);">${escapeHtml(f.actividades)}</p>

    ${renderEvidenciaSeguimientoHtml(f.evidencia_seguimiento)}
    ${renderCuadroMaestroHtml(f.datos_cuadro_maestro_cna)}
    ${renderComunidadFactor4Html(f.datos_comunidad_factor4)}
    ${renderConveniosHtml(f.datos_convenios)}
  `;

  document.getElementById('modal-overlay').classList.add('open');
}

function renderComunidadFactor4Html(datos) {
  if (!datos) return '';
  const egresados = datos.egresados;
  const estados = datos.estados;
  
  const anios = Object.keys(egresados.por_anio_estimado).sort();
  
  let matriculadosTotal = 0;
  estados.proyectos.forEach(p => {
    matriculadosTotal += (p.conteo_por_estado['Matriculado'] || 0);
  });
  
  const descarga = GOLD_DATA.comunidadEstudiantil.descarga;
  const fuenteHtml = descarga ? `
    <div style="margin-top:10px;">
      <a href="${fileHref(descarga.archivo)}" download class="btn btn-outline" style="text-decoration:none; font-size:11px; padding:4px 8px;">
        📥 Descargar fuente: ${escapeHtml(descarga.titulo)}
      </a>
    </div>
  ` : '';

  return `
    <div style="margin-bottom:16px;">
      <p style="margin-bottom:8px;"><strong>📊 Datos de Estudiantes y Graduados (2022-2026):</strong></p>
      
      <div style="margin-bottom:14px;">
        <div style="font-size:12.5px; font-weight:700; margin-bottom:6px;">Matriculados (Estado Actual)</div>
        <div class="plan-meta-box">
          <div class="plan-meta-row"><span>Total Estudiantes Matriculados Activos</span><span>${matriculadosTotal}</span></div>
        </div>
        <p style="font-size:10.5px; color: var(--text-soft); margin-top:6px;">Fuente: Consolidado Comunidad Estudiantil</p>
      </div>

      <div style="margin-bottom:14px;">
        <div style="font-size:12.5px; font-weight:700; margin-bottom:6px;">Graduados por año (Estimación 2022-2026)</div>
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse: collapse; font-size: 11.5px;">
            <thead>
              <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
                <th style="padding:6px 8px;">Métrica</th>
                ${anios.map(a => `<th style="padding:6px 8px; text-align:right;">${escapeHtml(a)}</th>`).join('')}
                <th style="padding:6px 8px; text-align:right;">Total</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td style="padding:5px 8px;">Graduados</td>
                ${anios.map(a => `<td style="padding:5px 8px; text-align:right;">${egresados.por_anio_estimado[a]}</td>`).join('')}
                <td style="padding:5px 8px; text-align:right; font-weight:700;">${anios.reduce((acc, a) => acc + egresados.por_anio_estimado[a], 0)}</td>
              </tr>
            </tbody>
          </table>
        </div>
        <p style="font-size:10.5px; color: var(--text-soft); margin-top:6px;">Fuente: Consolidado Comunidad Estudiantil</p>
      </div>
      ${fuenteHtml}
    </div>
  `;
}

function renderCuadroMaestroHtml(datos) {
  if (!datos) return '';
  let bloques = '';

  if (datos.graduacion) {
    const periodos = datos.graduacion.por_periodo;
    bloques += `
      <div style="margin-bottom:14px;">
        <div style="font-size:12.5px; font-weight:700; margin-bottom:6px;">Graduación por periodo (Cuadro CNA No. 04)</div>
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse: collapse; font-size: 11.5px; min-width:640px;">
            <thead>
              <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
                <th style="padding:6px 8px;">Periodo</th>
                ${periodos.map((p) => `<th style="padding:6px 8px; text-align:right;">${escapeHtml(p.periodo)}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              <tr><td style="padding:5px 8px;">Matriculados</td>${periodos.map((p) => `<td style="padding:5px 8px; text-align:right;">${p.matriculados ?? '—'}</td>`).join('')}</tr>
              <tr><td style="padding:5px 8px;">Graduados</td>${periodos.map((p) => `<td style="padding:5px 8px; text-align:right; font-weight:700;">${p.graduados ?? '—'}</td>`).join('')}</tr>
            </tbody>
          </table>
        </div>
        <p style="font-size:10.5px; color: var(--text-soft); margin-top:6px;">${fuenteHtml(datos.graduacion.fuente, `Descargar Cuadro Maestro (hoja "${datos.graduacion.fuente.hoja}")`)}</p>
      </div>
    `;
  }

  if (datos.enlace_modulo_egresados) {
    const link = datos.enlace_modulo_egresados;
    bloques += `
      <p style="margin-bottom:14px;">
        <a class="btn btn-outline" style="font-size:12px; padding:6px 12px;" href="${escapeHtml(link.url)}" target="_blank" rel="noopener noreferrer">🔗 ${escapeHtml(link.titulo)}</a>
      </p>
    `;
  }

  if (datos.grupos_produccion) {
    const gp = datos.grupos_produccion;
    bloques += `
      <div style="margin-bottom:14px;">
        <div style="font-size:12.5px; font-weight:700; margin-bottom:6px;">Producción investigativa por grupo (Cuadro CNA No. 08)</div>
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse: collapse; font-size: 11px; min-width:780px;">
            <thead>
              <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
                <th style="padding:6px 8px;">Grupo</th>
                <th style="padding:6px 8px; text-align:right;">Clasif.</th>
                <th style="padding:6px 8px; text-align:right;">Proy. internos</th>
                <th style="padding:6px 8px; text-align:right;">Proy. externos</th>
                <th style="padding:6px 8px; text-align:right;">Art. nacional</th>
                <th style="padding:6px 8px; text-align:right;">Art. internacional</th>
                <th style="padding:6px 8px; text-align:right;">Productos totales</th>
              </tr>
            </thead>
            <tbody>
              ${gp.grupos.map((gr) => `
                <tr style="border-bottom:1px solid var(--border-color);">
                  <td style="padding:5px 8px;">${escapeHtml(gr.nombre_grupo)}</td>
                  <td style="padding:5px 8px; text-align:right;">${escapeHtml(gr.clasificacion_minciencias ?? '—')}</td>
                  <td style="padding:5px 8px; text-align:right;">${gr.proyectos_recursos_internos ?? '—'}</td>
                  <td style="padding:5px 8px; text-align:right;">${gr.proyectos_recursos_externos ?? '—'}</td>
                  <td style="padding:5px 8px; text-align:right;">${gr.articulos_indexados_nacional ?? '—'}</td>
                  <td style="padding:5px 8px; text-align:right;">${gr.articulos_indexados_internacional ?? '—'}</td>
                  <td style="padding:5px 8px; text-align:right; font-weight:700;">${gr.productos_totales ?? '—'}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
        <p style="font-size:10.5px; color: var(--text-soft); margin-top:6px;">${fuenteHtml(gp.fuente, `Descargar Cuadro Maestro (hoja "${gp.fuente.hoja}")`)}</p>
      </div>
    `;
  }

  if (datos.resumen_grupos_investigacion) {
    const r = datos.resumen_grupos_investigacion;
    const fuentesHtmlText = r.fuentes_consultadas && r.fuentes_consultadas.length
      ? `<div style="font-size:10.5px; color: var(--text-soft); margin-top:8px; padding-top:6px; border-top:1px dashed var(--border-color);">
           📂 <strong>Bases de datos oficiales de trabajo de grado:</strong><br>
           ${r.fuentes_consultadas.map((f) => `• ${escapeHtml(f)}`).join('<br>')}
         </div>`
      : '';

    const consolidadoPrograma = r.total_programa_consolidado && r.total_programa_consolidado !== r.total_proyectos_grado_detectados
      ? `<div class="plan-meta-row" style="margin-top:8px; padding-top:6px; border-top:1px dashed var(--border-color); font-weight:600;">
           <span>Consolidado total programa MCIC (595 + 695)</span><span>${r.total_programa_consolidado}</span>
         </div>`
      : '';

    bloques += `
      <div class="plan-meta-box" style="margin-bottom:14px;">
        <div class="plan-meta-row"><span>Grupos de investigación registrados</span><span>${r.total_grupos}</span></div>
        <div class="plan-meta-row"><span>Docentes disponibles para dirección</span><span>${r.total_docentes_disponibles}</span></div>
        <div class="plan-meta-row" style="font-weight:600; color:var(--ud-blue);">
          <span>Proyectos de grado en esta modalidad</span><span>${r.total_proyectos_grado_detectados}</span>
        </div>
        ${Object.entries(r.proyectos_grado_por_etapa).map(([etapa, n]) => `
          <div class="plan-meta-row" style="padding-left:8px;">
            <span>${etapaBadge(etapa)} ${escapeHtml(etapa)}</span>
            <span style="font-weight:600;">${n}</span>
          </div>
        `).join('')}
        ${consolidadoPrograma}
        ${fuentesHtmlText}
      </div>
    `;
  }

  if (!bloques) return '';
  return `
    <div style="margin-bottom:16px;">
      <p style="margin-bottom:8px;"><strong>📊 Datos oficiales del Cuadro Maestro (CNA):</strong></p>
      ${bloques}
    </div>
  `;
}

function renderConveniosHtml(datos) {
  if (!datos) return '';
  const relacionados = datos.convenios_relacionados_facultad_ingenieria || [];
  const aplicables = relacionados.filter((c) => c.aplica_a_posgrado_mcic === true);

  const totalesHtml = `
    <div class="plan-meta-box" style="margin-bottom:12px;">
      <div class="plan-meta-row"><span>Convenios institucionales vigentes (toda la UD)</span><span>${datos.convenios_totales}</span></div>
      ${Object.entries(datos.por_nivel || {}).map(([nivel, n]) => `<div class="plan-meta-row"><span>· ${escapeHtml(nivel)}</span><span>${n}</span></div>`).join('')}
      <div class="plan-meta-row"><span>Relacionados con la Facultad de Ingeniería / áreas afines de la MCIC</span><span>${relacionados.length}</span></div>
      <div class="plan-meta-row"><span>Aplican realmente a un posgrado como la MCIC (los que se listan abajo)</span><span>${aplicables.length}</span></div>
    </div>
  `;

  const tablaHtml = aplicables.length ? `
    <div style="overflow-x:auto; margin-bottom:10px;">
      <table style="width:100%; border-collapse: collapse; font-size: 11.5px; min-width:640px;">
        <thead>
          <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
            <th style="padding:6px 8px;">Institución</th>
            <th style="padding:6px 8px;">País / Categoría</th>
            <th style="padding:6px 8px;">Tipo</th>
            <th style="padding:6px 8px;">Denominación</th>
            <th style="padding:6px 8px;">Vigencia hasta</th>
          </tr>
        </thead>
        <tbody>
          ${aplicables.map((c) => `
            <tr style="border-bottom:1px solid var(--border-color);" title="${escapeHtml(c.objeto || '')}">
              <td style="padding:5px 8px;">${escapeHtml(c.institucion)}</td>
              <td style="padding:5px 8px;">${escapeHtml(c.pais_categoria)}</td>
              <td style="padding:5px 8px;">${escapeHtml(c.tipo)}</td>
              <td style="padding:5px 8px;">${escapeHtml(c.denominacion)}</td>
              <td style="padding:5px 8px;">${escapeHtml(c.fecha_fin)} (${escapeHtml(c.estado)})</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  ` : '';

  const fuente = datos.fuente || {};
  return `
    <div style="margin-bottom:16px;">
      <p style="margin-bottom:8px;"><strong>🤝 Convenios institucionales (URELINTER):</strong></p>
      ${totalesHtml}
      ${tablaHtml}
      <p style="font-size:11.5px; color: var(--text-soft); margin-bottom:6px;">${escapeHtml(datos.limitacion)}</p>
      <a class="btn btn-outline" style="font-size:11.5px; padding:4px 10px;" href="${fileHref(fuente.archivo_descargado)}" target="_blank" rel="noopener noreferrer">📥 Ver listado completo (${escapeHtml(datos.convenios_totales)} convenios, descargado de URELINTER el ${escapeHtml(fuente.fecha_descarga)})</a>
    </div>
  `;
}

// ==========================================================================
// DOCUMENTOS FUENTE
// ==========================================================================
function initDocumentos() {
  const sharepointUrl = GOLD_DATA.meta.sharepointUrl;
  document.getElementById('sharepoint-link-btn').href = sharepointUrl;
  document.getElementById('sharepoint-link-banner').href = sharepointUrl;

  const grid = document.getElementById('documentos-grid');
  grid.innerHTML = GOLD_DATA.documentosPrincipales.map((doc) => {
    const ext = doc.extension;
    return `
    <div class="doc-card">
      <div>
        <div class="doc-card-top">
          <div class="doc-icon ${ext}">${ext.toUpperCase()}</div>
          <div>
            <div class="doc-name">${escapeHtml(doc.titulo)}</div>
            ${modalidadTagHtml(doc.modalidad)}
          </div>
        </div>
        <span class="doc-path-code">${escapeHtml(doc.archivo)}</span>
      </div>
      <div class="doc-footer">
        <span>${escapeHtml(doc.tamano_legible)}</span>
        <a class="btn btn-outline" style="font-size:12px; padding:6px 12px;" href="${fileHref(doc.archivo)}" target="_blank" rel="noopener noreferrer">Abrir archivo ↗</a>
      </div>
    </div>
  `;
  }).join('');

  const rcAacGrid = document.getElementById('evidencia-rc-aac-grid');
  const combinados = [];
  for (const modalidad of ['investigacion', 'profundizacion']) {
    for (const doc of GOLD_DATA.evidenciaDocumentosGenerales[modalidad]) {
      combinados.push({ ...doc, modalidad, categoria: 'Plan de Mejoramiento' });
    }
    for (const doc of GOLD_DATA.evidenciaProcesosRcAac[modalidad]) {
      combinados.push({ ...doc, modalidad, categoria: 'Procesos RC / AAC' });
    }
  }
  rcAacGrid.innerHTML = combinados.map((doc) => {
    const ext = (doc.nombre.split('.').pop() || '').toLowerCase();
    return `
    <div class="doc-card">
      <div>
        <div class="doc-card-top">
          <div class="doc-icon ${ext === 'docx' || ext === 'doc' ? 'docx' : (ext === 'xlsx' || ext === 'xls' ? 'xlsx' : '')}">${ext.toUpperCase()}</div>
          <div>
            <div class="doc-name">${escapeHtml(doc.nombre)}</div>
            <div style="display:flex; gap:6px; margin-top:4px;">${modalidadTagHtml(doc.modalidad)}<span class="bronze-ext-badge">${escapeHtml(doc.categoria)}</span></div>
          </div>
        </div>
        <span class="doc-path-code">${escapeHtml(doc.archivo)}</span>
      </div>
      <div class="doc-footer">
        <span>${escapeHtml(doc.tamano_legible)}</span>
        <a class="btn btn-outline" style="font-size:12px; padding:6px 12px;" href="${fileHref(doc.archivo)}" target="_blank" rel="noopener noreferrer">Abrir archivo ↗</a>
      </div>
    </div>
  `;
  }).join('');
}

// ==========================================================================
// CATÁLOGO COMPLETO DE Data/Bronze (data RAW)
// ==========================================================================
function initBronzeCatalog() {
  const manifest = GOLD_DATA.documentosBronze;
  const stats = GOLD_DATA.documentosBronzeStats;

  const carpetaSelect = document.getElementById('bronze-carpeta-filter');
  const carpetas = Object.keys(stats.por_carpeta_raiz).sort();
  carpetaSelect.innerHTML = '<option value="all">Todas las carpetas</option>' +
    carpetas.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)} (${stats.por_carpeta_raiz[c]})</option>`).join('');

  const extSelect = document.getElementById('bronze-extension-filter');
  const extensiones = Object.keys(stats.por_extension).sort();
  extSelect.innerHTML = '<option value="all">Todos los tipos de archivo</option>' +
    extensiones.map((e) => `<option value="${escapeHtml(e)}">.${escapeHtml(e)} (${stats.por_extension[e]})</option>`).join('');

  document.getElementById('bronze-modalidad-filter').addEventListener('change', (e) => {
    bronzeState.modalidad = e.target.value;
    renderBronzeTable();
  });
  carpetaSelect.addEventListener('change', (e) => {
    bronzeState.carpeta = e.target.value;
    renderBronzeTable();
  });
  extSelect.addEventListener('change', (e) => {
    bronzeState.extension = e.target.value;
    renderBronzeTable();
  });
  document.getElementById('bronze-search').addEventListener('input', (e) => {
    bronzeState.busqueda = e.target.value.trim().toLowerCase();
    renderBronzeTable();
  });

  renderBronzeTable();
}

function renderBronzeTable() {
  const manifest = GOLD_DATA.documentosBronze;
  const stats = GOLD_DATA.documentosBronzeStats;

  const filtrados = manifest.filter((doc) => {
    if (bronzeState.modalidad !== 'all' && doc.modalidad !== bronzeState.modalidad) return false;
    if (bronzeState.carpeta !== 'all' && doc.carpeta_raiz !== bronzeState.carpeta) return false;
    if (bronzeState.extension !== 'all' && doc.extension !== bronzeState.extension) return false;
    if (bronzeState.busqueda) {
      const haystack = (doc.archivo + ' ' + (doc.extracto || '')).toLowerCase();
      if (!haystack.includes(bronzeState.busqueda)) return false;
    }
    return true;
  });

  document.getElementById('bronze-count-label').textContent =
    `Mostrando ${filtrados.length} de ${stats.total} archivos — ${stats.con_texto_extraido} con texto buscable por contenido.`;

  const tbody = document.getElementById('bronze-table-body');
  if (filtrados.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" style="color: var(--text-soft); padding: 16px;">Ningún archivo coincide con el filtro/búsqueda actual.</td></tr>';
    return;
  }

  tbody.innerHTML = filtrados.map((doc) => `
    <tr>
      <td>
        <span class="bronze-file-name">${escapeHtml(doc.nombre)}</span>
        <span class="bronze-path">${escapeHtml(doc.carpeta_contenedora)}</span>
        ${doc.extracto ? `<span class="bronze-path" style="color: var(--text-muted);" title="Extracto de texto leído del archivo">📝 ${escapeHtml(doc.extracto)}</span>` : ''}
      </td>
      <td>${escapeHtml(doc.carpeta_raiz)}</td>
      <td>${modalidadTagHtml(doc.modalidad)}</td>
      <td><span class="bronze-ext-badge">${escapeHtml(doc.extension)}</span></td>
      <td>${escapeHtml(doc.tamano_legible)}</td>
      <td><a class="btn btn-outline" style="font-size:11.5px; padding:5px 10px;" href="${fileHref(doc.archivo)}" target="_blank" rel="noopener noreferrer">Abrir ↗</a></td>
    </tr>
  `).join('');
}

// ==========================================================================
// SOLICITUDES DE LOS PARES (material entregado por día de visita)
// ==========================================================================
function solicitudDocCardHtml(doc) {
  const ext = doc.extension;
  return `
    <div class="doc-card">
      <div>
        <div class="doc-card-top">
          <div class="doc-icon ${escapeHtml(ext)}">${escapeHtml(ext.toUpperCase())}</div>
          <div class="doc-name">${escapeHtml(doc.nombre)}</div>
        </div>
        <span class="doc-path-code">${escapeHtml(doc.archivo)}</span>
      </div>
      <div class="doc-footer">
        <span>${escapeHtml(doc.tamano_legible)}</span>
        <a class="btn btn-outline" style="font-size:12px; padding:6px 12px;" href="${fileHref(doc.archivo)}" target="_blank" rel="noopener noreferrer">Abrir archivo ↗</a>
      </div>
    </div>
  `;
}

function initSolicitudesPares() {
  const container = document.getElementById('solicitudes-dias-container');
  const dias = (GOLD_DATA.solicitudesPares || {}).dias || [];
  if (dias.length === 0) {
    container.innerHTML = '<p style="color: var(--text-soft);">Aún no hay material registrado en SolicitudesPares/.</p>';
    return;
  }

  container.innerHTML = dias.map((dia) => {
    const totalSolicitudes = dia.solicitudes.reduce((acc, a) => acc + a.documentos.length, 0);
    const presentacionesHtml = dia.presentaciones.length
      ? `<div class="doc-grid">${dia.presentaciones.map(solicitudDocCardHtml).join('')}</div>`
      : '<p style="color: var(--text-soft); font-size: 13px;">Sin presentaciones registradas.</p>';
    const solicitudesHtml = dia.solicitudes.length
      ? dia.solicitudes.map((area) => `
          <h4 class="solicitud-area-title">${escapeHtml(area.area)} <span class="bronze-ext-badge">${area.documentos.length}</span></h4>
          <div class="doc-grid" style="margin-bottom: 20px;">${area.documentos.map(solicitudDocCardHtml).join('')}</div>
        `).join('')
      : '<p style="color: var(--text-soft); font-size: 13px;">Sin solicitudes registradas.</p>';

    return `
      <div class="card">
        <h3 class="card-title">📅 Día ${escapeHtml(dia.numero)}</h3>

        <h4 class="solicitud-subtitle">1. Presentaciones <span class="bronze-ext-badge">${dia.presentaciones.length}</span></h4>
        ${presentacionesHtml}

        <h4 class="solicitud-subtitle" style="margin-top: 28px;">2. Solicitudes <span class="bronze-ext-badge">${totalSolicitudes}</span></h4>
        ${solicitudesHtml}
      </div>
    `;
  }).join('');
}
