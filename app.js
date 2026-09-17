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
      <strong>${MODALIDAD_LABEL[key]}:</strong> ${escapeHtml(cab.registro_calificado)}<br>
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
  renderTransparencyNote();
}

function renderTransparencyNote() {
  const cmp = GOLD_DATA.comparacionModalidades;
  const note = document.getElementById('plan-transparency-note');
  if (cmp.factores_con_diferencias.length === 0) {
    note.innerHTML = `
      ℹ️ <strong>Nota de transparencia:</strong> el texto de los 12 factores del plan de mejoramiento es
      <strong>idéntico</strong> en el archivo de Investigación y en el de Profundización — así están redactados
      en ambos <code>.xlsx</code> fuente. Lo único que cambia entre modalidades es el encabezado institucional
      (<em>${escapeHtml(cmp.campos_cabecera_distintos.join(', ') || 'ninguno')}</em>), visible arriba en el
      banner de cada modalidad. El botón "Ver fuente" de cada tarjeta te lleva siempre al archivo real de la
      modalidad activa (INV o PROF), aunque el contenido textual coincida.
    `;
  } else {
    note.innerHTML = `
      ℹ️ <strong>Nota de transparencia:</strong> ${cmp.factores_con_diferencias.length} factor(es) tienen
      contenido distinto entre Investigación y Profundización (calculado campo a campo contra ambos <code>.xlsx</code>).
    `;
  }
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
      <div class="modality-banner-item"><span>Registro Calificado</span><strong>${escapeHtml(cab.registro_calificado)}</strong></div>
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

  const todosEstados = Array.from(new Set(estado.proyectos.flatMap((p) => Object.keys(p.conteo_por_estado)))).sort();
  const estadoTable = `
    <div style="overflow-x:auto;">
    <table style="width:100%; border-collapse: collapse; font-size: 12.5px; min-width: 720px;">
      <thead>
        <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
          <th style="padding:8px 10px;">Proyecto Curricular (Cód.)</th>
          <th style="padding:8px 10px; text-align:right;">Total</th>
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
  const anios = Object.keys(egresados.por_anio).sort();
  const filaAnios = `
    <table style="width:100%; border-collapse: collapse; font-size: 13px; margin-bottom:16px;">
      <thead>
        <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
          <th style="padding:8px 12px;">Año de grado</th>
          ${anios.map((a) => `<th style="padding:8px 12px; text-align:right;">${escapeHtml(a)}</th>`).join('')}
          <th style="padding:8px 12px; text-align:right; font-weight:800;">Total (${escapeHtml(egresados.rango_presentado)})</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td style="padding:7px 12px;">Graduados</td>
          ${anios.map((a) => `<td style="padding:7px 12px; text-align:right;">${egresados.por_anio[a]}</td>`).join('')}
          <td style="padding:7px 12px; text-align:right; font-weight:700;">${egresados.total_graduados}</td>
        </tr>
      </tbody>
    </table>
  `;

  const porEnfasis = `
    <table style="width:100%; border-collapse: collapse; font-size: 12.5px;">
      <thead>
        <tr style="background: var(--bg-subtle); border-bottom: 2px solid var(--border-color); text-align:left;">
          <th style="padding:8px 10px;">Cód. Proyecto (énfasis)</th>
          <th style="padding:8px 10px; text-align:right;">Graduados</th>
          <th style="padding:8px 10px; text-align:right;">Promedio académico</th>
        </tr>
      </thead>
      <tbody>
        ${egresados.por_enfasis.map((p) => `
          <tr style="border-bottom:1px solid var(--border-color);">
            <td style="padding:7px 10px;">Cód. ${escapeHtml(p.cod_proyecto)}</td>
            <td style="padding:7px 10px; text-align:right;">${p.total_graduados}</td>
            <td style="padding:7px 10px; text-align:right;">${p.promedio_academico ?? '—'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
    <p style="font-size:11px; color: var(--text-soft); margin-top:10px;">
      Fuente: ${escapeHtml(fileLabel(egresados.fuente.archivo))} (no publicado por contener datos personales del egresado; solo llega hasta ${anios[anios.length - 1]}).
    </p>
  `;

  document.getElementById('egresados-container').innerHTML = filaAnios + porEnfasis;
}

function renderGruposInvestigacion(gruposData) {
  const grupos = gruposData.grupos.filter((g) => g.integrantes.length > 0 || g.nombre);
  document.getElementById('grupos-subtitle').innerHTML =
    `${grupos.length} grupos registrados, ${grupos.reduce((acc, g) => acc + g.integrantes.length, 0)} docentes con disponibilidad de dirección ` +
    fuenteHtml(grupos[0]?.fuente, 'Ver archivo fuente');

  document.getElementById('grupos-investigacion-container').innerHTML = `
    <div style="display:grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap:14px;">
      ${grupos.map((g) => `
        <div class="plan-meta-box" style="margin-bottom:0;">
          <div style="font-weight:700; color: var(--ud-blue); margin-bottom:6px;">${escapeHtml(g.nombre || g.sigla)} <span class="bronze-ext-badge">${escapeHtml(g.sigla)}</span></div>
          ${g.clasificacion ? `<div class="plan-meta-row"><span>Clasificación MinCiencias</span><span>${escapeHtml(g.clasificacion)}</span></div>` : ''}
          ${g.lider ? `<div class="plan-meta-row"><span>Líder</span><span>${escapeHtml(g.lider)}</span></div>` : ''}
          <div class="plan-meta-row"><span>Docentes disponibles</span><span>${g.integrantes.length}</span></div>
          ${g.integrantes.length ? `<div style="margin-top:8px; font-size:11.5px; color: var(--text-muted);">${g.integrantes.map((i) => escapeHtml(i.nombre)).join(', ')}</div>` : ''}
        </div>
      `).join('')}
    </div>
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
  `;

  document.getElementById('modal-overlay').classList.add('open');
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
    `Mostrando ${filtrados.length} de ${stats.total} archivos reales en Data/Bronze/ (recorrido automático por app/build_bronze_manifest.py) — ` +
    `${stats.con_texto_extraido} con texto extraído y buscable por contenido (app/extract_texto_bronze.py).`;

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
