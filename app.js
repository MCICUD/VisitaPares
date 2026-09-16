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

  const seguimientoTexto = (corte) => {
    const c = f.seguimiento[corte];
    const hayDatos = c.fecha_seguimiento || c.descripcion_avance_cualitativo || c.balance_cualitativo;
    return hayDatos
      ? `${escapeHtml(c.fecha_seguimiento || '')} — ${escapeHtml(c.descripcion_avance_cualitativo || c.balance_cualitativo || '')}`
      : '<em style="color: var(--text-soft);">Sin reportes registrados a la fecha en el archivo fuente.</em>';
  };

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

    <p style="margin-bottom:8px;"><strong>Seguimiento (Corte 1):</strong></p>
    <p style="margin-bottom:12px; font-size: 13px;">${seguimientoTexto('corte_1')}</p>
    <p style="margin-bottom:8px;"><strong>Seguimiento (Corte 2):</strong></p>
    <p style="margin-bottom:12px; font-size: 13px;">${seguimientoTexto('corte_2')}</p>
    <p style="margin-bottom:8px;"><strong>Balance acumulado / evaluación:</strong></p>
    <p style="font-size: 13px;">${seguimientoTexto('acumulado')}</p>
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
    if (bronzeState.busqueda && !doc.archivo.toLowerCase().includes(bronzeState.busqueda)) return false;
    return true;
  });

  document.getElementById('bronze-count-label').textContent =
    `Mostrando ${filtrados.length} de ${stats.total} archivos reales en Data/Bronze/ (recorrido automático por app/build_bronze_manifest.py).`;

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
      </td>
      <td>${escapeHtml(doc.carpeta_raiz)}</td>
      <td>${modalidadTagHtml(doc.modalidad)}</td>
      <td><span class="bronze-ext-badge">${escapeHtml(doc.extension)}</span></td>
      <td>${escapeHtml(doc.tamano_legible)}</td>
      <td><a class="btn btn-outline" style="font-size:11.5px; padding:5px 10px;" href="${fileHref(doc.archivo)}" target="_blank" rel="noopener noreferrer">Abrir ↗</a></td>
    </tr>
  `).join('');
}
