/**
 * APLICACIÓN PORTAL AUTOEVALUACIÓN Y VISITA DE PARES MCIC 2026
 * Universidad Distrital Francisco José de Caldas
 */

function initApp() {
  initNavigation();
  initHeroStats();
  initAgenda();
  initPlanMejoramiento();
  initEncuestas();
  initAnalisisCualitativo();
  initAutoevaluacion();
  initDocumentos();
  initModal();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}

// ==========================================================================
// 1. NAVEGACIÓN Y TABS
// ==========================================================================
function initNavigation() {
  const tabButtons = document.querySelectorAll('.nav-tab-btn');
  const tabSections = document.querySelectorAll('.tab-section');

  function switchTab(tabKey) {
    const cleanKey = tabKey.replace('tab-', '').replace('section-', '');
    tabButtons.forEach(btn => {
      btn.classList.toggle('active', btn.dataset.tab === cleanKey);
    });
    tabSections.forEach(sec => {
      sec.classList.toggle('active', sec.id === 'section-' + cleanKey);
    });
    if (window.history && window.history.replaceState) {
      window.history.replaceState(null, null, '#' + cleanKey);
    }
    window.scrollTo({ top: 0, behavior: 'instant' });
  }

  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      switchTab(btn.dataset.tab);
    });
  });

  // Handle URL hash on load
  const currentHash = window.location.hash.replace('#', '');
  if (currentHash) {
    switchTab(currentHash);
  }
}

// ==========================================================================
// 2. HERO STATS & METADATA
// ==========================================================================
function initHeroStats() {
  const meta = MCIC_DATA.meta;
  
  // Update peers card
  const peerListEl = document.getElementById('hero-peer-list');
  if (peerListEl) {
    peerListEl.innerHTML = meta.pares.slice(0, 2).map(p => `
      <div class="peer-item">
        <div class="peer-avatar">CNA</div>
        <div>
          <div class="peer-name">${p.nombre}</div>
          <div class="peer-tag">${p.rol} (${p.estado})</div>
        </div>
      </div>
    `).join('');
  }
}

// ==========================================================================
// 3. AGENDA VISITA DE PARES
// ==========================================================================
let currentAgendaVersion = 'oficial';
let currentAgendaDay = 'all';

function initAgenda() {
  const versionSelect = document.getElementById('agenda-version-select');
  const searchInput = document.getElementById('agenda-search');
  const dayButtons = document.querySelectorAll('.agenda-day-btn');

  if (versionSelect) {
    versionSelect.addEventListener('change', (e) => {
      currentAgendaVersion = e.target.value;
      renderAgenda();
    });
  }

  if (searchInput) {
    searchInput.addEventListener('input', () => {
      renderAgenda();
    });
  }

  dayButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      dayButtons.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentAgendaDay = btn.dataset.day;
      renderAgenda();
    });
  });

  renderAgenda();
}

function renderAgenda() {
  const timelineEl = document.getElementById('agenda-timeline');
  if (!timelineEl) return;

  const rawAgenda = currentAgendaVersion === 'oficial' ? MCIC_DATA.agendaOficial : MCIC_DATA.agendaPrevia;
  const searchTerm = (document.getElementById('agenda-search')?.value || '').toLowerCase();

  let filteredDays = rawAgenda;
  if (currentAgendaDay !== 'all') {
    const dayIdx = parseInt(currentAgendaDay, 10);
    filteredDays = [rawAgenda[dayIdx]].filter(Boolean);
  }

  let html = '';

  filteredDays.forEach(day => {
    const matchingItems = day.items.filter(item => {
      return item.actividad.toLowerCase().includes(searchTerm) ||
             item.participantes.toLowerCase().includes(searchTerm) ||
             item.lugar.toLowerCase().includes(searchTerm);
    });

    if (matchingItems.length === 0) return;

    html += `
      <div style="margin-bottom: 24px;">
        <div style="background: var(--ud-blue-soft); border-left: 4px solid var(--ud-blue); padding: 10px 16px; border-radius: var(--radius-sm); margin-bottom: 14px;">
          <h3 style="font-size: 15px; font-weight: 800; color: var(--ud-blue);">${day.dia}</h3>
          <p style="font-size: 12.5px; color: var(--text-muted);">${day.titulo}</p>
        </div>
        <div class="timeline">
    `;

    matchingItems.forEach(item => {
      html += `
        <div class="timeline-item">
          <div class="timeline-time-col">
            <div class="timeline-time">
              <span>🕒</span> ${item.hora}
            </div>
            <div class="timeline-badge-loc">
              📍 <strong>Lugar:</strong> ${item.lugar}
            </div>
          </div>
          <div class="timeline-content-col">
            <h4>${item.actividad}</h4>
            <div class="timeline-participants">
              <strong>Participantes:</strong> ${item.participantes}
            </div>
          </div>
        </div>
      `;
    });

    html += `</div></div>`;
  });

  if (!html) {
    html = `<div style="text-align: center; padding: 40px; color: var(--text-soft);">No se encontraron actividades con los filtros seleccionados.</div>`;
  }

  timelineEl.innerHTML = html;
}

// ==========================================================================
// 4. PLAN DE MEJORAMIENTO (CC-FR-001) - SEPARADO POR MODALIDAD
// ==========================================================================
let currentPlanModality = 'investigacion'; // 'investigacion' | 'profundizacion'

function switchPlanModality(modality) {
  if (modality !== 'investigacion' && modality !== 'profundizacion') return;
  currentPlanModality = modality;

  const btnInv = document.getElementById('plan-btn-investigacion');
  const btnProf = document.getElementById('plan-btn-profundizacion');

  if (btnInv) {
    btnInv.classList.toggle('active', modality === 'investigacion');
    if (modality === 'investigacion') {
      btnInv.style.backgroundColor = 'var(--ud-blue)';
      btnInv.style.color = '#FFFFFF';
      btnInv.style.borderColor = 'var(--ud-blue)';
    } else {
      btnInv.style.backgroundColor = '#FFFFFF';
      btnInv.style.color = 'var(--text-muted)';
      btnInv.style.borderColor = 'var(--border-color)';
    }
  }

  if (btnProf) {
    btnProf.classList.toggle('active', modality === 'profundizacion');
    if (modality === 'profundizacion') {
      btnProf.style.backgroundColor = 'var(--ud-blue)';
      btnProf.style.color = '#FFFFFF';
      btnProf.style.borderColor = 'var(--ud-blue)';
    } else {
      btnProf.style.backgroundColor = '#FFFFFF';
      btnProf.style.color = 'var(--text-muted)';
      btnProf.style.borderColor = 'var(--border-color)';
    }
  }

  updatePlanFactorFilter();
  renderPlanMejoramiento();
}
window.switchPlanModality = switchPlanModality;

function updatePlanFactorFilter() {
  const factorFilter = document.getElementById('plan-factor-filter');
  if (!factorFilter) return;
  const planGroup = getPlanGroup();
  const factors = [...new Set(planGroup.items.map(p => p.factor))];
  factorFilter.innerHTML = `<option value="all">Todos los Factores CNA (${factors.length})</option>` +
    factors.map((f, i) => `<option value="${f}">${f.split('.')[0]} - ${f.split('.')[1] ? f.split('.')[1].substring(0, 32) : f}...</option>`).join('');
}

function initPlanMejoramiento() {
  const searchInput = document.getElementById('plan-search');
  const factorFilter = document.getElementById('plan-factor-filter');
  const typeFilter = document.getElementById('plan-type-filter');
  const btnInv = document.getElementById('plan-btn-investigacion');
  const btnProf = document.getElementById('plan-btn-profundizacion');

  if (btnInv) {
    btnInv.addEventListener('click', (e) => {
      e.preventDefault();
      switchPlanModality('investigacion');
    });
  }

  if (btnProf) {
    btnProf.addEventListener('click', (e) => {
      e.preventDefault();
      switchPlanModality('profundizacion');
    });
  }

  [searchInput, factorFilter, typeFilter].forEach(el => {
    if (el) el.addEventListener('input', renderPlanMejoramiento);
  });

  updatePlanFactorFilter();
  renderPlanMejoramiento();
}

function getPlanGroup() {
  if (MCIC_DATA.planMejoramiento && MCIC_DATA.planMejoramiento[currentPlanModality]) {
    return MCIC_DATA.planMejoramiento[currentPlanModality];
  }
  return {
    modalidad: currentPlanModality === 'investigacion' ? 'Investigación' : 'Profundización',
    snies: currentPlanModality === 'investigacion' ? '17528' : '116070',
    resolucion: currentPlanModality === 'investigacion' ? 'Resolución 16163 (05/09/2023)' : 'Resolución 9925 (21/06/2023)',
    repEstudiantil: currentPlanModality === 'investigacion' ? 'Cristian Leonardo Alape' : 'Jhon Jairo Soler Umbarila',
    repDocente: currentPlanModality === 'investigacion' ? 'Dr. Leonardo Plazas Nossa' : 'Dr. Andrés Leonardo Jutinico',
    archivoSoporte: currentPlanModality === 'investigacion' ? 'CC-FR-001 Plan de mejoramiento INV.xlsx' : 'CC-FR-001 Plan de mejoramiento PROF.xlsx',
    rutaSoporte: currentPlanModality === 'investigacion' ? 'AUTOEVALUACION/MCIC- INVESTIGACIÓN/CC-FR-001 Plan de mejoramiento INV.xlsx' : 'AUTOEVALUACION/MCICI- PRODUNDIZACIÓN/CC-FR-001 Plan de mejoramiento PROF.xlsx',
    items: []
  };
}

function renderPlanMejoramiento() {
  const gridEl = document.getElementById('plan-grid');
  const bannerEl = document.getElementById('plan-modality-banner');
  if (!gridEl) return;

  const planGroup = getPlanGroup();
  const isInv = currentPlanModality === 'investigacion';

  // Render Modality Banner
  if (bannerEl) {
    bannerEl.innerHTML = `
      <div class="modality-banner-header">
        <div class="modality-banner-title">
          <span>${isInv ? '🔬' : '💼'}</span>
          <span>Plan de Mejoramiento Oficial — Modalidad ${planGroup.modalidad}</span>
          <span class="factor-tag" style="background: ${isInv ? 'var(--ud-blue-soft)' : '#FEF3C7'}; color: ${isInv ? 'var(--ud-blue)' : '#92400E'};">
            SNIES ${planGroup.snies}
          </span>
        </div>
        <a href="${encodeURI(planGroup.rutaSoporte)}" download="${planGroup.archivoSoporte}" class="btn btn-outline" style="font-size: 12px; font-weight: 700; color: var(--ud-blue); background: white; text-decoration: none;">
          📥 Descargar Matriz Oficial Excel (${planGroup.modalidad})
        </a>
      </div>

      <div class="modality-banner-grid">
        <div class="modality-banner-item">
          <span>Registro Calificado & Vigencia:</span>
          <strong>${planGroup.resolucion} (${planGroup.vigencia || '7 años'})</strong>
        </div>
        <div class="modality-banner-item">
          <span>Representante Docente:</span>
          <strong>${planGroup.repDocente}</strong>
        </div>
        <div class="modality-banner-item">
          <span>Representante Estudiantil:</span>
          <strong>${planGroup.repEstudiantil}</strong>
        </div>
        <div class="modality-banner-item">
          <span>Archivo Institucional:</span>
          <strong>${planGroup.archivoSoporte}</strong>
        </div>
      </div>
    `;
  }

  const searchTerm = (document.getElementById('plan-search')?.value || '').toLowerCase();
  const selectedFactor = document.getElementById('plan-factor-filter')?.value || 'all';
  const selectedType = document.getElementById('plan-type-filter')?.value || 'all';

  const items = planGroup.items.filter(p => {
    const matchesFactor = selectedFactor === 'all' || p.factor === selectedFactor;
    const matchesType = selectedType === 'all' || p.tipo.toLowerCase().includes(selectedType.toLowerCase());
    const matchesSearch = p.factor.toLowerCase().includes(searchTerm) ||
                          (p.proyecto && p.proyecto.toLowerCase().includes(searchTerm)) ||
                          (p.descripcion && p.descripcion.toLowerCase().includes(searchTerm)) ||
                          (p.meta && p.meta.toLowerCase().includes(searchTerm)) ||
                          (p.responsable && p.responsable.toLowerCase().includes(searchTerm));
    return matchesFactor && matchesType && matchesSearch;
  });

  if (items.length === 0) {
    gridEl.innerHTML = `<div style="grid-column: 1/-1; text-align: center; padding: 40px; color: var(--text-soft);">No se encontraron proyectos o acciones en el Plan de Mejoramiento de ${planGroup.modalidad} con los filtros dados.</div>`;
    return;
  }

  gridEl.innerHTML = items.map((p, index) => {
    const isFortaleza = p.tipo.toLowerCase().includes('fortaleza');
    const factorNum = p.factor.split('.')[0].trim();
    const itemId = p.id || `item-${index}`;
    
    return `
      <div class="plan-card ${isFortaleza ? 'fortaleza' : ''}">
        <div>
          <div class="plan-card-header">
            <span class="factor-tag">${factorNum}</span>
            <span class="plan-modality-tag ${isInv ? 'tag-inv' : 'tag-prof'}">
              ${isInv ? '🔬 INVESTIGACIÓN (SNIES 17528)' : '💼 PROFUNDIZACIÓN (SNIES 116070)'}
            </span>
            <span class="type-tag ${isFortaleza ? 'fortaleza' : 'oportunidad'}">${p.tipo}</span>
          </div>
          <h3>${p.proyecto || p.factor}</h3>
          <p class="plan-card-desc">${p.descripcion}</p>
          
          <div class="plan-meta-box">
            <div class="plan-meta-row">
              <span>Indicador:</span>
              <span title="${p.indicador}">${p.indicador.length > 55 ? p.indicador.substring(0, 55) + '...' : p.indicador}</span>
            </div>
            <div class="plan-meta-row">
              <span>Meta:</span>
              <span title="${p.meta}">${p.meta.length > 55 ? p.meta.substring(0, 55) + '...' : p.meta}</span>
            </div>
            <div class="plan-meta-row">
              <span>Responsable:</span>
              <span>${p.responsable}</span>
            </div>
            <div class="plan-meta-row">
              <span>Periodicidad:</span>
              <span>${p.periodicidad}</span>
            </div>
          </div>
        </div>

        <div class="plan-card-actions">
          <span style="font-size: 11px; color: var(--text-soft);">Prioridad: <strong>${p.prioridad || 'Alta'}</strong></span>
          <button class="btn btn-outline" onclick="openPlanModal('${itemId}')">
            Ver detalle completo ➔
          </button>
        </div>
      </div>
    `;
  }).join('');
}

window.openPlanModal = function(itemId) {
  const planGroup = getPlanGroup();
  let p = planGroup.items.find(item => item.id === itemId);
  if (!p && typeof itemId === 'number') {
    p = planGroup.items[itemId];
  }
  if (!p && planGroup.items.length > 0) {
    p = planGroup.items[0];
  }
  if (!p) return;

  const content = `
    <div style="display: flex; gap: 8px; margin-bottom: 12px; align-items: center; flex-wrap: wrap;">
      <span class="factor-tag">${p.factor.split('.')[0]}</span>
      <span class="type-tag ${p.tipo.toLowerCase().includes('fortaleza') ? 'fortaleza' : 'oportunidad'}">${p.tipo}</span>
      <span class="factor-tag" style="background: #FEF3C7; color: #92400E;">Modalidad: ${planGroup.modalidad} (SNIES ${planGroup.snies})</span>
      <span style="font-size: 12px; color: var(--text-soft); margin-left: auto;">Origen: ${p.origen}</span>
    </div>

    <h3 style="font-size: 17px; font-weight: 800; color: var(--ud-blue); margin-bottom: 8px;">${p.proyecto || p.factor}</h3>
    <h4 style="font-size: 13.5px; font-weight: 700; color: var(--text-muted); margin-bottom: 16px;">${p.factor}</h4>

    <div style="background: var(--bg-subtle); padding: 14px; border-radius: var(--radius-md); margin-bottom: 16px;">
      <strong style="font-size: 12px; text-transform: uppercase; color: var(--text-soft); display: block; margin-bottom: 4px;">Descripción de la Situación:</strong>
      <p style="font-size: 13.5px; color: var(--text-main); line-height: 1.5;">${p.descripcion}</p>
    </div>

    <div style="background: white; border: 1px solid var(--border-color); border-radius: var(--radius-md); padding: 14px; margin-bottom: 16px;">
      <strong style="font-size: 12px; text-transform: uppercase; color: var(--text-soft); display: block; margin-bottom: 4px;">Objetivo del Proyecto / Acción Global:</strong>
      <p style="font-size: 13.5px; color: var(--text-main); line-height: 1.5;">${p.objetivo}</p>
    </div>

    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 16px;">
      <div style="background: var(--ud-blue-soft); padding: 12px; border-radius: var(--radius-sm);">
        <strong style="font-size: 11.5px; color: var(--ud-blue-dark); text-transform: uppercase; display: block; margin-bottom: 4px;">Indicador de Cumplimiento:</strong>
        <p style="font-size: 13px; font-weight: 700; color: var(--ud-blue);">${p.indicador}</p>
        <span style="font-size: 11px; color: var(--text-soft);">Tipo: ${p.tipo_indicador || 'Resultado'}</span>
      </div>
      <div style="background: var(--success-bg); padding: 12px; border-radius: var(--radius-sm);">
        <strong style="font-size: 11.5px; color: var(--success); text-transform: uppercase; display: block; margin-bottom: 4px;">Meta Proyectada:</strong>
        <p style="font-size: 13px; font-weight: 700; color: #14532d;">${p.meta}</p>
        <span style="font-size: 11px; color: var(--text-soft);">Línea Base: ${p.linea_base || '2026-1'}</span>
      </div>
    </div>

    <div style="margin-bottom: 16px;">
      <strong style="font-size: 13px; color: var(--ud-blue-dark); display: block; margin-bottom: 6px;">Actividades Requeridas para el Logro:</strong>
      <ul style="padding-left: 20px; font-size: 13px; color: var(--text-muted); line-height: 1.6;">
        ${p.actividades.map(a => `<li>${a}</li>`).join('')}
      </ul>
    </div>

    <div style="background: var(--bg-app); border: 1px solid var(--border-color); border-radius: var(--radius-md); padding: 12px; font-size: 12.5px; display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px;">
      <div>
        <span style="color: var(--text-soft); display: block;">Responsable:</span>
        <strong>${p.responsable}</strong>
      </div>
      <div>
        <span style="color: var(--text-soft); display: block;">Periodicidad:</span>
        <strong>${p.periodicidad}</strong>
      </div>
      <div>
        <span style="color: var(--text-soft); display: block;">Recursos:</span>
        <strong>${p.recursos}</strong>
      </div>
    </div>
  `;

  showModal(`Detalle Plan de Mejoramiento (${planGroup.modalidad}) — ${p.factor.split('.')[0]}`, content);
};

// ==========================================================================
// 5. ENCUESTAS Y RESULTADOS CNA
// ==========================================================================
let currentStakeholder = 'estudiantes';

function initEncuestas() {
  const stakeholderBtns = document.querySelectorAll('.stakeholder-btn');
  const factorSelect = document.getElementById('survey-factor-select');
  const searchInput = document.getElementById('survey-search');

  stakeholderBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      stakeholderBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentStakeholder = btn.dataset.stakeholder;
      populateSurveyFactors();
      renderEncuestas();
    });
  });

  if (factorSelect) {
    factorSelect.addEventListener('change', renderEncuestas);
  }
  if (searchInput) {
    searchInput.addEventListener('input', renderEncuestas);
  }

  populateSurveyFactors();
  renderEncuestas();
}

function populateSurveyFactors() {
  const factorSelect = document.getElementById('survey-factor-select');
  if (!factorSelect) return;

  const items = MCIC_DATA.encuestas[currentStakeholder] || [];
  const factors = [...new Set(items.map(i => i.factor))];

  factorSelect.innerHTML = `<option value="all">Todos los Factores Evaluados (${factors.length})</option>` +
    factors.map(f => `<option value="${f}">${f.split('.')[0]} - ${f.split('.')[1] ? f.split('.')[1].substring(0, 32) : f}...</option>`).join('');
}

function renderEncuestas() {
  const listEl = document.getElementById('survey-items-list');
  const statsEl = document.getElementById('survey-stats-summary');
  if (!listEl) return;

  const items = MCIC_DATA.encuestas[currentStakeholder] || [];
  const selectedFactor = document.getElementById('survey-factor-select')?.value || 'all';
  const searchTerm = (document.getElementById('survey-search')?.value || '').toLowerCase();

  const filtered = items.filter(it => {
    const matchesFactor = selectedFactor === 'all' || it.factor === selectedFactor;
    const matchesSearch = it.item.toLowerCase().includes(searchTerm) || it.factor.toLowerCase().includes(searchTerm);
    return matchesFactor && matchesSearch;
  });

  // Calculate stats
  if (statsEl) {
    const totalResp = currentStakeholder === 'estudiantes' ? 35 : (currentStakeholder === 'docentes' ? 9 : 2);
    const avgFav = filtered.length > 0 ? (filtered.reduce((acc, c) => acc + c.favorablePct, 0) / filtered.length).toFixed(1) : 0;
    
    statsEl.innerHTML = `
      <div class="stat-box">
        <div class="stat-val">${filtered.length}</div>
        <div class="stat-lbl">Ítems Evaluados</div>
      </div>
      <div class="stat-box">
        <div class="stat-val">${totalResp}</div>
        <div class="stat-lbl">Muestra Participante</div>
      </div>
      <div class="stat-box">
        <div class="stat-val" style="color: ${avgFav >= 75 ? 'var(--success)' : (avgFav >= 60 ? 'var(--warning)' : 'var(--danger)')};">${avgFav}%</div>
        <div class="stat-lbl">Satisfacción Favorable Promedio (4 y 5)</div>
      </div>
    `;
  }

  if (filtered.length === 0) {
    listEl.innerHTML = `<div style="text-align: center; padding: 40px; color: var(--text-soft);">No se encontraron preguntas de la encuesta con los filtros dados.</div>`;
    return;
  }

  listEl.innerHTML = filtered.slice(0, 50).map((it, idx) => {
    let badgeClass = 'high';
    if (it.favorablePct < 60) badgeClass = 'low';
    else if (it.favorablePct < 75) badgeClass = 'medium';

    return `
      <div class="survey-item-card">
        <div class="survey-header-row">
          <span class="survey-factor-tag">${it.factor}</span>
          <span class="satisfaction-badge ${badgeClass}">
            ★ ${it.favorablePct}% Favorable (Escala 4 y 5)
          </span>
        </div>

        <div class="survey-question-text">${it.item}</div>

        <!-- Horizontal Distribution Bar -->
        <div class="distribution-bar">
          ${it.distribucion.map(d => {
            const pctNum = parseFloat(d.porcentaje.replace('%', '')) || 0;
            if (pctNum <= 0) return '';
            let optClass = 'opt-other';
            if (d.opcion === '1') optClass = 'opt-1';
            else if (d.opcion === '2') optClass = 'opt-2';
            else if (d.opcion === '3') optClass = 'opt-3';
            else if (d.opcion === '4') optClass = 'opt-4';
            else if (d.opcion === '5') optClass = 'opt-5';

            return `
              <div class="dist-slice ${optClass}" style="width: ${pctNum}%;" title="Opción ${d.opcion}: ${d.cantidad} respuestas (${d.porcentaje})">
                ${pctNum >= 12 ? d.opcion + ' (' + Math.round(pctNum) + '%)' : ''}
              </div>
            `;
          }).join('')}
        </div>

        <div class="dist-legend">
          <span>Respuestas: <strong>${it.respuestas}</strong></span>
          ${it.distribucion.map(d => `
            <span>
              <span class="legend-dot" style="background: ${getOptColor(d.opcion)};"></span>
              ${d.opcion}: ${d.cantidad} (${d.porcentaje})
            </span>
          `).join('')}
        </div>
      </div>
    `;
  }).join('') + (filtered.length > 50 ? `<div style="text-align: center; padding: 12px; color: var(--text-soft); font-size: 13px;">Mostrando los primeros 50 ítems de ${filtered.length}. Utilice los filtros para explorar factores específicos.</div>` : '');
}

function getOptColor(opt) {
  switch (opt) {
    case '1': return '#EF4444';
    case '2': return '#F97316';
    case '3': return '#FBBF24';
    case '4': return '#34D399';
    case '5': return '#10B981';
    default: return '#64748B';
  }
}

// ==========================================================================
// 6. ANÁLISIS CUALITATIVO
// ==========================================================================
function initAnalisisCualitativo() {
  const container = document.getElementById('analisis-cualitativo-container');
  if (!container) return;

  const estamentos = [
    { key: 'estudiantes', label: 'Estudiantes (35 participantes)', color: 'var(--ud-blue)' },
    { key: 'docentes', label: 'Docentes (9 participantes)', color: '#047857' },
    { key: 'directivos', label: 'Directivos (2 participantes)', color: '#B45309' }
  ];

  let html = '';

  estamentos.forEach(est => {
    const list = MCIC_DATA.analisisCualitativo[est.key] || [];
    html += `
      <div style="margin-bottom: 30px;">
        <div style="background: var(--bg-subtle); border-left: 4px solid ${est.color}; padding: 10px 18px; border-radius: var(--radius-sm); margin-bottom: 16px;">
          <h3 style="font-size: 16px; font-weight: 800; color: var(--text-main);">Apreciación de ${est.label}</h3>
          <p style="font-size: 12.5px; color: var(--text-muted);">Síntesis cualitativa extraída de los instrumentos aplicados en 2026-1</p>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(420px, 1fr)); gap: 16px;">
          ${list.map(f => `
            <div class="card" style="margin-bottom: 0;">
              <h4 style="font-size: 14px; font-weight: 800; color: var(--ud-blue); margin-bottom: 10px; border-bottom: 1px solid var(--border-color); padding-bottom: 6px;">
                ${f.factor}
              </h4>
              ${f.parrafos.map(p => `
                <p style="font-size: 13px; color: var(--text-muted); line-height: 1.55; margin-bottom: 8px;">
                  ${p}
                </p>
              `).join('')}
            </div>
          `).join('')}
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

// ==========================================================================
// 7. AUTOEVALUACIÓN, PONDERACIONES & SWOT
// ==========================================================================
function initAutoevaluacion() {
  renderPonderacionesCNA();
  renderSWOT();
  renderDocentes();
  renderProyectosRoadmap();
}

function renderPonderacionesCNA() {
  const container = document.getElementById('cna-weights-grid');
  if (!container) return;

  container.innerHTML = MCIC_DATA.factorsCNA.map(f => `
    <div class="weight-card">
      <div class="weight-header">
        <span class="weight-factor-num">Factor ${f.numero}</span>
        <span class="weight-pct-pill">${f.ponderacion}</span>
      </div>
      <h4>${f.nombre}</h4>
      <ul class="caract-list">
        ${f.caracteristicas.map(c => `
          <li>
            <span>${c.nombre}</span>
            <span>${c.ponderacion}</span>
          </li>
        `).join('')}
      </ul>
    </div>
  `).join('');
}

function renderSWOT() {
  const container = document.getElementById('swot-container');
  if (!container) return;

  container.innerHTML = MCIC_DATA.swot.map(s => `
    <div class="swot-factor-block">
      <div class="swot-factor-title">${s.titulo}</div>
      <div class="swot-columns">
        <div class="swot-col fortalezas">
          <h5><span>✓</span> Fortalezas Institucionales</h5>
          <ul class="swot-list">
            ${s.items.filter(i => i.fortaleza).map(i => `<li>${i.fortaleza}</li>`).join('')}
          </ul>
        </div>
        <div class="swot-col oportunidades">
          <h5><span>▲</span> Oportunidades de Mejoramiento</h5>
          <ul class="swot-list">
            ${s.items.filter(i => i.oportunidad).map(i => `<li>${i.oportunidad}</li>`).join('')}
          </ul>
        </div>
      </div>
    </div>
  `).join('');
}

function renderDocentes() {
  const container = document.getElementById('docentes-table-body');
  if (!container) return;

  container.innerHTML = MCIC_DATA.docentes.map((d, i) => `
    <tr style="border-bottom: 1px solid var(--border-color);">
      <td style="padding: 10px 12px; font-size: 13px; font-weight: 700; color: var(--text-main);">${i + 1}. ${d.nombre}</td>
      <td style="padding: 10px 12px; font-size: 12.5px; color: var(--ud-blue); font-weight: 600;">${d.vinculacion}</td>
    </tr>
  `).join('');
}

function renderProyectosRoadmap() {
  const container = document.getElementById('proyectos-table-body');
  if (!container) return;

  container.innerHTML = MCIC_DATA.proyectosEnfasis.map(p => `
    <tr style="border-bottom: 1px solid var(--border-color); font-weight: ${p.enfasis === 'Total' ? '800' : 'normal'};">
      <td style="padding: 8px 12px; font-size: 13px;">${p.enfasis}</td>
      <td style="padding: 8px 12px; font-size: 14px; text-align: right; color: var(--ud-blue); font-weight: 800;">${p.proyectos}</td>
    </tr>
  `).join('');

  const actContainer = document.getElementById('actividades-2026-container');
  if (actContainer) {
    actContainer.innerHTML = MCIC_DATA.actividades2026.map(a => `
      <div style="background: white; border: 1px solid var(--border-color); border-radius: var(--radius-sm); padding: 12px; margin-bottom: 10px;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
          <strong style="font-size: 13px; color: var(--ud-blue);">${a.actividad}</strong>
          <span style="font-size: 11.5px; font-weight: 700; color: var(--ud-gold-dark);">${a.fecha}</span>
        </div>
        <p style="font-size: 12.5px; color: var(--text-muted); margin-bottom: 4px;">${a.objetivo}</p>
        <span style="font-size: 11px; color: var(--text-soft);">Participantes: ${a.participantes}</span>
      </div>
    `).join('');
  }
}

// ==========================================================================
// 8. DOCUMENTOS REPOSITORIO
// ==========================================================================
function initDocumentos() {
  const gridEl = document.getElementById('documentos-grid');
  if (!gridEl) return;

  gridEl.innerHTML = MCIC_DATA.documentos.map(doc => {
    const isDocx = doc.tipo === 'DOCX';
    return `
      <div class="doc-card">
        <div>
          <div class="doc-card-top">
            <div class="doc-icon ${isDocx ? 'docx' : 'xlsx'}">${doc.tipo}</div>
            <div>
              <div class="doc-name">${doc.nombre}</div>
              <span style="font-size: 11px; color: var(--ud-blue); font-weight: 700;">${doc.categoria}</span>
            </div>
          </div>
          <p class="doc-desc">${doc.descripcion}</p>
          <code class="doc-path-code" title="${doc.ruta}">${doc.ruta}</code>
        </div>
        <div class="doc-footer">
          <span>Tamaño: <strong>${doc.tamano}</strong></span>
          <a href="${encodeURI(doc.ruta)}" download="${doc.nombre}" class="btn btn-outline" style="padding: 4px 10px; font-size: 11px; text-decoration: none; color: var(--ud-blue); font-weight: 700;">
            ⬇️ Abrir / Descargar
          </a>
        </div>
      </div>
    `;
  }).join('');
}

// ==========================================================================
// 9. MODAL COMPONENT
// ==========================================================================
function initModal() {
  const overlay = document.getElementById('modal-overlay');
  const closeBtn = document.getElementById('modal-close-btn');

  if (closeBtn && overlay) {
    closeBtn.addEventListener('click', () => {
      overlay.classList.remove('open');
    });
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) {
        overlay.classList.remove('open');
      }
    });
  }
}

function showModal(title, bodyHtml) {
  const overlay = document.getElementById('modal-overlay');
  const titleEl = document.getElementById('modal-title');
  const bodyEl = document.getElementById('modal-body');

  if (overlay && titleEl && bodyEl) {
    titleEl.textContent = title;
    bodyEl.innerHTML = bodyHtml;
    overlay.classList.add('open');
  }
}
