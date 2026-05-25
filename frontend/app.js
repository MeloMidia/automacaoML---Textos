/* =========================================================
   AutomacaoML — App Logic (app.js)
   ========================================================= */

const App = (() => {

  // ── Estado ──────────────────────────────────────────────
  let _clients = [];
  let _selected = new Set();
  let _running  = false;
  let _eventSource = null;

  let _clientSheets = {};   // clientId → string[]  (undefined = não carregado)
  let _selectedSheets = {}; // clientId → Set<string>  (undefined = todas)

  // ── Elementos DOM ───────────────────────────────────────
  const $ = (id) => document.getElementById(id);

  const el = {
    clientsList:    () => $('clients-list'),
    clientsCount:   () => $('clients-count'),
    selectedCount:  () => $('selected-count'),
    btnRun:         () => $('btn-run'),
    btnCancel:      () => $('btn-cancel'),
    terminal:       () => $('terminal-output'),
    statusBadge:    () => $('status-badge'),
    statusLabel:    () => $('status-badge').querySelector('.status-label'),
    resultsOverlay: () => $('results-overlay'),
    metricCreated:  () => $('metric-created'),
    metricSkipped:  () => $('metric-skipped'),
    errorToast:     () => $('error-toast'),
    errorMessage:   () => $('error-message'),
    btnReset:       () => $('btn-reset'),
  };

  // ── Inicialização ────────────────────────────────────────
  async function init() {
    await loadClients();
  }

  // ── Carregar clientes do Drive ───────────────────────────
  async function loadClients() {
    renderLoadingState();
    try {
      const res  = await fetch('/api/clients');
      const data = await res.json();

      if (!data.ok) {
        renderErrorState(data.error || 'Falha ao conectar com o Google Drive.');
        return;
      }

      _clients = data.clients || [];
      el.clientsCount().textContent = _clients.length;

      if (_clients.length === 0) {
        renderEmptyState();
      } else {
        renderClientsList();
      }
    } catch (_) {
      renderErrorState('Não foi possível conectar ao servidor. Verifique se o backend está rodando.');
    }
  }

  // ── Renderizar estados da lista ──────────────────────────
  function renderLoadingState() {
    el.clientsList().innerHTML = `
      <div class="loading-state">
        <div class="spinner"></div>
        <span>Conectando ao Google Drive...</span>
      </div>`;
  }

  function renderErrorState(msg) {
    el.clientsList().innerHTML = `
      <div class="error-state">
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
        </svg>
        <span>${escapeHtml(msg)}</span>
      </div>`;
  }

  function renderEmptyState() {
    el.clientsList().innerHTML = `
      <div class="empty-state">
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
        </svg>
        <span>Nenhum cliente encontrado na pasta MERCADO LIVRE do Drive.</span>
      </div>`;
  }

  function renderClientsList() {
    const list = el.clientsList();
    list.innerHTML = '';
    _clients.forEach(client => {
      const wrapper = document.createElement('div');
      wrapper.className = 'client-wrapper';
      wrapper.dataset.id = client.id;
      wrapper.innerHTML = `
        <div class="client-item" data-id="${client.id}">
          <input type="checkbox" class="client-checkbox" id="chk-${client.id}"
                 onchange="App.toggleClient('${client.id}')" />
          <label class="client-name" for="chk-${client.id}" title="${escapeHtml(client.name)}">
            ${escapeHtml(client.name)}
          </label>
        </div>
        <div class="client-sheets hidden" id="sheets-${client.id}"></div>`;

      wrapper.querySelector('.client-item').addEventListener('click', (e) => {
        if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'LABEL') {
          toggleClient(client.id);
        }
      });
      list.appendChild(wrapper);
    });
  }

  // ── Seleção de clientes ──────────────────────────────────
  async function toggleClient(id) {
    if (_selected.has(id)) {
      _selected.delete(id);
      delete _selectedSheets[id];
    } else {
      _selected.add(id);
    }
    updateSelectionUI();
    if (_selected.has(id)) {
      await fetchAndShowSheets(id);
    }
  }

  async function selectAll() {
    const newIds = _clients.filter(c => !_selected.has(c.id)).map(c => c.id);
    _clients.forEach(c => _selected.add(c.id));
    updateSelectionUI();
    await Promise.all(newIds.map(id => fetchAndShowSheets(id)));
  }

  function clearAll() {
    _selected.clear();
    _selectedSheets = {};
    updateSelectionUI();
  }

  function updateSelectionUI() {
    _clients.forEach(client => {
      const chk      = document.getElementById(`chk-${client.id}`);
      const item     = chk?.closest('.client-item');
      const sheetsEl = document.getElementById(`sheets-${client.id}`);
      const isSel    = _selected.has(client.id);

      if (chk)      chk.checked = isSel;
      if (item)     item.classList.toggle('selected', isSel);
      if (sheetsEl) sheetsEl.classList.toggle('hidden', !isSel);
    });

    el.selectedCount().textContent = _selected.size;
    el.btnRun().disabled = _selected.size === 0 || _running || hasEmptySheetSelection();
  }

  function hasEmptySheetSelection() {
    for (const id of _selected) {
      const sheets   = _clientSheets[id];
      const selected = _selectedSheets[id];
      if (sheets && sheets.length > 0 && selected && selected.size === 0) return true;
    }
    return false;
  }

  // ── Carregamento de abas ─────────────────────────────────
  async function fetchAndShowSheets(clientId) {
    const sheetsEl = document.getElementById(`sheets-${clientId}`);
    if (!sheetsEl) return;

    // Cache hit
    if (_clientSheets[clientId] !== undefined) {
      renderSheetsForClient(clientId, _clientSheets[clientId]);
      return;
    }

    sheetsEl.innerHTML = `
      <div class="sheets-loading">
        <div class="spinner-sm"></div>
        <span>Carregando páginas...</span>
      </div>`;

    try {
      const res  = await fetch(`/api/sheets?client_id=${encodeURIComponent(clientId)}`);
      const data = await res.json();

      if (!_selected.has(clientId)) return; // cliente foi desmarcado enquanto carregava

      _clientSheets[clientId] = (data.ok && data.sheets) ? data.sheets : [];
      renderSheetsForClient(clientId, _clientSheets[clientId]);
    } catch (_) {
      _clientSheets[clientId] = [];
      sheetsEl.innerHTML = `<div class="sheets-message">Não foi possível carregar páginas.</div>`;
    }
  }

  function renderSheetsForClient(clientId, sheets) {
    const sheetsEl = document.getElementById(`sheets-${clientId}`);
    const itemEl   = document.querySelector(`.client-wrapper[data-id="${clientId}"] .client-item`);
    if (!sheetsEl) return;

    if (!sheets || sheets.length === 0) {
      sheetsEl.innerHTML = `<div class="sheets-message">Nenhuma aba encontrada na planilha.</div>`;
      if (itemEl) itemEl.classList.remove('has-sheets');
      return;
    }

    if (itemEl) itemEl.classList.add('has-sheets');

    // Inicializa seleção com todas as abas se ainda não definido
    if (!_selectedSheets[clientId]) {
      _selectedSheets[clientId] = new Set(sheets);
    }

    const sel = _selectedSheets[clientId];

    sheetsEl.innerHTML = `
      <div class="sheets-header">
        <span class="sheets-label">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="3" y="3" width="18" height="18" rx="2"/>
            <line x1="3" y1="9" x2="21" y2="9"/>
            <line x1="9" y1="21" x2="9" y2="9"/>
          </svg>
          Páginas
        </span>
        <div class="sheets-quick">
          <button class="btn-sheet-quick" onclick="App.selectAllSheets('${clientId}')">Todas</button>
          <button class="btn-sheet-quick" onclick="App.clearAllSheets('${clientId}')">Nenhuma</button>
        </div>
      </div>
      <div class="sheets-grid">
        ${sheets.map((sheet, idx) => `
          <div class="sheet-item${sel.has(sheet) ? ' selected' : ''}">
            <input type="checkbox" class="sheet-checkbox" id="sh-${clientId}-${idx}"
                   ${sel.has(sheet) ? 'checked' : ''}
                   onchange="App.toggleSheetIdx('${clientId}', ${idx}, this.checked)" />
            <label class="sheet-name" for="sh-${clientId}-${idx}">${escapeHtml(sheet)}</label>
          </div>`).join('')}
      </div>`;

    updateSelectionUI(); // re-valida botão run
  }

  function toggleSheetIdx(clientId, idx, checked) {
    const sheetName = _clientSheets[clientId]?.[idx];
    if (sheetName === undefined) return;

    if (!_selectedSheets[clientId]) {
      _selectedSheets[clientId] = new Set(_clientSheets[clientId] || []);
    }
    const set = _selectedSheets[clientId];
    if (checked) { set.add(sheetName); } else { set.delete(sheetName); }

    const item = document.getElementById(`sh-${clientId}-${idx}`)?.closest('.sheet-item');
    if (item) item.classList.toggle('selected', checked);

    updateSelectionUI(); // re-valida botão run
  }

  function selectAllSheets(clientId) {
    _selectedSheets[clientId] = new Set(_clientSheets[clientId] || []);
    renderSheetsForClient(clientId, _clientSheets[clientId] || []);
  }

  function clearAllSheets(clientId) {
    _selectedSheets[clientId] = new Set();
    renderSheetsForClient(clientId, _clientSheets[clientId] || []);
  }

  // ── Execução ─────────────────────────────────────────────
  async function run() {
    if (_running || _selected.size === 0) return;

    // Valida seleção de abas
    for (const id of _selected) {
      const sheets   = _clientSheets[id];
      const selected = _selectedSheets[id];
      if (sheets && sheets.length > 0 && selected && selected.size === 0) {
        const name = _clients.find(c => c.id === id)?.name || id;
        showError(`"${name}": nenhuma página selecionada. Selecione ao menos uma aba.`);
        return;
      }
    }

    const allSheets = _clientSheets;
    const selectedClients = _clients
      .filter(c => _selected.has(c.id))
      .map(c => {
        const clientObj = { id: c.id, name: c.name };
        const selSheets = _selectedSheets[c.id];
        const allSh     = allSheets[c.id];
        // Envia filtro apenas se não são todas as abas
        if (selSheets && allSh && selSheets.size < allSh.length) {
          clientObj.sheets = [...selSheets];
        }
        return clientObj;
      });

    const delay = parseInt($('input-delay')?.value ?? '45', 10) || 45;

    _running = true;
    setStatus('running', 'Processando...');
    el.btnRun().style.display  = 'none';
    el.btnCancel().style.display = '';

    clearTerminal();
    appendLog(`▶  Iniciando para ${selectedClients.length} cliente(s)... (delay: ${delay}s/produto)`, 'log-info');
    appendLog('', 'log-dim');

    try {
      const res  = await fetch('/api/run', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ clients: selectedClients, delay_seconds: delay }),
      });
      const data = await res.json();

      if (!data.ok) {
        finishWithError(data.error || 'Erro ao iniciar automação.');
        return;
      }

      connectStream(data.job_id);
    } catch (_) {
      finishWithError('Não foi possível conectar ao servidor.');
    }
  }

  // ── SSE Stream ───────────────────────────────────────────
  function connectStream(jobId) {
    if (_eventSource) _eventSource.close();

    _eventSource = new EventSource(`/api/stream/${jobId}`);

    _eventSource.onmessage = (e) => {
      let event;
      try { event = JSON.parse(e.data); } catch { return; }

      switch (event.type) {
        case 'log':       appendLog(event.text); break;
        case 'done':      finishSuccess(event.created, event.skipped); break;
        case 'error':     finishWithError(event.message); break;
        case 'cancelled': finishCancelled(); break;
        case 'ping': break;
      }
    };

    _eventSource.onerror = () => {
      if (_running) finishWithError('Conexão com o servidor foi interrompida.');
    };
  }

  // ── Estados finais ───────────────────────────────────────
  function finishSuccess(created, skipped) {
    if (_eventSource) { _eventSource.close(); _eventSource = null; }
    _running = false;
    appendLog('', 'log-dim');
    appendLog(`✅  Concluído! ${created} criados | ${skipped} pulados`, 'log-success');
    setStatus('done', 'Concluído');
    resetRunButton();
    showResults(created, skipped);
  }

  function finishWithError(msg) {
    if (_eventSource) { _eventSource.close(); _eventSource = null; }
    _running = false;
    appendLog('', 'log-dim');
    appendLog(`❌  Erro: ${msg}`, 'log-error');
    setStatus('error', 'Erro');
    resetRunButton();
    showError(msg);
  }

  function finishCancelled() {
    if (_eventSource) { _eventSource.close(); _eventSource = null; }
    _running = false;
    appendLog('', 'log-dim');
    appendLog('⚠️  Automação cancelada pelo usuário.', 'log-warning');
    setStatus('idle', 'Aguardando');
    resetRunButton();
  }

  function resetRunButton() {
    const btnCancel = el.btnCancel();
    btnCancel.style.display = 'none';
    btnCancel.disabled = false;
    const span = btnCancel.querySelector('span');
    if (span) span.textContent = 'Cancelar Automação';
    el.btnRun().style.display    = '';
    el.btnRun().disabled = _selected.size === 0 || hasEmptySheetSelection();
  }

  async function cancel() {
    try { await fetch('/api/cancel', { method: 'POST' }); } catch (_) {}
    const btnCancel = el.btnCancel();
    btnCancel.disabled = true;
    const span = btnCancel.querySelector('span');
    if (span) span.textContent = 'Cancelando...';
  }

  // ── Terminal ─────────────────────────────────────────────
  function clearTerminal() {
    el.terminal().innerHTML = '';
  }

  function appendLog(text, forceClass = null) {
    const span = document.createElement('span');
    span.className = `log-line ${forceClass || classifyLine(text)}`;
    span.textContent = text;

    const term = el.terminal();
    const placeholder = term.querySelector('.terminal-placeholder');
    if (placeholder) placeholder.remove();

    term.appendChild(span);
    term.scrollTop = term.scrollHeight;
  }

  function classifyLine(text) {
    if (!text || text.trim() === '') return 'log-dim';
    if (/✅|Autenticado|Concluído|criados/.test(text)) return 'log-success';
    if (/❌|Erro|erro|Error/.test(text))               return 'log-error';
    if (/⏳|aguardando|Rate limit/.test(text))          return 'log-warning';
    if (/⏭️|Já existe|pulados/.test(text))             return 'log-warning';
    if (/📊|📁|📝|📦|👥|🔑/.test(text))               return 'log-info';
    if (/✨|Gerando/.test(text))                        return 'log-generating';
    if (/^[═─▶\s]+$/.test(text))                        return 'log-separator';
    return 'log-default';
  }

  // ── Status badge ─────────────────────────────────────────
  function setStatus(type, label) {
    const badge = el.statusBadge();
    badge.className = `status-badge status-${type}`;
    el.statusLabel().textContent = label;
  }

  // ── Resultados ───────────────────────────────────────────
  function showResults(created, skipped) {
    el.metricCreated().textContent = created;
    el.metricSkipped().textContent = skipped;
    el.resultsOverlay().classList.remove('hidden');
  }

  function closeResults() {
    el.resultsOverlay().classList.add('hidden');
  }

  // ── Toast de erro ─────────────────────────────────────────
  function showError(msg) {
    el.errorMessage().textContent = msg;
    el.errorToast().classList.remove('hidden');
    if (msg.includes('automação em execução')) {
      el.btnReset().classList.remove('hidden');
    } else {
      el.btnReset().classList.add('hidden');
    }
    setTimeout(() => el.errorToast().classList.add('hidden'), 12000);
  }

  function closeError() {
    el.errorToast().classList.add('hidden');
    el.btnReset().classList.add('hidden');
  }

  async function forceReset() {
    await fetch('/api/reset', { method: 'POST' });
    _running = false;
    closeError();
    resetRunButton();
    setStatus('idle', 'Pronto');
    appendLog('🔄 Estado resetado — pode iniciar uma nova automação.', 'log-warning');
  }

  // ── Utils ────────────────────────────────────────────────
  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ── Inicializa ao carregar ───────────────────────────────
  document.addEventListener('DOMContentLoaded', init);

  // ── API pública ──────────────────────────────────────────
  return {
    toggleClient,
    selectAll,
    clearAll,
    run,
    cancel,
    clearTerminal,
    closeResults,
    closeError,
    forceReset,
    toggleSheetIdx,
    selectAllSheets,
    clearAllSheets,
  };

})();
