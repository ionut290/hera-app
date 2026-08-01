(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  const M = window.HeraPreventiviModels;
  if (!PV || !M) return;

  const STORAGE_KEY = 'hera_preventivi_clienti_v1';
  const COLLECTION = 'preventiviClienti';
  const text = (value) => String(value ?? '').trim();
  const normalize = (value) => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const escape = (value) => PV.escapeHtml ? PV.escapeHtml(value) : text(value).replace(/[&<>"']/g, (char) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const now = () => new Date().toISOString();

  const columns = [
    ['codiceCliente', 'CODICE CLIENTE'],
    ['ragioneSociale', 'RAGIONE SOCIALE'],
    ['spettabile', 'SPETT.LE'],
    ['referente', 'C.A. / RICHIEDENTE'],
    ['email', 'EMAIL'],
    ['telefono', 'TELEFONO'],
    ['indirizzo', 'INDIRIZZO'],
    ['cap', 'CAP'],
    ['comune', 'COMUNE'],
    ['provincia', 'PROVINCIA'],
    ['partitaIva', 'PARTITA IVA'],
    ['codiceFiscale', 'CODICE FISCALE'],
    ['numeroContratto', 'NUMERO CONTRATTO'],
    ['dataInizioContratto', 'DATA INIZIO CONTRATTO'],
    ['dataScadenzaContratto', 'SCADENZA CONTRATTO'],
    ['pec', 'PEC'],
    ['codiceSdi', 'CODICE SDI'],
    ['note', 'NOTE']
  ];

  const state = {
    clients: [],
    loaded: false,
    syncing: false,
    viewInstalled: false,
    remoteLoadedAt: 0
  };
  const REMOTE_TTL_MS = 5 * 60 * 1000;

  function loadLocal() {
    try {
      const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
      state.clients = Array.isArray(parsed) ? parsed : [];
    } catch (_) {
      state.clients = [];
    }
  }

  function saveLocal() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state.clients)); } catch (_) { /* non bloccante */ }
  }

  function clientId(raw = {}) {
    const direct = text(raw.id || raw.uid || raw.codiceCliente);
    if (direct) return direct;
    const base = normalize(raw.ragioneSociale || raw.spettabile || raw.email || raw.partitaIva);
    return base ? `cliente-${base.replace(/\s+/g, '-')}` : `cliente-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
  }

  function normalizeClient(raw = {}) {
    const aliases = {
      codiceCliente: ['codiceCliente','CODICE CLIENTE','codice cliente','codice'],
      ragioneSociale: ['ragioneSociale','RAGIONE SOCIALE','ragione sociale','cliente','committente'],
      spettabile: ['spettabile','SPETT.LE','Spett.le','spett le'],
      referente: ['referente','C.A. / RICHIEDENTE','C.A.','RICHIEDENTE','richiedente','contatto'],
      email: ['email','EMAIL','e-mail'], telefono: ['telefono','TELEFONO','tel'],
      indirizzo: ['indirizzo','INDIRIZZO','via'], cap: ['cap','CAP'], comune: ['comune','COMUNE'], provincia: ['provincia','PROVINCIA'],
      partitaIva: ['partitaIva','PARTITA IVA','p iva','p.iva'], codiceFiscale: ['codiceFiscale','CODICE FISCALE','cf'],
      numeroContratto: ['numeroContratto','NUMERO CONTRATTO','contratto'],
      dataInizioContratto: ['dataInizioContratto','DATA INIZIO CONTRATTO','inizio contratto'],
      dataScadenzaContratto: ['dataScadenzaContratto','SCADENZA CONTRATTO','data scadenza contratto','scadenza'],
      pec: ['pec','PEC'], codiceSdi: ['codiceSdi','CODICE SDI','sdi'], note: ['note','NOTE']
    };
    const pick = (keys) => {
      for (const key of keys) if (raw[key] !== undefined && text(raw[key])) return text(raw[key]);
      return '';
    };
    const client = { id: clientId(raw), updatedAt: raw.updatedAt || now(), syncPending: raw.syncPending === true };
    Object.entries(aliases).forEach(([key, keys]) => { client[key] = pick(keys); });
    if (!client.spettabile) client.spettabile = client.ragioneSociale;
    return client;
  }

  function mergeClients(rows, { markPending = false } = {}) {
    const map = new Map(state.clients.map((item) => [item.id, item]));
    rows.map(normalizeClient).forEach((incoming) => {
      const match = [...map.values()].find((item) =>
        (incoming.codiceCliente && normalize(item.codiceCliente) === normalize(incoming.codiceCliente)) ||
        (incoming.partitaIva && normalize(item.partitaIva) === normalize(incoming.partitaIva)) ||
        (incoming.ragioneSociale && normalize(item.ragioneSociale) === normalize(incoming.ragioneSociale))
      );
      if (match?.syncPending && !markPending) return;
      const values = Object.fromEntries(Object.entries(incoming).filter(([key, value]) => key === 'syncPending' || text(value)));
      if (match) Object.assign(match, values, markPending ? { updatedAt: now(), syncPending: true } : {});
      else map.set(incoming.id, { ...incoming, syncPending: markPending || incoming.syncPending === true });
    });
    state.clients = [...map.values()].sort((a, b) => (a.ragioneSociale || a.spettabile).localeCompare(b.ragioneSociale || b.spettabile, 'it'));
    saveLocal();
  }

  async function syncRemote({ force = false } = {}) {
    const db = window.firebase?.firestore?.();
    const user = window.firebase?.auth?.()?.currentUser;
    if (!db || !user || state.syncing) return;
    state.syncing = true;
    try {
      if (force || !state.remoteLoadedAt || Date.now() - state.remoteLoadedAt >= REMOTE_TTL_MS) {
        const snapshot = await db.collection(COLLECTION).get();
        mergeClients(snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data(), syncPending: false })));
        state.remoteLoadedAt = Date.now();
      }
      const pending = state.clients.filter((client) => client.syncPending);
      if (pending.length) {
        const batch = db.batch();
        pending.forEach((client) => {
          const payload = { ...client, syncPending: false };
          batch.set(db.collection(COLLECTION).doc(client.id), payload, { merge: true });
        });
        await batch.commit();
        pending.forEach((client) => { client.syncPending = false; });
        saveLocal();
      }
    } catch (error) {
      console.warn('Clienti preventivi: sincronizzazione non riuscita.', error);
    } finally {
      state.syncing = false;
      renderClientsView();
    }
  }

  function ensureTab() {
    const page = PV.page?.() || document.getElementById('preventivi-page');
    const nav = page?.querySelector('.pv-nav');
    if (!nav || nav.querySelector('[data-pv-view="clients"]')) return;
    nav.insertAdjacentHTML('beforeend', '<button type="button" class="pv-tab" data-pv-view="clients">Clienti</button>');
  }

  function downloadMatrix() {
    if (!window.XLSX) return PV.setFeedback?.('Motore Excel non disponibile.', 'error');
    const data = [Object.fromEntries(columns.map(([, label]) => [label, '']))];
    const sheet = XLSX.utils.json_to_sheet(data, { header: columns.map(([, label]) => label) });
    sheet['!cols'] = columns.map(([, label]) => ({ wch: Math.max(16, label.length + 3) }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'CLIENTI');
    XLSX.writeFile(workbook, 'Matrice_Clienti_Hera_App.xlsx');
  }

  function importMatrix(file) {
    if (!file || !window.XLSX) return;
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: false });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const valid = rows.filter((row) => text(row['RAGIONE SOCIALE'] || row['SPETT.LE'] || row['PARTITA IVA']));
        mergeClients(valid, { markPending: true });
        await syncRemote({ force: true });
        PV.setFeedback?.(`Importati o aggiornati ${valid.length} clienti.`, 'success');
        renderClientsView();
      } catch (error) {
        console.error(error);
        PV.setFeedback?.('Errore durante importazione matrice clienti.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function renderClientsView() {
    const page = PV.page?.() || document.getElementById('preventivi-page');
    if (!page || PV.state.view !== 'clients') return;
    const current = page.querySelector('.pv-view, [data-pv-current-view]') || page.querySelector('.pv-content') || page;
    let host = page.querySelector('[data-pv-clients-view]');
    if (!host) {
      host = document.createElement('section');
      host.dataset.pvClientsView = '1';
      host.className = 'pv-view';
      current.replaceChildren(host);
    }
    host.innerHTML = `
      <div class="pv-section-head"><div><h2>Clienti</h2><p class="pv-muted">Scarica la matrice, compilala e ricaricala per aggiungere o aggiornare i clienti.</p></div></div>
      <div class="pv-form-card"><div class="pv-form-actions">
        <button type="button" class="pv-btn pv-btn-primary" data-clients-download>Scarica matrice clienti</button>
        <label class="pv-btn pv-btn-light">Importa matrice<input type="file" accept=".xlsx,.xls,.ods" data-clients-import hidden></label>
      </div></div>
      <div class="pv-form-card"><label class="pv-label"><span>Cerca cliente</span><input type="search" data-clients-search placeholder="Ragione sociale, codice, P.IVA o referente"></label></div>
      <div data-clients-list>${clientCards('')}</div>`;
  }

  function clientCards(query) {
    const q = normalize(query);
    const rows = state.clients.filter((client) => !q || normalize(Object.values(client).join(' ')).includes(q));
    if (!rows.length) return '<p class="pv-muted">Nessun cliente presente.</p>';
    return rows.map((client) => `
      <article class="pv-card"><h3>${escape(client.ragioneSociale || client.spettabile || 'Cliente')}</h3>
      <p>${escape(client.codiceCliente || '—')} • P.IVA ${escape(client.partitaIva || '—')}</p>
      <p><b>Spett.le:</b> ${escape(client.spettabile || client.ragioneSociale || '—')}<br><b>C.A.:</b> ${escape(client.referente || '—')}<br><b>Contratto:</b> ${escape(client.numeroContratto || '—')} ${client.dataScadenzaContratto ? `• scadenza ${escape(client.dataScadenzaContratto)}` : ''}</p>
      </article>`).join('');
  }

  function modelNeeds(form, names) {
    const model = M.selectedModel?.(form);
    const fields = model?.fields || [];
    return fields.some((field) => {
      const key = normalize(`${field.key || ''} ${field.label || ''} ${field.source || ''}`);
      return names.some((name) => key.includes(normalize(name)));
    });
  }

  function injectClientSelector(form) {
    if (!form || form.querySelector('[data-doc-client-select]')) return;
    const firstCard = form.querySelector('.pv-form-card');
    if (!firstCard) return;
    firstCard.insertAdjacentHTML('afterbegin', `
      <label class="pv-label"><span>Cliente</span><select data-doc-client-select name="clientRegistryId"><option value="">Seleziona cliente</option>${state.clients.map((client) => `<option value="${escape(client.id)}">${escape(client.ragioneSociale || client.spettabile || client.id)}</option>`).join('')}</select></label>`);
  }

  function setField(form, names, value, onlyIfEmpty = false) {
    for (const name of names) {
      const field = form.elements?.namedItem(name) || form.querySelector(`[data-pvm-field="${name}"]`);
      if (!field || !('value' in field)) continue;
      if (!onlyIfEmpty || !text(field.value)) field.value = value || '';
    }
  }

  function applyClient(form, client) {
    if (!form || !client) return;
    setField(form, ['clientName','cliente','committente'], client.ragioneSociale || client.spettabile);
    setField(form, ['clientTaxCode','partita_iva','partitaIva'], client.partitaIva || client.codiceFiscale);
    setField(form, ['clientAddress','indirizzo_cliente'], [client.indirizzo, client.cap, client.comune, client.provincia].filter(Boolean).join(', '));

    if (modelNeeds(form, ['spettabile','spett.le','committente'])) setField(form, ['spettabile','cliente','committente'], client.spettabile || client.ragioneSociale);
    if (modelNeeds(form, ['c.a.','richiedente','referente'])) setField(form, ['requester','richiedente','referente','c_a'], client.referente);
    if (modelNeeds(form, ['numero contratto','contratto'])) setField(form, ['contractNumber','numero_contratto','numeroContratto'], client.numeroContratto);
    if (modelNeeds(form, ['scadenza contratto','data scadenza'])) setField(form, ['contractExpiry','data_scadenza_contratto','dataScadenzaContratto'], client.dataScadenzaContratto);
    if (modelNeeds(form, ['email'])) setField(form, ['email_cliente','clientEmail'], client.email);
    if (modelNeeds(form, ['telefono'])) setField(form, ['telefono_cliente','clientPhone'], client.telefono);
  }

  function decorateForms() {
    const page = PV.page?.() || document.getElementById('preventivi-page');
    page?.querySelectorAll('[data-pv-quote-form], [data-cons-form]').forEach(injectClientSelector);
  }

  document.addEventListener('click', (event) => {
    if (event.target.closest('[data-clients-download]')) downloadMatrix();
  }, true);

  document.addEventListener('change', (event) => {
    const importInput = event.target.closest('[data-clients-import]');
    if (importInput) { importMatrix(importInput.files?.[0]); importInput.value = ''; return; }
    const select = event.target.closest('[data-doc-client-select]');
    if (select) applyClient(select.closest('form'), state.clients.find((client) => client.id === select.value));
  }, true);

  document.addEventListener('input', (event) => {
    const search = event.target.closest('[data-clients-search]');
    if (search) {
      const list = document.querySelector('[data-clients-list]');
      if (list) list.innerHTML = clientCards(search.value);
    }
  }, true);

  document.addEventListener('click', (event) => {
    const tab = event.target.closest('[data-pv-view="clients"]');
    if (!tab) return;
    event.preventDefault();
    PV.state.view = 'clients';
    renderClientsView();
    void syncRemote();
    document.querySelectorAll('[data-pv-view]').forEach((button) => button.classList.toggle('active', button.dataset.pvView === 'clients'));
  }, true);

  const originalRender = PV.renderCurrentView?.bind(PV);
  if (originalRender) PV.renderCurrentView = (...args) => {
    if (PV.state.view === 'clients') { renderClientsView(); return; }
    const result = originalRender(...args);
    queueMicrotask(() => { ensureTab(); decorateForms(); });
    return result;
  };

  const observer = new MutationObserver(() => { ensureTab(); decorateForms(); });
  const start = () => {
    loadLocal();
    ensureTab();
    decorateForms();
    observer.observe(document.body, { childList: true, subtree: true });
    state.loaded = true;
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();

  window.HeraPreventiviClienti = { state, normalizeClient, mergeClients, applyClient, downloadMatrix };
})();
