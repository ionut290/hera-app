(() => {
  'use strict';
  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');

  PV.renderPriceListOverview = () => {
    const content = PV.content();
    if (!content) return;
    const query = PV.normalizeText(PV.state.priceListSearch);
    const lists = [...PV.state.priceLists]
      .filter((list) => !query || PV.normalizeText([list.name, list.reference, list.description].join(' ')).includes(query))
      .sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'it'));

    content.innerHTML = `
      <div class="pv-shell">
        <div class="pv-toolbar">
          <div><h2>Prezziari</h2><p class="pv-muted">Inserisci manualmente le voci oppure importale da Excel/CSV.</p></div>
          <button type="button" class="pv-btn pv-btn-primary" data-pv-action="new-price-list">+ Nuovo prezziario</button>
        </div>
        <div class="pv-toolbar"><input class="pv-search" type="search" data-pv-price-list-search value="${PV.escapeHtml(PV.state.priceListSearch)}" placeholder="Cerca nome, riferimento o descrizione…"></div>
        <div class="pv-grid">
          ${lists.length ? lists.map(PV.renderPriceListCard).join('') : '<div class="pv-empty"><strong>Nessun prezziario presente</strong>Crea un prezziario e inserisci le voci con codice, descrizione, unità e prezzo.</div>'}
        </div>
        <p class="pv-feedback" data-pv-feedback role="status"></p>
      </div>`;
  };

  PV.renderPriceListCard = (list) => {
    const usedBy = PV.state.quotes.filter((quote) => (quote.priceListIds || []).includes(list.id)).length;
    return `
      <article class="pv-card">
        <div><span class="pv-chip">${PV.escapeHtml(list.reference || 'Prezziario')}</span><h3>${PV.escapeHtml(list.name)}</h3><p class="pv-muted">${PV.escapeHtml(list.description || 'Nessuna descrizione')}</p></div>
        <div class="pv-card-meta">
          <div>Voci<strong>${(list.items || []).length}</strong></div>
          <div>Preventivi collegati<strong>${usedBy}</strong></div>
          <div>Aggiornato<strong>${PV.formatDate(String(list.updatedAt || '').slice(0, 10))}</strong></div>
          <div>Sincronizzazione<strong>${list.syncPending ? 'Da sincronizzare' : 'Salvato'}</strong></div>
        </div>
        <div class="pv-card-actions">
          <button type="button" class="pv-btn pv-btn-secondary pv-btn-small" data-pv-action="edit-price-list" data-id="${PV.escapeHtml(list.id)}">Apri</button>
          <button type="button" class="pv-btn pv-btn-danger pv-btn-small" data-pv-action="delete-price-list" data-id="${PV.escapeHtml(list.id)}">Elimina</button>
        </div>
      </article>`;
  };

  PV.newPriceItem = () => ({ id: PV.uid('item'), code: '', description: '', unit: '', unitPrice: 0 });

  PV.renderPriceListEditor = (existingList) => {
    const content = PV.content();
    if (!content) return;
    const list = existingList ? JSON.parse(JSON.stringify(existingList)) : { id: '', name: '', reference: '', description: '', items: [] };
    content.innerHTML = `
      <div class="pv-shell">
        <div class="pv-section-head">
          <div><h2>${existingList ? PV.escapeHtml(list.name) : 'Nuovo prezziario'}</h2><p class="pv-muted">Ogni voce deve avere codice, lavorazione, unità di misura e prezzo unitario.</p></div>
          <button type="button" class="pv-btn pv-btn-light" data-pv-action="cancel-editor">← Elenco prezziari</button>
        </div>
        <form data-pv-price-list-form novalidate>
          <section class="pv-form-card">
            <div class="pv-form-grid">
              <label class="pv-label"><span class="pv-required">Nome prezziario</span><input name="name" value="${PV.escapeHtml(list.name)}" placeholder="Es. Prezziario Regione Emilia-Romagna 2026" required></label>
              <label class="pv-label"><span>Edizione / Riferimento</span><input name="reference" value="${PV.escapeHtml(list.reference)}" placeholder="Es. Edizione luglio 2026"></label>
              <label class="pv-label pv-span-2"><span>Descrizione</span><textarea name="description">${PV.escapeHtml(list.description)}</textarea></label>
            </div>
          </section>
          <section class="pv-form-card">
            <div class="pv-section-head">
              <div><h3>Voci del prezziario</h3><p class="pv-muted">Puoi aggiungerle a mano oppure importarle da Excel/CSV.</p></div>
              <button type="button" class="pv-btn pv-btn-primary" data-pv-action="add-price-item">+ Aggiungi voce</button>
            </div>
            <div class="pv-file-box">
              <input type="file" accept=".xlsx,.xls,.csv" data-pv-price-import>
              <span class="pv-muted">Colonne riconosciute: Codice, Lavorazione/Descrizione, U.M., Prezzo.</span>
            </div>
            <div class="pv-items-head" aria-hidden="true"><span>Codice</span><span>Lavorazione / Descrizione</span><span>U.M.</span><span>Prezzo unitario</span><span></span></div>
            <div data-pv-price-items></div>
            <p class="pv-feedback" data-pv-feedback role="status"></p>
          </section>
          <div class="pv-form-actions">
            <button type="button" class="pv-btn pv-btn-light" data-pv-action="cancel-editor">Annulla</button>
            <button type="submit" class="pv-btn pv-btn-primary">Salva prezziario</button>
          </div>
        </form>
      </div>`;
    const items = (list.items || []).length ? list.items : [PV.newPriceItem()];
    items.forEach(PV.appendPriceItem);
  };

  PV.appendPriceItem = (item = PV.newPriceItem()) => {
    const container = PV.page()?.querySelector('[data-pv-price-items]');
    if (!container) return;
    const row = document.createElement('div');
    row.className = 'pv-item-row';
    row.dataset.priceItemId = item.id || PV.uid('item');
    row.innerHTML = `
      <label class="pv-label"><span class="pv-required">Codice</span><input data-pv-item-code value="${PV.escapeHtml(item.code)}" required></label>
      <label class="pv-label"><span class="pv-required">Lavorazione / Descrizione</span><input data-pv-item-description value="${PV.escapeHtml(item.description)}" required></label>
      <label class="pv-label"><span class="pv-required">U.M.</span><input data-pv-item-unit value="${PV.escapeHtml(item.unit)}" placeholder="mq, ml, ora…" required></label>
      <label class="pv-label"><span class="pv-required">Prezzo €</span><input type="number" min="0" step="0.01" data-pv-item-price value="${PV.escapeHtml(item.unitPrice)}" required></label>
      <button type="button" class="pv-icon-btn" data-pv-action="remove-price-item" aria-label="Rimuovi voce">✕</button>`;
    container.appendChild(row);
  };

  PV.collectPriceItems = () => [...(PV.page()?.querySelectorAll('[data-price-item-id]') || [])].map((row) => ({
    id: row.dataset.priceItemId,
    code: row.querySelector('[data-pv-item-code]')?.value.trim() || '',
    description: row.querySelector('[data-pv-item-description]')?.value.trim() || '',
    unit: row.querySelector('[data-pv-item-unit]')?.value.trim() || '',
    unitPrice: PV.parseNumber(row.querySelector('[data-pv-item-price]')?.value)
  }));

  PV.validatePriceItems = (items) => {
    if (!items.length) return 'Inserisci almeno una voce nel prezziario.';
    const seenCodes = new Set();
    for (const item of items) {
      if (!item.code || !item.description || !item.unit) return 'Compila codice, lavorazione e unità di misura per ogni voce.';
      if (item.unitPrice < 0) return 'Il prezzo unitario non può essere negativo.';
      const code = PV.normalizeText(item.code);
      if (seenCodes.has(code)) return `Il codice ${item.code} è presente più di una volta.`;
      seenCodes.add(code);
    }
    return '';
  };

  PV.savePriceList = async (form) => {
    const data = new FormData(form);
    const name = String(data.get('name') || '').trim();
    if (!name) { PV.setFeedback('Inserisci il nome del prezziario.', 'error'); form.elements.namedItem('name')?.focus(); return; }
    const items = PV.collectPriceItems();
    const itemError = PV.validatePriceItems(items);
    if (itemError) { PV.setFeedback(itemError, 'error'); return; }

    const existing = PV.state.editingPriceListId !== 'new' ? PV.getPriceList(PV.state.editingPriceListId) : null;
    const list = {
      id: existing?.id || PV.uid('price-list'),
      name,
      reference: String(data.get('reference') || '').trim(),
      description: String(data.get('description') || '').trim(),
      items,
      createdAt: existing?.createdAt || PV.nowIso(),
      createdBy: existing?.createdBy || PV.currentUser(),
      updatedAt: PV.nowIso(),
      updatedBy: PV.currentUser(),
      syncPending: true,
      version: PV.version
    };
    if (PV.state.priceLists.some((item) => item.id !== list.id && PV.normalizeText(item.name) === PV.normalizeText(list.name))) {
      PV.setFeedback('Esiste già un prezziario con questo nome.', 'error'); return;
    }
    const index = PV.state.priceLists.findIndex((item) => item.id === list.id);
    if (index >= 0) PV.state.priceLists[index] = list;
    else PV.state.priceLists.push(list);
    PV.persistLocal();
    const saved = await PV.saveRemote(PV.collections.priceLists, { ...list, syncPending: false });
    if (saved) list.syncPending = false;
    PV.persistLocal();
    PV.state.editingPriceListId = '';
    PV.renderPriceListOverview();
    PV.setFeedback(saved || !PV.state.firestore ? 'Prezziario salvato correttamente.' : 'Prezziario salvato sul dispositivo; sincronizzazione online non disponibile.', saved || !PV.state.firestore ? 'success' : 'warning');
  };

  PV.findHeaderIndex = (headers, aliases) => {
    const normalized = headers.map(PV.normalizeText);
    return normalized.findIndex((header) => aliases.some((alias) => header === alias || header.includes(alias)));
  };

  PV.importPriceFile = async (file) => {
    if (!file) return;
    if (!window.XLSX) { PV.setFeedback('Libreria Excel non disponibile. Ricarica l’app e riprova.', 'error'); return; }
    try {
      const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '', raw: false });
      if (!rows.length) throw new Error('Il file è vuoto.');
      const detected = rows.findIndex((row) => {
        const text = PV.normalizeText(row.join(' '));
        return text.includes('cod') && (text.includes('prezz') || text.includes('import'));
      });
      const headerRow = detected >= 0 ? detected : 0;
      const headers = rows[headerRow].map((cell) => String(cell || ''));
      const codeIndex = PV.findHeaderIndex(headers, ['codice', 'cod.', 'code', 'articolo']);
      const descriptionIndex = PV.findHeaderIndex(headers, ['lavorazione', 'descrizione', 'voce', 'prestazione']);
      const unitIndex = PV.findHeaderIndex(headers, ['u.m.', 'um', 'unita di misura', 'unità di misura', 'unita']);
      const priceIndex = PV.findHeaderIndex(headers, ['prezzo unitario', 'prezzo', 'importo', 'tariffa']);
      if ([codeIndex, descriptionIndex, unitIndex, priceIndex].some((index) => index < 0)) {
        throw new Error('Non riconosco le colonne. Servono Codice, Lavorazione/Descrizione, U.M. e Prezzo.');
      }
      const imported = rows.slice(headerRow + 1).map((row) => ({
        id: PV.uid('item'),
        code: String(row[codeIndex] || '').trim(),
        description: String(row[descriptionIndex] || '').trim(),
        unit: String(row[unitIndex] || '').trim(),
        unitPrice: PV.parseNumber(row[priceIndex])
      })).filter((item) => item.code && item.description && item.unit);
      if (!imported.length) throw new Error('Non ho trovato righe valide da importare.');
      const container = PV.page()?.querySelector('[data-pv-price-items]');
      const existingRows = [...(container?.querySelectorAll('[data-price-item-id]') || [])];
      if (existingRows.length === 1 && !existingRows[0].querySelector('[data-pv-item-code]')?.value && !existingRows[0].querySelector('[data-pv-item-description]')?.value) existingRows[0].remove();
      imported.forEach(PV.appendPriceItem);
      PV.setFeedback(`Importate ${imported.length} voci dal file. Controlla i dati e salva il prezziario.`, 'success');
    } catch (error) {
      console.warn('Preventivi: importazione prezziario non riuscita.', error);
      PV.setFeedback(error?.message || 'Importazione non riuscita.', 'error');
    }
  };

  PV.deletePriceList = async (id) => {
    const list = PV.getPriceList(id);
    if (!list) return;
    const linked = PV.state.quotes.filter((quote) => (quote.priceListIds || []).includes(id));
    if (linked.length) {
      window.alert(`Il prezziario “${list.name}” è usato da ${linked.length} preventivi e non può essere eliminato. Conservandolo si evita di perdere i riferimenti delle lavorazioni.`);
      return;
    }
    if (!window.confirm(`Eliminare il prezziario “${list.name}”?`)) return;
    PV.state.priceLists = PV.state.priceLists.filter((item) => item.id !== id);
    PV.state.deletions.priceLists[id] = PV.nowIso();
    PV.persistLocal();
    if (await PV.deleteRemote(PV.collections.priceLists, id)) delete PV.state.deletions.priceLists[id];
    PV.persistLocal();
    PV.renderPriceListOverview();
    PV.setFeedback('Prezziario eliminato.', 'success');
  };
})();
