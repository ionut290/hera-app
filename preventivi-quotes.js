(() => {
  'use strict';
  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');

  PV.renderQuoteOverview = () => {
    const content = PV.content();
    if (!content) return;
    const query = PV.normalizeText(PV.state.quoteSearch);
    const quotes = [...PV.state.quotes]
      .filter((quote) => !query || PV.normalizeText([quote.number, quote.clientName, quote.subject, quote.workLocation].join(' ')).includes(query))
      .sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')));
    content.innerHTML = `
      <div class="pv-shell">
        <div class="pv-toolbar">
          <div><h2>Preventivi</h2><p class="pv-muted">Crea preventivi con più lavorazioni e prezzi automatici.</p></div>
          <button type="button" class="pv-btn pv-btn-primary" data-pv-action="new-quote">+ Nuovo preventivo</button>
        </div>
        <div class="pv-toolbar"><input class="pv-search" type="search" data-pv-quote-search value="${PV.escapeHtml(PV.state.quoteSearch)}" placeholder="Cerca numero, cliente, oggetto o luogo…"></div>
        <div class="pv-grid">
          ${quotes.length ? quotes.map(PV.renderQuoteCard).join('') : '<div class="pv-empty"><strong>Nessun preventivo presente</strong>Crea il primo preventivo dopo aver inserito almeno un prezziario.</div>'}
        </div>
        <p class="pv-feedback" data-pv-feedback role="status"></p>
      </div>`;
  };

  PV.renderQuoteCard = (quote) => {
    const totals = PV.quoteTotals(quote);
    const names = (quote.priceListIds || []).map((id) => PV.getPriceList(id)?.name).filter(Boolean);
    return `
      <article class="pv-card" data-pv-quote-card="${PV.escapeHtml(quote.id)}">
        <div><span class="pv-chip">${PV.escapeHtml(quote.number || 'Senza numero')}</span><h3>${PV.escapeHtml(quote.clientName || 'Cliente non indicato')}</h3><p class="pv-muted">${PV.escapeHtml(quote.subject || 'Nessun oggetto')}</p></div>
        <div class="pv-card-meta">
          <div>Data<strong>${PV.formatDate(quote.date)}</strong></div>
          <div>Lavorazioni<strong>${(quote.lines || []).length}</strong></div>
          <div>Prezziari<strong>${PV.escapeHtml(names.join(', ') || '—')}</strong></div>
          <div>Luogo<strong>${PV.escapeHtml(quote.workLocation || '—')}</strong></div>
        </div>
        <div class="pv-card-total">${PV.formatMoney(totals.total)}</div>
        <div class="pv-card-actions">
          <button type="button" class="pv-btn pv-btn-secondary pv-btn-small" data-pv-action="edit-quote" data-id="${PV.escapeHtml(quote.id)}">Apri</button>
          <button type="button" class="pv-btn pv-btn-light pv-btn-small" data-pv-action="print-quote" data-id="${PV.escapeHtml(quote.id)}">Stampa</button>
          <button type="button" class="pv-btn pv-btn-light pv-btn-small" data-pv-action="duplicate-quote" data-id="${PV.escapeHtml(quote.id)}">Duplica</button>
          <button type="button" class="pv-btn pv-btn-danger pv-btn-small" data-pv-action="delete-quote" data-id="${PV.escapeHtml(quote.id)}">Elimina</button>
        </div>
      </article>`;
  };

  PV.newQuote = () => ({
    id: '',
    number: PV.nextQuoteNumber(),
    date: PV.todayIso(),
    issuerName: PV.state.settings.lastIssuerName || '',
    clientName: '',
    clientTaxCode: '',
    workLocation: '',
    subject: '',
    notes: '',
    validityDays: 30,
    vatRate: 22,
    priceListIds: [],
    lines: []
  });
  PV.newQuoteLine = () => ({ id: PV.uid('line'), priceListId: '', priceItemId: '', code: '', description: '', unit: '', quantity: 1, unitPrice: 0 });

  PV.renderQuoteEditor = (existingQuote) => {
    const content = PV.content();
    if (!content) return;
    const quote = existingQuote ? JSON.parse(JSON.stringify(existingQuote)) : PV.newQuote();
    const choices = PV.state.priceLists.map((list) => `
      <label class="pv-check-card">
        <input type="checkbox" name="priceListIds" value="${PV.escapeHtml(list.id)}" ${(quote.priceListIds || []).includes(list.id) ? 'checked' : ''}>
        <span><strong>${PV.escapeHtml(list.name)}</strong><small>${PV.escapeHtml(list.reference || '')} · ${(list.items || []).length} voci</small></span>
      </label>`).join('');
    content.innerHTML = `
      <div class="pv-shell">
        <div class="pv-section-head">
          <div><h2>${existingQuote ? `Preventivo ${PV.escapeHtml(quote.number)}` : 'Nuovo preventivo'}</h2><p class="pv-muted">I campi con * sono obbligatori.</p></div>
          <button type="button" class="pv-btn pv-btn-light" data-pv-action="cancel-editor">← Elenco preventivi</button>
        </div>
        <form data-pv-quote-form novalidate>
          <section class="pv-form-card">
            <div class="pv-section-head"><h3>Intestazione</h3></div>
            <div class="pv-form-grid">
              <label class="pv-label"><span class="pv-required">Numero preventivo</span><input name="number" value="${PV.escapeHtml(quote.number)}" required></label>
              <label class="pv-label"><span class="pv-required">Data</span><input type="date" name="date" value="${PV.escapeHtml(quote.date)}" required></label>
              <label class="pv-label pv-span-2"><span class="pv-required">Intestazione azienda / fornitore</span><input name="issuerName" value="${PV.escapeHtml(quote.issuerName)}" placeholder="Es. Avola Coop Manutenzione Verde" required></label>
              <label class="pv-label"><span class="pv-required">Cliente / Ragione sociale</span><input name="clientName" value="${PV.escapeHtml(quote.clientName)}" required></label>
              <label class="pv-label"><span>Partita IVA / Codice fiscale cliente</span><input name="clientTaxCode" value="${PV.escapeHtml(quote.clientTaxCode)}"></label>
              <label class="pv-label"><span class="pv-required">Cantiere / Luogo dei lavori</span><input name="workLocation" value="${PV.escapeHtml(quote.workLocation)}" required></label>
              <label class="pv-label"><span class="pv-required">Oggetto del preventivo</span><input name="subject" value="${PV.escapeHtml(quote.subject)}" required></label>
              <label class="pv-label"><span>Validità (giorni)</span><input type="number" min="0" step="1" name="validityDays" value="${PV.escapeHtml(quote.validityDays)}"></label>
              <label class="pv-label"><span>IVA %</span><input type="number" min="0" step="0.01" name="vatRate" value="${PV.escapeHtml(quote.vatRate)}" data-pv-vat></label>
              <label class="pv-label pv-span-2"><span>Note e condizioni</span><textarea name="notes">${PV.escapeHtml(quote.notes)}</textarea></label>
            </div>
          </section>
          <section class="pv-form-card">
            <div class="pv-section-head">
              <div><h3>Prezziari di riferimento *</h3><p class="pv-muted">Seleziona uno o più prezziari prima di aggiungere le lavorazioni.</p></div>
              <button type="button" class="pv-btn pv-btn-light pv-btn-small" data-pv-action="go-price-lists">Gestisci prezziari</button>
            </div>
            <div class="pv-price-list-choice" data-pv-price-list-choices>${choices || '<p class="pv-muted">Non sono ancora presenti prezziari.</p>'}</div>
          </section>
          <section class="pv-form-card">
            <div class="pv-section-head">
              <div><h3>Lavorazioni *</h3><p class="pv-muted">Scrivi la lavorazione, scegli la voce trovata e indica la quantità.</p></div>
              <button type="button" class="pv-btn pv-btn-primary" data-pv-action="add-quote-line">+ Aggiungi lavorazione</button>
            </div>
            <div class="pv-lines" data-pv-lines></div>
            <p class="pv-feedback" data-pv-feedback role="status"></p>
          </section>
          <aside class="pv-summary" data-pv-summary></aside>
          <div class="pv-form-actions">
            ${existingQuote ? '<button type="button" class="pv-btn pv-btn-light" data-pv-action="print-current-quote">Stampa</button>' : ''}
            <button type="button" class="pv-btn pv-btn-light" data-pv-action="cancel-editor">Annulla</button>
            <button type="submit" class="pv-btn pv-btn-primary">Salva preventivo</button>
          </div>
        </form>
      </div>`;
    ((quote.lines || []).length ? quote.lines : [PV.newQuoteLine()]).forEach(PV.appendQuoteLine);
    PV.updateQuoteTotals();
  };

  PV.selectedPriceListIds = () => [...(PV.page()?.querySelectorAll('input[name="priceListIds"]:checked') || [])].map((input) => input.value);
  PV.availablePriceItems = (selectedIds, filter = '') => {
    const query = PV.normalizeText(filter);
    const result = [];
    selectedIds.forEach((listId) => {
      const list = PV.getPriceList(listId);
      (list?.items || []).forEach((item) => {
        if (!query || PV.normalizeText([item.code, item.description, item.unit].join(' ')).includes(query)) result.push({ list, item });
      });
    });
    return result.sort((a, b) => String(a.item.description || '').localeCompare(String(b.item.description || ''), 'it'));
  };
  PV.buildPriceItemOptions = (selectedValue = '', filter = '') => {
    const selectedIds = PV.selectedPriceListIds();
    const placeholder = selectedIds.length ? 'Seleziona la voce trovata' : 'Seleziona prima un prezziario';
    return `<option value="">${placeholder}</option>${PV.availablePriceItems(selectedIds, filter).map(({ list, item }) => {
      const value = `${list.id}::${item.id}`;
      return `<option value="${PV.escapeHtml(value)}" ${value === selectedValue ? 'selected' : ''}>${PV.escapeHtml(`[${list.name}] ${item.code} — ${item.description} (${item.unit || 'u.m.'})`)}</option>`;
    }).join('')}`;
  };

  PV.appendQuoteLine = (line = PV.newQuoteLine()) => {
    const container = PV.page()?.querySelector('[data-pv-lines]');
    if (!container) return;
    const row = document.createElement('div');
    row.className = 'pv-line';
    row.dataset.lineId = line.id || PV.uid('line');
    const selectedValue = line.priceListId && line.priceItemId ? `${line.priceListId}::${line.priceItemId}` : '';
    row.innerHTML = `
      <label class="pv-label pv-line-wide"><span>Cerca lavorazione</span><input type="search" data-pv-line-filter value="${PV.escapeHtml(line.description || '')}" placeholder="Es. sfalcio, potatura, raccolta…"></label>
      <label class="pv-label pv-line-picker"><span class="pv-required">Voce del prezziario</span><select data-pv-line-item required>${PV.buildPriceItemOptions(selectedValue)}</select></label>
      <label class="pv-label"><span>Codice</span><input data-pv-line-code value="${PV.escapeHtml(line.code || '')}" readonly></label>
      <label class="pv-label"><span>U.M.</span><input data-pv-line-unit value="${PV.escapeHtml(line.unit || '')}" readonly></label>
      <label class="pv-label"><span class="pv-required">Quantità</span><input type="number" min="0.0001" step="0.0001" data-pv-line-quantity value="${PV.escapeHtml(line.quantity || 1)}" required></label>
      <label class="pv-label"><span>Prezzo unitario</span><input type="number" step="0.01" data-pv-line-price value="${PV.escapeHtml(line.unitPrice || 0)}" readonly></label>
      <div class="pv-line-remove"><button type="button" class="pv-icon-btn" data-pv-action="remove-quote-line" aria-label="Rimuovi lavorazione">✕</button></div>
      <div class="pv-line-total pv-line-wide" data-pv-line-total>${PV.formatMoney((Number(line.quantity) || 0) * (Number(line.unitPrice) || 0))}</div>`;
    container.appendChild(row);
  };

  PV.resolvePriceItem = (value) => {
    const [listId, itemId] = String(value || '').split('::');
    const list = PV.getPriceList(listId);
    const item = (list?.items || []).find((entry) => entry.id === itemId) || null;
    return item ? { list, item } : null;
  };
  PV.clearQuoteLine = (row, keepFilter = false) => {
    if (!row) return;
    if (!keepFilter) row.querySelector('[data-pv-line-filter]').value = '';
    row.querySelector('[data-pv-line-code]').value = '';
    row.querySelector('[data-pv-line-unit]').value = '';
    row.querySelector('[data-pv-line-price]').value = '0';
    PV.updateQuoteTotals();
  };
  PV.populateQuoteLine = (row, value) => {
    const resolved = PV.resolvePriceItem(value);
    if (!resolved) { PV.clearQuoteLine(row, true); return; }
    row.querySelector('[data-pv-line-filter]').value = resolved.item.description || '';
    row.querySelector('[data-pv-line-code]').value = resolved.item.code || '';
    row.querySelector('[data-pv-line-unit]').value = resolved.item.unit || '';
    row.querySelector('[data-pv-line-price]').value = String(Number(resolved.item.unitPrice) || 0);
    PV.updateQuoteTotals();
  };
  PV.refreshLineSelectors = () => {
    PV.page()?.querySelectorAll('[data-line-id]').forEach((row) => {
      const select = row.querySelector('[data-pv-line-item]');
      if (!select) return;
      const current = select.value;
      select.innerHTML = PV.buildPriceItemOptions(current, row.querySelector('[data-pv-line-filter]')?.value || '');
      if (current && !select.value) PV.clearQuoteLine(row, true);
    });
    PV.updateQuoteTotals();
  };
  PV.collectQuoteLines = () => [...(PV.page()?.querySelectorAll('[data-line-id]') || [])].map((row) => {
    const selected = PV.resolvePriceItem(row.querySelector('[data-pv-line-item]')?.value);
    return {
      id: row.dataset.lineId,
      priceListId: selected?.list?.id || '',
      priceItemId: selected?.item?.id || '',
      code: row.querySelector('[data-pv-line-code]')?.value.trim() || '',
      description: selected?.item?.description || row.querySelector('[data-pv-line-filter]')?.value.trim() || '',
      unit: row.querySelector('[data-pv-line-unit]')?.value.trim() || '',
      quantity: PV.parseNumber(row.querySelector('[data-pv-line-quantity]')?.value),
      unitPrice: PV.parseNumber(row.querySelector('[data-pv-line-price]')?.value)
    };
  });
  PV.updateQuoteTotals = () => {
    const summary = PV.page()?.querySelector('[data-pv-summary]');
    if (!summary) return;
    const totals = PV.quoteTotals({ lines: PV.collectQuoteLines(), vatRate: Math.max(0, PV.parseNumber(PV.page()?.querySelector('[data-pv-vat]')?.value)) });
    summary.innerHTML = `
      <div class="pv-summary-row"><span>Subtotale</span><strong>${PV.formatMoney(totals.subtotal)}</strong></div>
      <div class="pv-summary-row"><span>IVA ${String(totals.vatRate).replace('.', ',')}%</span><strong>${PV.formatMoney(totals.vatAmount)}</strong></div>
      <div class="pv-summary-row"><span>Totale preventivo</span><strong>${PV.formatMoney(totals.total)}</strong></div>`;
    PV.page()?.querySelectorAll('[data-line-id]').forEach((row) => {
      const total = row.querySelector('[data-pv-line-total]');
      if (total) total.textContent = PV.formatMoney(PV.parseNumber(row.querySelector('[data-pv-line-quantity]')?.value) * PV.parseNumber(row.querySelector('[data-pv-line-price]')?.value));
    });
  };

  PV.validateQuote = (form, lines, priceListIds) => {
    for (const name of ['number', 'date', 'issuerName', 'clientName', 'workLocation', 'subject']) {
      const input = form.elements.namedItem(name);
      if (!input?.value?.trim()) { input?.focus(); return 'Compila tutti i campi obbligatori dell’intestazione.'; }
    }
    if (!priceListIds.length) return 'Seleziona almeno un prezziario di riferimento.';
    if (!lines.length) return 'Inserisci almeno una lavorazione.';
    for (const line of lines) {
      if (!line.priceListId || !line.priceItemId || !line.code) return 'Scegli una voce valida del prezziario per ogni lavorazione.';
      if (!(line.quantity > 0)) return 'La quantità di ogni lavorazione deve essere maggiore di zero.';
      if (line.unitPrice < 0) return 'Il prezzo unitario non può essere negativo.';
    }
    return '';
  };

  PV.saveQuote = async (form) => {
    const data = new FormData(form);
    const priceListIds = PV.selectedPriceListIds();
    const lines = PV.collectQuoteLines();
    const validation = PV.validateQuote(form, lines, priceListIds);
    if (validation) { PV.setFeedback(validation, 'error'); return; }
    const existing = PV.state.editingQuoteId !== 'new' ? PV.getQuote(PV.state.editingQuoteId) : null;
    const quote = {
      id: existing?.id || PV.uid('quote'),
      number: String(data.get('number') || '').trim(),
      date: String(data.get('date') || ''),
      issuerName: String(data.get('issuerName') || '').trim(),
      clientName: String(data.get('clientName') || '').trim(),
      clientTaxCode: String(data.get('clientTaxCode') || '').trim(),
      workLocation: String(data.get('workLocation') || '').trim(),
      subject: String(data.get('subject') || '').trim(),
      notes: String(data.get('notes') || '').trim(),
      validityDays: Math.max(0, Math.round(PV.parseNumber(data.get('validityDays')))),
      vatRate: Math.max(0, PV.parseNumber(data.get('vatRate'))),
      priceListIds,
      lines,
      createdAt: existing?.createdAt || PV.nowIso(),
      createdBy: existing?.createdBy || PV.currentUser(),
      updatedAt: PV.nowIso(),
      updatedBy: PV.currentUser(),
      syncPending: true,
      version: PV.version
    };
    if (PV.state.quotes.some((item) => item.id !== quote.id && PV.normalizeText(item.number) === PV.normalizeText(quote.number))) {
      PV.setFeedback('Esiste già un preventivo con questo numero.', 'error'); form.elements.namedItem('number')?.focus(); return;
    }
    const index = PV.state.quotes.findIndex((item) => item.id === quote.id);
    if (index >= 0) PV.state.quotes[index] = quote;
    else PV.state.quotes.push(quote);
    PV.state.settings.lastIssuerName = quote.issuerName;
    PV.persistLocal();
    const saved = await PV.saveRemote(PV.collections.quotes, { ...quote, syncPending: false });
    if (saved) quote.syncPending = false;
    PV.persistLocal();
    PV.state.editingQuoteId = '';
    PV.renderQuoteOverview();
    PV.setFeedback(saved || !PV.state.firestore ? 'Preventivo salvato correttamente.' : 'Preventivo salvato sul dispositivo; sincronizzazione online non disponibile.', saved || !PV.state.firestore ? 'success' : 'warning');
  };

  PV.deleteQuote = async (id) => {
    const quote = PV.getQuote(id);
    if (!quote || !window.confirm(`Eliminare il preventivo ${quote.number || ''}?`)) return;
    PV.state.quotes = PV.state.quotes.filter((item) => item.id !== id);
    PV.state.deletions.quotes[id] = PV.nowIso();
    PV.persistLocal();
    if (await PV.deleteRemote(PV.collections.quotes, id)) delete PV.state.deletions.quotes[id];
    PV.persistLocal();
    PV.renderQuoteOverview();
    PV.setFeedback('Preventivo eliminato.', 'success');
  };
  PV.duplicateQuote = (id) => {
    const source = PV.getQuote(id);
    if (!source) return;
    const duplicate = {
      ...JSON.parse(JSON.stringify(source)), id: PV.uid('quote'), number: PV.nextQuoteNumber(), date: PV.todayIso(),
      createdAt: PV.nowIso(), updatedAt: PV.nowIso(), createdBy: PV.currentUser(), updatedBy: PV.currentUser(), syncPending: true
    };
    PV.state.quotes.push(duplicate);
    PV.persistLocal();
    PV.scheduleSync();
    PV.state.editingQuoteId = duplicate.id;
    PV.renderCurrentView();
    PV.setFeedback('Copia creata. Controlla i dati e salva.', 'success');
  };

  PV.printQuote = (quote) => {
    if (!quote) return;
    const totals = PV.quoteTotals(quote);
    const listNames = (quote.priceListIds || []).map((id) => PV.getPriceList(id)?.name).filter(Boolean).join(', ');
    const rows = (quote.lines || []).map((line, index) => `<tr><td>${index + 1}</td><td>${PV.escapeHtml(line.code)}</td><td>${PV.escapeHtml(line.description)}</td><td>${PV.escapeHtml(line.unit)}</td><td class="num">${PV.escapeHtml(String(line.quantity).replace('.', ','))}</td><td class="num">${PV.formatMoney(line.unitPrice)}</td><td class="num">${PV.formatMoney((Number(line.quantity) || 0) * (Number(line.unitPrice) || 0))}</td></tr>`).join('');
    const popup = window.open('', '_blank', 'width=1100,height=800');
    if (!popup) { PV.setFeedback('Il browser ha bloccato la finestra di stampa. Consenti i popup e riprova.', 'error'); return; }
    try { popup.opener = null; } catch (error) { /* Nessuna azione necessaria. */ }
    popup.document.open();
    popup.document.write(`<!DOCTYPE html><html lang="it"><head><meta charset="UTF-8"><title>${PV.escapeHtml(quote.number)}</title><style>
      body{font-family:Arial,sans-serif;color:#172033;margin:34px;line-height:1.35}h1,h2,p{margin:0}.head{display:flex;justify-content:space-between;gap:30px;border-bottom:3px solid #166534;padding-bottom:16px;margin-bottom:20px}.issuer{font-size:22px;font-weight:800}.doc-title{text-align:right}.doc-title h1{font-size:27px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin:18px 0}.box{border:1px solid #cbd5e1;border-radius:8px;padding:12px}.label{font-size:11px;text-transform:uppercase;color:#667085;font-weight:700;margin-bottom:3px}table{width:100%;border-collapse:collapse;margin-top:18px;font-size:12px}th,td{border:1px solid #cbd5e1;padding:8px;text-align:left;vertical-align:top}th{background:#eaf5ee}.num{text-align:right;white-space:nowrap}.totals{width:360px;margin:18px 0 0 auto}.totals div{display:flex;justify-content:space-between;padding:7px;border-bottom:1px solid #d0d5dd}.totals .grand{font-size:18px;font-weight:800;color:#166534}.notes{margin-top:22px;white-space:pre-wrap}.small{font-size:11px;color:#667085;margin-top:6px}@media print{body{margin:12mm}}</style></head><body>
      <div class="head"><div><div class="issuer">${PV.escapeHtml(quote.issuerName)}</div><p class="small">Preventivo economico</p></div><div class="doc-title"><h1>PREVENTIVO</h1><p><strong>N. ${PV.escapeHtml(quote.number)}</strong></p><p>Data ${PV.formatDate(quote.date)}</p></div></div>
      <div class="grid"><div class="box"><div class="label">Cliente</div><strong>${PV.escapeHtml(quote.clientName)}</strong><p>${PV.escapeHtml(quote.clientTaxCode || '')}</p></div><div class="box"><div class="label">Cantiere / Luogo</div><strong>${PV.escapeHtml(quote.workLocation)}</strong></div><div class="box"><div class="label">Oggetto</div><strong>${PV.escapeHtml(quote.subject)}</strong></div><div class="box"><div class="label">Prezziari di riferimento</div><strong>${PV.escapeHtml(listNames || '—')}</strong></div></div>
      <table><thead><tr><th>#</th><th>Codice</th><th>Lavorazione</th><th>U.M.</th><th>Quantità</th><th>Prezzo unitario</th><th>Totale</th></tr></thead><tbody>${rows}</tbody></table>
      <div class="totals"><div><span>Subtotale</span><strong>${PV.formatMoney(totals.subtotal)}</strong></div><div><span>IVA ${PV.escapeHtml(String(totals.vatRate).replace('.', ','))}%</span><strong>${PV.formatMoney(totals.vatAmount)}</strong></div><div class="grand"><span>TOTALE</span><strong>${PV.formatMoney(totals.total)}</strong></div></div>
      ${quote.validityDays ? `<p class="small">Validità del preventivo: ${PV.escapeHtml(quote.validityDays)} giorni.</p>` : ''}
      ${quote.notes ? `<div class="notes"><div class="label">Note e condizioni</div>${PV.escapeHtml(quote.notes)}</div>` : ''}
      <script>window.addEventListener('load',()=>window.print());<\/script></body></html>`);
    popup.document.close();
  };
})();
