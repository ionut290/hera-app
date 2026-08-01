(() => {
  'use strict';
  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');

  const KEY = 'hera_preventivi_consuntivi_v1';
  const collection = 'commesse/__preventivi__/consuntivi';
  const read = () => { try { const v = JSON.parse(localStorage.getItem(KEY) || '[]'); return Array.isArray(v) ? v : []; } catch (_) { return []; } };
  const write = () => { try { localStorage.setItem(KEY, JSON.stringify(PV.state.consuntivi || [])); } catch (e) { console.warn('Consuntivi: salvataggio locale non riuscito.', e); } };
  PV.state.consuntivi = read();
  PV.state.editingConsuntivoId = '';
  PV.state.consuntiviSearch = '';
  PV.collections.consuntivi = collection;

  const first = (obj, names) => {
    for (const name of names) {
      const value = obj?.[name];
      if (value !== undefined && value !== null && String(value).trim() !== '') return value;
    }
    return '';
  };
  const arrayCandidates = () => {
    const found = [];
    const add = (value) => { if (Array.isArray(value) && value.length && typeof value[0] === 'object') found.push(value); };
    ['commesse','projects','cantieri','impianti','plants','sites'].forEach((key) => add(window[key]));
    try {
      for (let i = 0; i < localStorage.length; i += 1) {
        const key = localStorage.key(i);
        if (!/(commess|impiant|cantier|plant|site)/i.test(key || '')) continue;
        const value = JSON.parse(localStorage.getItem(key));
        add(value); add(value?.commesse); add(value?.impianti); add(value?.data);
      }
    } catch (_) { /* dati non leggibili ignorati */ }
    return found;
  };
  const normalizeCommessa = (raw, index) => ({
    raw,
    id: String(first(raw, ['id','uid','commessaId','projectId','codice','codiceCommessa']) || `commessa-${index}`),
    name: String(first(raw, ['nome','name','denominazione','titolo','commessa']) || 'Commessa senza nome'),
    code: String(first(raw, ['codiceCommessa','codice','code','numeroCommessa']) || ''),
    client: String(first(raw, ['cliente','committente','clientName','ragioneSociale']) || ''),
    contract: String(first(raw, ['numeroContratto','contratto','contractNumber']) || ''),
    requester: String(first(raw, ['richiedente','referente','referenteCliente']) || '')
  });
  const normalizePlant = (raw, index) => ({
    raw,
    id: String(first(raw, ['id','uid','impiantoId','plantId','idSap','ID SAP']) || `impianto-${index}`),
    commessaId: String(first(raw, ['commessaId','projectId','idCommessa','codiceCommessa','commessa']) || ''),
    name: String(first(raw, ['denominazione','Denominazione Impianto','nome','name','impianto']) || 'Impianto senza nome'),
    sap: String(first(raw, ['idSap','ID SAP','sap','codiceSap']) || ''),
    address: String(first(raw, ['indirizzo','Descrizione via','descrizioneVia','via','address']) || ''),
    city: String(first(raw, ['comune','Comune','city']) || ''),
    type: String(first(raw, ['tipologia','Tipologia impianto','tipo','type']) || ''),
    area: String(first(raw, ['area','AREA','competenza','Area/Competenza']) || ''),
    coordinates: String(first(raw, ['coordinate','coordinates','Coordinate GPS(X)/GPS(Y)']) || '')
  });
  const registry = () => {
    const commesse = [], plants = [];
    arrayCandidates().forEach((array) => array.forEach((item, index) => {
      const looksPlant = first(item, ['idSap','ID SAP','Denominazione Impianto','tipologia','Tipologia impianto','impiantoId']);
      const looksCommessa = first(item, ['codiceCommessa','numeroContratto','committente','commessaId']) || (!looksPlant && first(item, ['nome','name','denominazione']));
      if (looksPlant) plants.push(normalizePlant(item, index));
      else if (looksCommessa) commesse.push(normalizeCommessa(item, index));
      const nested = first(item, ['impianti','plants','sites']);
      if (Array.isArray(nested)) nested.forEach((plant, pIndex) => plants.push({ ...normalizePlant(plant, pIndex), commessaId: String(first(plant, ['commessaId']) || first(item, ['id','uid','codiceCommessa','codice'])) }));
    }));
    return {
      commesse: [...new Map(commesse.map((x) => [x.id, x])).values()],
      plants: [...new Map(plants.map((x) => [x.id, x])).values()]
    };
  };

  const nextNumber = () => {
    const year = new Date().getFullYear();
    const max = Math.max(0, ...(PV.state.consuntivi || []).map((x) => Number(String(x.number || '').match(/(\d+)(?!.*\d)/)?.[1] || 0)));
    return `CONS-${year}-${String(max + 1).padStart(3, '0')}`;
  };
  const totals = (doc) => PV.quoteTotals({ lines: doc.lines || [], vatRate: doc.vatRate || 0 });
  const getDoc = (id) => (PV.state.consuntivi || []).find((x) => x.id === id) || null;
  const lineHtml = (line = {}) => `<div class="pv-line" data-cons-line>
    <label class="pv-label pv-line-wide"><span>Descrizione *</span><input data-cons-desc required value="${PV.escapeHtml(line.description || '')}"></label>
    <label class="pv-label"><span>Codice</span><input data-cons-code value="${PV.escapeHtml(line.code || '')}"></label>
    <label class="pv-label"><span>U.M.</span><input data-cons-unit value="${PV.escapeHtml(line.unit || '')}"></label>
    <label class="pv-label"><span>Quantità *</span><input type="number" min="0" step="0.0001" data-cons-qty required value="${PV.escapeHtml(line.quantity ?? 1)}"></label>
    <label class="pv-label"><span>Prezzo unitario</span><input type="number" min="0" step="0.01" data-cons-price value="${PV.escapeHtml(line.unitPrice ?? 0)}"></label>
    <button type="button" class="pv-icon-btn" data-cons-remove aria-label="Rimuovi">✕</button>
  </div>`;

  function ensureTab() {
    const nav = PV.page()?.querySelector('.pv-nav');
    if (nav && !nav.querySelector('[data-pv-view="consuntivi"]')) nav.insertAdjacentHTML('beforeend', '<button type="button" class="pv-tab" data-pv-view="consuntivi">Consuntivi</button>');
  }
  function renderOverview() {
    const q = PV.normalizeText(PV.state.consuntiviSearch);
    const docs = [...(PV.state.consuntivi || [])].filter((d) => !q || PV.normalizeText([d.number,d.contractNumber,d.requester,d.commessaName,d.plantName,d.subject,d.workLocation].join(' ')).includes(q));
    PV.content().innerHTML = `<div class="pv-shell"><div class="pv-toolbar"><div><h2>Consuntivi</h2><p class="pv-muted">Registra lavori eseguiti, date, quantità e importi.</p></div><button class="pv-btn pv-btn-primary" type="button" data-cons-new>+ Nuovo consuntivo</button></div>
      <div class="pv-toolbar"><input class="pv-search" data-cons-search type="search" value="${PV.escapeHtml(PV.state.consuntiviSearch)}" placeholder="Cerca numero, contratto, richiedente, commessa o impianto…"></div>
      <div class="pv-grid">${docs.length ? docs.map((d) => { const t = totals(d); return `<article class="pv-card"><div><span class="pv-chip">${PV.escapeHtml(d.number)}</span><h3>${PV.escapeHtml(d.commessaName || 'Commessa non indicata')}</h3><p class="pv-muted">${PV.escapeHtml(d.subject || 'Nessun oggetto')}</p></div><div class="pv-card-meta"><div>Contratto<strong>${PV.escapeHtml(d.contractNumber || '—')}</strong></div><div>Richiedente<strong>${PV.escapeHtml(d.requester || '—')}</strong></div><div>Periodo<strong>${PV.formatDate(d.startDate)} – ${PV.formatDate(d.endDate)}</strong></div><div>Impianto<strong>${PV.escapeHtml(d.plantName || '—')}</strong></div></div><div class="pv-card-total">${PV.formatMoney(t.total)}</div><div class="pv-card-actions"><button class="pv-btn pv-btn-secondary pv-btn-small" data-cons-edit="${PV.escapeHtml(d.id)}">Apri</button><button class="pv-btn pv-btn-light pv-btn-small" data-cons-print="${PV.escapeHtml(d.id)}">Stampa</button><button class="pv-btn pv-btn-light pv-btn-small" data-cons-duplicate="${PV.escapeHtml(d.id)}">Duplica</button><button class="pv-btn pv-btn-danger pv-btn-small" data-cons-delete="${PV.escapeHtml(d.id)}">Elimina</button></div></article>`; }).join('') : '<div class="pv-empty"><strong>Nessun consuntivo presente</strong>Crea il primo consuntivo.</div>'}</div></div>`;
  }
  function renderEditor(existing) {
    const d = existing ? JSON.parse(JSON.stringify(existing)) : { number: nextNumber(), date: PV.todayIso(), startDate: PV.todayIso(), endDate: PV.todayIso(), status:'Bozza', vatRate:22, lines:[] };
    const r = registry();
    const commesseOptions = r.commesse.map((c) => `<option value="${PV.escapeHtml(c.id)}" ${c.id === d.commessaId ? 'selected':''}>${PV.escapeHtml(c.name)}${c.code ? ` — ${PV.escapeHtml(c.code)}`:''}</option>`).join('');
    PV.content().innerHTML = `<div class="pv-shell"><div class="pv-section-head"><div><h2>${existing ? `Consuntivo ${PV.escapeHtml(d.number)}` : 'Nuovo consuntivo'}</h2><p class="pv-muted">I campi con * sono obbligatori.</p></div><button type="button" class="pv-btn pv-btn-light" data-cons-cancel>← Elenco consuntivi</button></div>
      <form data-cons-form><section class="pv-form-card"><h3>Intestazione e collegamento</h3><div class="pv-form-grid">
      <label class="pv-label"><span>Numero consuntivo *</span><input name="number" required value="${PV.escapeHtml(d.number || '')}"></label><label class="pv-label"><span>Data *</span><input type="date" name="date" required value="${PV.escapeHtml(d.date || '')}"></label>
      <label class="pv-label"><span>Numero contratto *</span><input name="contractNumber" required value="${PV.escapeHtml(d.contractNumber || '')}"></label><label class="pv-label"><span>Richiedente *</span><input name="requester" required value="${PV.escapeHtml(d.requester || '')}"></label>
      <label class="pv-label"><span>Data inizio lavori *</span><input type="date" name="startDate" required value="${PV.escapeHtml(d.startDate || '')}"></label><label class="pv-label"><span>Data fine lavori *</span><input type="date" name="endDate" required value="${PV.escapeHtml(d.endDate || '')}"></label>
      <label class="pv-label"><span>Commessa *</span><select name="commessaId" data-doc-commessa required><option value="">Seleziona commessa</option>${commesseOptions}</select></label><label class="pv-label"><span>Impianto *</span><select name="plantId" data-doc-plant required><option value="">Seleziona prima la commessa</option></select></label>
      <label class="pv-label"><span>Codice commessa</span><input name="commessaCode" value="${PV.escapeHtml(d.commessaCode || '')}"></label><label class="pv-label"><span>ID SAP</span><input name="plantSap" value="${PV.escapeHtml(d.plantSap || '')}"></label>
      <label class="pv-label"><span>Committente *</span><input name="clientName" required value="${PV.escapeHtml(d.clientName || '')}"></label><label class="pv-label"><span>Stato</span><select name="status">${['Bozza','Da verificare','Approvato','Inviato'].map((s)=>`<option ${s===d.status?'selected':''}>${s}</option>`).join('')}</select></label>
      <label class="pv-label pv-span-2"><span>Indirizzo / luogo lavori *</span><input name="workLocation" required value="${PV.escapeHtml(d.workLocation || '')}"></label><label class="pv-label"><span>Comune</span><input name="city" value="${PV.escapeHtml(d.city || '')}"></label><label class="pv-label"><span>Tipologia impianto</span><input name="plantType" value="${PV.escapeHtml(d.plantType || '')}"></label>
      <label class="pv-label pv-span-2"><span>Oggetto lavori *</span><input name="subject" required value="${PV.escapeHtml(d.subject || '')}"></label><label class="pv-label pv-span-2"><span>Attività eseguite *</span><textarea name="description" required>${PV.escapeHtml(d.description || '')}</textarea></label>
      <label class="pv-label"><span>Ore lavorate</span><input type="number" min="0" step="0.25" name="hours" value="${PV.escapeHtml(d.hours || '')}"></label><label class="pv-label"><span>Numero operatori</span><input type="number" min="0" step="1" name="operators" value="${PV.escapeHtml(d.operators || '')}"></label>
      <label class="pv-label"><span>Mezzi utilizzati</span><input name="vehicles" value="${PV.escapeHtml(d.vehicles || '')}"></label><label class="pv-label"><span>Materiali utilizzati</span><input name="materials" value="${PV.escapeHtml(d.materials || '')}"></label>
      <label class="pv-label pv-span-2"><span>Link allegati / Google Drive</span><input name="attachments" value="${PV.escapeHtml(d.attachments || '')}"></label><label class="pv-label"><span>Compilatore *</span><input name="compiler" required value="${PV.escapeHtml(d.compiler || PV.currentUser().displayName)}"></label><label class="pv-label"><span>IVA %</span><input type="number" min="0" step="0.01" name="vatRate" value="${PV.escapeHtml(d.vatRate ?? 22)}"></label>
      <label class="pv-label pv-span-2"><span>Note operative</span><textarea name="notes">${PV.escapeHtml(d.notes || '')}</textarea></label></div></section>
      <section class="pv-form-card"><div class="pv-section-head"><h3>Lavorazioni economiche</h3><button type="button" class="pv-btn pv-btn-primary" data-cons-add-line>+ Aggiungi lavorazione</button></div><div data-cons-lines>${(d.lines?.length ? d.lines : [{}]).map(lineHtml).join('')}</div></section>
      <div class="pv-form-actions">${existing ? '<button type="button" class="pv-btn pv-btn-light" data-cons-print-current>Stampa</button>' : ''}<button type="button" class="pv-btn pv-btn-light" data-cons-cancel>Annulla</button><button type="submit" class="pv-btn pv-btn-primary">Salva consuntivo</button></div></form></div>`;
    populatePlants(d.commessaId, d.plantId);
  }
  function populatePlants(commessaId, selected='') {
    const select = PV.page()?.querySelector('[data-doc-plant]'); if (!select) return;
    const plants = registry().plants.filter((p) => !commessaId || !p.commessaId || p.commessaId === commessaId);
    select.innerHTML = '<option value="">Seleziona impianto</option>' + plants.map((p)=>`<option value="${PV.escapeHtml(p.id)}" ${p.id===selected?'selected':''}>${PV.escapeHtml(p.name)}${p.sap?` — ${PV.escapeHtml(p.sap)}`:''}</option>`).join('');
  }
  function fillFromCommessa(id) {
    const form = PV.page()?.querySelector('[data-cons-form], [data-pv-quote-form]'); const c = registry().commesse.find((x)=>x.id===id); if (!form || !c) return;
    const set = (name,val)=>{ const el=form.elements[name]; if(el && val) el.value=val; };
    set('commessaCode',c.code); set('clientName',c.client); set('contractNumber',c.contract); set('requester',c.requester); populatePlants(id);
  }
  function fillFromPlant(id) {
    const form = PV.page()?.querySelector('[data-cons-form], [data-pv-quote-form]'); const p = registry().plants.find((x)=>x.id===id); if (!form || !p) return;
    const set = (name,val)=>{ const el=form.elements[name]; if(el && val) el.value=val; };
    set('plantSap',p.sap); set('workLocation',[p.address,p.city].filter(Boolean).join(', ')); set('city',p.city); set('plantType',p.type); set('plantName',p.name);
  }
  function save(form) {
    if (!form.reportValidity()) return;
    const fd = new FormData(form); const existing = getDoc(PV.state.editingConsuntivoId); const id = existing?.id || PV.uid('cons');
    const c = registry().commesse.find((x)=>x.id===fd.get('commessaId')); const p = registry().plants.find((x)=>x.id===fd.get('plantId'));
    const lines = [...form.querySelectorAll('[data-cons-line]')].map((row)=>({ id:PV.uid('line'), description:row.querySelector('[data-cons-desc]').value.trim(), code:row.querySelector('[data-cons-code]').value.trim(), unit:row.querySelector('[data-cons-unit]').value.trim(), quantity:PV.parseNumber(row.querySelector('[data-cons-qty]').value), unitPrice:PV.parseNumber(row.querySelector('[data-cons-price]').value) })).filter((x)=>x.description);
    const doc = { ...(existing||{}), id, ...Object.fromEntries(fd.entries()), commessaName:c?.name||existing?.commessaName||'', plantName:p?.name||existing?.plantName||'', lines, updatedAt:PV.nowIso(), createdAt:existing?.createdAt||PV.nowIso(), updatedBy:PV.currentUser(), syncPending:true };
    doc.vatRate=PV.parseNumber(doc.vatRate); doc.hours=PV.parseNumber(doc.hours); doc.operators=PV.parseNumber(doc.operators);
    const i=PV.state.consuntivi.findIndex((x)=>x.id===id); if(i>=0) PV.state.consuntivi[i]=doc; else PV.state.consuntivi.push(doc); write();
    if (PV.state.storageMode !== 'device' && PV.state.firestore) PV.saveRemote(collection, doc).catch(()=>{});
    PV.state.editingConsuntivoId=''; renderOverview();
  }
  function printDoc(d) {
    if (!d) return; const t=totals(d); const popup=window.open('','_blank','noopener,noreferrer'); if(!popup) return;
    popup.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${PV.escapeHtml(d.number)}</title><style>body{font-family:Arial;padding:28px;color:#172033}h1{color:#126b36}.meta{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:18px 0}.box{border:1px solid #ccd5df;padding:10px}table{width:100%;border-collapse:collapse;margin-top:18px}th,td{border:1px solid #ccd5df;padding:8px;text-align:left}th{background:#eef5f0}.tot{text-align:right;font-weight:700;margin-top:16px}.sign{display:flex;justify-content:space-between;margin-top:70px}</style></head><body><h1>Consuntivo ${PV.escapeHtml(d.number)}</h1><div class="meta"><div class="box"><b>Contratto:</b> ${PV.escapeHtml(d.contractNumber)}</div><div class="box"><b>Richiedente:</b> ${PV.escapeHtml(d.requester)}</div><div class="box"><b>Commessa:</b> ${PV.escapeHtml(d.commessaName)} ${PV.escapeHtml(d.commessaCode||'')}</div><div class="box"><b>Impianto:</b> ${PV.escapeHtml(d.plantName)} ${PV.escapeHtml(d.plantSap||'')}</div><div class="box"><b>Periodo:</b> ${PV.formatDate(d.startDate)} – ${PV.formatDate(d.endDate)}</div><div class="box"><b>Luogo:</b> ${PV.escapeHtml(d.workLocation)}</div></div><h2>${PV.escapeHtml(d.subject)}</h2><p>${PV.escapeHtml(d.description)}</p><table><thead><tr><th>Codice</th><th>Descrizione</th><th>U.M.</th><th>Q.tà</th><th>Prezzo</th><th>Totale</th></tr></thead><tbody>${(d.lines||[]).map((l)=>`<tr><td>${PV.escapeHtml(l.code)}</td><td>${PV.escapeHtml(l.description)}</td><td>${PV.escapeHtml(l.unit)}</td><td>${l.quantity}</td><td>${PV.formatMoney(l.unitPrice)}</td><td>${PV.formatMoney(l.quantity*l.unitPrice)}</td></tr>`).join('')}</tbody></table><div class="tot">Imponibile ${PV.formatMoney(t.subtotal)}<br>IVA ${PV.formatMoney(t.vatAmount)}<br>Totale ${PV.formatMoney(t.total)}</div><p><b>Note:</b> ${PV.escapeHtml(d.notes||'—')}</p><p><b>Compilatore:</b> ${PV.escapeHtml(d.compiler||'')}</p><div class="sign"><span>Firma committente __________________</span><span>Firma impresa __________________</span></div><script>window.print()<\/script></body></html>`); popup.document.close();
  }

  const originalEnsurePage = PV.ensurePage.bind(PV);
  PV.ensurePage = () => { originalEnsurePage(); ensureTab(); };
  const originalRender = PV.renderCurrentView.bind(PV);
  PV.renderCurrentView = () => { ensureTab(); if (PV.state.view === 'consuntivi') { PV.page()?.querySelectorAll('[data-pv-view]').forEach((b)=>b.classList.toggle('is-active',b.dataset.pvView==='consuntivi')); return PV.state.editingConsuntivoId ? renderEditor(PV.state.editingConsuntivoId==='new'?null:getDoc(PV.state.editingConsuntivoId)) : renderOverview(); } return originalRender(); };

  document.addEventListener('click',(e)=>{
    const tab=e.target.closest('[data-pv-view="consuntivi"]'); if(tab){e.preventDefault();e.stopImmediatePropagation();PV.state.view='consuntivi';PV.state.editingQuoteId='';PV.state.editingPriceListId='';PV.state.editingConsuntivoId='';PV.renderCurrentView();return;}
    if(!e.target.closest(`#${PV.pageId}`))return;
    if(e.target.closest('[data-cons-new]')){PV.state.editingConsuntivoId='new';renderEditor(null);}
    else if(e.target.closest('[data-cons-cancel]')){PV.state.editingConsuntivoId='';renderOverview();}
    else if(e.target.closest('[data-cons-add-line]')){PV.page().querySelector('[data-cons-lines]').insertAdjacentHTML('beforeend',lineHtml());}
    else if(e.target.closest('[data-cons-remove]')){const rows=PV.page().querySelectorAll('[data-cons-line]'); if(rows.length>1)e.target.closest('[data-cons-line]').remove();}
    else if(e.target.closest('[data-cons-edit]')){PV.state.editingConsuntivoId=e.target.closest('[data-cons-edit]').dataset.consEdit;renderEditor(getDoc(PV.state.editingConsuntivoId));}
    else if(e.target.closest('[data-cons-print]'))printDoc(getDoc(e.target.closest('[data-cons-print]').dataset.consPrint));
    else if(e.target.closest('[data-cons-print-current]'))printDoc(getDoc(PV.state.editingConsuntivoId));
    else if(e.target.closest('[data-cons-duplicate]')){const d=getDoc(e.target.closest('[data-cons-duplicate]').dataset.consDuplicate);if(d){const copy={...JSON.parse(JSON.stringify(d)),id:PV.uid('cons'),number:nextNumber(),status:'Bozza',createdAt:PV.nowIso(),updatedAt:PV.nowIso(),syncPending:true};PV.state.consuntivi.push(copy);write();renderOverview();}}
    else if(e.target.closest('[data-cons-delete]')){const id=e.target.closest('[data-cons-delete]').dataset.consDelete;if(confirm('Eliminare questo consuntivo?')){PV.state.consuntivi=PV.state.consuntivi.filter((x)=>x.id!==id);write();if(PV.state.firestore)PV.deleteRemote(collection,id).catch(()=>{});renderOverview();}}
  },true);
  document.addEventListener('input',(e)=>{if(e.target.matches('[data-cons-search]')){PV.state.consuntiviSearch=e.target.value;renderOverview();PV.page().querySelector('[data-cons-search]')?.focus();}},true);
  document.addEventListener('change',(e)=>{if(e.target.matches('[data-doc-commessa]'))fillFromCommessa(e.target.value);else if(e.target.matches('[data-doc-plant]'))fillFromPlant(e.target.value);},true);
  document.addEventListener('submit',(e)=>{const form=e.target.closest('[data-cons-form]');if(form){e.preventDefault();e.stopImmediatePropagation();save(form);}},true);

  // Aggiunge la selezione commessa/impianto anche ai preventivi senza alterare i preventivi esistenti.
  const originalQuoteEditor = PV.renderQuoteEditor?.bind(PV);
  if (originalQuoteEditor) PV.renderQuoteEditor = (quote) => {
    originalQuoteEditor(quote);
    const form=PV.page()?.querySelector('[data-pv-quote-form]'); const grid=form?.querySelector('.pv-form-grid'); if(!grid || grid.querySelector('[data-doc-commessa]')) return;
    const r=registry();
    grid.insertAdjacentHTML('afterbegin',`<label class="pv-label"><span>Commessa</span><select name="commessaId" data-doc-commessa><option value="">Seleziona commessa</option>${r.commesse.map((c)=>`<option value="${PV.escapeHtml(c.id)}" ${c.id===quote?.commessaId?'selected':''}>${PV.escapeHtml(c.name)}</option>`).join('')}</select></label><label class="pv-label"><span>Impianto</span><select name="plantId" data-doc-plant><option value="">Seleziona prima la commessa</option></select></label><label class="pv-label"><span>Codice commessa</span><input name="commessaCode" value="${PV.escapeHtml(quote?.commessaCode||'')}"></label><label class="pv-label"><span>ID SAP</span><input name="plantSap" value="${PV.escapeHtml(quote?.plantSap||'')}"></label>`);
    populatePlants(quote?.commessaId||'',quote?.plantId||'');
  };
  const originalSaveQuote = PV.saveQuote?.bind(PV);
  if (originalSaveQuote) PV.saveQuote = (form) => {
    const editingId=PV.state.editingQuoteId; originalSaveQuote(form);
    const saved=PV.state.quotes.find((q)=>q.id===editingId) || [...PV.state.quotes].sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt)))[0];
    if(saved){const fd=new FormData(form);const c=registry().commesse.find((x)=>x.id===fd.get('commessaId'));const p=registry().plants.find((x)=>x.id===fd.get('plantId'));Object.assign(saved,{commessaId:fd.get('commessaId')||'',commessaName:c?.name||'',commessaCode:fd.get('commessaCode')||'',plantId:fd.get('plantId')||'',plantName:p?.name||'',plantSap:fd.get('plantSap')||''});PV.persistLocal();PV.scheduleSync?.();}
  };
})();
