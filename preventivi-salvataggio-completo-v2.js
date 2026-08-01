(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  const M = window.HeraPreventiviModels;
  if (!PV || !M) return;

  const text = (value) => String(value ?? '').trim();
  const clone = (value) => {
    try { return JSON.parse(JSON.stringify(value)); } catch (_) { return value; }
  };
  const isFilled = (value) => value !== undefined && value !== null && !(typeof value === 'string' && !value.trim());
  const pending = new Map();

  function localClients() {
    try {
      const rows = JSON.parse(localStorage.getItem('hera_preventivi_clienti_v1') || '[]');
      return Array.isArray(rows) ? rows : [];
    } catch (_) {
      return [];
    }
  }

  function selectedText(select) {
    return text(select?.selectedOptions?.[0]?.textContent || '');
  }

  function collect(form) {
    const data = new FormData(form);
    const modelSelect = form.querySelector('[data-pvm-model-select]');
    const clientSelect = form.querySelector('[data-doc-client-select]');
    const commessaSelect = form.querySelector('[data-doc-commessa]');
    const plantIdField = form.querySelector('[data-doc-plant-id]');
    const plantSearch = form.querySelector('[data-doc-plant-search]');
    const model = M.selectedModel?.(form) || M.getModel?.(modelSelect?.value);
    const clientRegistryId = text(clientSelect?.value || data.get('clientRegistryId'));
    const clientSnapshot = localClients().find((item) => text(item.id) === clientRegistryId) || null;
    const plantLabel = text(plantSearch?.value);

    const snapshot = {
      number: text(data.get('number')),
      clientRegistryId,
      clientSnapshot: clientSnapshot ? clone(clientSnapshot) : undefined,
      modelId: text(modelSelect?.value || model?.id),
      modelName: text(model?.name || selectedText(modelSelect)),
      modelVersion: model?.version || 1,
      modelFields: M.modelPayload?.(form)?.modelFields || {},
      commessaId: text(commessaSelect?.value || data.get('commessaId')),
      commessaName: selectedText(commessaSelect).split(' — ')[0] || text(data.get('commessaName')),
      commessaCode: text(data.get('commessaCode')),
      plantId: text(data.get('plantId') || plantIdField?.value),
      plantName: plantLabel.replace(/\s+—\s+ID SAP.*$/i, '') || text(data.get('plantName')),
      plantSap: text(data.get('plantSap')),
      workLocation: text(data.get('workLocation')),
      city: text(data.get('city')),
      plantType: text(data.get('plantType')),
      clientName: text(data.get('clientName')),
      clientTaxCode: text(data.get('clientTaxCode')),
      requester: text(data.get('requester')),
      contractNumber: text(data.get('contractNumber')),
      contractExpiry: text(data.get('contractExpiry') || data.get('dataScadenzaContratto'))
    };
    Object.keys(snapshot).forEach((key) => {
      if (!isFilled(snapshot[key])) delete snapshot[key];
    });
    return snapshot;
  }

  function merge(target, source) {
    Object.entries(source || {}).forEach(([key, value]) => {
      if (isFilled(value)) target[key] = clone(value);
    });
    return target;
  }

  async function persistRecord(record, snapshot) {
    if (!record || !snapshot) return false;
    merge(record, snapshot);
    record.updatedAt = PV.nowIso?.() || new Date().toISOString();
    record.syncPending = true;
    PV.persistLocal?.();

    try {
      const payload = { ...record, syncPending: false };
      const result = await PV.saveRemote?.(PV.collections?.quotes, payload);
      if (result !== false) record.syncPending = false;
      PV.persistLocal?.();
      return true;
    } catch (error) {
      console.warn('Preventivi: salvataggio completo non sincronizzato.', error);
      return false;
    }
  }

  async function finalize(snapshot, preferredId = '') {
    if (!snapshot?.number) return;
    for (let attempt = 0; attempt < 12; attempt += 1) {
      const record = preferredId && preferredId !== 'new'
        ? PV.getQuote?.(preferredId)
        : [...(PV.state?.quotes || [])].reverse().find((item) => text(item.number) === snapshot.number);
      if (record) {
        await persistRecord(record, snapshot);
        return;
      }
      await new Promise((resolve) => setTimeout(resolve, 250));
    }
    console.warn('Preventivi: preventivo salvato non trovato per completare i metadati.', snapshot.number);
  }

  document.addEventListener('submit', (event) => {
    const form = event.target.closest?.('[data-pv-quote-form]');
    if (!form) return;
    const snapshot = collect(form);
    const editingId = PV.state?.editingQuoteId || '';
    pending.set(snapshot.number, snapshot);
    setTimeout(() => finalize(snapshot, editingId), 0);
    setTimeout(() => finalize(snapshot, editingId), 800);
    setTimeout(() => finalize(snapshot, editingId), 2000);
  }, true);

  const originalSaveRemote = PV.saveRemote?.bind(PV);
  if (originalSaveRemote && !PV.__completeQuoteRemoteV2) {
    PV.__completeQuoteRemoteV2 = true;
    PV.saveRemote = (collection, record) => {
      if (record && collection === PV.collections?.quotes) {
        const existing = (PV.state?.quotes || []).find((item) => item.id === record.id);
        const draft = pending.get(text(record.number));
        if (existing) merge(record, existing);
        if (draft) merge(record, draft);
      }
      return originalSaveRemote(collection, record);
    };
  }

  function restore(form) {
    const id = PV.state?.editingQuoteId;
    if (!id || id === 'new') return;
    const quote = PV.getQuote?.(id);
    if (!quote) return;

    const clientSelect = form.querySelector('[data-doc-client-select]');
    if (clientSelect && quote.clientRegistryId && clientSelect.value !== quote.clientRegistryId) clientSelect.value = quote.clientRegistryId;

    const modelSelect = form.querySelector('[data-pvm-model-select]');
    if (modelSelect && quote.modelId && modelSelect.value !== quote.modelId) {
      modelSelect.value = quote.modelId;
      M.renderDynamic?.(form, quote);
    }

    const commessaSelect = form.querySelector('[data-doc-commessa]');
    if (commessaSelect && quote.commessaId) commessaSelect.value = quote.commessaId;

    const hiddenPlant = form.querySelector('[data-doc-plant-id]');
    if (hiddenPlant && quote.plantId) hiddenPlant.value = quote.plantId;
    const plantSearch = form.querySelector('[data-doc-plant-search]');
    if (plantSearch && quote.plantName) {
      plantSearch.value = `${quote.plantName}${quote.plantSap ? ` — ID SAP ${quote.plantSap}` : ''}`;
      plantSearch.setCustomValidity('');
    }

    const values = {
      commessaCode: quote.commessaCode,
      plantSap: quote.plantSap,
      workLocation: quote.workLocation,
      city: quote.city,
      plantType: quote.plantType,
      requester: quote.requester,
      contractNumber: quote.contractNumber,
      contractExpiry: quote.contractExpiry
    };
    Object.entries(values).forEach(([name, value]) => {
      const field = form.elements?.namedItem(name);
      if (field && isFilled(value)) field.value = value;
    });
  }

  let queued = false;
  const queueRestore = () => {
    if (queued) return;
    queued = true;
    requestAnimationFrame(() => {
      queued = false;
      document.querySelectorAll('[data-pv-quote-form]').forEach(restore);
    });
  };

  const observer = new MutationObserver(queueRestore);
  const start = () => {
    observer.observe(document.body, { childList: true, subtree: true });
    queueRestore();
    [200, 600, 1200, 2500].forEach((delay) => setTimeout(queueRestore, delay));
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();
