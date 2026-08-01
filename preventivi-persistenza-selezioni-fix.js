(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  const M = window.HeraPreventiviModels;
  if (!PV || !M) return;

  const text = (value) => String(value ?? '').trim();
  const clone = (value) => {
    try { return JSON.parse(JSON.stringify(value)); } catch (_) { return value; }
  };

  function currentQuote(form) {
    const id = PV.state?.editingQuoteId;
    if (id && id !== 'new') return PV.getQuote?.(id) || null;
    const number = text(new FormData(form).get('number'));
    return [...(PV.state?.quotes || [])].reverse().find((item) => text(item.number) === number) || null;
  }

  function selectedOptionText(select) {
    return text(select?.selectedOptions?.[0]?.textContent || '');
  }

  function collectSelectionMetadata(form) {
    if (!form) return {};
    const data = new FormData(form);
    const clientSelect = form.querySelector('[data-doc-client-select]');
    const modelSelect = form.querySelector('[data-pvm-model-select]');
    const commessaSelect = form.querySelector('[data-doc-commessa]');
    const plantId = text(data.get('plantId') || form.querySelector('[data-doc-plant-id]')?.value);
    const plantSearch = text(form.querySelector('[data-doc-plant-search]')?.value);
    const model = M.selectedModel?.(form);
    const clientId = text(clientSelect?.value || data.get('clientRegistryId'));
    const client = window.HeraPreventiviClienti?.getById?.(clientId)
      || window.HeraPreventiviClients?.getById?.(clientId)
      || null;

    return {
      clientRegistryId: clientId,
      clientSnapshot: client ? clone(client) : undefined,
      modelId: text(modelSelect?.value || model?.id),
      modelName: text(model?.name || selectedOptionText(modelSelect)),
      modelVersion: model?.version || 1,
      modelFields: M.modelPayload?.(form)?.modelFields || {},
      commessaId: text(commessaSelect?.value || data.get('commessaId')),
      commessaName: selectedOptionText(commessaSelect).split(' — ')[0] || text(data.get('commessaName')),
      commessaCode: text(data.get('commessaCode')),
      plantId,
      plantName: plantSearch.replace(/\s+—\s+ID SAP.*$/i, '') || text(data.get('plantName')),
      plantSap: text(data.get('plantSap')),
      workLocation: text(data.get('workLocation')),
      city: text(data.get('city')),
      plantType: text(data.get('plantType'))
    };
  }

  function mergeDefined(target, patch) {
    Object.entries(patch || {}).forEach(([key, value]) => {
      if (value === undefined || value === null) return;
      if (typeof value === 'string' && !value.trim()) return;
      target[key] = clone(value);
    });
    return target;
  }

  async function persistQuoteMetadata(form, quote) {
    if (!quote) return;
    mergeDefined(quote, collectSelectionMetadata(form));
    quote.updatedAt = PV.nowIso?.() || new Date().toISOString();
    quote.syncPending = true;
    PV.persistLocal?.();
    try {
      const saved = await PV.saveRemote?.(PV.collections?.quotes, { ...quote, syncPending: false });
      if (saved !== false) quote.syncPending = false;
      PV.persistLocal?.();
    } catch (error) {
      console.warn('Preventivi: metadati selezioni non sincronizzati.', error);
    }
  }

  const originalSaveQuote = PV.saveQuote?.bind(PV);
  if (originalSaveQuote && !PV.__selectionPersistenceSavePatched) {
    PV.__selectionPersistenceSavePatched = true;
    PV.saveQuote = async (form) => {
      const metadata = collectSelectionMetadata(form);
      const beforeId = PV.state?.editingQuoteId;
      const number = text(new FormData(form).get('number'));
      const result = await originalSaveQuote(form);
      const quote = beforeId && beforeId !== 'new'
        ? PV.getQuote?.(beforeId)
        : [...(PV.state?.quotes || [])].reverse().find((item) => text(item.number) === number);
      if (quote) {
        mergeDefined(quote, metadata);
        await persistQuoteMetadata(form, quote);
      }
      return result;
    };
  }

  const originalSaveRemote = PV.saveRemote?.bind(PV);
  if (originalSaveRemote && !PV.__selectionPersistenceRemotePatched) {
    PV.__selectionPersistenceRemotePatched = true;
    PV.saveRemote = (collection, record) => {
      if (record && collection === PV.collections?.quotes) {
        const existing = (PV.state?.quotes || []).find((item) => item.id === record.id);
        if (existing) {
          ['clientRegistryId','clientSnapshot','modelId','modelName','modelVersion','modelFields','commessaId','commessaName','commessaCode','plantId','plantName','plantSap','workLocation','city','plantType']
            .forEach((key) => {
              const missing = record[key] === undefined || record[key] === null || (typeof record[key] === 'string' && !record[key].trim());
              if (missing && existing[key] !== undefined) record[key] = clone(existing[key]);
            });
        }
      }
      return originalSaveRemote(collection, record);
    };
  }

  function restoreSelections(form) {
    if (!form) return;
    const quote = currentQuote(form);
    if (!quote) return;

    const clientSelect = form.querySelector('[data-doc-client-select]');
    if (clientSelect && quote.clientRegistryId) {
      clientSelect.value = quote.clientRegistryId;
      clientSelect.dataset.restored = '1';
    }

    const modelSelect = form.querySelector('[data-pvm-model-select]');
    if (modelSelect && quote.modelId) {
      modelSelect.value = quote.modelId;
      M.renderDynamic?.(form, quote);
    }

    const commessaSelect = form.querySelector('[data-doc-commessa]');
    if (commessaSelect && quote.commessaId) commessaSelect.value = quote.commessaId;

    const plantId = form.querySelector('[data-doc-plant-id]');
    if (plantId && quote.plantId) plantId.value = quote.plantId;
    const plantSearch = form.querySelector('[data-doc-plant-search]');
    if (plantSearch && quote.plantName) {
      plantSearch.value = `${quote.plantName}${quote.plantSap ? ` — ID SAP ${quote.plantSap}` : ''}`;
      plantSearch.setCustomValidity('');
    }

    const namedValues = {
      commessaCode: quote.commessaCode,
      plantSap: quote.plantSap,
      workLocation: quote.workLocation,
      city: quote.city,
      plantType: quote.plantType
    };
    Object.entries(namedValues).forEach(([name, value]) => {
      const field = form.elements?.namedItem(name);
      if (field && value !== undefined && !text(field.value)) field.value = value;
    });
  }

  let queued = false;
  function queueRestore() {
    if (queued) return;
    queued = true;
    requestAnimationFrame(() => {
      queued = false;
      const page = PV.page?.() || document.getElementById('preventivi-page');
      page?.querySelectorAll('[data-pv-quote-form]').forEach(restoreSelections);
    });
  }

  const observer = new MutationObserver(queueRestore);
  const start = () => {
    observer.observe(document.body, { childList: true, subtree: true });
    queueRestore();
    [250, 750, 1500].forEach((delay) => setTimeout(queueRestore, delay));
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();
