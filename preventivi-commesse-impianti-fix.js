(() => {
  'use strict';

  const normalize = (value) => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const first = (obj, keys) => {
    for (const key of keys) {
      const value = obj?.[key];
      if (value !== undefined && value !== null && String(value).trim() !== '') return value;
    }
    return '';
  };

  const COMMESSA_KEYS = ['id','uid','commessaId','projectId','codice','codiceCommessa','numeroCommessa'];
  const COMMESSA_NAME_KEYS = ['nome','name','denominazione','titolo','commessa','descrizione'];
  const PLANT_ID_KEYS = ['id','uid','impiantoId','plantId','idSap','ID SAP','sap','codiceSap'];
  const PLANT_NAME_KEYS = ['denominazione','Denominazione Impianto','nome','name','impianto','descrizione'];

  function collectRegistry() {
    const commesse = [];
    const plants = [];
    const seenObjects = new WeakSet();

    const addCommessa = (raw, isGlobal, path) => {
      if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return null;
      const id = String(first(raw, COMMESSA_KEYS) || `${isGlobal ? 'global' : 'commessa'}-${normalize(first(raw, COMMESSA_NAME_KEYS)) || commesse.length}`);
      const name = String(first(raw, COMMESSA_NAME_KEYS) || 'Commessa senza nome');
      if (!name || /impianto senza nome/i.test(name)) return null;
      const item = {
        id,
        name,
        code: String(first(raw, ['codiceCommessa','codice','code','numeroCommessa']) || ''),
        client: String(first(raw, ['cliente','committente','clientName','ragioneSociale']) || ''),
        contract: String(first(raw, ['numeroContratto','contratto','contractNumber']) || ''),
        requester: String(first(raw, ['richiedente','referente','referenteCliente']) || ''),
        isGlobal: Boolean(isGlobal),
        path,
        raw
      };
      const existing = commesse.find((entry) => entry.id === item.id && entry.isGlobal === item.isGlobal);
      if (!existing) commesse.push(item);
      return existing || item;
    };

    const addPlant = (raw, commessa, isGlobal, path) => {
      if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return;
      const name = String(first(raw, PLANT_NAME_KEYS) || '');
      const sap = String(first(raw, ['idSap','ID SAP','sap','codiceSap']) || '');
      if (!name && !sap) return;
      const id = String(first(raw, PLANT_ID_KEYS) || `${commessa?.id || 'impianto'}-${normalize(name || sap)}-${plants.length}`);
      const explicitCommessa = String(first(raw, ['commessaId','projectId','idCommessa','codiceCommessa','commessa','commessaUid']) || '');
      const item = {
        id,
        commessaId: explicitCommessa || commessa?.id || '',
        commessaName: commessa?.name || String(first(raw, ['nomeCommessa','commessaNome']) || ''),
        isGlobal: Boolean(isGlobal || commessa?.isGlobal),
        name: name || `Impianto ${sap}`,
        sap,
        address: String(first(raw, ['indirizzo','Descrizione via','descrizioneVia','via','address','ubicazione']) || ''),
        city: String(first(raw, ['comune','Comune','city','localita']) || ''),
        type: String(first(raw, ['tipologia','Tipologia impianto','tipo','type']) || ''),
        area: String(first(raw, ['area','AREA','competenza','Area/Competenza']) || ''),
        coordinates: String(first(raw, ['coordinate','coordinates','Coordinate GPS(X)/GPS(Y)','coordinateGps']) || ''),
        raw,
        path
      };
      if (!plants.some((entry) => entry.id === item.id && entry.commessaId === item.commessaId && entry.isGlobal === item.isGlobal)) plants.push(item);
    };

    const walk = (value, path = '', inheritedCommessa = null, inheritedGlobal = false, depth = 0) => {
      if (depth > 8 || value === null || value === undefined) return;
      const isGlobal = inheritedGlobal || /(^|[._\-/ ])global([._\-/ ]|$)/i.test(path);

      if (Array.isArray(value)) {
        value.forEach((item, index) => walk(item, `${path}[${index}]`, inheritedCommessa, isGlobal, depth + 1));
        return;
      }
      if (typeof value !== 'object') return;
      if (seenObjects.has(value)) return;
      seenObjects.add(value);

      const looksPlant = Boolean(first(value, ['idSap','ID SAP','impiantoId','plantId','Denominazione Impianto','Tipologia impianto'])) || /impiant|plant|site/i.test(path.split('.').pop() || '');
      const nestedPlants = first(value, ['impianti','plants','sites','elencoImpianti','impiantiGlobal']);
      const looksCommessa = Boolean(first(value, ['codiceCommessa','numeroContratto','committente','projectId'])) || (!looksPlant && Boolean(nestedPlants)) || /commess|project|global/i.test(path.split('.').pop() || '');

      let currentCommessa = inheritedCommessa;
      if (looksCommessa && first(value, COMMESSA_NAME_KEYS)) currentCommessa = addCommessa(value, isGlobal, path) || inheritedCommessa;
      if (looksPlant && (first(value, PLANT_NAME_KEYS) || first(value, ['idSap','ID SAP','sap']))) addPlant(value, currentCommessa, isGlobal, path);

      Object.entries(value).forEach(([key, child]) => {
        if (['raw'].includes(key)) return;
        walk(child, path ? `${path}.${key}` : key, currentCommessa, isGlobal || /global/i.test(key), depth + 1);
      });
    };

    ['commesse','projects','cantieri','impianti','plants','sites','global','globalCommesse','commesseGlobal','globalImpianti','impiantiGlobal'].forEach((key) => {
      try { walk(window[key], `window.${key}`, null, /global/i.test(key)); } catch (_) { /* sorgente opzionale */ }
    });

    try {
      for (let i = 0; i < localStorage.length; i += 1) {
        const key = localStorage.key(i) || '';
        if (!/(commess|impiant|cantier|plant|site|global)/i.test(key)) continue;
        try { walk(JSON.parse(localStorage.getItem(key)), `localStorage.${key}`, null, /global/i.test(key)); } catch (_) { /* valore non JSON */ }
      }
    } catch (_) { /* localStorage non disponibile */ }

    const uniqueCommesse = [...new Map(commesse.map((item) => [`${item.isGlobal ? 'G' : 'N'}:${item.id}`, item])).values()]
      .sort((a, b) => Number(a.isGlobal) - Number(b.isGlobal) || a.name.localeCompare(b.name, 'it'));
    const uniquePlants = [...new Map(plants.map((item) => [`${item.isGlobal ? 'G' : 'N'}:${item.commessaId}:${item.id}`, item])).values()];
    return { commesse: uniqueCommesse, plants: uniquePlants };
  }

  function fieldByLabel(root, labelText) {
    const wanted = normalize(labelText);
    return [...root.querySelectorAll('label')].find((label) => normalize(label.querySelector('span')?.textContent || label.textContent).startsWith(wanted))?.querySelector('input,select,textarea') || null;
  }

  function commessaField(root) {
    return root.querySelector('[data-doc-commessa], select[name="commessaId"], input[name="commessaId"], select[name="commessa"], input[name="commessa"]') || fieldByLabel(root, 'Commessa');
  }

  function plantField(root) {
    return root.querySelector('[data-doc-plant], select[name="plantId"], input[name="plantId"], input[name="impianto"], select[name="impianto"]') || fieldByLabel(root, 'Impianto');
  }

  function selectedCommessa(registry, field) {
    if (!field) return null;
    const option = field.tagName === 'SELECT' ? field.selectedOptions?.[0] : null;
    const id = String(field.value || '');
    const label = String(option?.textContent || field.value || '').replace(/^GLOBAL\s*[—-]?\s*/i, '').trim();
    const isGlobal = /^GLOBAL\b/i.test(option?.textContent || '') || field.dataset.global === '1';
    return registry.commesse.find((item) => item.id === id && (!isGlobal || item.isGlobal))
      || registry.commesse.find((item) => normalize(item.name) === normalize(label) && (!isGlobal || item.isGlobal))
      || null;
  }

  function fillFromPlant(root, plant) {
    if (!plant) return;
    const set = (name, value) => {
      if (!value) return;
      const input = root.querySelector(`[name="${name}"]`);
      if (input && !String(input.value || '').trim()) input.value = value;
    };
    set('plantSap', plant.sap);
    set('idSap', plant.sap);
    set('workLocation', [plant.address, plant.city].filter(Boolean).join(', '));
    set('city', plant.city);
    set('comune', plant.city);
    set('plantType', plant.type);
    set('tipologiaImpianto', plant.type);
    set('area', plant.area);
    set('coordinates', plant.coordinates);
    set('plantName', plant.name);
    set('nomeImpianto', plant.name);
  }

  function refreshPlantChoices(root, force = false) {
    const registry = collectRegistry();
    const commessa = selectedCommessa(registry, commessaField(root));
    const plant = plantField(root);
    if (!plant) return;

    const matches = registry.plants.filter((item) => {
      if (!commessa) return true;
      if (item.isGlobal !== commessa.isGlobal) return false;
      const byId = item.commessaId && (item.commessaId === commessa.id || normalize(item.commessaId) === normalize(commessa.code));
      const byName = item.commessaName && normalize(item.commessaName) === normalize(commessa.name);
      const byPath = item.path && commessa.path && item.path.startsWith(commessa.path);
      return byId || byName || byPath || (!item.commessaId && !item.commessaName);
    });

    if (plant.tagName === 'SELECT') {
      const previous = plant.value;
      plant.innerHTML = '<option value="">Seleziona impianto</option>' + matches.map((item) => `<option value="${String(item.id).replace(/"/g, '&quot;')}">${item.isGlobal ? 'GLOBAL — ' : ''}${item.name}${item.sap ? ` — ${item.sap}` : ''}</option>`).join('');
      if (matches.some((item) => item.id === previous)) plant.value = previous;
      return;
    }

    let listId = plant.getAttribute('list');
    let list = listId ? document.getElementById(listId) : null;
    if (!list) {
      listId = `pv-impianti-${Math.random().toString(36).slice(2)}`;
      list = document.createElement('datalist');
      list.id = listId;
      plant.setAttribute('list', listId);
      plant.insertAdjacentElement('afterend', list);
    }
    const query = normalize(plant.value);
    const filtered = query && !force ? matches.filter((item) => normalize(`${item.name} ${item.sap} ${item.city} ${item.address}`).includes(query)) : matches;
    list.innerHTML = filtered.slice(0, 300).map((item) => `<option value="${String(item.name).replace(/"/g, '&quot;')}">${item.isGlobal ? 'GLOBAL — ' : ''}${item.sap ? `${item.sap} — ` : ''}${item.city || ''}</option>`).join('');
    plant.dataset.registryCount = String(matches.length);
  }

  function augmentCommesse(root) {
    const field = commessaField(root);
    if (!field || field.tagName !== 'SELECT') return;
    const registry = collectRegistry();
    const previous = field.value;
    const current = [...field.options].map((option) => ({ value: option.value, text: option.textContent || '' }));
    const placeholder = current.find((item) => !item.value)?.text || 'Seleziona commessa';
    const options = registry.commesse.map((item) => ({ value: item.id, text: `${item.isGlobal ? 'GLOBAL — ' : ''}${item.name}${item.code ? ` — ${item.code}` : ''}` }));
    const merged = [...new Map(options.map((item) => [`${item.text}|${item.value}`, item])).values()];
    if (!merged.length) return;
    field.innerHTML = `<option value="">${placeholder}</option>` + merged.map((item) => `<option value="${String(item.value).replace(/"/g, '&quot;')}">${item.text}</option>`).join('');
    if ([...field.options].some((option) => option.value === previous)) field.value = previous;
  }

  function enhance(root = document) {
    const page = root.querySelector?.('#preventivi-page') || document.getElementById('preventivi-page');
    if (!page || page.classList.contains('hidden')) return;
    augmentCommesse(page);
    refreshPlantChoices(page, true);
  }

  document.addEventListener('change', (event) => {
    const page = event.target.closest?.('#preventivi-page');
    if (!page) return;
    if (event.target === commessaField(page)) {
      refreshPlantChoices(page, true);
      const registry = collectRegistry();
      const commessa = selectedCommessa(registry, event.target);
      if (commessa) {
        const set = (name, value) => { const input = page.querySelector(`[name="${name}"]`); if (input && value && !String(input.value || '').trim()) input.value = value; };
        set('commessaCode', commessa.code);
        set('clientName', commessa.client);
        set('contractNumber', commessa.contract);
        set('requester', commessa.requester);
      }
      return;
    }
    const plant = plantField(page);
    if (event.target === plant) {
      const registry = collectRegistry();
      const commessa = selectedCommessa(registry, commessaField(page));
      const match = registry.plants.find((item) => item.id === plant.value)
        || registry.plants.find((item) => normalize(item.name) === normalize(plant.value) && (!commessa || item.isGlobal === commessa.isGlobal));
      fillFromPlant(page, match);
    }
  }, true);

  document.addEventListener('input', (event) => {
    const page = event.target.closest?.('#preventivi-page');
    if (!page || event.target !== plantField(page)) return;
    refreshPlantChoices(page, false);
  }, true);

  const observer = new MutationObserver(() => enhance(document));
  const start = () => {
    observer.observe(document.body, { childList: true, subtree: true });
    enhance(document);
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();
