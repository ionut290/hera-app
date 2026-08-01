(() => {
  'use strict';

  if (window.__pvCommessaSearchBridgeV2) return;
  window.__pvCommessaSearchBridgeV2 = true;

  const PV = window.HeraPreventivi;
  if (!PV) return;

  const FORM_SELECTOR = '[data-pv-quote-form], [data-cons-form]';
  const COMMESSA_SELECTOR = '[data-matrix-commessa], [data-pvd-commessa]';
  const PLANT_SELECTOR = '[data-matrix-plant-search], [data-pvd-plant-search]';

  const state = {
    globals: [],
    globalAt: 0,
    plants: new Map(),
    rendered: new WeakMap(),
    selectionByEditor: new Map(),
    queued: false
  };

  const clean = (value) => String(value ?? '').trim().replace(/\s+/g, ' ');
  const normalize = (value) => clean(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
  const escapeHtml = (value) => String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  const first = (object, keys) => {
    for (const key of keys) {
      if (clean(object?.[key])) return object[key];
    }
    return '';
  };
  const db = () => window.firebase?.firestore?.() || null;
  const normalCollection = () => {
    try {
      return window.getCommesseCollectionName?.() || 'commesse';
    } catch (_) {
      return 'commesse';
    }
  };
  const invalidPlantName = (value) => !clean(value) || [
    'senza nome',
    'impianto senza nome',
    'impianto da identificare',
    'non disponibile',
    'non definito',
    'undefined',
    'null',
    '-'
  ].includes(normalize(value));

  function editorKey(form) {
    if (!form) return '';
    if (form.matches('[data-cons-form]')) {
      return `consuntivo:${PV.state?.editingConsuntivoId || 'new'}`;
    }
    return `preventivo:${PV.state?.editingQuoteId || 'new'}`;
  }

  function commessaInfo(select) {
    const value = clean(select?.value);
    const option = select?.selectedOptions?.[0];
    const isGlobal = value.startsWith('global::')
      || option?.dataset.scope === 'global'
      || /^GLOBAL\b/i.test(clean(option?.textContent));
    return {
      value,
      scope: isGlobal ? 'global' : 'operativa',
      id: isGlobal ? value.replace(/^global::/, '') : value,
      name: clean(option?.dataset.name || option?.textContent)
        .replace(/^GLOBAL\s*[—-]?\s*/i, '')
        .replace(/\s+[—-]\s+[^—]+$/, '')
        .trim(),
      code: clean(option?.dataset.code),
      label: clean(option?.textContent)
    };
  }

  function findBackingControl(form, name) {
    return [...form.querySelectorAll(`[name="${name}"]`)].find((control) => (
      !control.matches?.(COMMESSA_SELECTOR)
      && !control.matches?.(PLANT_SELECTOR)
    )) || null;
  }

  function setBackingValue(form, name, value, label = '') {
    let control = findBackingControl(form, name);
    if (!control) {
      control = document.createElement('input');
      control.type = 'hidden';
      control.name = name;
      control.dataset.commessaBridgeBacking = '1';
      form.appendChild(control);
    }
    if (control.tagName === 'SELECT' && value && ![...control.options].some((option) => option.value === value)) {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = label || value;
      option.dataset.commessaBridgeOption = '1';
      control.appendChild(option);
    }
    control.value = value ?? '';
    return control;
  }

  function getBackingValue(form, name) {
    return clean(findBackingControl(form, name)?.value);
  }

  function rememberSelection(form, select) {
    const info = commessaInfo(select);
    const key = editorKey(form);
    if (!info.id) {
      if (key) state.selectionByEditor.delete(key);
      delete form.dataset.commessaBridgeValue;
      delete form.dataset.commessaBridgeScope;
      delete form.dataset.commessaBridgeName;
      delete form.dataset.commessaBridgeCode;
      return null;
    }
    const saved = { ...info };
    if (key) state.selectionByEditor.set(key, saved);
    form.dataset.commessaBridgeValue = saved.value;
    form.dataset.commessaBridgeScope = saved.scope;
    form.dataset.commessaBridgeName = saved.name;
    form.dataset.commessaBridgeCode = saved.code;
    setBackingValue(form, 'commessaId', saved.id, saved.label || saved.name);
    setBackingValue(form, 'commessaSource', saved.scope);
    return saved;
  }

  function rememberedSelection(form) {
    const key = editorKey(form);
    const fromMap = key ? state.selectionByEditor.get(key) : null;
    if (fromMap?.id) return fromMap;

    const datasetValue = clean(form.dataset.commessaBridgeValue);
    if (datasetValue) {
      const scope = clean(form.dataset.commessaBridgeScope) || (datasetValue.startsWith('global::') ? 'global' : 'operativa');
      return {
        value: datasetValue,
        scope,
        id: scope === 'global' ? datasetValue.replace(/^global::/, '') : datasetValue,
        name: clean(form.dataset.commessaBridgeName),
        code: clean(form.dataset.commessaBridgeCode),
        label: ''
      };
    }

    const id = getBackingValue(form, 'commessaId');
    if (!id) return null;
    const scope = getBackingValue(form, 'commessaSource') || 'operativa';
    return {
      value: scope === 'global' ? `global::${id}` : id,
      scope,
      id,
      name: '',
      code: '',
      label: ''
    };
  }

  function restoreSelection(form, select) {
    if (!form || !select) return null;
    if (clean(select.value)) return rememberSelection(form, select);

    const saved = rememberedSelection(form);
    if (!saved?.id) return null;
    const targetValue = saved.value || (saved.scope === 'global' ? `global::${saved.id}` : saved.id);
    const option = [...select.options].find((item) => item.value === targetValue)
      || [...select.options].find((item) => item.value === saved.id)
      || [...select.options].find((item) => normalize(item.textContent) === normalize(saved.label));
    if (!option) return saved;

    select.value = option.value;
    const restored = rememberSelection(form, select) || saved;
    select.dataset.commessaRestored = '1';
    return restored;
  }

  function normalizePlant(raw, id, commessa) {
    const sap = clean(first(raw, ['idSap', 'idSAP', 'ID SAP', 'ID_SAP', 'id_sap', 'sap', 'sapId', 'codiceSap', 'codiceHera', 'codiceImpianto']));
    const city = clean(first(raw, ['comune', 'Comune', 'city', 'localita', 'municipality']));
    const address = clean(first(raw, ['indirizzo', 'Descrizione via', 'descrizioneVia', 'via', 'address', 'ubicazione']));
    const type = clean(first(raw, ['tipologia', 'Tipologia impianto', 'tipologiaImpianto', 'tipo', 'type']));
    const priceCode = clean(first(raw, ['codicePrezzo', 'Codice prezzo', 'CODICE PREZZO', 'codice_prezzo', 'priceCode', 'codicePrestazione']));
    let name = clean(first(raw, ['Denominazione Impianto', 'denominazioneImpianto', 'denominazione_impianto', 'nomeImpianto', 'Nome Impianto', 'plantName', 'siteName', 'denominazione', 'nome', 'name', 'impianto']));
    if (invalidPlantName(name)) {
      name = sap ? `Impianto SAP ${sap}` : city ? `Impianto di ${city}` : type || 'Impianto da identificare';
    }
    const distance = Number(first(raw, ['distance', 'distanza', 'distanceMeters', 'distanzaMetri']));
    return {
      raw,
      id: clean(id || first(raw, ['id', 'uid', 'impiantoId', 'plantId']) || sap || name),
      name,
      sap,
      city,
      address,
      type,
      priceCode,
      distance: Number.isFinite(distance) ? distance : null,
      scope: commessa.scope,
      commessaId: commessa.id
    };
  }

  function dedupePlants(items) {
    const map = new Map();
    const score = (item) => (invalidPlantName(item.name) ? 0 : 10)
      + (item.sap ? 5 : 0)
      + (item.city ? 2 : 0)
      + (item.address ? 2 : 0);
    for (const item of items) {
      const key = normalize(item.sap || item.id || `${item.name}|${item.city}|${item.address}`);
      if (!key) continue;
      const previous = map.get(key);
      if (!previous || score(item) > score(previous)) map.set(key, item);
    }
    return [...map.values()].sort((a, b) => (
      (a.distance ?? Number.MAX_SAFE_INTEGER) - (b.distance ?? Number.MAX_SAFE_INTEGER)
      || a.name.localeCompare(b.name, 'it')
    ));
  }

  async function loadGlobals(force = false) {
    if (!force && state.globals.length && Date.now() - state.globalAt < 60_000) return state.globals;
    const output = [];
    if (db()) {
      try {
        const snapshot = await db().collection('globalArchive').doc('commesse').collection('items').get();
        snapshot.docs.forEach((document) => {
          const raw = document.data() || {};
          output.push({
            id: document.id,
            name: clean(first(raw, ['nome', 'name', 'denominazione', 'titolo'])) || `Commessa ${document.id}`,
            code: clean(first(raw, ['codice', 'codiceCommessa', 'numeroCommessa'])),
            client: clean(first(raw, ['cliente', 'committente'])),
            contract: clean(first(raw, ['numeroContratto', 'contratto'])),
            requester: clean(first(raw, ['richiedente', 'referente']))
          });
        });
      } catch (error) {
        console.warn('Preventivi: archivio Global non leggibile.', error);
      }
    }
    try {
      const registry = window.HeraPreventiviRegistry?.registry?.();
      (registry?.commesse || [])
        .filter((item) => item.global === true || item.scope === 'global' || item.archivioPermanente === true)
        .forEach((item) => output.push({
          id: clean(item.id || item.sourceCommessaId),
          name: clean(item.name || item.nome),
          code: clean(item.code || item.codice),
          client: clean(item.client),
          contract: clean(item.contract),
          requester: clean(item.requester)
        }));
    } catch (_) {
      // Fallback non bloccante.
    }
    state.globals = [...new Map(output.filter((item) => item.id && item.name).map((item) => [item.id, item])).values()]
      .sort((a, b) => a.name.localeCompare(b.name, 'it'));
    state.globalAt = Date.now();
    return state.globals;
  }

  async function enhanceSelect(form, select, force = false) {
    if (!form || !select) return;
    const current = clean(select.value) ? rememberSelection(form, select) : rememberedSelection(form);

    select.querySelectorAll('option').forEach((option) => {
      if (!option.value || option.value.startsWith('global::')) return;
      option.dataset.scope = 'operativa';
      option.dataset.name = clean(option.textContent).split(' — ')[0];
      option.dataset.code = clean(option.textContent).split(' — ')[1] || '';
    });

    const globals = await loadGlobals(force);
    const signature = globals.map((item) => `${item.id}:${item.name}:${item.code}`).join('|');
    if (select.dataset.globalSignature !== signature) {
      select.querySelector('optgroup[data-global-group]')?.remove();
      if (globals.length) {
        const group = document.createElement('optgroup');
        group.label = 'GLOBAL';
        group.dataset.globalGroup = '1';
        group.innerHTML = globals.map((item) => (
          `<option value="global::${escapeHtml(item.id)}" data-scope="global" data-name="${escapeHtml(item.name)}" data-code="${escapeHtml(item.code)}">GLOBAL — ${escapeHtml(item.name)}${item.code ? ` — ${escapeHtml(item.code)}` : ''}</option>`
        )).join('');
        select.appendChild(group);
      }
      select.dataset.globalSignature = signature;
    }

    if (current?.id) {
      const target = current.value || (current.scope === 'global' ? `global::${current.id}` : current.id);
      const available = [...select.options].some((option) => option.value === target);
      if (available) select.value = target;
    }
    restoreSelection(form, select);
    select.dataset.commessaBridge = '1';
  }

  function setFieldValues(form, selector, value) {
    form.querySelectorAll(selector).forEach((control) => {
      if ('value' in control) control.value = value ?? '';
    });
  }

  function syncCommessa(form, commessa) {
    setBackingValue(form, 'commessaId', commessa.id, commessa.label || commessa.name);
    setBackingValue(form, 'commessaSource', commessa.scope);
    setFieldValues(form, '[data-pvm-field="commessa"]', commessa.name);
    setFieldValues(form, '[data-pvm-field="codice_commessa"], [name="commessaCode"]', commessa.code);
    const global = state.globals.find((item) => item.id === commessa.id);
    if (global) {
      setFieldValues(form, '[data-pvm-field="cliente"], [data-pvm-field="committente"], [name="clientName"]', global.client);
      setFieldValues(form, '[data-pvm-field="numero_contratto"], [name="contractNumber"]', global.contract);
      setFieldValues(form, '[data-pvm-field="richiedente"], [name="requester"]', global.requester);
    }
  }

  function clearPlant(form) {
    setBackingValue(form, 'plantId', '');
    const input = form.querySelector(PLANT_SELECTOR);
    if (input) {
      input.value = '';
      input.setCustomValidity('Seleziona un impianto dai risultati.');
    }
    setFieldValues(form, '[data-pvm-field="impianto"], [data-pvm-field="denominazione_impianto"], [data-pvm-field="id_sap"], [data-pvm-field="comune"], [data-pvm-field="indirizzo"], [data-pvm-field="tipologia_impianto"]', '');
  }

  function fallbackPlants(commessa) {
    const output = [];
    try {
      const registry = window.HeraPreventiviRegistry?.registry?.();
      for (const item of registry?.plants || []) {
        const scope = item.global === true || item.scope === 'global' ? 'global' : 'operativa';
        if (scope !== commessa.scope) continue;
        const relations = [item.commessaId, item.commessaName].map(normalize).filter(Boolean);
        const aliases = [commessa.id, commessa.name, commessa.code].map(normalize).filter(Boolean);
        if (relations.length && !relations.some((value) => aliases.some((alias) => value === alias || value.includes(alias) || alias.includes(value)))) continue;
        output.push(normalizePlant(item.raw || item, item.id, commessa));
      }
    } catch (_) {
      // Fallback non bloccante.
    }
    return output;
  }

  async function loadPlants(commessa, force = false) {
    const key = `${commessa.scope}::${commessa.id}`;
    if (!force && state.plants.has(key)) return state.plants.get(key);
    const output = [];
    if (db()) {
      if (commessa.scope === 'global') {
        try {
          const snapshot = await db().collection('globalArchive').doc('commesse').collection('items').doc(commessa.id).collection('impianti').get();
          snapshot.docs.forEach((document) => output.push(normalizePlant(document.data() || {}, document.id, commessa)));
        } catch (error) {
          console.warn('Preventivi: impianti Global non leggibili.', error);
        }
      } else {
        const reference = db().collection(normalCollection()).doc(commessa.id);
        for (const collectionName of ['impiantiFisici', 'impianti']) {
          try {
            const snapshot = await reference.collection(collectionName).get();
            snapshot.docs.forEach((document) => output.push(normalizePlant(document.data() || {}, document.id, commessa)));
          } catch (error) {
            console.warn(`Preventivi: ${collectionName} non leggibile.`, error);
          }
        }
      }
    }
    output.push(...fallbackPlants(commessa));
    const result = dedupePlants(output);
    state.plants.set(key, result);
    return result;
  }

  function resultsBox(form) {
    return form.querySelector('[data-matrix-plant-results], [data-pvd-results]');
  }

  async function renderResults(input, force = false) {
    const form = input?.closest?.(FORM_SELECTOR);
    const select = form?.querySelector(COMMESSA_SELECTOR);
    const box = form ? resultsBox(form) : null;
    if (!form || !select || !box) return;

    const restored = restoreSelection(form, select);
    const commessa = commessaInfo(select);
    if (!commessa.id && restored?.id) {
      commessa.id = restored.id;
      commessa.scope = restored.scope;
      commessa.name = restored.name;
      commessa.code = restored.code;
      commessa.value = restored.value;
    }
    if (!commessa.id) {
      box.innerHTML = '<p class="pv-muted">Seleziona prima una commessa operativa oppure GLOBAL.</p>';
      box.classList.remove('hidden');
      return;
    }

    syncCommessa(form, commessa);
    box.innerHTML = '<p class="pv-muted">Caricamento impianti…</p>';
    box.classList.remove('hidden');

    const items = await loadPlants(commessa, force);
    restoreSelection(form, select);
    const query = normalize(input.value);
    const filtered = items.filter((item) => (
      !query || normalize(`${item.name} ${item.sap} ${item.city} ${item.address} ${item.priceCode}`).includes(query)
    )).slice(0, 150);
    const map = new Map();
    state.rendered.set(form, map);
    box.innerHTML = filtered.length ? filtered.map((item, index) => {
      const key = `${item.id}::${index}`;
      map.set(key, item);
      const details = [
        item.sap ? `ID SAP ${item.sap}` : '',
        item.city,
        item.address,
        item.priceCode ? `Codice prezzo ${item.priceCode}` : ''
      ].filter(Boolean).join(' • ');
      return `<button type="button" class="pv-plant-result" data-commessa-plant="${escapeHtml(key)}"><strong>${escapeHtml(item.name)}</strong><span>${escapeHtml(details || 'Dati impianto disponibili')}</span></button>`;
    }).join('') : '<p class="pv-muted">Nessun impianto trovato. Cerca per nome, comune, indirizzo, ID SAP o codice prezzo.</p>';
    box.classList.remove('hidden');

    const feedback = form.querySelector('[data-matrix-feedback]');
    if (feedback) {
      feedback.textContent = items.length
        ? `${items.length} impianti disponibili nella commessa ${commessa.scope === 'global' ? 'GLOBAL ' : ''}${commessa.name}.`
        : 'Nessun impianto disponibile nella commessa selezionata.';
      feedback.dataset.type = items.length ? 'success' : 'warning';
    }
  }

  function choosePlant(button) {
    const form = button.closest(FORM_SELECTOR);
    const item = state.rendered.get(form)?.get(button.dataset.commessaPlant);
    if (!form || !item) return;

    const select = form.querySelector(COMMESSA_SELECTOR);
    restoreSelection(form, select);
    setBackingValue(form, 'plantId', item.id, item.name);
    setBackingValue(form, 'plantSource', item.scope);
    const input = form.querySelector(PLANT_SELECTOR);
    if (input) {
      input.value = `${item.name} — ID SAP ${item.sap || '—'}`;
      input.setCustomValidity('');
    }
    setFieldValues(form, '[data-pvm-field="impianto"], [data-pvm-field="denominazione_impianto"], [name="plantName"], [name="nomeImpianto"]', item.name);
    setFieldValues(form, '[data-pvm-field="id_sap"], [name="plantSap"], [name="idSap"]', item.sap);
    setFieldValues(form, '[data-pvm-field="comune"], [name="city"], [name="comune"]', item.city);
    setFieldValues(form, '[data-pvm-field="indirizzo"]', item.address);
    setFieldValues(form, '[data-pvm-field="tipologia_impianto"], [name="plantType"]', item.type);
    setFieldValues(form, '[name="workLocation"]', [item.address, item.city].filter(Boolean).join(', '));
    resultsBox(form)?.classList.add('hidden');
  }

  async function enhance(force = false) {
    await loadGlobals(force);
    for (const form of document.querySelectorAll(FORM_SELECTOR)) {
      const select = form.querySelector(COMMESSA_SELECTOR);
      if (select) await enhanceSelect(form, select, force);
      const input = form.querySelector(PLANT_SELECTOR);
      if (input) {
        input.dataset.commessaBridge = '1';
        input.placeholder = 'Cerca impianto, comune, indirizzo, ID SAP o codice prezzo…';
        if (invalidPlantName(input.value)) input.value = '';
      }
    }
  }

  window.addEventListener('change', (event) => {
    if (!event.target.matches?.(`${COMMESSA_SELECTOR}[data-commessa-bridge="1"]`)) return;
    event.stopImmediatePropagation();
    const form = event.target.closest(FORM_SELECTOR);
    if (!form) return;
    if (!clean(event.target.value)) {
      rememberSelection(form, event.target);
      clearPlant(form);
      return;
    }
    const commessa = rememberSelection(form, event.target) || commessaInfo(event.target);
    syncCommessa(form, commessa);
    clearPlant(form);
    state.plants.delete(`${commessa.scope}::${commessa.id}`);
    const input = form.querySelector(PLANT_SELECTOR);
    if (input) setTimeout(() => renderResults(input, true), 0);
  }, true);

  window.addEventListener('focusin', (event) => {
    if (!event.target.matches?.(`${PLANT_SELECTOR}[data-commessa-bridge="1"]`)) return;
    event.stopImmediatePropagation();
    const form = event.target.closest(FORM_SELECTOR);
    restoreSelection(form, form?.querySelector(COMMESSA_SELECTOR));
    setTimeout(() => renderResults(event.target, false), 0);
  }, true);

  window.addEventListener('input', (event) => {
    if (!event.target.matches?.(`${PLANT_SELECTOR}[data-commessa-bridge="1"]`)) return;
    event.stopImmediatePropagation();
    const form = event.target.closest(FORM_SELECTOR);
    restoreSelection(form, form?.querySelector(COMMESSA_SELECTOR));
    setBackingValue(form, 'plantId', '');
    event.target.setCustomValidity('Seleziona un impianto dai risultati.');
    setTimeout(() => renderResults(event.target, false), 0);
  }, true);

  window.addEventListener('click', (event) => {
    const button = event.target.closest?.('[data-commessa-plant]');
    if (button) {
      event.preventDefault();
      event.stopImmediatePropagation();
      choosePlant(button);
      return;
    }
    if (!event.target.matches?.(`${PLANT_SELECTOR}[data-commessa-bridge="1"]`)) return;
    event.stopImmediatePropagation();
    const form = event.target.closest(FORM_SELECTOR);
    restoreSelection(form, form?.querySelector(COMMESSA_SELECTOR));
    setTimeout(() => renderResults(event.target, false), 0);
  }, true);

  function queueEnhance(force = false) {
    if (state.queued) return;
    state.queued = true;
    requestAnimationFrame(() => {
      state.queued = false;
      enhance(force).catch((error) => console.warn('Preventivi: ricerca commessa non collegata.', error));
    });
  }

  const observer = new MutationObserver(() => queueEnhance(false));
  const start = () => {
    observer.observe(document.body, { childList: true, subtree: true });
    queueEnhance(true);
    [700, 1800, 4000].forEach((delay) => setTimeout(() => queueEnhance(false), delay));
    window.firebase?.auth?.()?.onAuthStateChanged?.((user) => {
      if (!user) return;
      state.globalAt = 0;
      queueEnhance(true);
    });
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();

  window.HeraPreventiviCommessaSearch = {
    refresh: () => {
      state.globalAt = 0;
      state.plants.clear();
      return enhance(true);
    },
    loadPlants,
    restoreSelection,
    version: '20260801b'
  };
})();