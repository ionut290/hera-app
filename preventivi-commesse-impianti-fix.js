(() => {
  'use strict';

  const norm = (v) => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const get = (o, keys) => {
    for (const k of keys) {
      const v = o?.[k];
      if (v !== undefined && v !== null && String(v).trim()) return v;
    }
    return '';
  };
  const esc = (v) => String(v ?? '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  function registry() {
    const commesse = [];
    const plants = [];
    const seen = new WeakSet();

    const addCommessa = (raw, global, path) => {
      const name = String(get(raw, ['nome','name','denominazione','titolo','commessa','descrizione']) || '');
      if (!name) return null;
      const id = String(get(raw, ['id','uid','commessaId','projectId','codice','codiceCommessa','numeroCommessa']) || `${global ? 'global' : 'commessa'}-${norm(name)}`);
      let item = commesse.find((x) => x.id === id && x.global === global);
      if (!item) {
        item = {
          id, name, global, path,
          code: String(get(raw, ['codiceCommessa','codice','code','numeroCommessa']) || ''),
          client: String(get(raw, ['cliente','committente','clientName','ragioneSociale']) || ''),
          contract: String(get(raw, ['numeroContratto','contratto','contractNumber']) || ''),
          requester: String(get(raw, ['richiedente','referente','referenteCliente']) || '')
        };
        commesse.push(item);
      }
      return item;
    };

    const addPlant = (raw, parent, global, path) => {
      const name = String(get(raw, ['denominazione','Denominazione Impianto','nome','name','impianto','descrizione']) || '');
      const sap = String(get(raw, ['idSap','ID SAP','sap','codiceSap']) || '');
      if (!name && !sap) return;
      const id = String(get(raw, ['id','uid','impiantoId','plantId','idSap','ID SAP','sap']) || `${parent?.id || 'impianto'}-${norm(name || sap)}`);
      const commessaId = String(get(raw, ['commessaId','projectId','idCommessa','codiceCommessa','commessa','commessaUid']) || parent?.id || '');
      const item = {
        id, commessaId, commessaName: parent?.name || String(get(raw, ['nomeCommessa','commessaNome']) || ''),
        global: Boolean(global || parent?.global), path, name: name || `Impianto ${sap}`, sap,
        address: String(get(raw, ['indirizzo','Descrizione via','descrizioneVia','via','address','ubicazione']) || ''),
        city: String(get(raw, ['comune','Comune','city','localita']) || ''),
        type: String(get(raw, ['tipologia','Tipologia impianto','tipo','type']) || ''),
        area: String(get(raw, ['area','AREA','competenza','Area/Competenza']) || ''),
        coordinates: String(get(raw, ['coordinate','coordinates','Coordinate GPS(X)/GPS(Y)','coordinateGps']) || '')
      };
      if (!plants.some((x) => x.id === id && x.commessaId === commessaId && x.global === item.global)) plants.push(item);
    };

    const walk = (value, path = '', parent = null, global = false, depth = 0) => {
      if (depth > 8 || value == null) return;
      global = global || /(^|[._\-/ ])global([._\-/ ]|$)/i.test(path);
      if (Array.isArray(value)) {
        value.forEach((x, i) => walk(x, `${path}[${i}]`, parent, global, depth + 1));
        return;
      }
      if (typeof value !== 'object' || seen.has(value)) return;
      seen.add(value);

      const last = path.split('.').pop() || '';
      const looksPlant = Boolean(get(value, ['idSap','ID SAP','impiantoId','plantId','Denominazione Impianto','Tipologia impianto'])) || /impiant|plant|site/i.test(last);
      const nested = get(value, ['impianti','plants','sites','elencoImpianti','impiantiGlobal']);
      const looksCommessa = Boolean(get(value, ['codiceCommessa','numeroContratto','committente','projectId'])) || (!looksPlant && Array.isArray(nested)) || /commess|project/i.test(last);
      let current = parent;
      if (looksCommessa) current = addCommessa(value, global, path) || parent;
      if (looksPlant) addPlant(value, current, global, path);
      Object.entries(value).forEach(([k, child]) => walk(child, path ? `${path}.${k}` : k, current, global || /global/i.test(k), depth + 1));
    };

    ['commesse','projects','cantieri','impianti','plants','sites','global','globalCommesse','commesseGlobal','globalImpianti','impiantiGlobal'].forEach((k) => {
      try { walk(window[k], `window.${k}`, null, /global/i.test(k)); } catch (_) {}
    });
    try {
      for (let i = 0; i < localStorage.length; i += 1) {
        const k = localStorage.key(i) || '';
        if (!/(commess|impiant|cantier|plant|site|global)/i.test(k)) continue;
        try { walk(JSON.parse(localStorage.getItem(k)), `localStorage.${k}`, null, /global/i.test(k)); } catch (_) {}
      }
    } catch (_) {}

    return {
      commesse: [...new Map(commesse.map((x) => [`${x.global}:${x.id}`, x])).values()].sort((a,b) => Number(a.global)-Number(b.global) || a.name.localeCompare(b.name,'it')),
      plants: [...new Map(plants.map((x) => [`${x.global}:${x.commessaId}:${x.id}`, x])).values()]
    };
  }

  function byLabel(root, text) {
    const wanted = norm(text);
    const label = [...root.querySelectorAll('label')].find((x) => norm(x.querySelector('span')?.textContent || x.textContent).startsWith(wanted));
    return label?.querySelector('input,select,textarea') || null;
  }
  const commessaField = (root) => root.querySelector('[data-doc-commessa],select[name="commessaId"],input[name="commessaId"],select[name="commessa"],input[name="commessa"]') || byLabel(root, 'Commessa');
  const plantField = (root) => root.querySelector('[data-doc-plant],select[name="plantId"],input[name="plantId"],select[name="impianto"],input[name="impianto"]') || byLabel(root, 'Impianto');

  function chosenCommessa(data, field) {
    if (!field) return null;
    const option = field.tagName === 'SELECT' ? field.selectedOptions?.[0] : null;
    const global = /^GLOBAL\b/i.test(option?.textContent || '');
    const value = String(field.value || '');
    const label = String(option?.textContent || value).replace(/^GLOBAL\s*[—-]?\s*/i, '').replace(/\s+[—-]\s+\S+$/, '').trim();
    return data.commesse.find((x) => x.id === value && (!global || x.global)) || data.commesse.find((x) => norm(x.name) === norm(label) && (!global || x.global)) || null;
  }

  function setEmpty(root, name, value) {
    if (!value) return;
    const el = root.querySelector(`[name="${name}"]`);
    if (el && !String(el.value || '').trim()) el.value = value;
  }
  function fillPlant(root, item) {
    if (!item) return;
    setEmpty(root, 'plantSap', item.sap); setEmpty(root, 'idSap', item.sap);
    setEmpty(root, 'plantName', item.name); setEmpty(root, 'nomeImpianto', item.name);
    setEmpty(root, 'workLocation', [item.address,item.city].filter(Boolean).join(', '));
    setEmpty(root, 'city', item.city); setEmpty(root, 'comune', item.city);
    setEmpty(root, 'plantType', item.type); setEmpty(root, 'tipologiaImpianto', item.type);
    setEmpty(root, 'area', item.area); setEmpty(root, 'coordinates', item.coordinates);
  }

  function matchesFor(data, commessa) {
    return data.plants.filter((x) => {
      if (!commessa) return true;
      if (x.global !== commessa.global) return false;
      return x.commessaId === commessa.id || norm(x.commessaId) === norm(commessa.code) || norm(x.commessaName) === norm(commessa.name) || (x.path && commessa.path && x.path.startsWith(commessa.path)) || (!x.commessaId && !x.commessaName);
    });
  }

  function refreshCommesse(page, data) {
    const field = commessaField(page);
    if (!field || field.tagName !== 'SELECT' || !data.commesse.length) return;
    const previous = field.value;
    const placeholder = [...field.options].find((o) => !o.value)?.textContent || 'Seleziona commessa';
    const html = `<option value="">${esc(placeholder)}</option>` + data.commesse.map((x) => `<option value="${esc(x.id)}">${x.global ? 'GLOBAL — ' : ''}${esc(x.name)}${x.code ? ` — ${esc(x.code)}` : ''}</option>`).join('');
    const signature = data.commesse.map((x) => `${x.global}:${x.id}:${x.name}:${x.code}`).join('|');
    if (field.dataset.registrySignature === signature) return;
    field.innerHTML = html;
    field.dataset.registrySignature = signature;
    if ([...field.options].some((o) => o.value === previous)) field.value = previous;
  }

  function refreshPlants(page, data, force = false) {
    const field = plantField(page);
    if (!field) return;
    const commessa = chosenCommessa(data, commessaField(page));
    const items = matchesFor(data, commessa);

    if (field.tagName === 'SELECT') {
      const previous = field.value;
      const signature = items.map((x) => `${x.global}:${x.id}:${x.name}:${x.sap}`).join('|');
      if (field.dataset.registrySignature === signature) return;
      field.innerHTML = '<option value="">Seleziona impianto</option>' + items.map((x) => `<option value="${esc(x.id)}">${x.global ? 'GLOBAL — ' : ''}${esc(x.name)}${x.sap ? ` — ${esc(x.sap)}` : ''}</option>`).join('');
      field.dataset.registrySignature = signature;
      if (items.some((x) => x.id === previous)) field.value = previous;
      return;
    }

    let list = field.list;
    if (!list) {
      const id = `pv-impianti-${Math.random().toString(36).slice(2)}`;
      list = document.createElement('datalist'); list.id = id; field.setAttribute('list', id); field.insertAdjacentElement('afterend', list);
    }
    const q = norm(field.value);
    const filtered = q && !force ? items.filter((x) => norm(`${x.name} ${x.sap} ${x.city} ${x.address}`).includes(q)) : items;
    const signature = filtered.map((x) => `${x.global}:${x.id}:${x.name}:${x.sap}`).join('|');
    if (list.dataset.registrySignature === signature) return;
    list.innerHTML = filtered.slice(0,300).map((x) => `<option value="${esc(x.name)}">${x.global ? 'GLOBAL — ' : ''}${x.sap ? `${esc(x.sap)} — ` : ''}${esc(x.city)}</option>`).join('');
    list.dataset.registrySignature = signature;
  }

  function enhance() {
    const page = document.getElementById('preventivi-page');
    if (!page || page.classList.contains('hidden')) return;
    const data = registry();
    refreshCommesse(page, data);
    refreshPlants(page, data, true);
  }

  document.addEventListener('change', (event) => {
    const page = event.target.closest?.('#preventivi-page');
    if (!page) return;
    const data = registry();
    if (event.target === commessaField(page)) {
      event.target.removeAttribute('data-registry-signature');
      const c = chosenCommessa(data, event.target);
      refreshPlants(page, data, true);
      if (c) {
        setEmpty(page, 'commessaCode', c.code); setEmpty(page, 'clientName', c.client);
        setEmpty(page, 'contractNumber', c.contract); setEmpty(page, 'requester', c.requester);
      }
    } else if (event.target === plantField(page)) {
      const c = chosenCommessa(data, commessaField(page));
      const item = data.plants.find((x) => x.id === event.target.value) || data.plants.find((x) => norm(x.name) === norm(event.target.value) && (!c || x.global === c.global));
      fillPlant(page, item);
    }
  }, true);

  document.addEventListener('input', (event) => {
    const page = event.target.closest?.('#preventivi-page');
    if (!page || event.target !== plantField(page)) return;
    refreshPlants(page, registry(), false);
  }, true);

  let timer = 0;
  const observer = new MutationObserver(() => {
    clearTimeout(timer);
    timer = setTimeout(enhance, 60);
  });
  const start = () => { observer.observe(document.body, { childList:true, subtree:true }); enhance(); };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once:true }); else start();
})();
