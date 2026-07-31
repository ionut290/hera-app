(() => {
  'use strict';

  const PV = {
    version: '20260731a',
    menuId: 'open-preventivi-btn',
    pageId: 'preventivi-page',
    collections: {
      priceLists: 'preventiviPriceLists',
      quotes: 'preventiviQuotes'
    },
    keys: {
      priceLists: 'hera_preventivi_prezziari_v1',
      quotes: 'hera_preventivi_documenti_v1',
      deletions: 'hera_preventivi_eliminati_v1',
      settings: 'hera_preventivi_impostazioni_v1'
    },
    state: {
      view: 'quotes',
      priceLists: [],
      quotes: [],
      deletions: { priceLists: {}, quotes: {} },
      settings: {},
      editingQuoteId: '',
      editingPriceListId: '',
      quoteSearch: '',
      priceListSearch: '',
      firestore: null,
      remoteConnected: false,
      remoteDenied: false,
      unsubscribers: [],
      syncTimer: null
    }
  };

  const escapeMap = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };

  PV.escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => escapeMap[char]);
  PV.uid = (prefix) => `${prefix}-${window.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`}`;
  PV.nowIso = () => new Date().toISOString();
  PV.todayIso = () => {
    const now = new Date();
    return new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  };
  PV.normalizeText = (value) => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  PV.parseNumber = (value) => {
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    let text = String(value ?? '').trim().replace(/\s/g, '').replace(/[^0-9,.-]/g, '');
    if (!text) return 0;
    const comma = text.lastIndexOf(',');
    const dot = text.lastIndexOf('.');
    if (comma > dot) text = text.replace(/\./g, '').replace(',', '.');
    else if (dot > comma && comma >= 0) text = text.replace(/,/g, '');
    else if (comma >= 0) text = text.replace(',', '.');
    const parsed = Number(text);
    return Number.isFinite(parsed) ? parsed : 0;
  };
  PV.roundMoney = (value) => Math.round((Number(value) + Number.EPSILON) * 100) / 100;
  PV.formatMoney = (value) => new Intl.NumberFormat('it-IT', {
    style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2
  }).format(Number(value) || 0);
  PV.formatDate = (value) => {
    if (!value) return '—';
    const date = new Date(`${value}T12:00:00`);
    return Number.isNaN(date.getTime()) ? PV.escapeHtml(value) : new Intl.DateTimeFormat('it-IT').format(date);
  };
  PV.currentUser = () => {
    const user = window.firebase?.auth?.()?.currentUser;
    return { uid: user?.uid || '', email: user?.email || '', displayName: user?.displayName || user?.email || 'Utente' };
  };
  PV.stripUndefined = (value) => {
    if (Array.isArray(value)) return value.map(PV.stripUndefined);
    if (!value || typeof value !== 'object') return value;
    return Object.fromEntries(Object.entries(value)
      .filter(([, item]) => item !== undefined)
      .map(([key, item]) => [key, PV.stripUndefined(item)]));
  };

  PV.readJson = (key, fallback) => {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; }
    catch (error) { console.warn(`Preventivi: dati locali non leggibili (${key}).`, error); return fallback; }
  };
  PV.writeJson = (key, value) => {
    try { localStorage.setItem(key, JSON.stringify(value)); return true; }
    catch (error) { console.warn(`Preventivi: salvataggio locale non riuscito (${key}).`, error); return false; }
  };
  PV.loadLocal = () => {
    const priceLists = PV.readJson(PV.keys.priceLists, []);
    const quotes = PV.readJson(PV.keys.quotes, []);
    const deletions = PV.readJson(PV.keys.deletions, { priceLists: {}, quotes: {} });
    const settings = PV.readJson(PV.keys.settings, {});
    PV.state.priceLists = Array.isArray(priceLists) ? priceLists : [];
    PV.state.quotes = Array.isArray(quotes) ? quotes : [];
    PV.state.deletions = {
      priceLists: deletions?.priceLists && typeof deletions.priceLists === 'object' ? deletions.priceLists : {},
      quotes: deletions?.quotes && typeof deletions.quotes === 'object' ? deletions.quotes : {}
    };
    PV.state.settings = settings && typeof settings === 'object' ? settings : {};
  };
  PV.persistLocal = () => {
    PV.writeJson(PV.keys.priceLists, PV.state.priceLists);
    PV.writeJson(PV.keys.quotes, PV.state.quotes);
    PV.writeJson(PV.keys.deletions, PV.state.deletions);
    PV.writeJson(PV.keys.settings, PV.state.settings);
  };

  PV.page = () => document.getElementById(PV.pageId);
  PV.content = () => PV.page()?.querySelector('[data-pv-content]') || null;
  PV.setFeedback = (message = '', type = '') => {
    const element = PV.page()?.querySelector('[data-pv-feedback]');
    if (!element) return;
    element.textContent = message;
    element.dataset.type = type;
  };
  PV.setSyncBadge = (message, type = '') => {
    const badge = PV.page()?.querySelector('[data-pv-sync]');
    if (!badge) return;
    badge.textContent = message;
    badge.dataset.state = type;
  };
  PV.getPriceList = (id) => PV.state.priceLists.find((item) => item.id === id) || null;
  PV.getQuote = (id) => PV.state.quotes.find((item) => item.id === id) || null;
  PV.quoteTotals = (quote) => {
    const subtotal = PV.roundMoney((quote.lines || []).reduce((sum, line) => sum + PV.roundMoney((Number(line.quantity) || 0) * (Number(line.unitPrice) || 0)), 0));
    const vatRate = Math.max(0, Number(quote.vatRate) || 0);
    const vatAmount = PV.roundMoney(subtotal * vatRate / 100);
    return { subtotal, vatRate, vatAmount, total: PV.roundMoney(subtotal + vatAmount) };
  };
  PV.nextQuoteNumber = () => {
    const year = new Date().getFullYear();
    const numbers = PV.state.quotes
      .map((quote) => String(quote.number || ''))
      .filter((number) => number.includes(String(year)))
      .map((number) => Number(number.match(/(\d+)(?!.*\d)/)?.[1] || 0));
    return `PREV-${year}-${String(Math.max(0, ...numbers) + 1).padStart(3, '0')}`;
  };

  PV.ensureMenuButton = () => {
    if (document.getElementById(PV.menuId)) return;
    const section = document.getElementById('menu-operativo-title')?.closest('.menu-section');
    if (!section) return;
    const button = document.createElement('button');
    button.id = PV.menuId;
    button.className = 'btn menu-title-btn';
    button.type = 'button';
    button.innerHTML = '<span class="menu-item-icon" aria-hidden="true">🧾</span>Preventivi';
    const anchor = document.getElementById('open-segnalazioni-btn');
    if (anchor?.parentElement === section) anchor.insertAdjacentElement('afterend', button);
    else section.appendChild(button);
  };

  PV.ensurePage = () => {
    if (PV.page()) return;
    const page = document.createElement('section');
    page.id = PV.pageId;
    page.className = 'pv-page hidden';
    page.setAttribute('aria-hidden', 'true');
    page.innerHTML = `
      <header class="pv-header">
        <button type="button" class="pv-back" data-pv-close>← Home</button>
        <div class="pv-header-title"><h1>🧾 Preventivi</h1><p>Prezziari, lavorazioni e calcoli automatici</p></div>
        <span class="pv-sync-badge" data-pv-sync>💾 Dati sul dispositivo</span>
      </header>
      <nav class="pv-nav" aria-label="Sezioni Preventivi">
        <button type="button" class="pv-tab is-active" data-pv-view="quotes">Preventivi</button>
        <button type="button" class="pv-tab" data-pv-view="priceLists">Prezziari</button>
      </nav>
      <div class="pv-content" data-pv-content></div>`;
    document.body.appendChild(page);
  };

  PV.open = () => {
    PV.ensureMenuButton();
    PV.ensurePage();
    PV.page().classList.remove('hidden');
    PV.page().setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';
    document.getElementById('side-menu')?.classList.add('hidden');
    document.getElementById('side-menu')?.setAttribute('aria-hidden', 'true');
    document.getElementById('menu-overlay')?.classList.add('hidden');
    PV.state.editingQuoteId = '';
    PV.state.editingPriceListId = '';
    PV.renderCurrentView();
  };
  PV.close = () => {
    PV.page()?.classList.add('hidden');
    PV.page()?.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
  };
  PV.renderCurrentView = () => {
    const page = PV.page();
    if (!page || page.classList.contains('hidden')) return;
    page.querySelectorAll('[data-pv-view]').forEach((button) => button.classList.toggle('is-active', button.dataset.pvView === PV.state.view));
    if (PV.state.editingQuoteId) PV.renderQuoteEditor?.(PV.state.editingQuoteId === 'new' ? null : PV.getQuote(PV.state.editingQuoteId));
    else if (PV.state.editingPriceListId) PV.renderPriceListEditor?.(PV.state.editingPriceListId === 'new' ? null : PV.getPriceList(PV.state.editingPriceListId));
    else if (PV.state.view === 'priceLists') PV.renderPriceListOverview?.();
    else PV.renderQuoteOverview?.();
  };

  PV.saveRemote = async (collectionName, record) => {
    if (!PV.state.firestore || !record?.id) return false;
    try {
      await PV.state.firestore.collection(collectionName).doc(record.id).set(PV.stripUndefined(record), { merge: true });
      PV.state.remoteConnected = true;
      PV.state.remoteDenied = false;
      PV.setSyncBadge('☁️ Sincronizzato', 'ok');
      return true;
    } catch (error) {
      if (error?.code === 'permission-denied') PV.state.remoteDenied = true;
      console.warn(`Preventivi: scrittura Firebase non riuscita (${collectionName}).`, error);
      PV.setSyncBadge('💾 Salvato sul dispositivo', 'warning');
      return false;
    }
  };
  PV.deleteRemote = async (collectionName, id) => {
    if (!PV.state.firestore || !id) return false;
    try { await PV.state.firestore.collection(collectionName).doc(id).delete(); return true; }
    catch (error) {
      if (error?.code === 'permission-denied') PV.state.remoteDenied = true;
      console.warn(`Preventivi: eliminazione Firebase non riuscita (${collectionName}).`, error);
      return false;
    }
  };
  PV.scheduleSync = () => {
    window.clearTimeout(PV.state.syncTimer);
    PV.state.syncTimer = window.setTimeout(() => PV.syncPending().catch((error) => console.warn('Preventivi: sincronizzazione differita non riuscita.', error)), 350);
  };
  PV.syncPending = async () => {
    if (!PV.state.firestore) return;
    for (const item of PV.state.priceLists.filter((entry) => entry.syncPending)) {
      if (await PV.saveRemote(PV.collections.priceLists, { ...item, syncPending: false })) item.syncPending = false;
    }
    for (const item of PV.state.quotes.filter((entry) => entry.syncPending)) {
      if (await PV.saveRemote(PV.collections.quotes, { ...item, syncPending: false })) item.syncPending = false;
    }
    for (const id of Object.keys(PV.state.deletions.priceLists)) {
      if (await PV.deleteRemote(PV.collections.priceLists, id)) delete PV.state.deletions.priceLists[id];
    }
    for (const id of Object.keys(PV.state.deletions.quotes)) {
      if (await PV.deleteRemote(PV.collections.quotes, id)) delete PV.state.deletions.quotes[id];
    }
    PV.persistLocal();
  };
  PV.mergeRemote = (localRecords, remoteRecords, deletionMap) => {
    const merged = new Map(localRecords.map((record) => [record.id, record]));
    remoteRecords.forEach((remote) => {
      if (!remote?.id || deletionMap[remote.id]) return;
      const local = merged.get(remote.id);
      const localTime = Date.parse(local?.updatedAt || '') || 0;
      const remoteTime = Date.parse(remote?.updatedAt || '') || 0;
      if (!local || (!local.syncPending && remoteTime >= localTime)) merged.set(remote.id, { ...remote, syncPending: false });
    });
    return [...merged.values()].filter((record) => !deletionMap[record.id]);
  };
  PV.subscribeCollection = (collectionName, stateKey, deletionKey) => {
    const unsubscribe = PV.state.firestore.collection(collectionName).onSnapshot((snapshot) => {
      const remote = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      PV.state[stateKey] = PV.mergeRemote(PV.state[stateKey], remote, PV.state.deletions[deletionKey]);
      PV.state.remoteConnected = true;
      PV.state.remoteDenied = false;
      PV.persistLocal();
      PV.setSyncBadge('☁️ Sincronizzato', 'ok');
      if (!PV.state.editingQuoteId && !PV.state.editingPriceListId) PV.renderCurrentView();
      PV.scheduleSync();
    }, (error) => {
      if (error?.code === 'permission-denied') PV.state.remoteDenied = true;
      console.warn(`Preventivi: lettura Firebase non riuscita (${collectionName}).`, error);
      PV.setSyncBadge('💾 Modalità dispositivo', 'warning');
    });
    PV.state.unsubscribers.push(unsubscribe);
  };
  PV.connectFirebase = () => {
    try {
      if (!window.firebase?.firestore) { PV.setSyncBadge('💾 Modalità dispositivo', 'warning'); return; }
      PV.state.firestore = window.firebase.firestore();
      PV.subscribeCollection(PV.collections.priceLists, 'priceLists', 'priceLists');
      PV.subscribeCollection(PV.collections.quotes, 'quotes', 'quotes');
      PV.scheduleSync();
    } catch (error) {
      console.warn('Preventivi: Firebase non disponibile.', error);
      PV.state.firestore = null;
      PV.setSyncBadge('💾 Modalità dispositivo', 'warning');
    }
  };

  window.HeraPreventivi = PV;
})();
