(() => {
  'use strict';

  if (window.__heraFirestoreUsageControlInstalled) return;
  window.__heraFirestoreUsageControlInstalled = true;

  const ENDPOINT = '/.netlify/functions/firestore-usage';
  const API_ENDPOINT = '/.netlify/functions/api-usage';
  const CARD_ID = 'firestore-usage-control-card';
  const STREET_VIEW_CARD_ID = 'street-view-usage-control-card';
  const API_CARD_ID = 'api-usage-control-card';
  const CACHE_KEY = 'hera_firestore_usage_cache_v1';
  const API_CACHE_KEY = 'hera_api_usage_cache_v1';
  const CACHE_MS = 15 * 60 * 1000;
  const STREET_VIEW_LIMIT = 4800;
  let loading = false;
  let apiLoading = false;
  let streetViewUnsubscribe = null;

  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[char]));
  }

  function number(value) {
    return new Intl.NumberFormat('it-IT').format(Math.max(0, Number(value || 0)));
  }

  function dateTime(value) {
    if (value?.toDate) value = value.toDate();
    const date = new Date(value || 0);
    return Number.isNaN(date.getTime()) ? 'Non disponibile' : date.toLocaleString('it-IT', { dateStyle: 'short', timeStyle: 'short' });
  }

  function level(percent) {
    if (percent >= 90) return { label: 'Vicino al limite', tone: 'danger' };
    if (percent >= 80) return { label: 'Consumo elevato', tone: 'orange' };
    if (percent >= 60) return { label: 'Attenzione', tone: 'warning' };
    return { label: 'Consumo regolare', tone: 'safe' };
  }

  function streetViewLevel(used, limit) {
    if (used >= limit) return { label: '⛔ Limite raggiunto · Street View bloccato', tone: 'danger' };
    if (used >= 4500) return { label: '🔴 Quasi esaurito', tone: 'danger' };
    if (used >= 4000) return { label: '🟡 Attenzione', tone: 'warning' };
    return { label: '🟢 Consumo regolare', tone: 'safe' };
  }

  function cachedValue(key = CACHE_KEY) {
    try {
      const parsed = JSON.parse(sessionStorage.getItem(key) || 'null');
      if (parsed?.savedAt && Date.now() - parsed.savedAt < CACHE_MS && parsed.data?.ok) return parsed.data;
    } catch (_) { /* cache facoltativa */ }
    return null;
  }

  function saveCache(data, key = CACHE_KEY) {
    try { sessionStorage.setItem(key, JSON.stringify({ savedAt: Date.now(), data })); } catch (_) { /* cache facoltativa */ }
  }

  function ensureStyle() {
    if (document.getElementById('firestore-usage-control-style')) return;
    const style = document.createElement('style');
    style.id = 'firestore-usage-control-style';
    style.textContent = `
      #${CARD_ID},#${STREET_VIEW_CARD_ID},#${API_CARD_ID}{grid-column:1/-1}.fs-usage-head{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}.fs-usage-title{display:flex;align-items:center;gap:8px;font-weight:800}.fs-usage-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin-top:14px}.fs-usage-item{border:1px solid rgba(148,163,184,.35);border-radius:14px;padding:14px;background:rgba(255,255,255,.72)}.fs-usage-label{font-weight:800;margin-bottom:6px}.fs-usage-value{font-size:1.25rem;font-weight:900}.fs-usage-meta{font-size:.82rem;opacity:.75;margin-top:5px}.fs-usage-bar{height:10px;border-radius:999px;background:#e5e7eb;overflow:hidden;margin-top:10px}.fs-usage-fill{height:100%;border-radius:inherit;background:#22c55e}.fs-usage-fill.warning{background:#eab308}.fs-usage-fill.orange{background:#f97316}.fs-usage-fill.danger{background:#dc2626}.fs-usage-status{font-weight:800;margin-top:8px}.fs-usage-error{padding:12px;border-radius:12px;background:#fff7ed;border:1px solid #fdba74;color:#9a3412;margin-top:12px}.fs-usage-note{font-size:.8rem;opacity:.72;margin-top:12px}.fs-usage-refresh{min-width:130px}.sv-usage-main{margin-top:14px;border:1px solid rgba(148,163,184,.35);border-radius:14px;padding:14px;background:rgba(255,255,255,.72)}.sv-usage-value{font-size:1.55rem;font-weight:900}.sv-usage-meta{font-size:.92rem;opacity:.78;margin-top:5px}.sv-usage-details{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px 16px;margin-top:12px;font-size:.84rem;opacity:.8}.sv-usage-live{display:inline-flex;align-items:center;gap:6px;font-size:.78rem;font-weight:800;color:#15803d}.sv-usage-live::before{content:'';width:8px;height:8px;border-radius:50%;background:#22c55e;box-shadow:0 0 0 3px rgba(34,197,94,.14)}.api-usage-summary{margin-top:14px;padding:14px;border-radius:14px;background:rgba(37,99,235,.08);border:1px solid rgba(37,99,235,.18)}.api-usage-summary strong{font-size:1.35rem}.api-usage-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin-top:12px}.api-usage-row{border:1px solid rgba(148,163,184,.35);border-radius:13px;padding:12px;background:rgba(255,255,255,.72)}.api-usage-row-name{font-weight:850}.api-usage-row-count{font-size:1.18rem;font-weight:900;margin-top:5px}.api-usage-row-service{font-size:.72rem;opacity:.62;margin-top:3px;word-break:break-all}.api-external{margin-top:12px;padding-top:10px;border-top:1px dashed rgba(148,163,184,.45);font-size:.82rem;opacity:.8}@media(max-width:720px){.fs-usage-grid,.api-usage-list{grid-template-columns:1fr}.fs-usage-item{padding:12px}.sv-usage-details{grid-template-columns:1fr}}
    `;
    document.head.appendChild(style);
  }

  function cardHost() {
    return document.getElementById('control-center-content');
  }

  function ensureCard() {
    const host = cardHost();
    if (!host) return null;
    let card = document.getElementById(CARD_ID);
    if (!card) {
      card = document.createElement('section');
      card.id = CARD_ID;
      card.className = 'card control-center-card';
      host.prepend(card);
    }
    return card;
  }

  function ensureStreetViewCard() {
    const host = cardHost();
    if (!host) return null;
    let card = document.getElementById(STREET_VIEW_CARD_ID);
    if (!card) {
      card = document.createElement('section');
      card.id = STREET_VIEW_CARD_ID;
      card.className = 'card control-center-card';
      const firestoreCard = ensureCard();
      if (firestoreCard?.parentNode) firestoreCard.insertAdjacentElement('afterend', card);
      else host.prepend(card);
    }
    return card;
  }

  function ensureApiCard() {
    const host = cardHost();
    if (!host) return null;
    let card = document.getElementById(API_CARD_ID);
    if (!card) {
      card = document.createElement('section');
      card.id = API_CARD_ID;
      card.className = 'card control-center-card';
      const streetCard = ensureStreetViewCard();
      if (streetCard?.parentNode) streetCard.insertAdjacentElement('afterend', card);
      else host.appendChild(card);
    }
    return card;
  }

  function usageItem(label, used, limit, percent) {
    const status = level(percent);
    const width = Math.min(100, Math.max(0, percent));
    return `<article class="fs-usage-item"><div class="fs-usage-label">${escapeHtml(label)}</div><div class="fs-usage-value">${number(used)} / ${number(limit)}</div><div class="fs-usage-meta">${percent.toFixed(1).replace('.', ',')}% utilizzato • ${number(Math.max(0, limit - used))} rimanenti</div><div class="fs-usage-bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${width.toFixed(0)}"><div class="fs-usage-fill ${status.tone}" style="width:${width}%"></div></div><div class="fs-usage-status">${escapeHtml(status.label)}</div></article>`;
  }

  function render(data) {
    const card = ensureCard();
    if (!card) return;
    const usage = data.usage || {};
    const limits = data.limits || {};
    const percentages = data.percentages || {};
    card.innerHTML = `<div class="fs-usage-head"><div class="fs-usage-title"><span aria-hidden="true">📊</span><span>Consumo Firestore giornaliero</span></div><button class="btn fs-usage-refresh" type="button" data-firestore-usage-refresh>AGGIORNA</button></div><div class="fs-usage-grid">${usageItem('Letture', usage.reads, limits.reads, Number(percentages.reads || 0))}${usageItem('Scritture', usage.writes, limits.writes, Number(percentages.writes || 0))}${usageItem('Eliminazioni', usage.deletes, limits.deletes, Number(percentages.deletes || 0))}</div><div class="fs-usage-note">Fonte: ${escapeHtml(data.source || 'Google Cloud Monitoring')} • aggiornato ${escapeHtml(dateTime(data.generatedAt))}${data.cached ? ' • dati in cache' : ''}. Il conteggio segue il giorno Firestore, che riparte a mezzanotte del Pacifico.</div>`;
    card.querySelector('[data-firestore-usage-refresh]')?.addEventListener('click', () => load(true));
  }

  function renderLoading() {
    const card = ensureCard();
    if (!card) return;
    card.innerHTML = '<div class="fs-usage-head"><div class="fs-usage-title"><span aria-hidden="true">📊</span><span>Consumo Firestore giornaliero</span></div></div><p>Caricamento dati ufficiali…</p>';
  }

  function renderError(message) {
    const card = ensureCard();
    if (!card) return;
    card.innerHTML = `<div class="fs-usage-head"><div class="fs-usage-title"><span aria-hidden="true">📊</span><span>Consumo Firestore giornaliero</span></div><button class="btn fs-usage-refresh" type="button" data-firestore-usage-refresh>RIPROVA</button></div><div class="fs-usage-error">${escapeHtml(message)}</div><div class="fs-usage-note">Il controllo non aggiunge letture, scritture o listener Firestore.</div>`;
    card.querySelector('[data-firestore-usage-refresh]')?.addEventListener('click', () => load(true));
  }

  function currentMonthKey(date = new Date()) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  }

  function resolveFirestore() {
    try { if (typeof db !== 'undefined' && db?.collection) return db; } catch (_) {}
    try {
      if (window.firebase?.firestore && typeof window.firebase.firestore === 'function') return window.firebase.firestore();
    } catch (_) {}
    return null;
  }

  function renderStreetViewUsage(data = {}) {
    const card = ensureStreetViewCard();
    if (!card) return;
    const used = Math.max(0, Number(data.count || 0));
    const limit = Math.max(1, Number(data.limit || STREET_VIEW_LIMIT));
    const remaining = Math.max(0, limit - used);
    const percent = Math.min(100, Math.max(0, used / limit * 100));
    const status = streetViewLevel(used, limit);
    const lastUser = data.lastUserEmail || data.lastUserUid || 'Nessun utilizzo registrato';
    card.innerHTML = `<div class="fs-usage-head"><div class="fs-usage-title"><span aria-hidden="true">🌐</span><span>Consumo Street View 360° mensile</span></div><span class="sv-usage-live">LIVE</span></div><div class="sv-usage-main"><div class="sv-usage-value">${number(used)} / ${number(limit)}</div><div class="sv-usage-meta">${percent.toFixed(1).replace('.', ',')}% utilizzato • ${number(remaining)} rimanenti</div><div class="fs-usage-bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${percent.toFixed(0)}"><div class="fs-usage-fill ${status.tone}" style="width:${percent}%"></div></div><div class="fs-usage-status">${escapeHtml(status.label)}</div><div class="sv-usage-details"><div><strong>Mese:</strong> ${escapeHtml(data.month || currentMonthKey())}</div><div><strong>Ultimo utilizzo:</strong> ${escapeHtml(data.updatedAt ? dateTime(data.updatedAt) : 'Non ancora registrato')}</div><div><strong>Ultimo utente:</strong> ${escapeHtml(lastUser)}</div><div><strong>Blocco automatico:</strong> ${used >= limit ? 'Attivo' : `a ${number(limit)}`}</div></div></div><div class="fs-usage-note">Dato condiviso Firestore: si aggiorna automaticamente quando Street View viene aperto da qualsiasi dispositivo.</div>`;
  }

  function renderStreetViewError(message) {
    const card = ensureStreetViewCard();
    if (!card) return;
    card.innerHTML = `<div class="fs-usage-head"><div class="fs-usage-title"><span aria-hidden="true">🌐</span><span>Consumo Street View 360° mensile</span></div></div><div class="fs-usage-error">${escapeHtml(message)}</div>`;
  }

  function renderApiUsage(data = {}) {
    const card = ensureApiCard();
    if (!card) return;
    const services = Array.isArray(data.services) ? data.services : [];
    const rows = services.length
      ? services.map((item) => `<article class="api-usage-row"><div class="api-usage-row-name">${escapeHtml(item.name || item.service)}</div><div class="api-usage-row-count">${number(item.requests)} richieste</div><div class="api-usage-row-service">${escapeHtml(item.service)}</div></article>`).join('')
      : '<div class="fs-usage-note">Nessuna richiesta Google API rilevata nel periodo.</div>';
    const external = Array.isArray(data.externalProviders) && data.externalProviders.length
      ? `<div class="api-external"><strong>API esterne all Google Cloud:</strong> ${data.externalProviders.map((item) => `${escapeHtml(item.name)} (${escapeHtml(item.note || item.tracking || '')})`).join(' • ')}</div>`
      : '';
    card.innerHTML = `<div class="fs-usage-head"><div class="fs-usage-title"><span aria-hidden="true">🧩</span><span>Utilizzo API dell'app · mese corrente</span></div><button class="btn fs-usage-refresh" type="button" data-api-usage-refresh>AGGIORNA</button></div><div class="api-usage-summary"><div>Richieste Google/Firebase rilevate</div><strong>${number(data.totalRequests || 0)}</strong></div><div class="api-usage-list">${rows}</div>${external}<div class="fs-usage-note">Fonte: ${escapeHtml(data.source || 'Google Cloud Monitoring')} • aggiornato ${escapeHtml(dateTime(data.generatedAt))}${data.cached ? ' • dati in cache' : ''}. Le API vengono mostrate automaticamente quando registrano traffico nel progetto. Il blocco automatico resta attivo solo su Street View 360°, dove abbiamo un contatore condiviso affidabile a 4.800/mese.</div>`;
    card.querySelector('[data-api-usage-refresh]')?.addEventListener('click', () => loadApiUsage(true));
  }

  function renderApiLoading() {
    const card = ensureApiCard();
    if (!card) return;
    card.innerHTML = '<div class="fs-usage-head"><div class="fs-usage-title"><span aria-hidden="true">🧩</span><span>Utilizzo API dell\'app · mese corrente</span></div></div><p>Caricamento API utilizzate…</p>';
  }

  function renderApiError(message) {
    const card = ensureApiCard();
    if (!card) return;
    card.innerHTML = `<div class="fs-usage-head"><div class="fs-usage-title"><span aria-hidden="true">🧩</span><span>Utilizzo API dell'app · mese corrente</span></div><button class="btn fs-usage-refresh" type="button" data-api-usage-refresh>RIPROVA</button></div><div class="fs-usage-error">${escapeHtml(message)}</div>`;
    card.querySelector('[data-api-usage-refresh]')?.addEventListener('click', () => loadApiUsage(true));
  }

  function stopStreetViewListener() {
    if (typeof streetViewUnsubscribe === 'function') {
      try { streetViewUnsubscribe(); } catch (_) {}
    }
    streetViewUnsubscribe = null;
  }

  function startStreetViewListener() {
    if (location.hash !== '#centro-controllo' || streetViewUnsubscribe) return;
    const firestore = resolveFirestore();
    if (!firestore) {
      renderStreetViewError('Contatore Street View non disponibile: Firestore non inizializzato.');
      return;
    }
    const ref = firestore.collection('appConfig').doc(`streetViewUsage_${currentMonthKey()}`);
    try {
      streetViewUnsubscribe = ref.onSnapshot((snap) => {
        renderStreetViewUsage(snap.exists ? (snap.data() || {}) : { count: 0, limit: STREET_VIEW_LIMIT, month: currentMonthKey() });
      }, (error) => {
        console.warn('[CONTROL CENTER][STREET VIEW]', error);
        renderStreetViewError('Impossibile leggere in tempo reale il consumo Street View.');
      });
    } catch (error) {
      renderStreetViewError(error?.message || 'Impossibile attivare il monitoraggio Street View.');
    }
  }

  async function load(force = false) {
    if (loading || location.hash !== '#centro-controllo') return;
    if (!force) {
      const cached = cachedValue();
      if (cached) { render(cached); return; }
    }
    const user = window.firebase?.auth?.()?.currentUser;
    if (!user) { renderError('Accedi come amministratore per vedere il consumo Firestore.'); return; }
    loading = true;
    renderLoading();
    try {
      const token = await user.getIdToken();
      const response = await fetch(ENDPOINT, { method: 'GET', headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }, cache: 'no-store' });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.ok) throw new Error(data.error || `Servizio non disponibile (HTTP ${response.status}).`);
      saveCache(data);
      render(data);
    } catch (error) {
      renderError(error?.message || 'Impossibile recuperare il consumo Firestore.');
    } finally {
      loading = false;
    }
  }

  async function loadApiUsage(force = false) {
    if (apiLoading || location.hash !== '#centro-controllo') return;
    if (!force) {
      const cached = cachedValue(API_CACHE_KEY);
      if (cached) { renderApiUsage(cached); return; }
    }
    const user = window.firebase?.auth?.()?.currentUser;
    if (!user) { renderApiError('Accedi come amministratore per vedere il consumo API.'); return; }
    apiLoading = true;
    renderApiLoading();
    try {
      const token = await user.getIdToken();
      const response = await fetch(API_ENDPOINT, { method: 'GET', headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }, cache: 'no-store' });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.ok) throw new Error(data.error || `Servizio API non disponibile (HTTP ${response.status}).`);
      saveCache(data, API_CACHE_KEY);
      renderApiUsage(data);
    } catch (error) {
      renderApiError(error?.message || 'Impossibile recuperare il consumo delle API.');
    } finally {
      apiLoading = false;
    }
  }

  function scheduleLoad() {
    if (location.hash !== '#centro-controllo') {
      stopStreetViewListener();
      return;
    }
    window.setTimeout(() => load(false), 0);
    window.setTimeout(() => load(false), 300);
    window.setTimeout(() => loadApiUsage(false), 0);
    window.setTimeout(() => loadApiUsage(false), 300);
    window.setTimeout(startStreetViewListener, 0);
    window.setTimeout(startStreetViewListener, 300);
  }

  ensureStyle();
  window.addEventListener('hashchange', scheduleLoad);
  document.addEventListener('click', (event) => {
    if (event.target.closest('#open-control-center-btn')) window.setTimeout(scheduleLoad, 50);
  }, true);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', scheduleLoad, { once: true });
  else scheduleLoad();
})();
