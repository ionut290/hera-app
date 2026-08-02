(() => {
  'use strict';

  if (window.__heraFirestoreUsageControlInstalled) return;
  window.__heraFirestoreUsageControlInstalled = true;

  const ENDPOINT = '/.netlify/functions/firestore-usage';
  const CARD_ID = 'firestore-usage-control-card';
  const CACHE_KEY = 'hera_firestore_usage_cache_v1';
  const CACHE_MS = 15 * 60 * 1000;
  let loading = false;

  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[char]));
  }

  function number(value) {
    return new Intl.NumberFormat('it-IT').format(Math.max(0, Number(value || 0)));
  }

  function dateTime(value) {
    const date = new Date(value || 0);
    return Number.isNaN(date.getTime()) ? 'Non disponibile' : date.toLocaleString('it-IT', { dateStyle: 'short', timeStyle: 'short' });
  }

  function level(percent) {
    if (percent >= 90) return { label: 'Vicino al limite', tone: 'danger' };
    if (percent >= 80) return { label: 'Consumo elevato', tone: 'orange' };
    if (percent >= 60) return { label: 'Attenzione', tone: 'warning' };
    return { label: 'Consumo regolare', tone: 'safe' };
  }

  function cachedValue() {
    try {
      const parsed = JSON.parse(sessionStorage.getItem(CACHE_KEY) || 'null');
      if (parsed?.savedAt && Date.now() - parsed.savedAt < CACHE_MS && parsed.data?.ok) return parsed.data;
    } catch (_) { /* cache facoltativa */ }
    return null;
  }

  function saveCache(data) {
    try { sessionStorage.setItem(CACHE_KEY, JSON.stringify({ savedAt: Date.now(), data })); } catch (_) { /* cache facoltativa */ }
  }

  function ensureStyle() {
    if (document.getElementById('firestore-usage-control-style')) return;
    const style = document.createElement('style');
    style.id = 'firestore-usage-control-style';
    style.textContent = `
      #${CARD_ID}{grid-column:1/-1}.fs-usage-head{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}.fs-usage-title{display:flex;align-items:center;gap:8px;font-weight:800}.fs-usage-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin-top:14px}.fs-usage-item{border:1px solid rgba(148,163,184,.35);border-radius:14px;padding:14px;background:rgba(255,255,255,.72)}.fs-usage-label{font-weight:800;margin-bottom:6px}.fs-usage-value{font-size:1.25rem;font-weight:900}.fs-usage-meta{font-size:.82rem;opacity:.75;margin-top:5px}.fs-usage-bar{height:10px;border-radius:999px;background:#e5e7eb;overflow:hidden;margin-top:10px}.fs-usage-fill{height:100%;border-radius:inherit;background:#22c55e}.fs-usage-fill.warning{background:#eab308}.fs-usage-fill.orange{background:#f97316}.fs-usage-fill.danger{background:#dc2626}.fs-usage-status{font-weight:800;margin-top:8px}.fs-usage-error{padding:12px;border-radius:12px;background:#fff7ed;border:1px solid #fdba74;color:#9a3412;margin-top:12px}.fs-usage-note{font-size:.8rem;opacity:.72;margin-top:12px}.fs-usage-refresh{min-width:130px}@media(max-width:720px){.fs-usage-grid{grid-template-columns:1fr}.fs-usage-item{padding:12px}}
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

  function scheduleLoad() {
    if (location.hash !== '#centro-controllo') return;
    window.setTimeout(() => load(false), 0);
    window.setTimeout(() => load(false), 300);
  }

  ensureStyle();
  window.addEventListener('hashchange', scheduleLoad);
  document.addEventListener('click', (event) => {
    if (event.target.closest('#open-control-center-btn')) window.setTimeout(scheduleLoad, 50);
  }, true);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', scheduleLoad, { once: true });
  else scheduleLoad();
})();
