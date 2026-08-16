(() => {
  'use strict';

  if (window.HeraEmailCounterControlCenter?.installed) return;

  const REGION = 'europe-west1';
  const FUNCTION_NAME = 'getErrorEmailUsage';
  const CARD_ATTR = 'data-email-counter-card';
  const CACHE_MS = 60 * 1000;
  let lastLoadedAt = 0;
  let lastData = null;
  let loading = null;

  function root() {
    return document.getElementById('control-center-content') || document.getElementById('control-center-page');
  }

  function callable() {
    if (!window.firebase?.apps?.length || !window.firebase?.functions) return null;
    try {
      return window.firebase.app().functions(REGION).httpsCallable(FUNCTION_NAME);
    } catch (_) {
      return null;
    }
  }

  function formatMonth(value) {
    const match = String(value || '').match(/^(\d{4})-(\d{2})$/);
    if (!match) return value || 'Mese corrente';
    const date = new Date(Number(match[1]), Number(match[2]) - 1, 1);
    return new Intl.DateTimeFormat('it-IT', { month: 'long', year: 'numeric' }).format(date);
  }

  function formatUpdatedAt(value) {
    if (!value) return 'Non ancora aggiornato';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return 'Non disponibile';
    return date.toLocaleString('it-IT');
  }

  function createCard() {
    const card = document.createElement('details');
    card.className = 'control-card';
    card.setAttribute(CARD_ATTR, '1');
    card.innerHTML = `
      <summary>📧 Contatore e-mail inviati</summary>
      <div style="display:grid;gap:10px;padding:10px 2px 2px">
        <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px">
          <div style="padding:12px;border-radius:13px;background:#f8fafc;border:1px solid #dbe4ed">
            <small style="display:block;color:#64748b;font-weight:700">INVIATE QUESTO MESE</small>
            <strong data-email-counter-sent style="display:block;margin-top:3px;font-size:1.45rem">—</strong>
          </div>
          <div style="padding:12px;border-radius:13px;background:#f8fafc;border:1px solid #dbe4ed">
            <small style="display:block;color:#64748b;font-weight:700">RIMANENTI</small>
            <strong data-email-counter-remaining style="display:block;margin-top:3px;font-size:1.45rem">—</strong>
          </div>
        </div>
        <div>
          <div style="display:flex;justify-content:space-between;gap:8px;font-size:.78rem;font-weight:800;color:#475569">
            <span data-email-counter-month>Mese corrente</span>
            <span data-email-counter-limit>Limite 2.500</span>
          </div>
          <div style="height:10px;margin-top:6px;border-radius:999px;background:#e2e8f0;overflow:hidden">
            <div data-email-counter-progress style="width:0%;height:100%;background:#2563eb;transition:width .2s ease"></div>
          </div>
          <small data-email-counter-percent style="display:block;margin-top:5px;color:#64748b">0% del limite mensile</small>
        </div>
        <div class="control-center-row"><span>Ultimo aggiornamento</span><strong data-email-counter-updated>—</strong></div>
        <div data-email-counter-status style="font-size:.78rem;color:#64748b">Apri questa voce per caricare il contatore.</div>
        <button type="button" data-email-counter-refresh style="min-height:42px;border:1px solid #cbd5e1;border-radius:12px;background:#fff;font-weight:900">AGGIORNA CONTATORE</button>
      </div>`;
    return card;
  }

  function render(card, data) {
    if (!card || !data) return;
    card.querySelector('[data-email-counter-sent]').textContent = Number(data.sent || 0).toLocaleString('it-IT');
    card.querySelector('[data-email-counter-remaining]').textContent = Number(data.remaining || 0).toLocaleString('it-IT');
    card.querySelector('[data-email-counter-month]').textContent = formatMonth(data.month);
    card.querySelector('[data-email-counter-limit]').textContent = `Limite ${Number(data.limit || 2500).toLocaleString('it-IT')}`;
    const percentage = Math.max(0, Math.min(100, Number(data.percentage || 0)));
    card.querySelector('[data-email-counter-progress]').style.width = `${percentage}%`;
    card.querySelector('[data-email-counter-percent]').textContent = `${percentage.toLocaleString('it-IT')}% del limite mensile`;
    card.querySelector('[data-email-counter-updated]').textContent = formatUpdatedAt(data.updatedAt);
    card.querySelector('[data-email-counter-status]').textContent = `Contatore mensile aggiornato. ${Number(data.remaining || 0).toLocaleString('it-IT')} e-mail ancora disponibili.`;
  }

  function renderError(card, message) {
    const status = card?.querySelector('[data-email-counter-status]');
    if (status) status.textContent = message || 'Contatore non disponibile.';
  }

  async function load(card, force = false) {
    if (!card) return;
    const now = Date.now();
    if (!force && lastData && now - lastLoadedAt < CACHE_MS) {
      render(card, lastData);
      return;
    }
    if (loading) return loading;
    const invoke = callable();
    if (!invoke) {
      renderError(card, 'Servizio contatore non disponibile in questo momento.');
      return;
    }

    const status = card.querySelector('[data-email-counter-status]');
    const button = card.querySelector('[data-email-counter-refresh]');
    if (status) status.textContent = 'Caricamento contatore…';
    if (button) button.disabled = true;

    loading = invoke({}).then((result) => {
      lastData = result?.data || null;
      lastLoadedAt = Date.now();
      if (lastData) render(card, lastData);
      else renderError(card, 'Nessun dato disponibile.');
    }).catch((error) => {
      const code = String(error?.code || '');
      renderError(card, code.includes('permission-denied') ? 'Contatore riservato all’amministratore.' : 'Impossibile leggere il contatore e-mail.');
    }).finally(() => {
      loading = null;
      if (button) button.disabled = false;
    });
    return loading;
  }

  function install() {
    const container = root();
    if (!container || container.querySelector(`[${CARD_ATTR}]`)) return Boolean(container);
    const card = createCard();
    container.appendChild(card);
    card.addEventListener('toggle', () => {
      if (card.open) void load(card, false);
    });
    card.querySelector('[data-email-counter-refresh]')?.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      void load(card, true);
    });
    return true;
  }

  function boot() {
    if (install()) return;
    const observer = new MutationObserver(() => {
      if (!install()) return;
      observer.disconnect();
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot, { once: true });
  else boot();

  window.HeraEmailCounterControlCenter = Object.freeze({ installed: true, version: '1.0.0' });
})();
