(() => {
  'use strict';

  if (window.__heraUpdateCenterInstalled) return;
  window.__heraUpdateCenterInstalled = true;

  const escapeHtml = (value) => String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');

  function isAdmin() {
    try {
      return typeof window.canManageData === 'function' && window.canManageData();
    } catch (_) {
      return false;
    }
  }

  function installStyles() {
    if (document.getElementById('update-center-styles')) return;
    const style = document.createElement('style');
    style.id = 'update-center-styles';
    style.textContent = `
      .update-center-page{position:fixed;inset:0;z-index:15000;overflow:auto;background:#f1f5f9;padding:clamp(12px,3vw,28px)}
      .update-center-shell{width:min(1100px,100%);margin:0 auto;display:grid;gap:16px}
      .update-center-head{display:grid;grid-template-columns:auto 1fr auto;gap:16px;align-items:center;padding:20px;border-radius:20px;background:linear-gradient(135deg,#eff6ff,#fff);box-shadow:0 12px 35px rgba(15,23,42,.1)}
      .update-center-head h1{margin:2px 0;font-size:clamp(1.7rem,4vw,2.8rem)}
      .update-center-kicker{margin:0;color:#1d4ed8;font-weight:900;letter-spacing:.05em}
      .update-center-summary{padding:16px;border:1px solid #bfdbfe;border-radius:16px;background:#eff6ff;color:#0f172a}
      .update-center-summary.is-urgent{border-color:#fca5a5;background:#fef2f2;color:#991b1b}
      .update-center-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:14px}
      .update-center-card{overflow:hidden;border-radius:18px;background:#fff;box-shadow:0 8px 24px rgba(15,23,42,.08)}
      .update-center-card h2{display:flex;align-items:center;gap:9px;margin:0;padding:16px;border-bottom:1px solid #e2e8f0;font-size:1rem}
      .update-center-dot{width:12px;height:12px;border-radius:50%;background:#94a3b8}.update-center-dot.ok{background:#22c55e}.update-center-dot.planned{background:#facc15}.update-center-dot.urgent{background:#ef4444}
      .update-center-row{display:grid;grid-template-columns:1fr 1.15fr;gap:10px;padding:10px 16px;border-bottom:1px solid #e2e8f0}.update-center-row span{color:#64748b}.update-center-row strong{text-align:right;overflow-wrap:anywhere}
      @media(max-width:720px){.update-center-head{grid-template-columns:1fr}.update-center-row{grid-template-columns:1fr}.update-center-row strong{text-align:left}}
    `;
    document.head.appendChild(style);
  }

  function installSurface() {
    if (document.getElementById('update-center-page')) return;
    const page = document.createElement('section');
    page.id = 'update-center-page';
    page.className = 'update-center-page hidden';
    page.setAttribute('aria-live', 'polite');
    page.innerHTML = `
      <div class="update-center-shell">
        <header class="update-center-head">
          <button id="update-center-back" class="btn" type="button">← Home</button>
          <div><p class="update-center-kicker">🔄 SOLO AMMINISTRATORI</p><h1>Centro aggiornamenti</h1><p class="muted">Dipendenze, sicurezza e scadenze tecniche</p></div>
          <button id="update-center-check" class="btn btn-primary" type="button">CONTROLLA ORA</button>
        </header>
        <div id="update-center-summary" class="update-center-summary">Controllo aggiornamenti in attesa.</div>
        <div id="update-center-grid" class="update-center-grid"></div>
      </div>`;
    document.body.appendChild(page);
    document.getElementById('update-center-back')?.addEventListener('click', closeCenter);
    document.getElementById('update-center-check')?.addEventListener('click', checkUpdates);
  }

  function installMenuButton() {
    const anchor = document.getElementById('open-control-center-btn');
    if (!anchor || document.getElementById('open-update-center-btn')) return;
    const button = document.createElement('button');
    button.id = 'open-update-center-btn';
    button.className = 'btn menu-title-btn hidden';
    button.type = 'button';
    button.innerHTML = '<span class="menu-item-icon" aria-hidden="true">🔄</span>Centro aggiornamenti <span id="update-center-menu-badge" class="pending-users-badge hidden">0</span>';
    button.addEventListener('click', openCenter);
    anchor.insertAdjacentElement('afterend', button);
  }

  function syncAdminVisibility() {
    document.getElementById('open-update-center-btn')?.classList.toggle('hidden', !isAdmin());
    if (!isAdmin()) closeCenter();
  }

  function openCenter() {
    if (!isAdmin()) return;
    document.getElementById('side-menu')?.classList.add('hidden');
    document.getElementById('menu-overlay')?.classList.add('hidden');
    document.getElementById('update-center-page')?.classList.remove('hidden');
    void checkUpdates();
  }

  function closeCenter() {
    document.getElementById('update-center-page')?.classList.add('hidden');
  }

  function renderReport(report) {
    const items = Array.isArray(report?.items) ? report.items : [];
    const attention = items.filter((item) => item.status !== 'ok').length;
    const urgent = items.filter((item) => item.status === 'urgent').length;
    const labels = { ok: 'Aggiornato', planned: 'Da pianificare', urgent: 'Urgente' };
    const grid = document.getElementById('update-center-grid');
    const summary = document.getElementById('update-center-summary');
    if (grid) grid.innerHTML = items.map((item) => `
      <article class="update-center-card">
        <h2><span class="update-center-dot ${escapeHtml(item.status)}"></span>${escapeHtml(labels[item.status] || 'Informazione')} · ${escapeHtml(item.name)}</h2>
        <div class="update-center-row"><span>Versione installata</span><strong>${escapeHtml(item.current || 'Non applicabile')}</strong></div>
        <div class="update-center-row"><span>Versione disponibile</span><strong>${escapeHtml(item.latest || 'Non applicabile')}</strong></div>
        <div class="update-center-row"><span>Scadenza</span><strong>${escapeHtml(item.deadline || 'Nessuna')}</strong></div>
        <div class="update-center-row"><span>Indicazione</span><strong>${escapeHtml(item.message || 'Nessun intervento necessario')}</strong></div>
      </article>`).join('');
    if (summary) {
      const checkedAt = report?.checkedAt ? new Date(report.checkedAt).toLocaleString('it-IT') : 'adesso';
      summary.innerHTML = `<strong>${urgent ? 'Aggiornamenti urgenti presenti' : attention ? 'Aggiornamenti da pianificare' : 'Tutto aggiornato'}</strong><br>Ultimo controllo: ${escapeHtml(checkedAt)}`;
      summary.classList.toggle('is-urgent', urgent > 0);
    }
    const badge = document.getElementById('update-center-menu-badge');
    if (badge) {
      badge.textContent = String(attention);
      badge.classList.toggle('hidden', attention === 0);
    }
  }

  async function checkUpdates() {
    if (!isAdmin()) return;
    const summary = document.getElementById('update-center-summary');
    const button = document.getElementById('update-center-check');
    if (summary) summary.textContent = 'Controllo aggiornamenti in corso…';
    button?.setAttribute('disabled', '');
    try {
      const response = await fetch('/.netlify/functions/dependency-updates', { cache: 'no-store', headers: { Accept: 'application/json' } });
      if (!response.ok) throw new Error(`Servizio non disponibile (${response.status})`);
      renderReport(await response.json());
    } catch (error) {
      if (summary) summary.innerHTML = `<strong>Controllo online non riuscito</strong><br>${escapeHtml(error?.message || 'Riprova più tardi')}`;
    } finally {
      button?.removeAttribute('disabled');
    }
  }

  installStyles();
  installSurface();
  installMenuButton();
  syncAdminVisibility();
  const sideMenu = document.getElementById('side-menu');
  if (sideMenu) new MutationObserver(syncAdminVisibility).observe(sideMenu, { attributes: true, attributeFilter: ['class'] });
  try {
    window.firebase?.auth?.().onAuthStateChanged(syncAdminVisibility);
  } catch (_) {
    // Il menu resta nascosto finché il ruolo amministratore non è verificabile.
  }
})();
