(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  if (!PV) return;

  const RETENTION_DAYS = 30;
  const WARNING_DAYS = 5;
  const CHECK_INTERVAL_MS = 24 * 60 * 60 * 1000;
  const MIN_CHECK_GAP_MS = 12 * 60 * 60 * 1000;
  const LAST_CHECK_KEY = 'hera_preventivi_temporary_archive_last_check_v1';
  const DAY_MS = 24 * 60 * 60 * 1000;
  let cleanupRunning = false;
  let timer = null;

  const text = (value) => String(value ?? '').trim();
  const nowIso = () => new Date().toISOString();
  const addDaysIso = (value, days) => {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return new Date(Date.now() + days * DAY_MS).toISOString();
    return new Date(date.getTime() + days * DAY_MS).toISOString();
  };

  function collectionEntries() {
    const entries = [];
    if (PV.collections?.quotes && Array.isArray(PV.state?.quotes)) {
      entries.push({ type: 'preventivo', collection: PV.collections.quotes, stateKey: 'quotes', deletionKey: 'quotes' });
    }
    if (PV.collections?.consuntivi && Array.isArray(PV.state?.consuntivi)) {
      entries.push({ type: 'consuntivo', collection: PV.collections.consuntivi, stateKey: 'consuntivi', deletionKey: 'consuntivi' });
    }
    return entries;
  }

  function ensureRetention(record) {
    if (!record || typeof record !== 'object') return record;
    const base = text(record.temporaryStoredAt || record.createdAt || record.updatedAt) || nowIso();
    if (!record.temporaryStoredAt) record.temporaryStoredAt = base;
    if (!record.expiresAt) record.expiresAt = addDaysIso(base, RETENTION_DAYS);
    record.retentionDays = RETENTION_DAYS;
    record.archiveTemporary = true;
    return record;
  }

  function expiryTime(record) {
    return Date.parse(record?.expiresAt || '') || 0;
  }

  function daysRemaining(record) {
    const expires = expiryTime(record);
    if (!expires) return RETENTION_DAYS;
    return Math.max(0, Math.ceil((expires - Date.now()) / DAY_MS));
  }

  function isExpired(record) {
    const expires = expiryTime(record);
    return Boolean(expires && expires <= Date.now());
  }

  function driveSaved(record) {
    return Boolean(record?.driveSaved || record?.driveFileId || record?.driveUrl);
  }

  function applyRetentionToLocal() {
    let changed = false;
    collectionEntries().forEach(({ stateKey }) => {
      (PV.state[stateKey] || []).forEach((record) => {
        const before = `${record.temporaryStoredAt || ''}|${record.expiresAt || ''}`;
        ensureRetention(record);
        const after = `${record.temporaryStoredAt || ''}|${record.expiresAt || ''}`;
        if (before !== after) {
          record.syncPending = true;
          changed = true;
        }
      });
    });
    if (changed) {
      PV.persistLocal?.();
      PV.scheduleSync?.();
    }
  }

  async function deleteExpiredRecord(entry, record) {
    const records = PV.state[entry.stateKey] || [];
    PV.state[entry.stateKey] = records.filter((item) => item.id !== record.id);
    if (PV.state.deletions?.[entry.deletionKey]) {
      PV.state.deletions[entry.deletionKey][record.id] = nowIso();
    }
    PV.persistLocal?.();
    if (PV.deleteRemote && entry.collection) {
      await PV.deleteRemote(entry.collection, record.id);
    }
  }

  function lastCheckAt() {
    try { return Number(localStorage.getItem(LAST_CHECK_KEY) || 0); }
    catch (_) { return 0; }
  }

  function markChecked() {
    try { localStorage.setItem(LAST_CHECK_KEY, String(Date.now())); }
    catch (_) { /* La cache temporale è facoltativa. */ }
  }

  async function cleanupExpired({ force = false } = {}) {
    if (cleanupRunning) return false;
    if (!force && Date.now() - lastCheckAt() < MIN_CHECK_GAP_MS) return false;
    cleanupRunning = true;
    try {
      applyRetentionToLocal();
      for (const entry of collectionEntries()) {
        const expired = [...(PV.state[entry.stateKey] || [])].filter(isExpired);
        for (const record of expired) await deleteExpiredRecord(entry, record);
      }
      markChecked();
      if (PV.state?.isOpen) {
        decorateWarnings();
        PV.renderCurrentView?.();
      }
      return true;
    } catch (error) {
      console.warn('Archivio temporaneo: eliminazione documenti scaduti non riuscita.', error);
      return false;
    } finally {
      cleanupRunning = false;
    }
  }

  function bannerHtml() {
    return `<aside class="pv-retention-banner" data-pv-retention-banner role="status">
      <strong>Archivio temporaneo di 30 giorni</strong>
      <span>Preventivi e consuntivi vengono salvati nell’app anche senza Google Drive. Dopo 30 giorni vengono eliminati automaticamente dall’archivio dell’app. Collega Google Drive per conservarli definitivamente.</span>
    </aside>`;
  }

  function decorateWarnings() {
    const page = PV.page?.() || document.getElementById('preventivi-page');
    if (!page || !PV.state?.isOpen) return;
    const content = page.querySelector('[data-pv-content]');
    if (content && !content.querySelector('[data-pv-retention-banner]')) {
      content.insertAdjacentHTML('afterbegin', bannerHtml());
    }

    const all = [...(PV.state?.quotes || []), ...(PV.state?.consuntivi || [])];
    page.querySelectorAll('[data-pv-quote-card], [data-cons-card], [data-consuntivo-card]').forEach((card) => {
      const id = card.dataset.pvQuoteCard || card.dataset.consCard || card.dataset.consuntivoCard;
      const record = all.find((item) => item.id === id);
      if (!record) return;
      let badge = card.querySelector('[data-pv-retention-status]');
      if (!badge) {
        badge = document.createElement('p');
        badge.dataset.pvRetentionStatus = '1';
        badge.className = 'pv-retention-status';
        card.appendChild(badge);
      }
      const days = daysRemaining(record);
      if (driveSaved(record)) {
        badge.textContent = `✓ Copia salvata su Google Drive · eliminazione dall’app tra ${days} giorni`;
        badge.dataset.state = 'drive';
      } else if (days <= WARNING_DAYS) {
        badge.textContent = `⚠ Verrà eliminato dall’app tra ${days} ${days === 1 ? 'giorno' : 'giorni'}. Salvalo su Google Drive.`;
        badge.dataset.state = 'danger';
      } else {
        badge.textContent = `Documento temporaneo · ${days} giorni rimanenti`;
        badge.dataset.state = 'warning';
      }
    });
  }

  function installStyles() {
    if (document.querySelector('style[data-pv-retention-style]')) return;
    const style = document.createElement('style');
    style.dataset.pvRetentionStyle = '1';
    style.textContent = `
      .pv-retention-banner{display:grid;gap:5px;margin:0 0 14px;padding:12px 14px;border:1px solid #f59e0b;border-radius:12px;background:#fff7ed;color:#7c2d12}
      .pv-retention-banner strong{font-size:1rem}.pv-retention-banner span{line-height:1.4}
      .pv-retention-status{margin:10px 0 0;padding:8px 10px;border-radius:9px;font-size:.86rem;font-weight:700;background:#fff7ed;color:#9a3412}
      .pv-retention-status[data-state="danger"]{background:#fef2f2;color:#b91c1c}
      .pv-retention-status[data-state="drive"]{background:#ecfdf5;color:#047857}
    `;
    document.head.appendChild(style);
  }

  const originalSaveRemote = PV.saveRemote?.bind(PV);
  if (originalSaveRemote && !PV.__temporaryArchiveSavePatched) {
    PV.__temporaryArchiveSavePatched = true;
    PV.saveRemote = (collectionName, record) => {
      if (record && collectionEntries().some((entry) => entry.collection === collectionName)) ensureRetention(record);
      return originalSaveRemote(collectionName, record);
    };
  }

  const originalPersistLocal = PV.persistLocal?.bind(PV);
  if (originalPersistLocal && !PV.__temporaryArchiveLocalPatched) {
    PV.__temporaryArchiveLocalPatched = true;
    PV.persistLocal = () => {
      collectionEntries().forEach(({ stateKey }) => (PV.state[stateKey] || []).forEach(ensureRetention));
      return originalPersistLocal();
    };
  }

  const originalRender = PV.renderCurrentView?.bind(PV);
  if (originalRender && !PV.__temporaryArchiveRenderPatched) {
    PV.__temporaryArchiveRenderPatched = true;
    PV.renderCurrentView = (...args) => {
      const result = originalRender(...args);
      if (PV.state?.isOpen) requestAnimationFrame(decorateWarnings);
      return result;
    };
  }

  function start() {
    installStyles();
    applyRetentionToLocal();
    void cleanupExpired();
    if (!timer) timer = window.setInterval(() => void cleanupExpired(), CHECK_INTERVAL_MS);
    document.addEventListener('visibilitychange', () => {
      if (!document.hidden && PV.state?.isOpen) void cleanupExpired();
    });
  }

  window.HeraPreventiviTemporaryArchive = Object.freeze({
    retentionDays: RETENTION_DAYS,
    warningDays: WARNING_DAYS,
    cleanupExpired,
    ensureRetention,
    daysRemaining
  });

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();
