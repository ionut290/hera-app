(() => {
  'use strict';

  const DB_NAME = 'varga-cantieri-admin-backups';
  const STORE_NAME = 'backups';
  const MAX_BACKUPS = 10;

  function clone(value) {
    return JSON.parse(JSON.stringify(value, (_key, item) => {
      if (item instanceof Map) return Object.fromEntries(item);
      if (item instanceof Set) return Array.from(item);
      if (item?.toDate && typeof item.toDate === 'function') return item.toDate().toISOString();
      if (item instanceof Date) return item.toISOString();
      return item;
    }));
  }

  function safeRead(factory, fallback) {
    try { return clone(factory()); } catch (_) { return fallback; }
  }

  function loadedSnapshot() {
    const commesse = safeRead(() => Array.from(commesseById.entries()), []);
    const impianti = safeRead(() => Array.from(impiantiByCommessaId.entries()), []);
    const current = safeRead(() => currentImpianti, []);
    const users = safeRead(() => platformUsers, []);
    const pendingDone = safeRead(() => pendingImpiantoActions, []);
    const pendingOffline = safeRead(() => typeof loadPendingOfflineMutations === 'function' ? loadPendingOfflineMutations() : [], []);
    const user = safeRead(() => ({ uid: currentUser?.uid || '', email: currentUser?.email || '', displayName: currentUser?.displayName || '' }), {});
    const recordCount = commesse.length + current.length + users.length + pendingDone.length + pendingOffline.length
      + impianti.reduce((sum, entry) => sum + (Array.isArray(entry?.[1]) ? entry[1].length : 0), 0);
    return {
      schemaVersion: 1,
      app: 'VARGA CANTIERI',
      createdAt: new Date().toISOString(),
      createdBy: user,
      source: 'Dati già caricati nella sessione; nessuna lettura Firestore aggiuntiva',
      recordCount,
      data: { commesse, impiantiByCommessa: impianti, currentImpianti: current, platformUsers: users, pendingImpiantoActions: pendingDone, pendingOfflineMutations: pendingOffline }
    };
  }

  function openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error('Archivio backup non disponibile.'));
    });
  }

  async function listBackups() {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const request = tx.objectStore(STORE_NAME).getAll();
      request.onsuccess = () => resolve((request.result || []).sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt))));
      request.onerror = () => reject(request.error);
      tx.oncomplete = () => db.close();
    });
  }

  async function saveBackup(snapshot) {
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put({ id: snapshot.createdAt, createdAt: snapshot.createdAt, recordCount: snapshot.recordCount, snapshot });
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
    db.close();
    const backups = await listBackups();
    if (backups.length > MAX_BACKUPS) {
      const cleanupDb = await openDb();
      await new Promise((resolve) => {
        const tx = cleanupDb.transaction(STORE_NAME, 'readwrite');
        backups.slice(MAX_BACKUPS).forEach((item) => tx.objectStore(STORE_NAME).delete(item.id));
        tx.oncomplete = resolve;
      });
      cleanupDb.close();
    }
  }

  function download(snapshot) {
    const stamp = snapshot.createdAt.replace(/[:.]/g, '-');
    const blob = new Blob([JSON.stringify(snapshot, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `VARGA-CANTIERI-backup-${stamp}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function backupCard() {
    return Array.from(document.querySelectorAll('#control-center-content details, #control-center-page details'))
      .find((card) => /Backup dati/i.test(card.querySelector('summary')?.textContent || '')) || null;
  }

  function buttons(card) {
    const all = Array.from(card?.querySelectorAll('button') || []);
    return {
      execute: all.find((button) => /ESEGUI BACKUP/i.test(button.textContent)),
      download: all.find((button) => /SCARICA BACKUP/i.test(button.textContent)),
      restore: all.find((button) => /RIPRISTINA BACKUP/i.test(button.textContent)),
      history: all.find((button) => /VISUALIZZA BACKUP PRECEDENTI/i.test(button.textContent))
    };
  }

  function setRow(card, label, value) {
    const row = Array.from(card.querySelectorAll('.control-center-row')).find((item) => (item.querySelector('span')?.textContent || '').trim() === label);
    const strong = row?.querySelector('strong');
    if (strong) strong.textContent = value;
  }

  async function refreshCard() {
    const card = backupCard();
    if (!card) return;
    const list = await listBackups().catch(() => []);
    const latest = list[0];
    setRow(card, 'Ultimo backup', latest ? new Date(latest.createdAt).toLocaleString('it-IT') : 'Mai eseguito');
    setRow(card, 'Stato', 'Attivo');
    setRow(card, 'Dimensione dati', latest ? `${Math.max(1, Math.round(JSON.stringify(latest.snapshot).length / 1024))} KB` : 'n/d');
    setRow(card, 'Record salvati', latest ? String(latest.recordCount || 0) : '0');
    setRow(card, 'Destinazione', 'Dispositivo amministratore');
    setRow(card, 'Errori', 'Nessuno');
    const btn = buttons(card);
    if (btn.restore) {
      btn.restore.disabled = true;
      btn.restore.title = 'Ripristino disattivato per proteggere FATTO, data, ora, operatore e WhatsApp.';
      btn.restore.textContent = 'RIPRISTINO PROTETTO';
    }
  }

  async function executeBackup() {
    const card = backupCard();
    const btn = buttons(card).execute;
    if (btn) btn.disabled = true;
    try {
      const snapshot = loadedSnapshot();
      await saveBackup(snapshot);
      download(snapshot);
      await refreshCard();
      alert(`Backup creato correttamente. ${snapshot.recordCount} record salvati nello storico locale e scaricati in formato JSON.`);
    } catch (error) {
      if (card) setRow(card, 'Errori', error?.message || 'Backup non riuscito');
      alert(error?.message || 'Backup non riuscito.');
    } finally {
      if (btn) btn.disabled = false;
    }
  }

  async function downloadLatest() {
    const latest = (await listBackups())[0];
    if (!latest) return alert('Nessun backup disponibile. Premi prima ESEGUI BACKUP.');
    download(latest.snapshot);
  }

  async function showHistory() {
    const list = await listBackups();
    if (!list.length) return alert('Nessun backup precedente disponibile.');
    const text = list.map((item, index) => `${index + 1}. ${new Date(item.createdAt).toLocaleString('it-IT')} — ${item.recordCount || 0} record`).join('\n');
    alert(`BACKUP PRECEDENTI (massimo ${MAX_BACKUPS})\n\n${text}\n\nPer scaricare l’ultimo backup usa SCARICA BACKUP.`);
  }

  function bind() {
    const card = backupCard();
    if (!card || card.dataset.backupSafeBound === '1') return;
    card.dataset.backupSafeBound = '1';
    const btn = buttons(card);
    btn.execute?.addEventListener('click', (event) => { event.preventDefault(); executeBackup(); });
    btn.download?.addEventListener('click', (event) => { event.preventDefault(); downloadLatest(); });
    btn.history?.addEventListener('click', (event) => { event.preventDefault(); showHistory(); });
    btn.restore?.addEventListener('click', (event) => event.preventDefault());
    refreshCard();
  }

  function init() {
    bind();
    const root = document.getElementById('control-center-content') || document.getElementById('control-center-page');
    if (!root) return;
    const observer = new MutationObserver(() => bind());
    observer.observe(root, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();

(() => {
  'use strict';
  if (document.querySelector('script[data-control-center-organizer]')) return;
  const script = document.createElement('script');
  script.src = './control-center-organizer.js?v=20260816c';
  script.defer = true;
  script.dataset.controlCenterOrganizer = '1';
  script.addEventListener('error', () => console.warn('Organizzazione Centro di controllo non caricata.'), { once: true });
  document.head.appendChild(script);
})();