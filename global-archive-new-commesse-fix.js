(() => {
  'use strict';

  const INTERVAL_MS = 15000;
  let timer = null;
  let unsubscribe = null;
  let running = false;

  const api = () => window.HeraGlobalArchive || null;
  const db = () => window.firebase?.firestore?.() || null;
  const authUser = () => window.firebase?.auth?.()?.currentUser || null;

  async function syncInMemoryCommesse() {
    const archive = api();
    if (!archive) return;
    const sources = [window.commesseById, window.commesse, window.projects, window.cantieri];
    for (const source of sources) {
      if (!source) continue;
      const rows = source instanceof Map ? [...source.entries()].map(([id, value]) => ({ id, ...(value || {}) })) : Array.isArray(source) ? source : [];
      for (const row of rows) {
        const id = String(row?.id || row?.uid || row?.commessaId || row?.projectId || '').trim();
        if (!id) continue;
        await archive.archiveCommessa(id, row);
      }
    }
  }

  async function reconcile() {
    if (running || !authUser() || !api()) return;
    running = true;
    try {
      await api().migrateAll();
      await syncInMemoryCommesse();
    } catch (error) {
      console.warn('Global archive: controllo nuove commesse non riuscito.', error);
    } finally {
      running = false;
    }
  }

  function installDirectListener() {
    if (unsubscribe || !db() || !api()) return;
    try {
      unsubscribe = db().collection('commesse').onSnapshot((snapshot) => {
        snapshot.docChanges().forEach((change) => {
          if (change.type === 'removed') return;
          const data = change.doc.data() || {};
          api().archiveCommessa(change.doc.id, data)
            .then(() => api().syncCommessa(change.doc.id, data))
            .catch((error) => console.warn('Global archive: nuova commessa non sincronizzata.', error));
        });
      }, (error) => console.warn('Global archive: listener nuove commesse non disponibile.', error));
    } catch (error) {
      console.warn('Global archive: listener nuove commesse non installato.', error);
    }
  }

  function startChecks() {
    installDirectListener();
    reconcile();
    if (!timer) timer = window.setInterval(reconcile, INTERVAL_MS);
  }

  function start() {
    if (!window.firebase?.auth) return;
    window.firebase.auth().onAuthStateChanged((user) => {
      if (!user) return;
      const waitForArchive = () => {
        if (api()) startChecks();
        else window.setTimeout(waitForArchive, 500);
      };
      waitForArchive();
    });
    window.addEventListener('focus', reconcile);
    document.addEventListener('visibilitychange', () => {
      if (!document.hidden) reconcile();
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();
