(() => {
  'use strict';

  const RUN_KEY = 'personale-clear-20260802-v1';
  if (window.__personaleClear20260802 || localStorage.getItem(RUN_KEY) === 'done') return;
  window.__personaleClear20260802 = true;

  const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  async function clearPersonale() {
    for (let attempt = 0; attempt < 120; attempt += 1) {
      const firestore = window.firebase?.firestore?.();
      const user = window.firebase?.auth?.()?.currentUser;
      const isManager = typeof window.canManageData === 'function' && window.canManageData();

      if (!firestore || !user || !isManager) {
        await wait(250);
        continue;
      }

      const collectionName = typeof window.getPersonaleCollectionName === 'function'
        ? window.getPersonaleCollectionName()
        : 'personale';
      const collection = firestore.collection(collectionName);
      let deleted = 0;

      while (true) {
        const snapshot = await collection.limit(400).get();
        if (snapshot.empty) break;

        const batch = firestore.batch();
        snapshot.docs.forEach((doc) => batch.delete(doc.ref));
        await batch.commit();
        deleted += snapshot.size;
      }

      localStorage.setItem(RUN_KEY, 'done');
      if (Array.isArray(window.personaleRecords)) window.personaleRecords.splice(0);
      if (typeof window.renderPersonale === 'function') {
        try { window.renderPersonale(); } catch (_) {}
      }
      console.info(`Elenco Personale svuotato: ${deleted} documenti eliminati.`);
      return;
    }

    throw new Error('Firestore o permessi amministratore non disponibili.');
  }

  clearPersonale().catch((error) => {
    window.__personaleClear20260802 = false;
    console.error('Svuotamento elenco Personale non riuscito:', error);
  });
})();
