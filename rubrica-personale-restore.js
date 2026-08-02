(() => {
  'use strict';
  if (window.__vargaPersonaleRestoreV2) return;
  window.__vargaPersonaleRestoreV2 = true;

  const RECORDS = __RECORDS__;
  const text = value => String(value ?? '').trim();
  const normalize = value => text(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleUpperCase('it-IT')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const personKey = record => [record?.cognome, record?.nome].map(normalize).filter(Boolean).join('|');
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));

  async function restoreHistoricalIds() {
    for (let attempt = 0; attempt < 80; attempt += 1) {
      const firestore = window.firebase?.firestore?.();
      const currentUser = window.firebase?.auth?.()?.currentUser;
      const isManager = typeof window.canManageData === 'function' && window.canManageData();
      if (!firestore || !currentUser || !isManager) {
        await wait(250);
        continue;
      }

      const collectionName = typeof window.getPersonaleCollectionName === 'function'
        ? window.getPersonaleCollectionName()
        : 'personale';
      const collection = firestore.collection(collectionName);
      const currentSnapshot = await collection.get();
      const currentByKey = new Map();
      currentSnapshot.docs.forEach(doc => {
        const data = doc.data() || {};
        const key = personKey(data);
        if (key && !currentByKey.has(key)) currentByKey.set(key, { doc, data });
      });

      const timestamp = window.firebase.firestore.FieldValue.serverTimestamp();
      let changed = 0;
      for (let start = 0; start < RECORDS.length; start += 400) {
        const batch = firestore.batch();
        let batchChanges = 0;
        RECORDS.slice(start, start + 400).forEach(record => {
          const historicalId = text(record.idOperatore);
          const key = personKey(record);
          if (!historicalId || !key) return;
          const current = currentByKey.get(key);
          const source = current?.data || {};
          const payload = {
            ...source,
            id: historicalId,
            idOperatore: historicalId,
            ID_OPERATORE: historicalId,
            nome: text(record.nome) || text(source.nome),
            cognome: text(record.cognome) || text(source.cognome),
            fullName: `${text(record.cognome)} ${text(record.nome)}`.trim(),
            codiceOperatore: text(record.codiceOperatore) || text(source.codiceOperatore),
            emailAccessoApp: text(record.emailAccessoApp) || text(source.emailAccessoApp),
            linkedUserId: text(record.linkedUserId) || text(source.linkedUserId),
            linkedUserEmail: text(record.linkedUserEmail) || text(source.linkedUserEmail),
            profiloCollegato: text(record.profiloCollegato) || text(source.profiloCollegato),
            restoredHistoricalIdAt: timestamp,
            updatedAt: timestamp,
            updatedBy: currentUser.uid,
            source: 'ripristino-id-matrice-2026-08-02'
          };
          batch.set(collection.doc(historicalId), payload, { merge: true });
          batchChanges += 1;
          changed += 1;
        });
        if (batchChanges) await batch.commit();
      }

      console.info(`Ripristino ID personale completato: ${changed} operatori elaborati.`);
      if (typeof window.subscribePersonale === 'function') {
        try { window.subscribePersonale(); } catch (_) {}
      }
      return;
    }
  }

  restoreHistoricalIds().catch(error => console.error('Ripristino ID personale non riuscito:', error));
})();
".replace("__RECORDS__