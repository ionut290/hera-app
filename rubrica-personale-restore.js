(() => {
  'use strict';

  const RUN_KEY = 'personale-safe-dedup-20260803-v1';
  const RESTORE_MARKER = 'Matrice_Personale_STATO_ATTIVO_2026-08-03';
  const MAX_EXPECTED_DELETIONS = 170;

  if (window.__vargaPersonaleSafeDedup20260803 || localStorage.getItem(RUN_KEY) === 'done') return;
  window.__vargaPersonaleSafeDedup20260803 = true;
  window.__personaleDestructiveClearDisabled = true;

  const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const normalize = (value) => String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]+/gi, ' ')
    .trim()
    .replace(/\s+/g, ' ')
    .toLocaleUpperCase('it-IT');

  const isEmpty = (value) => value == null
    || value === ''
    || (Array.isArray(value) && value.length === 0);

  const personKey = (data = {}) => {
    const nome = normalize(data.nome || data.name || data.firstName || '');
    const cognome = normalize(data.cognome || data.surname || data.lastName || '');
    if (nome && cognome) return `${nome}|${cognome}`;

    const fullName = normalize(
      data.nomeCompleto
      || data.fullName
      || data.displayName
      || data.nominativo
      || data.operatore
      || ''
    );
    return fullName ? `FULL|${fullName}` : '';
  };

  const toList = (value) => {
    if (Array.isArray(value)) return value.map((item) => String(item || '').trim()).filter(Boolean);
    return String(value || '').split(/[;,|\n]+/).map((item) => item.trim()).filter(Boolean);
  };

  const unionList = (first, second) => {
    const values = [...toList(first), ...toList(second)];
    return [...new Map(values.map((value) => [normalize(value), value])).values()];
  };

  function mergeDuplicateIntoCanonical(canonicalData, duplicateData) {
    const merged = { ...canonicalData };
    let changed = false;
    const protectedFields = new Set(['idOperatore', 'nome', 'cognome', 'restoredFrom']);

    Object.entries(duplicateData || {}).forEach(([key, value]) => {
      if (protectedFields.has(key) || isEmpty(value)) return;

      if (key === 'commesseAbilitate' || key === 'abilitazioni') {
        const combined = unionList(merged[key], value);
        const before = JSON.stringify(toList(merged[key]));
        const after = JSON.stringify(combined);
        if (before !== after) {
          merged[key] = combined;
          changed = true;
        }
        return;
      }

      if (isEmpty(merged[key])) {
        merged[key] = value;
        changed = true;
      }
    });

    merged.restoredFrom = RESTORE_MARKER;
    return { merged, changed };
  }

  async function commitOperations(firestore, operations) {
    for (let start = 0; start < operations.length; start += 400) {
      const batch = firestore.batch();
      operations.slice(start, start + 400).forEach((operation) => {
        if (operation.type === 'set') batch.set(operation.ref, operation.data, { merge: true });
        if (operation.type === 'delete') batch.delete(operation.ref);
      });
      await batch.commit();
    }
  }

  async function removeSafeDuplicates() {
    for (let attempt = 0; attempt < 120; attempt += 1) {
      const firestore = window.firebase?.firestore?.();
      const user = window.firebase?.auth?.()?.currentUser;

      if (!firestore || !user) {
        await wait(250);
        continue;
      }

      const canCheckManager = typeof window.canManageData === 'function';
      if (!canCheckManager) {
        await wait(250);
        continue;
      }

      const isManager = Boolean(window.canManageData());
      if (!isManager) {
        console.info('Pulizia sicura duplicati Personale saltata: utente senza permessi amministratore.');
        return {
          skipped: true,
          reason: 'not-admin'
        };
      }

      const collectionName = typeof window.getPersonaleCollectionName === 'function'
        ? window.getPersonaleCollectionName()
        : 'personale';
      const collection = firestore.collection(collectionName);
      const snapshot = typeof window.runFirestoreGetWithRetry === 'function'
        ? await window.runFirestoreGetWithRetry(collection, {
            label: 'RIMOZIONE DUPLICATI PERSONALE',
            timeoutMs: 15000,
            retries: 2
          })
        : await collection.get();

      const groups = new Map();
      snapshot.docs.forEach((doc) => {
        const data = doc.data() || {};
        const key = personKey(data);
        if (!key) return;
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push({ doc, data });
      });

      const operations = [];
      const deletedIds = [];
      let canonicalUpdates = 0;
      let skippedUnsafeGroups = 0;

      groups.forEach((entries) => {
        if (entries.length < 2) return;

        const canonicalEntries = entries.filter(({ data }) => data.restoredFrom === RESTORE_MARKER);
        if (canonicalEntries.length !== 1) {
          skippedUnsafeGroups += 1;
          return;
        }

        const canonical = canonicalEntries[0];
        const duplicates = entries.filter(({ doc }) => doc.id !== canonical.doc.id);
        if (!duplicates.length) return;

        let canonicalData = { ...canonical.data };
        let canonicalChanged = false;

        duplicates.forEach(({ doc, data }) => {
          const result = mergeDuplicateIntoCanonical(canonicalData, data);
          canonicalData = result.merged;
          canonicalChanged = canonicalChanged || result.changed;
          operations.push({ type: 'delete', ref: doc.ref });
          deletedIds.push(doc.id);
        });

        if (canonicalChanged) {
          operations.push({ type: 'set', ref: canonical.doc.ref, data: canonicalData });
          canonicalUpdates += 1;
        }
      });

      if (deletedIds.length > MAX_EXPECTED_DELETIONS) {
        throw new Error(`Pulizia bloccata: rilevate ${deletedIds.length} eliminazioni, oltre il limite di sicurezza ${MAX_EXPECTED_DELETIONS}.`);
      }

      if (operations.length) await commitOperations(firestore, operations);

      localStorage.setItem(RUN_KEY, 'done');
      localStorage.setItem(`${RUN_KEY}:result`, JSON.stringify({
        completedAt: new Date().toISOString(),
        collection: collectionName,
        initialRecords: snapshot.size,
        duplicatesDeleted: deletedIds.length,
        canonicalRecordsUpdated: canonicalUpdates,
        unsafeGroupsSkipped: skippedUnsafeGroups
      }));

      console.info('Pulizia sicura Personale completata', {
        recordIniziali: snapshot.size,
        duplicatiEliminati: deletedIds.length,
        recordCorrettiAggiornati: canonicalUpdates,
        gruppiNonSicuriIgnorati: skippedUnsafeGroups
      });

      window.dispatchEvent(new CustomEvent('varga-personale-dedup-complete', {
        detail: {
          initialRecords: snapshot.size,
          duplicatesDeleted: deletedIds.length,
          canonicalRecordsUpdated: canonicalUpdates,
          unsafeGroupsSkipped: skippedUnsafeGroups
        }
      }));
      return {
        skipped: false,
        duplicatesDeleted: deletedIds.length,
        canonicalRecordsUpdated: canonicalUpdates,
        unsafeGroupsSkipped: skippedUnsafeGroups
      };
    }

    console.info('Pulizia sicura duplicati Personale saltata: Firebase, autenticazione o controllo permessi non ancora disponibili.');
    window.__vargaPersonaleSafeDedup20260803 = false;
    return {
      skipped: true,
      reason: 'dependencies-unavailable'
    };
  }

  removeSafeDuplicates().catch((error) => {
    window.__vargaPersonaleSafeDedup20260803 = false;
    console.error('Pulizia sicura duplicati Personale non riuscita:', error);
  });
})();