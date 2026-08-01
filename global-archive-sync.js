(() => {
  'use strict';

  const VERSION = '20260801a';
  const ARCHIVE_ROOT = 'globalArchive';
  const MIGRATION_KEY = `hera_global_archive_migrated_${VERSION}`;
  const syncLocks = new Map();

  const db = () => window.firebase?.firestore?.() || null;
  const authUser = () => window.firebase?.auth?.()?.currentUser || null;
  const text = (value) => String(value ?? '').trim();
  const norm = (value) => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  const first = (obj, keys) => {
    for (const key of keys) {
      const value = obj?.[key];
      if (value !== undefined && value !== null && text(value)) return value;
    }
    return '';
  };
  const cleanObject = (source) => Object.fromEntries(Object.entries(source).filter(([, value]) => value !== undefined && value !== null && text(value) !== ''));
  const serverTimestamp = () => window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || new Date().toISOString();

  function commesseCollectionName() {
    try {
      if (typeof window.getCommesseCollectionName === 'function') return window.getCommesseCollectionName();
    } catch (_) { /* fallback */ }
    return 'commesse';
  }

  function commessaSnapshot(raw = {}, id = '') {
    return cleanObject({
      sourceCommessaId: text(id || first(raw, ['id','uid','commessaId','projectId'])),
      nome: text(first(raw, ['nome','name','denominazione','titolo','commessa'])),
      codice: text(first(raw, ['codiceCommessa','codice','code','numeroCommessa'])),
      cliente: text(first(raw, ['cliente','committente','clientName','ragioneSociale'])),
      numeroContratto: text(first(raw, ['numeroContratto','contratto','contractNumber'])),
      richiedente: text(first(raw, ['richiedente','referente','referenteCliente'])),
      archivioPermanente: true,
      updatedAt: serverTimestamp(),
      updatedByUid: authUser()?.uid || ''
    });
  }

  function plantSnapshot(raw = {}, commessaId = '', plantId = '') {
    return cleanObject({
      sourceCommessaId: text(commessaId),
      sourceImpiantoId: text(plantId || first(raw, ['id','uid','impiantoId','plantId'])),
      globalImpiantoId: text(first(raw, ['globalImpiantoId'])),
      idSap: text(first(raw, ['idSap','idSAP','ID SAP','sap','codiceSap','codiceHera'])),
      denominazione: text(first(raw, ['denominazione','Denominazione Impianto','nome','name','impianto'])),
      comune: text(first(raw, ['comune','Comune','city','localita'])),
      indirizzo: text(first(raw, ['indirizzo','Descrizione via','descrizioneVia','via','address','ubicazione'])),
      tipologia: text(first(raw, ['tipologia','Tipologia impianto','tipologiaImpianto','tipo','type'])),
      area: text(first(raw, ['area','AREA','competenza','Area/Competenza'])),
      coordinate: text(first(raw, ['coordinate','coordinates','Coordinate GPS(X)/GPS(Y','coordinateGps'])),
      gpsX: first(raw, ['gpsX','longitude','longitudine']),
      gpsY: first(raw, ['gpsY','latitude','latitudine']),
      dittaEsecutrice: text(first(raw, ['dittaEsecutrice','Ditta esecutrice'])),
      archivioPermanente: true,
      updatedAt: serverTimestamp(),
      updatedByUid: authUser()?.uid || ''
    });
  }

  function archiveCommessaRef(commessaId) {
    return db()?.collection(ARCHIVE_ROOT).doc('commesse').collection('items').doc(text(commessaId));
  }

  function archivePlantsRef(commessaId) {
    return archiveCommessaRef(commessaId)?.collection('impianti');
  }

  function plantKey(raw = {}, fallback = '') {
    const sap = norm(first(raw, ['idSap','idSAP','ID SAP','sap','codiceSap','codiceHera']));
    if (sap) return `sap-${sap}`;
    const linked = norm(first(raw, ['globalImpiantoId']));
    if (linked) return `global-${linked}`;
    const composite = norm(`${first(raw, ['denominazione','Denominazione Impianto','nome','name','impianto'])}|${first(raw, ['comune','Comune','city'])}|${first(raw, ['indirizzo','Descrizione via','descrizioneVia','via','address'])}`);
    return composite ? `anag-${composite}` : `source-${norm(fallback || first(raw, ['id','uid','impiantoId','plantId']))}`;
  }

  async function mergeArchiveDocument(ref, incoming) {
    if (!ref || !incoming) return;
    const current = await ref.get();
    if (!current.exists) {
      await ref.set(incoming, { merge: true });
      return;
    }
    const existing = current.data() || {};
    const merged = { ...existing };
    Object.entries(incoming).forEach(([key, value]) => {
      if (value !== undefined && value !== null && text(value) !== '') merged[key] = value;
    });
    merged.archivioPermanente = true;
    merged.updatedAt = incoming.updatedAt || serverTimestamp();
    await ref.set(merged, { merge: true });
  }

  async function archiveCommessa(commessaId, raw = {}) {
    if (!db() || !text(commessaId)) return;
    await mergeArchiveDocument(archiveCommessaRef(commessaId), commessaSnapshot(raw, commessaId));
  }

  async function archivePlant(commessaId, plantId, raw = {}) {
    if (!db() || !text(commessaId)) return;
    await archiveCommessa(commessaId, { id: commessaId });
    const key = plantKey(raw, plantId);
    if (!key) return;
    await mergeArchiveDocument(archivePlantsRef(commessaId).doc(key), plantSnapshot(raw, commessaId, plantId));
  }

  async function syncCommessa(commessaId, rawCommessa = null) {
    if (!db() || !text(commessaId)) return;
    if (syncLocks.has(commessaId)) return syncLocks.get(commessaId);
    const task = (async () => {
      const ref = db().collection(commesseCollectionName()).doc(commessaId);
      const commessaDoc = rawCommessa ? { exists: true, data: () => rawCommessa } : await ref.get();
      if (commessaDoc.exists) await archiveCommessa(commessaId, commessaDoc.data() || {});
      for (const collectionName of ['impiantiFisici','impianti']) {
        const snapshot = await ref.collection(collectionName).get();
        for (const doc of snapshot.docs) await archivePlant(commessaId, doc.id, doc.data() || {});
      }
    })().finally(() => syncLocks.delete(commessaId));
    syncLocks.set(commessaId, task);
    return task;
  }

  async function migrateAll() {
    if (!db() || !authUser()) return;
    try {
      if (localStorage.getItem(MIGRATION_KEY) === '1') return;
    } catch (_) { /* continua */ }
    const snapshot = await db().collection(commesseCollectionName()).get();
    for (const doc of snapshot.docs) await syncCommessa(doc.id, doc.data() || {});
    try { localStorage.setItem(MIGRATION_KEY, '1'); } catch (_) { /* non bloccante */ }
  }

  function installRealtimeArchive() {
    if (!db()) return;
    const root = db().collection(commesseCollectionName());
    root.onSnapshot((snapshot) => {
      snapshot.docChanges().forEach((change) => {
        if (change.type === 'removed') return;
        archiveCommessa(change.doc.id, change.doc.data() || {}).catch((error) => console.warn('Global archive: commessa non archiviata.', error));
        syncCommessa(change.doc.id, change.doc.data() || {}).catch((error) => console.warn('Global archive: sincronizzazione commessa non riuscita.', error));
      });
    }, (error) => console.warn('Global archive: listener commesse non disponibile.', error));
  }

  function patchWrites() {
    const proto = window.firebase?.firestore?.DocumentReference?.prototype;
    if (!proto || proto.__heraGlobalArchivePatched) return;
    proto.__heraGlobalArchivePatched = true;
    const originalSet = proto.set;
    const originalUpdate = proto.update;
    const afterWrite = (ref, data) => {
      try {
        const segments = ref.path.split('/');
        const collection = segments[0];
        const commessaId = segments[1];
        const childCollection = segments[2];
        const childId = segments[3];
        if (collection !== commesseCollectionName() || !commessaId) return;
        if (!childCollection) archiveCommessa(commessaId, data || {}).catch(() => {});
        else if (['impiantiFisici','impianti'].includes(childCollection)) archivePlant(commessaId, childId, data || {}).catch(() => {});
      } catch (_) { /* mai bloccare la scrittura originale */ }
    };
    proto.set = function patchedSet(data, options) {
      return originalSet.call(this, data, options).then((result) => { afterWrite(this, data); return result; });
    };
    proto.update = function patchedUpdate(data) {
      return originalUpdate.call(this, data).then((result) => { afterWrite(this, data); return result; });
    };
  }

  function exposeApi() {
    window.HeraGlobalArchive = Object.freeze({
      version: VERSION,
      archiveCommessa,
      archivePlant,
      syncCommessa,
      migrateAll,
      rootCollection: ARCHIVE_ROOT
    });
  }

  async function start() {
    exposeApi();
    patchWrites();
    if (!window.firebase?.auth) return;
    window.firebase.auth().onAuthStateChanged((user) => {
      if (!user) return;
      installRealtimeArchive();
      migrateAll().catch((error) => console.warn('Global archive: migrazione iniziale non riuscita.', error));
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();
