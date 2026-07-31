(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');
  if (PV.firestoreChunkingInstalled) return;
  PV.firestoreChunkingInstalled = true;

  const FORMAT = 'chunked-v1';
  const CHUNKS_COLLECTION = 'chunks';
  const TARGET_CHUNK_BYTES = 600 * 1024;
  const MAX_ITEMS_PER_CHUNK = 350;
  const MAX_BATCH_WRITES = 400;
  const MIGRATION_KEY = 'hera_preventivi_firestore_chunks_v1';

  const originalSaveRemote = PV.saveRemote.bind(PV);
  const originalDeleteRemote = PV.deleteRemote.bind(PV);
  const originalSubscribeCollection = PV.subscribeCollection.bind(PV);
  const originalLoadLocal = PV.loadLocal.bind(PV);

  const encoder = typeof TextEncoder !== 'undefined' ? new TextEncoder() : null;

  function byteLength(value) {
    const text = JSON.stringify(value ?? null);
    return encoder ? encoder.encode(text).length : unescape(encodeURIComponent(text)).length;
  }

  function cleanItem(item) {
    return PV.stripUndefined({
      id: String(item?.id || ''),
      chapter: String(item?.chapter || ''),
      code: String(item?.code || ''),
      description: String(item?.description || ''),
      unit: String(item?.unit || ''),
      unitPrice: PV.parseNumber(item?.unitPrice),
      discount: PV.parseNumber(item?.discount)
    });
  }

  function splitItems(items) {
    const chunks = [];
    let current = [];
    let estimatedBytes = 256;

    (Array.isArray(items) ? items : []).forEach((sourceItem) => {
      const item = cleanItem(sourceItem);
      const itemBytes = byteLength(item) + 96;
      const exceedsTarget = current.length > 0 && estimatedBytes + itemBytes > TARGET_CHUNK_BYTES;
      const exceedsCount = current.length >= MAX_ITEMS_PER_CHUNK;
      if (exceedsTarget || exceedsCount) {
        chunks.push(current);
        current = [];
        estimatedBytes = 256;
      }
      current.push(item);
      estimatedBytes += itemBytes;
    });

    if (current.length || !chunks.length) chunks.push(current);
    return chunks;
  }

  function generationId() {
    const random = window.crypto?.randomUUID?.().replace(/-/g, '').slice(0, 10)
      || Math.random().toString(36).slice(2, 12);
    return `g${Date.now()}-${random}`;
  }

  function chunkDocumentId(generation, index) {
    return `${generation}-${String(index).padStart(5, '0')}`;
  }

  async function commitOperations(firestore, operations) {
    for (let start = 0; start < operations.length; start += MAX_BATCH_WRITES) {
      const batch = firestore.batch();
      operations.slice(start, start + MAX_BATCH_WRITES).forEach((operation) => {
        if (operation.type === 'delete') batch.delete(operation.ref);
        else batch.set(operation.ref, operation.data, operation.options || {});
      });
      await batch.commit();
    }
  }

  function showProgress(message, type = '') {
    PV.setSyncBadge(message, type);
    const button = PV.page()?.querySelector('[data-pv-sync]');
    if (button && message) {
      button.textContent = message;
      button.title = `${message}. Premi per scegliere dove salvare i dati.`;
      button.setAttribute('aria-label', button.title);
    }
  }

  async function saveChunkedPriceList(record) {
    if (!PV.state.firestore || !record?.id) return false;

    const firestore = PV.state.firestore;
    const collectionRef = firestore.collection(PV.collections.priceLists);
    const listRef = collectionRef.doc(record.id);
    const chunksRef = listRef.collection(CHUNKS_COLLECTION);
    const items = Array.isArray(record.items) ? record.items : [];
    const chunks = splitItems(items);
    const generation = generationId();
    const oldChunksSnapshot = await chunksRef.get();
    const oldChunkRefs = oldChunksSnapshot.docs.map((doc) => doc.ref);

    showProgress(`☁️ Firestore: preparo ${chunks.length} parti`);

    const writeOperations = chunks.map((chunkItems, index) => ({
      type: 'set',
      ref: chunksRef.doc(chunkDocumentId(generation, index)),
      data: {
        generation,
        index,
        itemCount: chunkItems.length,
        estimatedBytes: byteLength(chunkItems),
        items: chunkItems,
        updatedAt: PV.nowIso()
      }
    }));

    await commitOperations(firestore, writeOperations);

    const metadata = PV.stripUndefined({ ...record });
    delete metadata.items;
    metadata.storageFormat = FORMAT;
    metadata.chunkGeneration = generation;
    metadata.chunkCount = chunks.length;
    metadata.itemCount = items.length;
    metadata.chunkTargetBytes = TARGET_CHUNK_BYTES;
    metadata.syncPending = false;

    await listRef.set(metadata, { merge: false });

    if (oldChunkRefs.length) {
      await commitOperations(firestore, oldChunkRefs.map((ref) => ({ type: 'delete', ref })));
    }

    PV.state.remoteConnected = true;
    PV.state.remoteDenied = false;
    showProgress('☁️ Firestore sincronizzato', 'ok');
    return true;
  }

  async function hydratePriceList(doc) {
    const data = doc.data() || {};
    if (data.storageFormat !== FORMAT) {
      return { id: doc.id, ...data, items: Array.isArray(data.items) ? data.items : [] };
    }

    const snapshot = await doc.ref.collection(CHUNKS_COLLECTION).get();
    const generation = String(data.chunkGeneration || '');
    const chunks = snapshot.docs
      .map((chunkDoc) => chunkDoc.data() || {})
      .filter((chunk) => !generation || chunk.generation === generation)
      .sort((a, b) => Number(a.index || 0) - Number(b.index || 0));
    const items = chunks.flatMap((chunk) => Array.isArray(chunk.items) ? chunk.items : []);

    if (Number.isFinite(Number(data.itemCount)) && items.length !== Number(data.itemCount)) {
      throw new Error(`Prezziario ${data.name || doc.id}: attese ${data.itemCount} voci, lette ${items.length}.`);
    }

    return { id: doc.id, ...data, items, syncPending: false };
  }

  function subscribeChunkedPriceLists(stateKey, deletionKey) {
    let readSequence = 0;
    const unsubscribe = PV.state.firestore.collection(PV.collections.priceLists).onSnapshot(async (snapshot) => {
      const currentSequence = ++readSequence;
      try {
        showProgress('☁️ Firestore: leggo i prezziari…');
        const remote = await Promise.all(snapshot.docs.map(hydratePriceList));
        if (currentSequence !== readSequence) return;
        PV.state[stateKey] = PV.mergeRemote(PV.state[stateKey], remote, PV.state.deletions[deletionKey]);
        PV.state.remoteConnected = true;
        PV.state.remoteDenied = false;
        PV.persistLocal();
        showProgress('☁️ Firestore sincronizzato', 'ok');
        if (!PV.state.editingQuoteId && !PV.state.editingPriceListId) PV.renderCurrentView();
        PV.scheduleSync();
      } catch (error) {
        if (currentSequence !== readSequence) return;
        if (error?.code === 'permission-denied') PV.state.remoteDenied = true;
        console.warn('Preventivi: lettura prezziari divisi non riuscita.', error);
        showProgress('⚠️ Firestore: sincronizzazione non riuscita', 'warning');
        PV.setFeedback?.(error?.message || 'Sincronizzazione prezziari non riuscita.', 'warning');
      }
    }, (error) => {
      if (error?.code === 'permission-denied') PV.state.remoteDenied = true;
      console.warn('Preventivi: ascolto prezziari Firestore non riuscito.', error);
      showProgress('⚠️ Firestore non autorizzato', 'warning');
    });
    PV.state.unsubscribers.push(unsubscribe);
  }

  async function deleteChunkedPriceList(id) {
    if (!PV.state.firestore || !id) return false;
    try {
      const firestore = PV.state.firestore;
      const listRef = firestore.collection(PV.collections.priceLists).doc(id);
      const chunksSnapshot = await listRef.collection(CHUNKS_COLLECTION).get();
      const operations = chunksSnapshot.docs.map((doc) => ({ type: 'delete', ref: doc.ref }));
      operations.push({ type: 'delete', ref: listRef });
      await commitOperations(firestore, operations);
      return true;
    } catch (error) {
      if (error?.code === 'permission-denied') PV.state.remoteDenied = true;
      console.warn('Preventivi: eliminazione prezziario diviso non riuscita.', error);
      return false;
    }
  }

  PV.saveRemote = async (collectionName, record) => {
    if (collectionName !== PV.collections.priceLists) return originalSaveRemote(collectionName, record);
    try {
      return await saveChunkedPriceList(record);
    } catch (error) {
      if (error?.code === 'permission-denied') PV.state.remoteDenied = true;
      console.warn('Preventivi: salvataggio prezziario diviso non riuscito.', error);
      showProgress('⚠️ Firestore: dati in attesa', 'warning');
      PV.setFeedback?.(`Prezziario salvato sul dispositivo. Firestore: ${error?.message || 'errore di sincronizzazione'}`, 'warning');
      return false;
    }
  };

  PV.deleteRemote = async (collectionName, id) => {
    if (collectionName !== PV.collections.priceLists) return originalDeleteRemote(collectionName, id);
    return deleteChunkedPriceList(id);
  };

  PV.subscribeCollection = (collectionName, stateKey, deletionKey) => {
    if (collectionName !== PV.collections.priceLists) {
      return originalSubscribeCollection(collectionName, stateKey, deletionKey);
    }
    return subscribeChunkedPriceLists(stateKey, deletionKey);
  };

  PV.retryFirestoreSync = () => {
    PV.state.priceLists = (PV.state.priceLists || []).map((item) => ({ ...item, syncPending: true }));
    PV.state.quotes = (PV.state.quotes || []).map((item) => ({ ...item, syncPending: true }));
    PV.persistLocal();
    PV.scheduleSync();
    showProgress('☁️ Firestore: nuovo tentativo…');
  };

  document.addEventListener('click', (event) => {
    const button = event.target.closest('[data-pv-retry-sync]');
    if (!button) return;
    event.preventDefault();
    PV.retryFirestoreSync();
  });

  function ensureRetryButton() {
    const modal = document.getElementById('pv-storage-modal');
    const actions = modal?.querySelector('.pv-storage-actions');
    if (!actions || actions.querySelector('[data-pv-retry-sync]')) return;
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'pv-storage-close';
    button.dataset.pvRetrySync = '1';
    button.textContent = 'RIPROVA SINCRONIZZAZIONE';
    actions.prepend(button);
  }

  const modalObserver = new MutationObserver(ensureRetryButton);
  if (document.body) modalObserver.observe(document.body, { childList: true, subtree: true });
  ensureRetryButton();

  function migrateLocalPriceLists() {
    try {
      const alreadyMigrated = localStorage.getItem(MIGRATION_KEY) === '1';
      if (alreadyMigrated) return;
      PV.state.priceLists = (PV.state.priceLists || []).map((item) => ({ ...item, syncPending: true }));
      PV.persistLocal();
      localStorage.setItem(MIGRATION_KEY, '1');
    } catch (error) {
      console.warn('Preventivi: stato migrazione prezziari non memorizzato.', error);
    }
  }

  PV.loadLocal = () => {
    originalLoadLocal();
    migrateLocalPriceLists();
  };

  // Se il vecchio collegamento era già partito, lo riavvia con il formato a blocchi.
  if (PV.state.storageMode !== 'device' && window.firebase?.auth?.()?.currentUser) {
    (PV.state.unsubscribers || []).forEach((unsubscribe) => {
      try { unsubscribe(); } catch (_) { /* Listener già chiuso. */ }
    });
    PV.state.unsubscribers = [];
    PV.state.firestore = null;
    PV.state.remoteConnected = false;
    window.setTimeout(() => PV.connectFirebase?.(), 0);
  }
})();
