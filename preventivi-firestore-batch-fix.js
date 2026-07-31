(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');
  if (PV.firestoreBatchSizeFixInstalled) return;
  PV.firestoreBatchSizeFixInstalled = true;

  const FORMAT = 'chunked-v1';
  const CHUNKS_COLLECTION = 'chunks';
  const TARGET_CHUNK_BYTES = 600 * 1024;
  const MAX_ITEMS_PER_CHUNK = 350;
  const MAX_BATCH_WRITES = 5;
  const MAX_BATCH_BYTES = 4 * 1024 * 1024;
  const OPERATION_OVERHEAD_BYTES = 16 * 1024;
  const MIGRATION_KEY = 'varga_preventivi_firestore_small_batches_v1';

  const originalSaveRemote = PV.saveRemote.bind(PV);
  const originalLoadLocal = PV.loadLocal.bind(PV);
  const inFlight = new Map();
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
      const exceedsBytes = current.length > 0 && estimatedBytes + itemBytes > TARGET_CHUNK_BYTES;
      const exceedsCount = current.length >= MAX_ITEMS_PER_CHUNK;
      if (exceedsBytes || exceedsCount) {
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

  function operationSize(operation) {
    if (operation.type === 'delete') return 2048;
    return byteLength(operation.data) + OPERATION_OVERHEAD_BYTES;
  }

  function groupOperations(operations) {
    const groups = [];
    let current = [];
    let currentBytes = 0;

    operations.forEach((operation) => {
      const estimatedBytes = operationSize(operation);
      const exceedsCount = current.length >= MAX_BATCH_WRITES;
      const exceedsBytes = current.length > 0 && currentBytes + estimatedBytes > MAX_BATCH_BYTES;
      if (exceedsCount || exceedsBytes) {
        groups.push(current);
        current = [];
        currentBytes = 0;
      }
      current.push(operation);
      currentBytes += estimatedBytes;
    });

    if (current.length) groups.push(current);
    return groups;
  }

  function showProgress(message, type = '') {
    PV.setSyncBadge(message, type);
    const button = PV.page()?.querySelector('[data-pv-sync]');
    if (!button || !message) return;
    button.textContent = message;
    button.title = `${message}. Non chiudere la pagina durante la sincronizzazione.`;
    button.setAttribute('aria-label', button.title);
  }

  async function commitOperations(firestore, operations, phase) {
    const groups = groupOperations(operations);
    for (let index = 0; index < groups.length; index += 1) {
      showProgress(`☁️ Firestore: ${phase} ${index + 1}/${groups.length}`);
      const batch = firestore.batch();
      groups[index].forEach((operation) => {
        if (operation.type === 'delete') batch.delete(operation.ref);
        else batch.set(operation.ref, operation.data, operation.options || {});
      });
      await batch.commit();
    }
    return groups.length;
  }

  async function savePriceListInSmallBatches(record) {
    if (!PV.state.firestore || !record?.id) return false;

    const firestore = PV.state.firestore;
    const listRef = firestore.collection(PV.collections.priceLists).doc(record.id);
    const chunksRef = listRef.collection(CHUNKS_COLLECTION);
    const items = Array.isArray(record.items) ? record.items : [];
    const chunks = splitItems(items);
    const generation = generationId();
    const oldChunksSnapshot = await chunksRef.get();
    const oldChunkRefs = oldChunksSnapshot.docs.map((doc) => doc.ref);

    const writeOperations = chunks.map((chunkItems, index) => ({
      type: 'set',
      ref: chunksRef.doc(`${generation}-${String(index).padStart(5, '0')}`),
      data: {
        generation,
        index,
        itemCount: chunkItems.length,
        estimatedBytes: byteLength(chunkItems),
        items: chunkItems,
        updatedAt: PV.nowIso()
      }
    }));

    showProgress(`☁️ Firestore: preparo ${chunks.length} parti`);
    const batchCount = await commitOperations(firestore, writeOperations, 'invio');

    showProgress('☁️ Firestore: confermo il prezziario');
    const metadata = PV.stripUndefined({ ...record });
    delete metadata.items;
    metadata.storageFormat = FORMAT;
    metadata.chunkGeneration = generation;
    metadata.chunkCount = chunks.length;
    metadata.itemCount = items.length;
    metadata.chunkTargetBytes = TARGET_CHUNK_BYTES;
    metadata.batchMaxWrites = MAX_BATCH_WRITES;
    metadata.batchCount = batchCount;
    metadata.syncPending = false;
    await listRef.set(metadata, { merge: false });

    if (oldChunkRefs.length) {
      await commitOperations(
        firestore,
        oldChunkRefs.map((ref) => ({ type: 'delete', ref })),
        'pulizia'
      );
    }

    PV.state.remoteConnected = true;
    PV.state.remoteDenied = false;
    showProgress('☁️ Firestore sincronizzato', 'ok');
    return true;
  }

  PV.saveRemote = async (collectionName, record) => {
    if (collectionName !== PV.collections.priceLists) {
      return originalSaveRemote(collectionName, record);
    }
    if (!record?.id) return false;
    if (inFlight.has(record.id)) return inFlight.get(record.id);

    const task = savePriceListInSmallBatches(record).catch((error) => {
      if (error?.code === 'permission-denied') PV.state.remoteDenied = true;
      console.warn('Preventivi: invio Firestore in piccoli batch non riuscito.', error);
      showProgress('⚠️ Firestore: dati in attesa', 'warning');
      PV.setFeedback?.(
        `Prezziario salvato sul dispositivo. Firestore: ${error?.message || 'errore di sincronizzazione'}`,
        'warning'
      );
      return false;
    }).finally(() => {
      inFlight.delete(record.id);
    });

    inFlight.set(record.id, task);
    return task;
  };

  PV.loadLocal = () => {
    originalLoadLocal();
    try {
      if (localStorage.getItem(MIGRATION_KEY) === '1') return;
      PV.state.priceLists = (PV.state.priceLists || []).map((item) => ({
        ...item,
        syncPending: true
      }));
      PV.persistLocal();
      localStorage.setItem(MIGRATION_KEY, '1');
    } catch (error) {
      console.warn('Preventivi: coda per piccoli batch non memorizzata.', error);
    }
  };
})();
