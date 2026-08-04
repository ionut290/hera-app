(function () {
  "use strict";

  if (window.HeraRegistryDeviceCache) return;

  const DB_NAME = "hera-registry-cache";
  const DB_VERSION = 1;
  const STORE = "datasets";
  const MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;
  const KEYS = { personale: "personale", mezzi: "mezzi" };

  function isValidKey(key) {
    return key === KEYS.personale || key === KEYS.mezzi;
  }

  function recordId(record) {
    return String(record?.__heraDocId || record?.id || record?.docId || record?._id || "").trim();
  }

  function openDb() {
    return new Promise((resolve, reject) => {
      if (!("indexedDB" in window)) return reject(new Error("IndexedDB non disponibile"));
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: "key" });
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Apertura cache non riuscita"));
    });
  }

  async function read(key) {
    if (!isValidKey(key)) return null;
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readonly");
      const request = tx.objectStore(STORE).get(key);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error || new Error("Lettura cache non riuscita"));
      tx.oncomplete = () => db.close();
    });
  }

  async function readFresh(key, maxAgeMs = MAX_AGE_MS) {
    const cached = await read(key);
    if (!cached || !Array.isArray(cached.records) || !cached.records.length) return null;
    const age = Date.now() - Number(cached.savedAt || 0);
    if (!Number.isFinite(age) || age < 0 || age > maxAgeMs) return null;
    return cached;
  }

  async function write(key, records) {
    if (!isValidKey(key) || !Array.isArray(records) || !records.length) return false;
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).put({ key, records, savedAt: Date.now() });
      tx.oncomplete = () => { db.close(); resolve(true); };
      tx.onerror = () => { db.close(); reject(tx.error || new Error("Salvataggio cache non riuscito")); };
    });
  }

  async function remove(key) {
    if (!isValidKey(key)) return false;
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).delete(key);
      tx.oncomplete = () => { db.close(); resolve(true); };
      tx.onerror = () => { db.close(); reject(tx.error || new Error("Eliminazione cache non riuscita")); };
    });
  }

  function sameDataset(a, b) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    try { return JSON.stringify(a) === JSON.stringify(b); } catch (_) { return false; }
  }

  async function writeIfChanged(key, records) {
    if (!isValidKey(key) || !Array.isArray(records) || !records.length) return false;
    const cached = await read(key).catch(() => null);
    if (cached && sameDataset(cached.records, records)) return false;
    await write(key, records);
    return true;
  }

  async function patchRecord(key, id, patch) {
    if (!isValidKey(key) || !id || !patch || typeof patch !== "object") return false;
    const cached = await read(key).catch(() => null);
    if (!cached?.records?.length) return false;
    const index = cached.records.findIndex((record) => recordId(record) === String(id));
    if (index < 0) return false;

    const current = cached.records[index] || {};
    const next = { ...current, ...patch };
    try {
      if (JSON.stringify(current) === JSON.stringify(next)) return false;
    } catch (_) {}

    const records = cached.records.slice();
    records[index] = next;
    await write(key, records);
    return true;
  }

  function notify(type, count) {
    window.dispatchEvent(new CustomEvent("hera:registry-cache-restored", { detail: { type, count } }));
  }

  async function restore() {
    try {
      const [personale, mezzi] = await Promise.all([
        readFresh(KEYS.personale),
        readFresh(KEYS.mezzi)
      ]);
      const personaleCorrente = Array.isArray(window.personaleRecords) ? window.personaleRecords : [];
      const mezziCorrenti = Array.isArray(window.mezziRecords) ? window.mezziRecords : [];

      if (!personaleCorrente.length && personale?.records?.length) {
        window.personaleRecords = personale.records;
        notify("personale", personale.records.length);
      }
      if (!mezziCorrenti.length && mezzi?.records?.length) {
        window.mezziRecords = mezzi.records;
        notify("mezzi", mezzi.records.length);
      }
    } catch (error) {
      console.warn("Cache personale/mezzi non ripristinata:", error);
    }
  }

  async function persistCurrent() {
    try {
      const tasks = [];
      if (Array.isArray(window.personaleRecords) && window.personaleRecords.length) {
        tasks.push(writeIfChanged(KEYS.personale, window.personaleRecords));
      }
      if (Array.isArray(window.mezziRecords) && window.mezziRecords.length) {
        tasks.push(writeIfChanged(KEYS.mezzi, window.mezziRecords));
      }
      await Promise.all(tasks);
    } catch (error) {
      console.warn("Cache personale/mezzi non aggiornata:", error);
    }
  }

  window.HeraRegistryDeviceCache = {
    restore,
    persistCurrent,
    read,
    readFresh,
    write,
    writeIfChanged,
    patchRecord,
    remove,
    maxAgeMs: MAX_AGE_MS
  };
  restore();

  const observer = new MutationObserver(() => {
    clearTimeout(observer.timer);
    observer.timer = setTimeout(persistCurrent, 500);
  });

  function startObserving() {
    if (!document.body || observer.started) return;
    observer.started = true;
    observer.observe(document.body, { childList: true, subtree: true });
    setTimeout(persistCurrent, 1500);
  }

  if (document.readyState === "loading") {
    window.addEventListener("DOMContentLoaded", startObserving, { once: true });
  } else {
    startObserving();
  }
})();
