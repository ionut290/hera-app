(function () {
  "use strict";

  if (window.HeraRegistryDeviceCache) return;

  const DB_NAME = "hera-registry-cache";
  const DB_VERSION = 1;
  const STORE = "datasets";
  const MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;
  const KEYS = { personale: "personale", mezzi: "mezzi" };

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
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readonly");
      const request = tx.objectStore(STORE).get(key);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error || new Error("Lettura cache non riuscita"));
      tx.oncomplete = () => db.close();
    });
  }

  async function write(key, records) {
    if (!Array.isArray(records)) return;
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).put({ key, records, savedAt: Date.now() });
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error || new Error("Salvataggio cache non riuscito")); };
    });
  }

  function sameDataset(a, b) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    try { return JSON.stringify(a) === JSON.stringify(b); } catch (_) { return false; }
  }

  function notify(type, count) {
    window.dispatchEvent(new CustomEvent("hera:registry-cache-restored", { detail: { type, count } }));
  }

  async function restore() {
    try {
      const [personale, mezzi] = await Promise.all([read(KEYS.personale), read(KEYS.mezzi)]);
      const personaleCorrente = Array.isArray(window.personaleRecords) ? window.personaleRecords : [];
      const mezziCorrenti = Array.isArray(window.mezziRecords) ? window.mezziRecords : [];

      if (!personaleCorrente.length && personale?.records?.length && Date.now() - personale.savedAt <= MAX_AGE_MS) {
        window.personaleRecords = personale.records;
        notify("personale", personale.records.length);
      }
      if (!mezziCorrenti.length && mezzi?.records?.length && Date.now() - mezzi.savedAt <= MAX_AGE_MS) {
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
        const cached = await read(KEYS.personale).catch(() => null);
        if (!cached || !sameDataset(cached.records, window.personaleRecords)) tasks.push(write(KEYS.personale, window.personaleRecords));
      }
      if (Array.isArray(window.mezziRecords) && window.mezziRecords.length) {
        const cached = await read(KEYS.mezzi).catch(() => null);
        if (!cached || !sameDataset(cached.records, window.mezziRecords)) tasks.push(write(KEYS.mezzi, window.mezziRecords));
      }
      await Promise.all(tasks);
    } catch (error) {
      console.warn("Cache personale/mezzi non aggiornata:", error);
    }
  }

  window.HeraRegistryDeviceCache = { restore, persistCurrent, read, write };
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
