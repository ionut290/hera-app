(function () {
  "use strict";

  const DB_NAME = "hera-registry-cache";
  const DB_VERSION = 1;
  const STORE = "datasets";
  const MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;
  const KEYS = {
    personale: "personale",
    mezzi: "mezzi"
  };

  function openDb() {
    return new Promise((resolve, reject) => {
      if (!("indexedDB" in window)) {
        reject(new Error("IndexedDB non disponibile"));
        return;
      }
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: "key" });
        }
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

  async function write(key, records, version) {
    if (!Array.isArray(records)) return;
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).put({
        key,
        records,
        version: version || "",
        savedAt: Date.now()
      });
      tx.oncomplete = () => {
        db.close();
        resolve();
      };
      tx.onerror = () => {
        db.close();
        reject(tx.error || new Error("Salvataggio cache non riuscito"));
      };
    });
  }

  function sameDataset(a, b) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    try {
      return JSON.stringify(a) === JSON.stringify(b);
    } catch (_) {
      return false;
    }
  }

  function dispatch(name, detail) {
    window.dispatchEvent(new CustomEvent(name, { detail }));
  }

  async function restore() {
    try {
      const [personale, mezzi] = await Promise.all([
        read(KEYS.personale),
        read(KEYS.mezzi)
      ]);

      if (personale?.records?.length && Date.now() - personale.savedAt <= MAX_AGE_MS) {
        window.personaleRecords = personale.records;
        dispatch("hera:registry-cache-restored", { type: "personale", count: personale.records.length });
      }

      if (mezzi?.records?.length && Date.now() - mezzi.savedAt <= MAX_AGE_MS) {
        window.mezziRecords = mezzi.records;
        dispatch("hera:registry-cache-restored", { type: "mezzi", count: mezzi.records.length });
      }
    } catch (error) {
      console.warn("Cache personale/mezzi non ripristinata:", error);
    }
  }

  async function persistCurrent() {
    try {
      const tasks = [];
      if (Array.isArray(window.personaleRecords)) {
        const cached = await read(KEYS.personale).catch(() => null);
        if (!cached || !sameDataset(cached.records, window.personaleRecords)) {
          tasks.push(write(KEYS.personale, window.personaleRecords));
        }
      }
      if (Array.isArray(window.mezziRecords)) {
        const cached = await read(KEYS.mezzi).catch(() => null);
        if (!cached || !sameDataset(cached.records, window.mezziRecords)) {
          tasks.push(write(KEYS.mezzi, window.mezziRecords));
        }
      }
      await Promise.all(tasks);
    } catch (error) {
      console.warn("Cache personale/mezzi non aggiornata:", error);
    }
  }

  restore();

  const observer = new MutationObserver(() => {
    clearTimeout(observer.timer);
    observer.timer = setTimeout(persistCurrent, 350);
  });

  window.addEventListener("DOMContentLoaded", () => {
    observer.observe(document.body, { childList: true, subtree: true });
    setTimeout(persistCurrent, 1200);
  }, { once: true });

  window.addEventListener("beforeunload", persistCurrent);
  window.HeraRegistryDeviceCache = { restore, persistCurrent, read, write };
})();
