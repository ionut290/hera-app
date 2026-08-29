(() => {
  "use strict";

  const VERSION = "1.1.0";
  const COMMESSE_CACHE_KEY = "heraCommesseCache";
  const MAX_COMMESSE_CACHE_BYTES = 1_500_000;
  const RESERVE_BYTES = 700_000;
  const SAFE_STORAGE_LIMIT_BYTES = 4_500_000;
  const TARGET_STORAGE_BYTES = SAFE_STORAGE_LIMIT_BYTES - RESERVE_BYTES;
  const DISPOSABLE_KEYS = [
    { priority: 1, prefix: "varga_fs_diag_v4_" },
    { priority: 1, prefix: "varga_fs_diag_v3_" },
    { priority: 1, prefix: "varga_fs_diag_v2_" },
    { priority: 2, exact: "heraTrafficRouteCacheV1" },
    { priority: 2, exact: "heraTrafficWeatherCacheV1" },
    { priority: 3, prefix: "hera-shared-static-view:" },
    { priority: 4, exact: COMMESSE_CACHE_KEY },
    { priority: 5, prefix: "heraImpiantiPersistentCacheV1:" }
  ];

  const state = {
    installed: false,
    removedOversizedCommesseCache: false,
    skippedOversizedCommesseWrites: 0,
    quotaRecoveries: 0,
    disposableKeysRemoved: 0,
    disposableBytesRemoved: 0,
    proactiveCleanups: 0,
    lastError: ""
  };

  function byteSize(value) {
    try {
      return new Blob([String(value ?? "")]).size;
    } catch (_) {
      return String(value ?? "").length * 2;
    }
  }

  function isQuotaError(error) {
    return Boolean(error && (
      error.name === "QuotaExceededError" ||
      error.name === "NS_ERROR_DOM_QUOTA_REACHED" ||
      error.code === 22 ||
      error.code === 1014
    ));
  }

  function safeRemove(storage, key) {
    try {
      storage.removeItem(key);
      return true;
    } catch (_) {
      return false;
    }
  }

  function disposablePriority(key) {
    const normalized = String(key || "");
    const rule = DISPOSABLE_KEYS.find((item) => item.exact === normalized || (item.prefix && normalized.startsWith(item.prefix)));
    return rule?.priority || 0;
  }

  function collectDisposableEntries(storage, excludedKey = "") {
    const entries = [];
    try {
      for (let index = 0; index < storage.length; index += 1) {
        const key = storage.key(index) || "";
        const priority = disposablePriority(key);
        if (!priority || key === excludedKey) continue;
        const value = storage.getItem(key) || "";
        entries.push({ key, priority, bytes: byteSize(key) + byteSize(value) });
      }
    } catch (_) {}
    return entries.sort((left, right) => left.priority - right.priority || right.bytes - left.bytes || left.key.localeCompare(right.key));
  }

  function removeDisposableEntry(storage, entry) {
    if (!entry || !safeRemove(storage, entry.key)) return false;
    state.disposableKeysRemoved += 1;
    state.disposableBytesRemoved += entry.bytes;
    return true;
  }

  function recoverQuotaWrite(storage, key, value, originalSetItem) {
    let lastError = null;
    for (const entry of collectDisposableEntries(storage, key)) {
      if (!removeDisposableEntry(storage, entry)) continue;
      try {
        const result = originalSetItem.call(storage, key, value);
        state.quotaRecoveries += 1;
        return { written: true, result };
      } catch (error) {
        if (!isQuotaError(error)) throw error;
        lastError = error;
      }
    }
    return { written: false, error: lastError };
  }

  function estimatedStorageBytes(storage) {
    let estimated = 0;
    try {
      for (let index = 0; index < storage.length; index += 1) {
        const key = storage.key(index) || "";
        estimated += byteSize(key) + byteSize(storage.getItem(key));
      }
    } catch (_) {}
    return estimated;
  }

  function ensureFirestoreReserve(storage) {
    let estimated = estimatedStorageBytes(storage);
    if (estimated + RESERVE_BYTES <= SAFE_STORAGE_LIMIT_BYTES) return;
    for (const entry of collectDisposableEntries(storage)) {
      if (estimated <= TARGET_STORAGE_BYTES) break;
      if (!removeDisposableEntry(storage, entry)) continue;
      estimated = Math.max(0, estimated - entry.bytes);
      state.proactiveCleanups += 1;
    }
  }

  function cleanupExistingCommesseCache() {
    try {
      const current = window.localStorage?.getItem(COMMESSE_CACHE_KEY);
      if (!current) return;
      if (byteSize(current) <= MAX_COMMESSE_CACHE_BYTES) return;
      safeRemove(window.localStorage, COMMESSE_CACHE_KEY);
      state.removedOversizedCommesseCache = true;
      console.warn("[STORAGE QUOTA GUARD] Cache commesse troppo grande rimossa; i dati saranno ricaricati da Firestore.");
    } catch (error) {
      state.lastError = String(error?.message || error || "");
    }
  }

  function install() {
    if (state.installed) return true;
    const StoragePrototype = window.Storage?.prototype;
    if (!StoragePrototype || typeof StoragePrototype.setItem !== "function") return false;
    if (StoragePrototype.setItem.__heraStorageQuotaGuardWrapped) {
      state.installed = true;
      return true;
    }

    const originalSetItem = StoragePrototype.setItem;

    const wrappedSetItem = function heraQuotaAwareSetItem(key, value) {
      const normalizedKey = String(key ?? "");
      const isLocalStorage = (() => {
        try { return this === window.localStorage; } catch (_) { return false; }
      })();

      if (isLocalStorage && normalizedKey === COMMESSE_CACHE_KEY) {
        const bytes = byteSize(value);
        if (bytes > MAX_COMMESSE_CACHE_BYTES) {
          safeRemove(this, COMMESSE_CACHE_KEY);
          state.skippedOversizedCommesseWrites += 1;
          console.warn(
            `[STORAGE QUOTA GUARD] Cache commesse non salvata: ${(bytes / 1024 / 1024).toFixed(2)} MB supera il limite sicuro di ${(MAX_COMMESSE_CACHE_BYTES / 1024 / 1024).toFixed(2)} MB.`
          );
          return;
        }
      }

      try {
        return originalSetItem.call(this, key, value);
      } catch (error) {
        if (!isLocalStorage || !isQuotaError(error)) throw error;
        const recovery = recoverQuotaWrite(this, normalizedKey, value, originalSetItem);
        if (recovery.written) return recovery.result;

        if (disposablePriority(normalizedKey)) {
          safeRemove(this, normalizedKey);
          if (normalizedKey === COMMESSE_CACHE_KEY) state.skippedOversizedCommesseWrites += 1;
          console.warn("[STORAGE QUOTA GUARD] Cache ricostruibile non salvata per proteggere Firestore.", normalizedKey);
          return;
        }

        const finalError = recovery.error || error;
        state.lastError = String(finalError?.message || finalError || "");
        throw finalError;
      }
    };

    Object.defineProperty(wrappedSetItem, "__heraStorageQuotaGuardWrapped", { value: true });
    Object.defineProperty(wrappedSetItem, "__heraStorageQuotaGuardOriginal", { value: originalSetItem });
    StoragePrototype.setItem = wrappedSetItem;

    cleanupExistingCommesseCache();

    try { ensureFirestoreReserve(window.localStorage); } catch (_) {}

    state.installed = true;
    return true;
  }

  window.HeraStorageQuotaGuard = {
    version: VERSION,
    installed: false,
    install,
    getState: () => ({ ...state })
  };

  install();
  window.HeraStorageQuotaGuard.installed = state.installed;
})();
