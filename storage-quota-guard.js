(() => {
  "use strict";

  const VERSION = "1.0.0";
  const COMMESSE_CACHE_KEY = "heraCommesseCache";
  const MAX_COMMESSE_CACHE_BYTES = 1_500_000;
  const RESERVE_BYTES = 350_000;

  const state = {
    installed: false,
    removedOversizedCommesseCache: false,
    skippedOversizedCommesseWrites: 0,
    quotaRecoveries: 0,
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

        const removed = safeRemove(this, COMMESSE_CACHE_KEY);
        if (removed) state.quotaRecoveries += 1;

        try {
          return originalSetItem.call(this, key, value);
        } catch (retryError) {
          if (!isQuotaError(retryError)) throw retryError;

          if (normalizedKey === COMMESSE_CACHE_KEY) {
            state.skippedOversizedCommesseWrites += 1;
            console.warn("[STORAGE QUOTA GUARD] Cache commesse saltata per quota browser esaurita.");
            return;
          }

          state.lastError = String(retryError?.message || retryError || "");
          throw retryError;
        }
      }
    };

    Object.defineProperty(wrappedSetItem, "__heraStorageQuotaGuardWrapped", { value: true });
    Object.defineProperty(wrappedSetItem, "__heraStorageQuotaGuardOriginal", { value: originalSetItem });
    StoragePrototype.setItem = wrappedSetItem;

    cleanupExistingCommesseCache();

    try {
      let estimated = 0;
      for (let i = 0; i < window.localStorage.length; i += 1) {
        const key = window.localStorage.key(i);
        estimated += byteSize(key) + byteSize(window.localStorage.getItem(key));
      }
      if (estimated + RESERVE_BYTES > 4_500_000) {
        safeRemove(window.localStorage, COMMESSE_CACHE_KEY);
      }
    } catch (_) {}

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
