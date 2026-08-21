(function installPdfCallableIosBridge(global) {
  "use strict";

  const TARGET_URL = "https://europe-west1-hera-app-6cd2b.cloudfunctions.net/uploadWhazzupPdfToDrive";
  const FUNCTION_NAME = "uploadWhazzupPdfToDrive";
  const PDF_SOURCE = "whazzup-impianto-pdf";
  const STORAGE_PATCH_RETRY_MS = 250;
  const STORAGE_PATCH_MAX_RETRIES = 80;
  let storagePatchRetries = 0;

  if (!global.__heraPdfCallableIosBridgeInstalled && typeof global.fetch === "function") {
    const nativeFetch = global.fetch.bind(global);

    function makeResponse(ok, status, payload) {
      return {
        ok,
        status,
        json: async () => payload,
        text: async () => JSON.stringify(payload)
      };
    }

    async function bridgedFetch(input, init) {
      const url = typeof input === "string" ? input : String(input?.url || "");
      if (url !== TARGET_URL) return nativeFetch(input, init);

      try {
        const firebase = global.firebase;
        if (!firebase?.apps?.length || typeof firebase.functions !== "function") {
          return nativeFetch(input, init);
        }

        let body = {};
        try {
          body = init?.body ? JSON.parse(init.body) : {};
        } catch (_) {}
        const data = body?.data || {};

        // Evita che l'SDK Firebase rientri nel bridge se internamente usa fetch.
        global.fetch = nativeFetch;
        let callableResult;
        try {
          const functions = firebase.app().functions("europe-west1");
          const callable = functions.httpsCallable(FUNCTION_NAME);
          callableResult = await callable(data);
        } finally {
          global.fetch = bridgedFetch;
        }

        return makeResponse(true, 200, { result: callableResult?.data || {} });
      } catch (error) {
        console.error("Bridge PDF callable non riuscito:", error);
        const code = String(error?.code || "functions/internal");
        const message = String(error?.message || "Caricamento PDF su Drive non riuscito.");
        return makeResponse(false, 500, { error: { status: code, message } });
      }
    }

    global.fetch = bridgedFetch;
    global.__heraPdfCallableIosBridgeInstalled = true;
  }

  // Firebase Storage del progetto sta rispondendo con 404/retry-limit per i PDF.
  // Per evitare attese lunghe e errori all'utente, intercettiamo SOLO gli upload
  // Whazzup PDF e facciamo scattare immediatamente il fallback Drive già previsto.
  function wrapStorageRef(ref) {
    if (!ref || ref.__heraWhazzupPdfDriveBypassWrapped) return ref;
    try {
      Object.defineProperty(ref, "__heraWhazzupPdfDriveBypassWrapped", {
        value: true,
        configurable: true
      });
    } catch (_) {
      try { ref.__heraWhazzupPdfDriveBypassWrapped = true; } catch (_) {}
    }

    if (typeof ref.child === "function") {
      const nativeChild = ref.child.bind(ref);
      try {
        ref.child = function heraPdfChild(path) {
          return wrapStorageRef(nativeChild(path));
        };
      } catch (_) {}
    }

    if (typeof ref.put === "function") {
      const nativePut = ref.put.bind(ref);
      try {
        ref.put = function heraPdfPut(data, metadata) {
          const source = String(metadata?.customMetadata?.source || "");
          if (source === PDF_SOURCE) {
            const error = new Error("Firebase Storage PDF non disponibile: uso Drive centrale.");
            error.code = "storage/unknown";
            return Promise.reject(error);
          }
          return nativePut(data, metadata);
        };
      } catch (_) {}
    }
    return ref;
  }

  function installStoragePdfBypass() {
    if (global.__heraWhazzupPdfStorageBypassInstalled) return true;
    const firebase = global.firebase;
    if (!firebase?.apps?.length || typeof firebase.storage !== "function") return false;

    try {
      const storage = firebase.storage();
      if (!storage || typeof storage.ref !== "function") return false;
      const nativeRef = storage.ref.bind(storage);
      storage.ref = function heraPdfStorageRef(path) {
        return wrapStorageRef(nativeRef(path));
      };
      // Se qualche chiamata Storage non-PDF resta attiva, non cambiamo il suo comportamento.
      global.__heraWhazzupPdfStorageBypassInstalled = true;
      console.info("PDF Whazzup: Drive centrale impostato come percorso affidabile.");
      return true;
    } catch (error) {
      console.warn("PDF Whazzup: bypass Storage non ancora installato.", error);
      return false;
    }
  }

  function retryStoragePatch() {
    if (installStoragePdfBypass()) return;
    storagePatchRetries += 1;
    if (storagePatchRetries < STORAGE_PATCH_MAX_RETRIES) {
      global.setTimeout(retryStoragePatch, STORAGE_PATCH_RETRY_MS);
    }
  }

  retryStoragePatch();
})(globalThis);
