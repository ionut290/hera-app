(function installPdfCallableIosBridge(global) {
  "use strict";

  const TARGET_URL = "https://europe-west1-hera-app-6cd2b.cloudfunctions.net/uploadWhazzupPdfToDrive";
  const FUNCTION_NAME = "uploadWhazzupPdfToDrive";
  if (global.__heraPdfCallableIosBridgeInstalled) return;
  if (typeof global.fetch !== "function") return;

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
})(globalThis);
