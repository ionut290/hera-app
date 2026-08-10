(() => {
  "use strict";

  if (!window.firebase || typeof firebase.initializeApp !== "function" || !window.firebaseConfig) return;
  if (!firebase.apps?.length) firebase.initializeApp(window.firebaseConfig);
  if (typeof firebase.auth !== "function") return;

  const auth = firebase.auth();
  if (!auth || auth.__heraCompatiblePersistenceInstalled || typeof auth.setPersistence !== "function") return;

  const originalSetPersistence = auth.setPersistence.bind(auth);
  const persistence = firebase.auth.Auth?.Persistence || {};
  const unsupportedCodes = new Set(["auth/unsupported-persistence-type", "auth/invalid-persistence-type"]);

  auth.setPersistence = async (requestedPersistence) => {
    try {
      return await originalSetPersistence(requestedPersistence);
    } catch (error) {
      const code = String(error?.code || "").toLowerCase();
      if (!unsupportedCodes.has(code)) throw error;
      console.warn("Persistenza Firebase richiesta non supportata; applico un fallback compatibile con Opera.", error);
    }

    if (persistence.SESSION && requestedPersistence !== persistence.SESSION) {
      try {
        return await originalSetPersistence(persistence.SESSION);
      } catch (error) {
        const code = String(error?.code || "").toLowerCase();
        if (!unsupportedCodes.has(code)) throw error;
      }
    }

    if (persistence.NONE) {
      try {
        return await originalSetPersistence(persistence.NONE);
      } catch (error) {
        console.warn("Persistenza Firebase temporanea non disponibile; continuo con la modalità del browser.", error);
      }
    }
    return undefined;
  };

  Object.defineProperty(auth, "__heraCompatiblePersistenceInstalled", {
    value: true,
    configurable: false,
    enumerable: false
  });
})();
