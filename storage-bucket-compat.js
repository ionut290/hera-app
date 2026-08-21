(function installStorageBucketCompatibility(global) {
  "use strict";

  const LEGACY_PROJECT_ID = "hera-app-6cd2b";
  const EXPECTED_LEGACY_BUCKET = `${LEGACY_PROJECT_ID}.appspot.com`;
  const NEW_STYLE_BUCKET = `${LEGACY_PROJECT_ID}.firebasestorage.app`;

  function normalizeConfig(value) {
    if (!value || typeof value !== "object") return value;
    const projectId = String(value.projectId || "").trim();
    const bucket = String(value.storageBucket || "").trim();
    if (projectId === LEGACY_PROJECT_ID && bucket === NEW_STYLE_BUCKET) {
      return { ...value, storageBucket: EXPECTED_LEGACY_BUCKET };
    }
    return value;
  }

  if (global.firebaseConfig) {
    global.firebaseConfig = normalizeConfig(global.firebaseConfig);
    return;
  }

  let interceptedValue;
  Object.defineProperty(global, "firebaseConfig", {
    configurable: true,
    enumerable: true,
    get() {
      return interceptedValue;
    },
    set(value) {
      interceptedValue = normalizeConfig(value);
      Object.defineProperty(global, "firebaseConfig", {
        configurable: true,
        enumerable: true,
        writable: true,
        value: interceptedValue
      });
    }
  });
})(globalThis);
