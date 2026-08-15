#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

function replaceOnce(path, before, after) {
  const source = fs.readFileSync(path, "utf8");
  if (!source.includes(before)) throw new Error(`Blocco atteso non trovato in ${path}`);
  if (source.indexOf(before) !== source.lastIndexOf(before)) throw new Error(`Blocco ambiguo in ${path}`);
  fs.writeFileSync(path, source.replace(before, after));
}

replaceOnce(
  "active-commesse-first-boot-guard.js",
  `  function loadFallbackState(query) {\n    if (fallbackStatePromise) return fallbackStatePromise;\n    const firestore = query?.firestore;\n    if (!firestore?.collection) {\n      return Promise.resolve({ explicit: false, ids: [] });\n    }\n\n    const indexRead = firestore.collection("appConfig").doc("activeCommesse").get();`,
  `  function loadSharedIndexState(firestore) {\n    if (fallbackStatePromise) return fallbackStatePromise;\n    if (!firestore?.collection) {\n      return Promise.resolve({ explicit: false, ids: [] });\n    }\n\n    const indexRead = firestore.collection("appConfig").doc("activeCommesse").get();`
);

replaceOnce(
  "active-commesse-first-boot-guard.js",
  `    return fallbackStatePromise;\n  }\n\n  function shouldBlockLegacyUserAlerts(path) {`,
  `    return fallbackStatePromise;\n  }\n\n  function loadFallbackState(query) {\n    return loadSharedIndexState(query?.firestore);\n  }\n\n  function shouldBlockLegacyUserAlerts(path) {`
);

replaceOnce(
  "active-commesse-first-boot-guard.js",
  `    install,\n    getState: () => ({ ...state })`,
  `    install,\n    getActiveIndexState: () => loadSharedIndexState(window.firebase?.firestore?.()),\n    getState: () => ({ ...state })`
);

replaceOnce(
  "activity-logs-read-disable.js",
  `  state.ready = db.collection("appConfig").doc("activeCommesse").get()\n    .then((snapshot) => {\n      if (!snapshot.exists) return;\n      const data = snapshot.data() || {};\n      if (!Array.isArray(data.ids)) return;\n      state.explicit = true;\n      state.activeIds = new Set(normalizeIds(data.ids));\n    })`,
  `  const sharedIndexPromise = window.HeraActiveCommesseFirstBootGuard?.getActiveIndexState?.();\n  state.ready = (sharedIndexPromise || db.collection("appConfig").doc("activeCommesse").get())\n    .then((result) => {\n      if (result && Object.prototype.hasOwnProperty.call(result, "explicit")) {\n        state.explicit = result.explicit === true;\n        state.activeIds = new Set(normalizeIds(result.ids));\n        return;\n      }\n      if (!result?.exists) return;\n      const data = result.data() || {};\n      if (!Array.isArray(data.ids)) return;\n      state.explicit = true;\n      state.activeIds = new Set(normalizeIds(data.ids));\n    })`
);

console.log("Lettura appConfig/activeCommesse condivisa tra le due guardie.");
