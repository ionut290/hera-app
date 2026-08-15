#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = "admin-user-access-tools.js";
let source = fs.readFileSync(path, "utf8");

function replaceOnce(before, after) {
  if (!source.includes(before)) throw new Error("Blocco atteso non trovato");
  if (source.indexOf(before) !== source.lastIndexOf(before)) throw new Error("Blocco ambiguo");
  source = source.replace(before, after);
}

replaceOnce(
  `  let activeProfile = null;\n  let observer = null;\n  let profilesCache = [];`,
  `  let activeProfile = null;\n  let observer = null;\n  let panelActivated = false;\n  let profilesCache = [];`
);

replaceOnce(
  `  async function enhance() {\n    ensureSearch();\n    await enhanceUserCards();\n    applySearch();\n  }`,
  `  async function enhance() {\n    if (!panelActivated || !isManager()) return;\n    ensureSearch();\n    await enhanceUserCards();\n    applySearch();\n  }`
);

replaceOnce(
  `  function installObserver() {\n    if (observer || !document.body) return;\n    let scheduled = false;\n    observer = new MutationObserver(() => {\n      if (scheduled) return;\n      scheduled = true;\n      window.setTimeout(() => {\n        scheduled = false;\n        void enhance();\n      }, 80);\n    });\n    observer.observe(document.body, { childList: true, subtree: true });\n  }`,
  `  function installObserver() {\n    if (observer || !panelActivated) return;\n    const panel = document.getElementById("panel-utenti");\n    if (!panel) return;\n    let scheduled = false;\n    observer = new MutationObserver(() => {\n      if (scheduled) return;\n      scheduled = true;\n      window.setTimeout(() => {\n        scheduled = false;\n        void enhance();\n      }, 80);\n    });\n    observer.observe(panel, { childList: true, subtree: true });\n  }`
);

replaceOnce(
  `  function initialize() {\n    document.addEventListener("click", interceptLegacyPasswordButton, true);\n    document.addEventListener("click", (event) => {\n      if (event.target?.closest?.("#open-panel-utenti")) {\n        window.setTimeout(() => void enhance(), 120);\n      }\n    }, true);\n    void enhance();\n    installObserver();\n  }`,
  `  function activateUserPanel() {\n    panelActivated = true;\n    installObserver();\n    void enhance();\n  }\n\n  function initialize() {\n    document.addEventListener("click", interceptLegacyPasswordButton, true);\n    document.addEventListener("click", (event) => {\n      if (event.target?.closest?.("#open-panel-utenti")) {\n        window.setTimeout(activateUserPanel, 120);\n      }\n    }, true);\n    const panel = document.getElementById("panel-utenti");\n    if (panel && !panel.classList.contains("hidden")) activateUserPanel();\n  }`
);

replaceOnce(
  `    refresh: enhance,\n    search: applySearch,\n    openPassword: openDialog`,
  `    refresh: enhance,\n    activate: activateUserPanel,\n    search: applySearch,\n    openPassword: openDialog`
);

fs.writeFileSync(path, source);
console.log("Gestione profili admin resa on-demand sul pannello utenti.");
