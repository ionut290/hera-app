// Keep management navigation usable after a registry panel refresh.
// This is a narrow UI repair: it does not read or write Firestore and does not
// touch FATTO, WHAZZUP, plant records or historical links.
(() => {
  "use strict";
  const panels = Object.freeze({
    "open-panel-commesse": "commesse",
    "open-panel-squadre": "squadre",
    "open-panel-personale": "personale",
    "open-panel-mezzi": "mezzi",
    "open-panel-utenti": "utenti",
    "open-panel-global": "global",
    "open-panel-programmazione": "programmazione"
  });
  document.addEventListener("click", (event) => {
    const button = event.target?.closest?.("#open-panel-commesse,#open-panel-squadre,#open-panel-personale,#open-panel-mezzi,#open-panel-utenti,#open-panel-global,#open-panel-programmazione");
    if (!button) return;
    const panel = panels[button.id];
    if (!panel || typeof window.openManagementPanel !== "function") return;
    queueMicrotask(() => {
      const page = document.getElementById("management-page");
      const target = document.getElementById("panel-" + panel);
      if (page?.classList.contains("hidden") || target?.classList.contains("hidden")) {
        try { window.openManagementPanel(panel); } catch (error) {
          console.warn("Navigazione pannello non riuscita:", error);
        }
      }
    });
  });
})();
