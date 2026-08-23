(() => {
  "use strict";

  if (window.HeraCantiereDocumentsLoaderInstalled) return;
  window.HeraCantiereDocumentsLoaderInstalled = true;

  const VERSION = "20260823-docs3";
  const COMMESSA_ID = "lavori-occasionali";
  const normalize = (value) => String(value || "").trim().replace(/\s+/g, " ").toLocaleUpperCase("it-IT");
  let listObserver = null;
  let observedList = null;

  function ensureStyle() {
    if (document.querySelector('link[data-cantiere-doc-style]')) return;
    const style = document.createElement("link");
    style.rel = "stylesheet";
    style.href = `documentazione-cantiere.css?v=${VERSION}`;
    style.dataset.cantiereDocStyle = "1";
    document.head.appendChild(style);
  }

  function ensureScript() {
    if (window.HeraCantiereDocuments?.installed) return Promise.resolve();
    const existing = document.querySelector('script[data-cantiere-doc-script]');
    if (existing) {
      return new Promise((resolve) => {
        if (window.HeraCantiereDocuments?.installed) return resolve();
        existing.addEventListener("load", resolve, { once: true });
        window.setTimeout(resolve, 1800);
      });
    }
    return new Promise((resolve) => {
      const script = document.createElement("script");
      script.src = `documentazione-cantiere.js?v=${VERSION}`;
      script.async = false;
      script.dataset.cantiereDocScript = "1";
      script.onload = resolve;
      script.onerror = () => {
        console.error("Caricamento Documentazione cantiere non riuscito.");
        resolve();
      };
      document.head.appendChild(script);
    });
  }

  function isAdmin() {
    try {
      if (typeof window.canManageData === "function" && window.canManageData()) return true;
    } catch (_) {}
    try {
      if (typeof canManageData === "function" && canManageData()) return true;
    } catch (_) {}
    try {
      const email = String(window.firebase?.auth?.().currentUser?.email || "").trim().toLowerCase();
      if (email === "ionut29019@gmail.com") return true;
    } catch (_) {}
    return false;
  }

  function occasionalPlants() {
    const result = new Map();
    const add = (plant) => {
      if (!plant || plant.lavoroOccasionale !== true) return;
      const id = String(plant.id || plant.docId || plant.impiantoId || plant.denominazione || plant.nome || "");
      if (id) result.set(id, plant);
    };
    try { if (Array.isArray(window.currentImpianti)) window.currentImpianti.forEach(add); } catch (_) {}
    try { if (Array.isArray(currentImpianti)) currentImpianti.forEach(add); } catch (_) {}
    try {
      const source = window.impiantiByCommessaId || (typeof impiantiByCommessaId !== "undefined" ? impiantiByCommessaId : null);
      const cached = source instanceof Map ? source.get(COMMESSA_ID) : null;
      if (Array.isArray(cached)) cached.forEach(add);
    } catch (_) {}
    return [...result.values()];
  }

  function findPlantForCard(card) {
    const cardText = normalize(card?.textContent);
    if (!cardText) return null;
    return occasionalPlants().find((plant) => {
      const name = normalize(plant.denominazione || plant.nome || plant.impianto);
      const id = normalize(plant.id || plant.docId || plant.impiantoId || plant.idSap || plant["ID SAP"]);
      return (name && cardText.includes(name)) || (id && cardText.includes(id));
    }) || null;
  }

  function cardFromManagementStack(stack) {
    if (!stack) return null;
    let node = stack.parentElement;
    for (let depth = 0; node && node !== document.body && depth < 8; depth += 1, node = node.parentElement) {
      if (node.querySelector?.('.impianto-primary-actions [data-action-key="navigate"]')) return node;
    }
    return null;
  }

  function addDocumentationButton(stack) {
    if (!(stack instanceof HTMLElement)) return;
    if (!isAdmin() || !window.HeraCantiereDocuments?.open) return;
    if (stack.querySelector("[data-cantiere-doc-admin], [data-cantiere-doc-fallback]")) return;

    const gear = stack.querySelector(".gestione-toggle-btn");
    const managementActions = stack.querySelector(".item-actions-gestione");
    if (!gear || !managementActions) return;

    const card = cardFromManagementStack(stack);
    const plant = findPlantForCard(card);
    if (!plant) return;

    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn cantiere-doc-admin";
    button.dataset.cantiereDocFallback = "1";
    button.innerHTML = "📁 ALLEGA DOCUMENTAZIONE";
    button.title = "Allega o gestisci documentazione cantiere";
    button.setAttribute("aria-label", "Allega documentazione cantiere");
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      window.HeraCantiereDocuments.open(plant);
    });

    // Dentro il pannello Gestione: resta nascosto finché non si preme l'ingranaggio.
    managementActions.appendChild(button);
  }

  function decorateList(list = document.getElementById("impianti-lista")) {
    if (!list) return;
    list.querySelectorAll(".impianto-management-stack").forEach(addDocumentationButton);
  }

  function bindList() {
    const list = document.getElementById("impianti-lista");
    if (!list) return false;
    if (list === observedList) {
      decorateList(list);
      return true;
    }

    listObserver?.disconnect();
    observedList = list;
    decorateList(list);
    listObserver = new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        for (const node of mutation.addedNodes) {
          if (!(node instanceof HTMLElement)) continue;
          if (node.matches?.(".impianto-management-stack")) addDocumentationButton(node);
          node.querySelectorAll?.(".impianto-management-stack").forEach(addDocumentationButton);
        }
      }
    });
    listObserver.observe(list, { childList: true, subtree: true });
    return true;
  }

  function installAdminFallback() {
    if (!isAdmin() || !window.HeraCantiereDocuments?.open) return;
    bindList();
  }

  function installObserver() {
    bindList();
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      if (bindList() || attempts >= 20) window.clearInterval(timer);
    }, 250);
  }

  ensureStyle();
  ensureScript().finally(() => {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", () => {
        installObserver();
        installAdminFallback();
      }, { once: true });
    } else {
      installObserver();
      installAdminFallback();
    }
  });
})();