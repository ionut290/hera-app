(() => {
  "use strict";

  if (window.HeraCantiereDocumentsLoaderInstalled) return;
  window.HeraCantiereDocumentsLoaderInstalled = true;

  const VERSION = "20260823-docs2";
  const COMMESSA_ID = "lavori-occasionali";
  const normalize = (value) => String(value || "").trim().replace(/\s+/g, " ").toLocaleUpperCase("it-IT");

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
      return name && cardText.includes(name);
    }) || null;
  }

  function findCardFromGear(gear) {
    let node = gear?.parentElement || null;
    for (let depth = 0; node && node !== document.body && depth < 10; depth += 1, node = node.parentElement) {
      const buttons = [...node.querySelectorAll("button")];
      const hasFatto = buttons.some((button) => normalize(button.textContent) === "FATTO");
      const hasNav = buttons.some((button) => normalize(button.textContent).includes("NAVIGA"));
      if (hasFatto && hasNav) return node;
    }
    return null;
  }

  function isGear(button) {
    const label = normalize([button?.textContent, button?.title, button?.getAttribute?.("aria-label")].filter(Boolean).join(" "));
    return label.includes("⚙") || label.includes("INGRAN") || label.includes("GESTIONE");
  }

  function installAdminFallback() {
    if (!isAdmin() || !window.HeraCantiereDocuments?.open) return;
    [...document.querySelectorAll("button")].filter(isGear).forEach((gear) => {
      const card = findCardFromGear(gear);
      if (!card || card.querySelector("[data-cantiere-doc-admin], [data-cantiere-doc-fallback]")) return;
      const plant = findPlantForCard(card);
      if (!plant) return;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "cantiere-doc-admin";
      button.dataset.cantiereDocFallback = "1";
      button.innerHTML = "📁<br>DOCUMENTAZIONE";
      button.title = "Documentazione cantiere";
      button.setAttribute("aria-label", "Apri documentazione cantiere");
      button.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        window.HeraCantiereDocuments.open(plant);
      });
      gear.insertAdjacentElement("afterend", button);
    });
  }

  function installObserver() {
    let timer = 0;
    const observer = new MutationObserver(() => {
      window.clearTimeout(timer);
      timer = window.setTimeout(installAdminFallback, 120);
    });
    observer.observe(document.body, { childList: true, subtree: true });
    [0, 500, 1500, 3500, 7000].forEach((delay) => window.setTimeout(installAdminFallback, delay));
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