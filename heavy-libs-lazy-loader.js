(() => {
  "use strict";

  if (window.HeraHeavyLibs?.installed) return;

  const DEFINITIONS = {
    xlsx: {
      src: "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
      ready: () => Boolean(window.XLSX)
    },
    exceljs: {
      src: "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js",
      ready: () => Boolean(window.ExcelJS?.Workbook)
    },
    html2canvas: {
      src: "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js",
      ready: () => typeof window.html2canvas === "function"
    },
    jspdf: {
      src: "https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js",
      ready: () => Boolean(window.jspdf?.jsPDF)
    }
  };

  const pending = new Map();
  const stats = {
    installedAt: new Date().toISOString(),
    requested: {},
    loaded: {},
    failures: {}
  };

  function increment(bucket, key) {
    bucket[key] = (Number(bucket[key]) || 0) + 1;
  }

  function findExistingScript(src) {
    const wanted = new URL(src, document.baseURI).href;
    return Array.from(document.scripts || []).find((script) => {
      try { return new URL(script.src, document.baseURI).href === wanted; } catch (_) { return false; }
    }) || null;
  }

  function ensure(name) {
    const definition = DEFINITIONS[name];
    if (!definition) return Promise.reject(new Error(`Libreria sconosciuta: ${name}`));
    increment(stats.requested, name);
    if (definition.ready()) return Promise.resolve(true);
    if (pending.has(name)) return pending.get(name);

    const promise = new Promise((resolve, reject) => {
      const settleReady = () => {
        if (!definition.ready()) {
          const error = new Error(`Libreria ${name} caricata ma non disponibile`);
          increment(stats.failures, name);
          reject(error);
          return;
        }
        increment(stats.loaded, name);
        resolve(true);
      };
      const fail = () => {
        const error = new Error(`Caricamento libreria ${name} non riuscito`);
        increment(stats.failures, name);
        reject(error);
      };

      const existing = findExistingScript(definition.src);
      if (existing) {
        existing.addEventListener("load", settleReady, { once: true });
        existing.addEventListener("error", fail, { once: true });
        if (definition.ready()) settleReady();
        return;
      }

      const script = document.createElement("script");
      script.src = definition.src;
      script.async = true;
      script.dataset.heraHeavyLib = name;
      script.addEventListener("load", settleReady, { once: true });
      script.addEventListener("error", fail, { once: true });
      document.head.appendChild(script);
    }).finally(() => pending.delete(name));

    pending.set(name, promise);
    return promise;
  }

  function ensureMany(names) {
    return Promise.all(names.map(ensure));
  }

  const replayed = new WeakSet();

  async function interceptClick(event) {
    const target = event.target?.closest?.("button, [role='button']");
    if (!target || replayed.has(target)) return;

    let libraries = null;
    if (["hours-table-export-btn", "hours-table-export-global-btn"].includes(target.id)) {
      libraries = ["exceljs"];
    } else if ([
      "download-excel-template-btn", "export-all-impianti-btn", "download-price-template-btn",
      "export-price-list-btn"
    ].includes(target.id)) {
      libraries = ["xlsx"];
    } else if (target.id === "business-card-share-btn") {
      libraries = ["html2canvas"];
    }

    if (!libraries || libraries.every((name) => DEFINITIONS[name].ready())) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    try {
      await ensureMany(libraries);
      replayed.add(target);
      target.click();
    } catch (error) {
      console.error("Caricamento funzione avanzata non riuscito:", error);
      window.alert("Funzione non disponibile in questo momento. Controlla la connessione e riprova.");
    } finally {
      replayed.delete(target);
    }
  }

  const fileInputIds = new Set([
    "excel-file", "price-file", "global-excel-file", "personale-excel-file", "mezzi-excel-file"
  ]);
  const replayedChanges = new WeakSet();

  async function interceptFileChange(event) {
    const input = event.target;
    if (!(input instanceof HTMLInputElement) || input.type !== "file" || !fileInputIds.has(input.id)) return;
    if (replayedChanges.has(input) || DEFINITIONS.xlsx.ready()) return;
    event.stopImmediatePropagation();
    try {
      await ensure("xlsx");
      replayedChanges.add(input);
      input.dispatchEvent(new Event("change", { bubbles: true }));
    } catch (error) {
      console.error("Caricamento import Excel non riuscito:", error);
      window.alert("Import Excel non disponibile in questo momento. Controlla la connessione e riprova.");
    } finally {
      replayedChanges.delete(input);
    }
  }

  const replayedForms = new WeakSet();

  async function interceptSubmit(event) {
    const form = event.target;
    if (!(form instanceof HTMLFormElement) || form.id !== "segnalazione-form" || replayedForms.has(form)) return;
    if (DEFINITIONS.html2canvas.ready() && DEFINITIONS.jspdf.ready()) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    try {
      await ensureMany(["html2canvas", "jspdf"]);
      replayedForms.add(form);
      if (typeof form.requestSubmit === "function") form.requestSubmit();
      else form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
    } catch (error) {
      console.error("Caricamento generazione PDF non riuscito:", error);
      window.alert("Generazione PDF non disponibile in questo momento. Controlla la connessione e riprova.");
    } finally {
      replayedForms.delete(form);
    }
  }

  document.addEventListener("click", interceptClick, true);
  document.addEventListener("change", interceptFileChange, true);
  document.addEventListener("submit", interceptSubmit, true);

  window.HeraHeavyLibs = {
    installed: true,
    version: "1.0.0",
    ensure,
    ensureMany,
    isReady: (name) => Boolean(DEFINITIONS[name]?.ready?.()),
    getStats: () => JSON.parse(JSON.stringify(stats))
  };
})();