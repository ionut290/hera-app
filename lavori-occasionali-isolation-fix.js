(() => {
  "use strict";
  if (window.HeraOccasionalIsolationFix?.installed) return;

  const COMMESSA_ID = "lavori-occasionali";
  const state = {
    plants: new Map(),
    loading: null,
    loaded: false,
    lastLoadAt: 0,
    observer: null
  };

  const text = (value) => String(value ?? "").trim();
  const normalize = (value) => text(value).replace(/\s+/g, " ").toLocaleUpperCase("it-IT");
  const collectionName = () => typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";

  function strictOccasionalPlant(plant) {
    if (!plant || typeof plant !== "object") return false;
    const commessaId = text(plant.commessaId || plant.parentCommessaId || plant.commessa?.id).toLowerCase();
    return commessaId === COMMESSA_ID && plant.lavoroOccasionale === true;
  }

  function normalizePlant(doc) {
    const data = typeof doc?.data === "function" ? doc.data() : doc || {};
    const id = text(doc?.id || data.id || data.docId);
    return { ...data, id: id || data.id, docId: id || data.docId, commessaId: COMMESSA_ID, lavoroOccasionale: true };
  }

  function cachePlants(items) {
    const clean = (Array.isArray(items) ? items : [])
      .map(normalizePlant)
      .filter(strictOccasionalPlant);
    state.plants = new Map(clean.map((plant) => [text(plant.id || plant.docId), plant]));
    state.loaded = true;
    state.lastLoadAt = Date.now();

    try {
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId instanceof Map) {
        impiantiByCommessaId.set(COMMESSA_ID, clean.slice());
      }
    } catch (_) {}

    window.dispatchEvent(new CustomEvent("hera:occasional-plants-updated", { detail: { count: clean.length } }));
    return clean;
  }

  async function loadDirect(force = false) {
    if (!force && state.loaded && Date.now() - state.lastLoadAt < 30000) return getPlants();
    if (state.loading) return state.loading;
    if (typeof db === "undefined" || !db) return [];

    state.loading = db.collection(collectionName()).doc(COMMESSA_ID).collection("impianti").get()
      .then((snapshot) => cachePlants(snapshot.docs))
      .catch((error) => {
        console.error("[LAVORI OCCASIONALI] caricamento dedicato fallito:", error);
        return getPlants();
      })
      .finally(() => { state.loading = null; });
    return state.loading;
  }

  function getPlants() {
    const merged = new Map(state.plants);

    try {
      const cached = impiantiByCommessaId instanceof Map ? impiantiByCommessaId.get(COMMESSA_ID) : null;
      if (Array.isArray(cached)) {
        cached.filter(strictOccasionalPlant).forEach((plant) => merged.set(text(plant.id || plant.docId), plant));
      }
    } catch (_) {}

    // currentImpianti viene consultato solo se il record dichiara esplicitamente la commessa corretta.
    // Non usare mai la commessa selezionata come fallback: evita contaminazioni da altre commesse.
    try {
      if (Array.isArray(currentImpianti)) {
        currentImpianti.filter(strictOccasionalPlant).forEach((plant) => merged.set(text(plant.id || plant.docId), plant));
      }
    } catch (_) {}

    return Array.from(merged.values());
  }

  function cleanForeignOccasionalCache() {
    try {
      if (!(impiantiByCommessaId instanceof Map)) return;
      for (const [commessaId, list] of impiantiByCommessaId.entries()) {
        if (text(commessaId).toLowerCase() === COMMESSA_ID || !Array.isArray(list)) continue;
        const contaminated = list.some((plant) => plant?.lavoroOccasionale === true && text(plant?.commessaId).toLowerCase() === COMMESSA_ID);
        if (contaminated) {
          impiantiByCommessaId.set(commessaId, list.filter((plant) => !(plant?.lavoroOccasionale === true && text(plant?.commessaId).toLowerCase() === COMMESSA_ID)));
        }
      }
    } catch (_) {}
  }

  function refreshDependentModules() {
    cleanForeignOccasionalCache();
    try { window.HeraLavoriOccasionali?.refresh?.(); } catch (_) {}
    try { window.HeraOccasionalMultiSiteHours?.refresh?.(); } catch (_) {}
    try { window.HeraOccasionalPdfStorage?.refresh?.(); } catch (_) {}
  }

  window.addEventListener("hera:occasional-plants-updated", () => queueMicrotask(refreshDependentModules));
  document.addEventListener("change", (event) => {
    if (event.target?.id !== "squadra-commessa") return;
    if (text(event.target.value).toLowerCase() === COMMESSA_ID) void loadDirect(true);
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => void loadDirect(true), { once: true });
  } else {
    void loadDirect(true);
  }

  window.HeraOccasionalIsolationFix = {
    installed: true,
    version: "1.0.0",
    COMMESSA_ID,
    isStrictOccasionalPlant: strictOccasionalPlant,
    getPlants,
    loadDirect,
    refresh: () => loadDirect(true)
  };
})();
