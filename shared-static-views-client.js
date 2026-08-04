(() => {
  "use strict";

  const original = {
    personale: subscribePersonale,
    mezzi: subscribeMezzi
  };
  const FALLBACK_MS = 9000;
  const registry = {
    unsubscribe: null,
    active: new Set(),
    sourceFallback: new Set(),
    timers: new Map(),
    pending: new Map(),
    lastView: null
  };

  function settle(kind, value) {
    const pending = registry.pending.get(kind) || [];
    registry.pending.delete(kind);
    pending.forEach((resolve) => resolve(value));
  }

  function clearFallbackTimer(kind) {
    const timer = registry.timers.get(kind);
    if (timer) clearTimeout(timer);
    registry.timers.delete(kind);
  }

  function release(kind) {
    clearFallbackTimer(kind);
    registry.active.delete(kind);
    if (registry.active.size || !registry.unsubscribe) return;
    registry.unsubscribe();
    registry.unsubscribe = null;
    registry.lastView = null;
  }

  function useOriginal(kind, reason) {
    if (registry.sourceFallback.has(kind)) return;
    registry.sourceFallback.add(kind);
    clearFallbackTimer(kind);
    registry.active.delete(kind);
    console.warn("[SHARED REGISTRIES] fallback sorgente", { kind, reason });
    Promise.resolve(original[kind]())
      .then(() => settle(kind, false))
      .catch((error) => {
        console.error("[SHARED REGISTRIES] fallback fallito", { kind, error });
        settle(kind, false);
      });
    if (!registry.active.size && registry.unsubscribe) {
      registry.unsubscribe();
      registry.unsubscribe = null;
      registry.lastView = null;
    }
  }

  function validView(view) {
    return Boolean(
      view &&
      view.payload &&
      Array.isArray(view.payload.personale) &&
      Array.isArray(view.payload.mezzi)
    );
  }

  function applyPersonale(records) {
    personaleRecords = records.map((record) =>
      normalizePersonaleDocument({ id: record.id, data: () => record })
    );
    personaleLoadState = { status: "loaded", message: "" };
    renderPersonaleList(ui.personaleLista, personaleRecords, deletePersonale);
    refreshResolvedUserIdentity();
    renderHoursOperatoriOptions();
    renderSquadre();
    renderTodaySummary();
    updateSquadraHintFromSources();
    updateSuggestionLists();
  }

  function applyMezzi(records) {
    mezziRecords = records.map((record) =>
      normalizeMezzoDocument({ id: record.id, data: () => record })
    );
    mezziLoadState = { status: "loaded", message: "" };
    renderMezziList(ui.mezziLista, mezziRecords, deleteMezzo);
    renderTodaySummary();
    updateSquadraHintFromSources();
    updateSuggestionLists();
  }

  function applyView(view, metadata = {}) {
    if (!validView(view)) {
      [...registry.active].forEach((kind) => useOriginal(kind, "payload-non-valido"));
      return;
    }
    registry.lastView = view;
    if (registry.active.has("personale") && !registry.sourceFallback.has("personale")) {
      applyPersonale(view.payload.personale);
      clearFallbackTimer("personale");
      settle("personale", true);
    }
    if (registry.active.has("mezzi") && !registry.sourceFallback.has("mezzi")) {
      applyMezzi(view.payload.mezzi);
      clearFallbackTimer("mezzi");
      settle("mezzi", true);
    }
    console.debug("[SHARED REGISTRIES] registri", {
      personale: view.payload.personale.length,
      mezzi: view.payload.mezzi.length,
      source: metadata.source || "firestore"
    });
  }

  function ensureListener() {
    if (registry.unsubscribe) return;
    if (typeof db === "undefined" || !db?.collection) {
      [...registry.active].forEach((kind) => useOriginal(kind, "firestore-non-disponibile"));
      return;
    }
    registry.unsubscribe = db.collection("sharedStaticViews").doc("registri__corrente")
      .onSnapshot((snapshot) => {
        if (!snapshot.exists) {
          [...registry.active].forEach((kind) => useOriginal(kind, "documento-mancante"));
          return;
        }
        applyView({ id: snapshot.id, ...(snapshot.data() || {}) }, {
          source: snapshot.metadata?.fromCache ? "firestore-cache" : "firestore"
        });
      }, (error) => {
        console.warn("[SHARED REGISTRIES] vista non disponibile", error);
        [...registry.active].forEach((kind) => useOriginal(kind, "errore-listener"));
      });
  }

  function subscribe(kind) {
    if (!currentUser) return Promise.resolve(false);
    if (kind === "personale") stopPersonaleSubscription();
    else stopMezziSubscription();

    registry.active.add(kind);
    registry.sourceFallback.delete(kind);
    const promise = new Promise((resolve) => {
      const pending = registry.pending.get(kind) || [];
      pending.push(resolve);
      registry.pending.set(kind, pending);
    });

    if (kind === "personale") unsubscribePersonale = () => release(kind);
    else unsubscribeMezzi = () => release(kind);

    registry.timers.set(kind, setTimeout(
      () => useOriginal(kind, "timeout"),
      FALLBACK_MS
    ));

    if (registry.lastView) applyView(registry.lastView, { source: "memory" });
    else ensureListener();
    return promise;
  }

  subscribePersonale = () => subscribe("personale");
  subscribeMezzi = () => subscribe("mezzi");

  window.HeraSharedRegistries = {
    installed: true,
    getState: () => ({
      active: [...registry.active],
      fallback: [...registry.sourceFallback],
      listening: Boolean(registry.unsubscribe)
    })
  };
})();
