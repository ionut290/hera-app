(() => {
  "use strict";

  if (!window.HeraActivityLogsReadGuard?.installed && typeof loadActiveUsersLogs === "function") {
    loadActiveUsersLogs = async function loadActiveUsersLogsDisabled() {
      activeUsersLogs = [];

      if (ui.activeUsersFilterOperator) ui.activeUsersFilterOperator.innerHTML = '<option value="">Tutti operatori</option>';
      if (ui.activeUsersFilterAction) ui.activeUsersFilterAction.innerHTML = '<option value="">Tutte azioni</option>';
      if (ui.activeUsersLogList) {
        ui.activeUsersLogList.classList.remove("hidden");
        ui.activeUsersLogList.innerHTML = '<p class="muted">Registro attività disattivato per evitare letture Firestore.</p>';
      }
      if (ui.activeUsersLogToggle) {
        ui.activeUsersLogToggle.classList.add("hidden");
        ui.activeUsersLogToggle.setAttribute("aria-hidden", "true");
      }

      renderActiveUsersDetail();
      return true;
    };

    window.HeraActivityLogsReadGuard = {
      installed: true,
      collection: "activityLogs",
      mode: "reads-disabled"
    };
  }

  if (window.HeraActiveCommesse?.installed) return;
  if (typeof db === "undefined" || !db || typeof firebase === "undefined" || !firebase.firestore) return;

  const INDEX_PATH = "appConfig/activeCommesse";
  const MAX_ACTIVE_IDS_PER_QUERY = 30;
  const state = {
    explicit: false,
    activeIds: new Set(),
    ready: null,
    lastError: null,
    savingIds: new Set(),
    managerRequested: false,
    managerLoaded: false,
    managerLoading: null,
    managerCommesse: new Map()
  };

  function normalizeIds(value) {
    return [...new Set((Array.isArray(value) ? value : [])
      .map((id) => String(id || "").trim())
      .filter(Boolean))];
  }

  function isActive(commessaId) {
    // Un indice vuoto o non sincronizzato non deve mai svuotare la Home.
    if (!state.explicit || state.activeIds.size === 0) return true;
    return state.activeIds.has(String(commessaId || ""));
  }

  function getQueryPath(query) {
    if (typeof query?.path === "string") return query.path;
    const path = query?._query?.path || query?._delegate?._query?.path;
    if (path && typeof path.canonicalString === "function") return path.canonicalString();
    return path ? String(path) : "";
  }

  function isRootCommesseQuery(path) {
    return String(path || "") === "commesse";
  }

  function extractImpiantiCommessaId(path) {
    const match = String(path || "").match(/^commesse\/([^/]+)\/impianti$/);
    return match ? match[1] : "";
  }

  function findNextCallback(args) {
    for (const arg of args) {
      if (typeof arg === "function") return arg;
      if (arg && typeof arg.next === "function") return arg.next.bind(arg);
    }
    return null;
  }

  function emptySnapshot(query) {
    return {
      docs: [],
      empty: true,
      size: 0,
      query,
      metadata: { fromCache: true, hasPendingWrites: false },
      forEach() {},
      docChanges() { return []; }
    };
  }

  function buildActiveCommesseQuery() {
    if (!state.explicit) return null;
    const ids = [...state.activeIds];
    // Fallback sicuro: se l'indice non contiene ID usa la raccolta completa.
    if (!ids.length) return null;
    if (ids.length > MAX_ACTIVE_IDS_PER_QUERY) {
      console.warn(
        `Indice commesse attive con ${ids.length} ID: superato il limite di ${MAX_ACTIVE_IDS_PER_QUERY}. `
        + "Mantengo temporaneamente la query completa per non interrompere l'app."
      );
      return null;
    }
    return db.collection("commesse").where(
      firebase.firestore.FieldPath.documentId(),
      "in",
      ids
    );
  }

  const sharedIndexPromise = window.HeraActiveCommesseFirstBootGuard?.getActiveIndexState?.();
  state.ready = (sharedIndexPromise || db.collection("appConfig").doc("activeCommesse").get())
    .then((result) => {
      if (result && Object.prototype.hasOwnProperty.call(result, "explicit")) {
        state.explicit = result.explicit === true;
        state.activeIds = new Set(normalizeIds(result.ids));
        return;
      }
      if (!result?.exists) return;
      const data = result.data() || {};
      if (!Array.isArray(data.ids)) return;
      state.explicit = true;
      state.activeIds = new Set(normalizeIds(data.ids));
    })
    .catch((error) => {
      state.lastError = error;
      state.explicit = false;
      console.warn("Indice commesse attive non disponibile: mantengo tutte le commesse attive.", error);
    });

  const QueryPrototype = firebase.firestore.Query?.prototype;
  if (QueryPrototype && !QueryPrototype.__heraActiveCommesseOriginalOnSnapshot) {
    const originalOnSnapshot = QueryPrototype.onSnapshot;
    Object.defineProperty(QueryPrototype, "__heraActiveCommesseOriginalOnSnapshot", {
      value: originalOnSnapshot,
      configurable: false,
      enumerable: false,
      writable: false
    });

    QueryPrototype.onSnapshot = function activeCommesseOnSnapshotGuard(...args) {
      const path = getQueryPath(this);
      const commessaId = extractImpiantiCommessaId(path);
      const rootCommesse = isRootCommesseQuery(path);
      if (!commessaId && !rootCommesse) return originalOnSnapshot.apply(this, args);

      let cancelled = false;
      let unsubscribe = () => {};
      state.ready.finally(() => {
        if (cancelled) return;

        if (rootCommesse) {
          const activeQuery = buildActiveCommesseQuery();
          if (activeQuery === false) {
            const next = findNextCallback(args);
            if (next) queueMicrotask(() => !cancelled && next(emptySnapshot(this)));
            console.info("Listener commesse iniziale evitato: nessuna commessa attiva nell'indice.");
            return;
          }
          unsubscribe = originalOnSnapshot.apply(activeQuery || this, args);
          if (activeQuery) {
            console.info("Listener commesse iniziale limitato agli ID attivi:", [...state.activeIds]);
          }
          return;
        }

        if (isActive(commessaId)) {
          unsubscribe = originalOnSnapshot.apply(this, args);
          return;
        }
        const next = findNextCallback(args);
        if (next) queueMicrotask(() => !cancelled && next(emptySnapshot(this)));
        console.info("Listener impianti evitato per commessa disattivata:", commessaId);
      });

      return () => {
        cancelled = true;
        try { unsubscribe(); } catch (error) { console.warn("Errore chiusura listener commessa:", error); }
      };
    };
  }

  function normalizeCommessaRecord(id, value) {
    return {
      id: String(id || ""),
      nome: String(value?.nome || value?.name || id || "Commessa"),
      codice: String(value?.codice || value?.code || "")
    };
  }

  function rememberManagerCommesse(snapshot) {
    snapshot?.docs?.forEach((doc) => {
      state.managerCommesse.set(String(doc.id), normalizeCommessaRecord(doc.id, doc.data() || {}));
    });
  }

  function allKnownCommesse() {
    const merged = new Map(state.managerCommesse);
    try {
      if (commesseById instanceof Map) {
        [...commesseById.entries()].forEach(([id, value]) => {
          merged.set(String(id), normalizeCommessaRecord(id, value));
        });
      }
    } catch (_) {}
    return [...merged.values()].sort((a, b) => a.nome.localeCompare(b.nome, "it"));
  }

  async function loadAllCommesseForManager() {
    if (state.managerLoaded) return allKnownCommesse();
    if (state.managerLoading) return state.managerLoading;

    state.managerLoading = db.collection("commesse").get()
      .then((snapshot) => {
        rememberManagerCommesse(snapshot);
        state.managerLoaded = true;
        state.lastError = null;
        return allKnownCommesse();
      })
      .catch((error) => {
        state.lastError = error;
        console.error("Caricamento amministrativo completo delle commesse non riuscito:", error);
        throw error;
      })
      .finally(() => {
        state.managerLoading = null;
      });

    return state.managerLoading;
  }

  async function saveActiveIds(ids) {
    const normalized = normalizeIds(ids);
    const user = firebase.auth?.().currentUser;
    const ref = db.collection("appConfig").doc("activeCommesse");

    await ref.set({
      ids: normalized,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedByUid: user?.uid || "",
      updatedByEmail: user?.email || "",
      mode: "active-commesse-startup-query-filter-v3"
    }, { merge: true });

    const verification = await ref.get({ source: "server" });
    const savedIds = normalizeIds(verification.data()?.ids);
    if (savedIds.length !== normalized.length || savedIds.some((id, index) => id !== normalized[index])) {
      throw new Error("Verifica salvataggio non riuscita. Lo stato non è stato applicato.");
    }

    state.explicit = true;
    state.activeIds = new Set(savedIds);
    state.lastError = null;
  }

  function syncOperationalCards() {
    // Il listener Firestore restituisce già soltanto le commesse operative.
    // Non applicare un secondo filtro al DOM: alcune schede usano un codice
    // legacy diverso dal document ID e verrebbero nascoste pur essendo valide.
    document.querySelectorAll("[data-commessa-hidden-by-active-index='true']").forEach((node) => {
      node.classList.remove("hidden");
      delete node.dataset.commessaHiddenByActiveIndex;
    });
  }

  function updateManagerRow(row, commessa) {
    const active = isActive(commessa.id);
    const button = row.querySelector("button[data-active-commessa-toggle]");
    const status = row.querySelector("[data-active-commessa-status]");
    if (status) status.textContent = `${commessa.codice ? `${commessa.codice} · ` : ""}${active ? "Attiva" : "Disattivata"}`;
    if (button) {
      button.textContent = active ? "Disattiva" : "Riattiva";
      button.classList.toggle("btn-primary", !active);
      button.setAttribute("aria-pressed", String(!active));
      button.disabled = state.savingIds.has(commessa.id);
    }
  }

  function managerSignature(commesse) {
    return commesse.map((item) => `${item.id}:${item.nome}:${item.codice}:${isActive(item.id) ? 1 : 0}`).join("|");
  }

  function renderManager() {
    if (!state.managerRequested) return;
    const panel = document.getElementById("panel-commesse");
    const host = document.getElementById("commesse-manage-list");
    if (!panel || !host) return;

    const commesse = allKnownCommesse();
    const signature = managerSignature(commesse);
    const existing = document.getElementById("active-commesse-manager");
    if (existing?.dataset.signature === signature && existing?.dataset.loaded === String(state.managerLoaded)) return;
    existing?.remove();

    const section = document.createElement("section");
    section.id = "active-commesse-manager";
    section.className = "card";
    section.style.margin = "12px 0";
    section.dataset.signature = signature;
    section.dataset.loaded = String(state.managerLoaded);
    section.innerHTML = `
      <div class="section-head">
        <div>
          <h3>Commesse caricate all'avvio</h3>
          <p class="muted">All'avvio vengono lette soltanto le commesse attive. L'elenco completo viene letto esclusivamente quando apri questa gestione. Le ore storiche restano nel calendario personale.</p>
        </div>
      </div>
      <div id="active-commesse-manager-list" class="simple-list"></div>
      <p id="active-commesse-manager-feedback" class="muted" role="status" aria-live="polite"></p>
    `;
    host.parentNode.insertBefore(section, host);
    const list = section.querySelector("#active-commesse-manager-list");
    const feedback = section.querySelector("#active-commesse-manager-feedback");

    if (!state.managerLoaded) {
      feedback.textContent = state.managerLoading
        ? "Caricamento dell'elenco completo delle commesse…"
        : "Apro l'elenco amministrativo completo…";
    }

    if (!commesse.length) {
      list.innerHTML = '<p class="muted">Nessuna commessa disponibile.</p>';
      return;
    }

    commesse.forEach((commessa) => {
      const row = document.createElement("div");
      row.className = "simple-list-item";
      row.dataset.activeCommessaId = commessa.id;
      row.style.display = "flex";
      row.style.alignItems = "center";
      row.style.justifyContent = "space-between";
      row.style.gap = "12px";
      row.innerHTML = `
        <div><strong></strong><div class="muted" data-active-commessa-status></div></div>
        <button type="button" class="btn" data-active-commessa-toggle></button>
      `;
      row.querySelector("strong").textContent = commessa.nome;
      updateManagerRow(row, commessa);

      row.querySelector("button").addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (state.savingIds.has(commessa.id)) return;

        const currentlyActive = isActive(commessa.id);
        if (currentlyActive) {
          const confirmed = window.confirm(
            "La commessa verrà nascosta dalle attività operative e i suoi impianti non saranno caricati automaticamente. "
            + "Le ore già registrate resteranno visibili nel calendario personale. Nessun dato verrà eliminato."
          );
          if (!confirmed) return;
        }

        state.savingIds.add(commessa.id);
        updateManagerRow(row, commessa);
        feedback.textContent = "Salvataggio in corso…";

        try {
          const current = state.explicit
            ? new Set(state.activeIds)
            : new Set(allKnownCommesse().map((item) => item.id));
          if (current.has(commessa.id)) current.delete(commessa.id);
          else current.add(commessa.id);

          await saveActiveIds([...current]);
          updateManagerRow(row, commessa);
          syncOperationalCards();
          feedback.textContent = isActive(commessa.id)
            ? `Commessa “${commessa.nome}” riattivata correttamente. Ricarico i dati operativi…`
            : `Commessa “${commessa.nome}” disattivata correttamente. Le ore storiche restano disponibili. Ricarico i dati operativi…`;
          setTimeout(() => window.location.reload(), 350);
        } catch (error) {
          state.lastError = error;
          feedback.textContent = `Salvataggio non riuscito: ${error?.message || error}`;
          console.error("Cambio stato commessa non riuscito:", commessa.id, error);
        } finally {
          state.savingIds.delete(commessa.id);
          updateManagerRow(row, commessa);
        }
      });
      list.appendChild(row);
    });
  }

  function requestManagerLoad() {
    state.managerRequested = true;
    renderManager();
    loadAllCommesseForManager()
      .then(() => renderManager())
      .catch(() => renderManager());
  }

  state.ready.finally(() => {
    const observer = new MutationObserver(() => {
      renderManager();
      syncOperationalCards();
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });

    const openButton = document.getElementById("open-panel-commesse");
    openButton?.addEventListener("click", () => queueMicrotask(requestManagerLoad));

    syncOperationalCards();
  });

  window.HeraActiveCommesse = {
    installed: true,
    indexPath: INDEX_PATH,
    isActive,
    ready: state.ready,
    refreshUi: syncOperationalCards,
    loadAllForManager: requestManagerLoad,
    getState: () => ({
      explicit: state.explicit,
      activeIds: [...state.activeIds],
      savingIds: [...state.savingIds],
      managerLoaded: state.managerLoaded,
      managerCount: state.managerCommesse.size,
      lastError: state.lastError ? String(state.lastError?.message || state.lastError) : ""
    })
  };
})();
