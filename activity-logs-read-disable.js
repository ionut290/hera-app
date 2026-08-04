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
    window.HeraActivityLogsReadGuard = { installed: true, collection: "activityLogs", mode: "reads-disabled" };
  }

  if (window.HeraActiveCommesse?.installed) return;
  if (typeof db === "undefined" || !db || typeof firebase === "undefined" || !firebase.firestore) return;

  const INDEX_PATH = "appConfig/activeCommesse";
  const indexRef = db.collection("appConfig").doc("activeCommesse");
  const state = { explicit: false, activeIds: new Set(), ready: null, lastError: null };

  const normalizeIds = (value) => [...new Set((Array.isArray(value) ? value : [])
    .map((id) => String(id || "").trim()).filter(Boolean))];
  const isActive = (commessaId) => !state.explicit || state.activeIds.has(String(commessaId || ""));

  function getQueryPath(query) {
    if (typeof query?.path === "string") return query.path;
    const path = query?._query?.path || query?._delegate?._query?.path;
    if (path && typeof path.canonicalString === "function") return path.canonicalString();
    return path ? String(path) : "";
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
      docs: [], empty: true, size: 0, query,
      metadata: { fromCache: true, hasPendingWrites: false },
      forEach() {}, docChanges() { return []; }
    };
  }

  async function loadIndex() {
    try {
      const snapshot = await indexRef.get({ source: "server" }).catch(() => indexRef.get());
      if (!snapshot.exists || !Array.isArray(snapshot.data()?.ids)) {
        state.explicit = false;
        state.activeIds = new Set();
        return;
      }
      state.explicit = true;
      state.activeIds = new Set(normalizeIds(snapshot.data().ids));
    } catch (error) {
      state.lastError = error;
      state.explicit = false;
      console.warn("Indice commesse attive non disponibile: mantengo tutte le commesse attive.", error);
    }
  }
  state.ready = loadIndex();

  const QueryPrototype = firebase.firestore.Query?.prototype;
  if (QueryPrototype && !QueryPrototype.__heraActiveCommesseOriginalOnSnapshot) {
    const originalOnSnapshot = QueryPrototype.onSnapshot;
    Object.defineProperty(QueryPrototype, "__heraActiveCommesseOriginalOnSnapshot", {
      value: originalOnSnapshot, configurable: false, enumerable: false, writable: false
    });
    QueryPrototype.onSnapshot = function activeCommesseOnSnapshotGuard(...args) {
      const commessaId = extractImpiantiCommessaId(getQueryPath(this));
      if (!commessaId) return originalOnSnapshot.apply(this, args);
      let cancelled = false;
      let unsubscribe = () => {};
      state.ready.finally(() => {
        if (cancelled) return;
        if (isActive(commessaId)) unsubscribe = originalOnSnapshot.apply(this, args);
        else {
          const next = findNextCallback(args);
          if (next) queueMicrotask(() => !cancelled && next(emptySnapshot(this)));
          console.info("Listener impianti evitato per commessa disattivata:", commessaId);
        }
      });
      return () => {
        cancelled = true;
        try { unsubscribe(); } catch (error) { console.warn("Errore chiusura listener commessa:", error); }
      };
    };
  }

  function allKnownCommesse() {
    try {
      if (!(commesseById instanceof Map)) return [];
      return [...commesseById.entries()].map(([id, value]) => ({
        id: String(id), nome: String(value?.nome || value?.name || id), codice: String(value?.codice || value?.code || "")
      })).sort((a, b) => a.nome.localeCompare(b.nome, "it"));
    } catch (_) { return []; }
  }

  async function saveActiveIds(ids) {
    const normalized = normalizeIds(ids);
    const user = firebase.auth?.().currentUser;
    await indexRef.set({
      ids: normalized,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedByUid: user?.uid || "",
      updatedByEmail: user?.email || "",
      mode: "non-destructive-listener-filter-v2"
    }, { merge: true });

    const verification = await indexRef.get({ source: "server" }).catch(() => indexRef.get());
    const savedIds = normalizeIds(verification.data()?.ids);
    if (!verification.exists || savedIds.length !== normalized.length || savedIds.some((id, i) => id !== normalized[i])) {
      throw new Error("Firestore non ha confermato lo stato della commessa.");
    }
    state.explicit = true;
    state.activeIds = new Set(savedIds);
    return savedIds;
  }

  function hideInactiveHomeCards() {
    const home = document.getElementById("home-page");
    if (!home) return;
    home.querySelectorAll("[data-commessa-id]").forEach((node) => {
      const id = node.getAttribute("data-commessa-id");
      if (!id) return;
      node.classList.toggle("hidden", state.explicit && !isActive(id));
    });
  }

  function removeManager() {
    document.getElementById("active-commesse-manager")?.remove();
  }

  function renderManager() {
    const host = document.getElementById("commesse-manage-list");
    if (!host || document.getElementById("active-commesse-manager")) return;
    const commesse = allKnownCommesse();
    if (!commesse.length) return;

    const section = document.createElement("section");
    section.id = "active-commesse-manager";
    section.className = "card";
    section.style.margin = "12px 0";
    section.innerHTML = `
      <div class="section-head"><div>
        <h3>Commesse caricate all'avvio</h3>
        <p class="muted">La disattivazione non elimina o modifica dati. Le ore storiche restano nel calendario personale.</p>
      </div></div>
      <div id="active-commesse-manager-list" class="simple-list"></div>
      <p id="active-commesse-manager-feedback" class="muted" role="status"></p>`;
    host.parentNode.insertBefore(section, host);
    const list = section.querySelector("#active-commesse-manager-list");

    commesse.forEach((commessa) => {
      const row = document.createElement("div");
      row.className = "simple-list-item";
      row.style.cssText = "display:flex;align-items:center;justify-content:space-between;gap:12px";
      const active = isActive(commessa.id);
      row.innerHTML = `<div><strong></strong><div class="muted"></div></div>
        <button type="button" class="btn ${active ? "" : "btn-primary"}">${active ? "Disattiva" : "Riattiva"}</button>`;
      row.querySelector("strong").textContent = commessa.nome;
      row.querySelector(".muted").textContent = `${commessa.codice ? `${commessa.codice} · ` : ""}${active ? "Attiva" : "Disattivata"}`;

      row.querySelector("button").addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        const feedback = section.querySelector("#active-commesse-manager-feedback");
        const button = event.currentTarget;
        if (button.disabled) return;
        button.disabled = true;
        feedback.textContent = active ? "Disattivazione in corso…" : "Riattivazione in corso…";
        try {
          const current = state.explicit ? new Set(state.activeIds) : new Set(commesse.map((item) => item.id));
          if (active) current.delete(commessa.id); else current.add(commessa.id);
          await saveActiveIds([...current].sort());
          feedback.textContent = active
            ? "Commessa disattivata e confermata da Firestore. La riduzione dei listener sarà completa alla prossima apertura dell’app."
            : "Commessa riattivata e confermata da Firestore. Sarà caricata alla prossima apertura dell’app.";
          removeManager();
          renderManager();
          hideInactiveHomeCards();
        } catch (error) {
          button.disabled = false;
          feedback.textContent = `Salvataggio non riuscito: ${error?.message || error}`;
        }
      }, true);
      list.appendChild(row);
    });
  }

  state.ready.finally(() => {
    const observer = new MutationObserver(() => { renderManager(); hideInactiveHomeCards(); });
    observer.observe(document.documentElement, { childList: true, subtree: true });
    renderManager();
    hideInactiveHomeCards();
  });

  window.HeraActiveCommesse = {
    installed: true,
    indexPath: INDEX_PATH,
    isActive,
    reloadIndex: loadIndex,
    getState: () => ({ explicit: state.explicit, activeIds: [...state.activeIds], lastError: state.lastError ? String(state.lastError?.message || state.lastError) : "" })
  };
})();
