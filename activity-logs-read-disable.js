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
  const state = {
    explicit: false,
    activeIds: new Set(),
    ready: null,
    lastError: null,
    savingIds: new Set()
  };

  function normalizeIds(value) {
    return [...new Set((Array.isArray(value) ? value : [])
      .map((id) => String(id || "").trim())
      .filter(Boolean))];
  }

  function isActive(commessaId) {
    return !state.explicit || state.activeIds.has(String(commessaId || ""));
  }

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
      docs: [],
      empty: true,
      size: 0,
      query,
      metadata: { fromCache: true, hasPendingWrites: false },
      forEach() {},
      docChanges() { return []; }
    };
  }

  state.ready = db.collection("appConfig").doc("activeCommesse").get()
    .then((snapshot) => {
      if (!snapshot.exists) return;
      const data = snapshot.data() || {};
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
      const commessaId = extractImpiantiCommessaId(getQueryPath(this));
      if (!commessaId) return originalOnSnapshot.apply(this, args);

      let cancelled = false;
      let unsubscribe = () => {};
      state.ready.finally(() => {
        if (cancelled) return;
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

  function allKnownCommesse() {
    try {
      if (!(commesseById instanceof Map)) return [];
      return [...commesseById.entries()].map(([id, value]) => ({
        id: String(id),
        nome: String(value?.nome || value?.name || id),
        codice: String(value?.codice || value?.code || "")
      })).sort((a, b) => a.nome.localeCompare(b.nome, "it"));
    } catch (_) {
      return [];
    }
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
      mode: "non-destructive-listener-filter-v2"
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
    if (!state.explicit) return;
    document.querySelectorAll("[data-commessa-id]").forEach((node) => {
      const id = String(node.getAttribute("data-commessa-id") || "");
      if (!id || node.closest("#active-commesse-manager")) return;
      node.classList.toggle("hidden", !isActive(id));
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

  function renderManager() {
    const panel = document.getElementById("panel-commesse");
    const host = document.getElementById("commesse-manage-list");
    if (!panel || !host || document.getElementById("active-commesse-manager")) return;
    const commesse = allKnownCommesse();
    if (!commesse.length) return;

    const section = document.createElement("section");
    section.id = "active-commesse-manager";
    section.className = "card";
    section.style.margin = "12px 0";
    section.innerHTML = `
      <div class="section-head">
        <div>
          <h3>Commesse caricate all'avvio</h3>
          <p class="muted">Disattivare una commessa non elimina, sposta o modifica alcun dato. Le ore storiche restano nel calendario personale.</p>
        </div>
      </div>
      <div id="active-commesse-manager-list" class="simple-list"></div>
      <p id="active-commesse-manager-feedback" class="muted" role="status" aria-live="polite"></p>
    `;
    host.parentNode.insertBefore(section, host);
    const list = section.querySelector("#active-commesse-manager-list");

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

        const feedback = section.querySelector("#active-commesse-manager-feedback");
        state.savingIds.add(commessa.id);
        updateManagerRow(row, commessa);
        feedback.textContent = "Salvataggio in corso…";

        try {
          const current = state.explicit
            ? new Set(state.activeIds)
            : new Set(commesse.map((item) => item.id));
          if (current.has(commessa.id)) current.delete(commessa.id);
          else current.add(commessa.id);

          await saveActiveIds([...current]);
          updateManagerRow(row, commessa);
          syncOperationalCards();
          feedback.textContent = isActive(commessa.id)
            ? `Commessa “${commessa.nome}” riattivata correttamente.`
            : `Commessa “${commessa.nome}” disattivata correttamente. Le ore storiche restano disponibili.`;
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

  state.ready.finally(() => {
    const observer = new MutationObserver(() => {
      renderManager();
      syncOperationalCards();
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });
    renderManager();
    syncOperationalCards();
  });

  window.HeraActiveCommesse = {
    installed: true,
    indexPath: INDEX_PATH,
    isActive,
    refreshUi: syncOperationalCards,
    getState: () => ({
      explicit: state.explicit,
      activeIds: [...state.activeIds],
      savingIds: [...state.savingIds],
      lastError: state.lastError ? String(state.lastError?.message || state.lastError) : ""
    })
  };
})();
