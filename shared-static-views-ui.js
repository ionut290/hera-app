(() => {
  "use strict";

  if (window.__heraSharedStaticViewsUiInstalled) return;
  window.__heraSharedStaticViewsUiInstalled = true;

  let unsubscribe = null;
  let activeType = "squadre";
  let activeKey = "";

  const escapeHtml = (value) => String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

  function todayKey() {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "Europe/Rome",
      year: "numeric",
      month: "2-digit",
      day: "2-digit"
    }).format(new Date());
  }

  function currentMonthKey() {
    return todayKey().slice(0, 7);
  }

  function isAdmin() {
    const user = typeof currentUser !== "undefined" ? currentUser : window.currentUser;
    const role = String(user?.role || user?.ruolo || "").toLowerCase();
    return user?.isAdmin === true || role === "admin" || role === "superadmin" || role === "super_admin";
  }

  function ensureStyles() {
    if (document.getElementById("shared-static-views-ui-style")) return;
    const style = document.createElement("style");
    style.id = "shared-static-views-ui-style";
    style.textContent = `
      .shared-view-dialog{width:min(94vw,760px);max-height:90vh;border:0;border-radius:18px;padding:0;background:#fff;color:#171717;box-shadow:0 18px 60px rgba(0,0,0,.35)}
      .shared-view-dialog::backdrop{background:rgba(0,0,0,.62)}
      .shared-view-head{position:sticky;top:0;z-index:2;background:#fff;padding:16px;border-bottom:1px solid #ddd;display:grid;gap:10px}
      .shared-view-head-row{display:flex;align-items:center;justify-content:space-between;gap:10px}
      .shared-view-controls{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
      .shared-view-controls input{min-height:42px;padding:8px;border:1px solid #aaa;border-radius:10px}
      .shared-view-body{padding:16px;overflow:auto}
      .shared-view-meta{font-size:.85rem;color:#666;margin-bottom:12px}
      .shared-view-empty{padding:24px;border:1px dashed #aaa;border-radius:12px;text-align:center;color:#666}
      .shared-view-card{border:1px solid #ddd;border-radius:12px;padding:12px;margin-bottom:10px;background:#fafafa}
      .shared-view-card h4{margin:0 0 8px}
      .shared-view-list{display:grid;gap:8px}
      .shared-view-json{white-space:pre-wrap;overflow-wrap:anywhere;font:13px/1.45 ui-monospace,SFMono-Regular,Menlo,monospace;background:#f4f4f4;padding:12px;border-radius:10px}
      .shared-view-feedback{min-height:20px;font-size:.9rem}
    `;
    document.head.appendChild(style);
  }

  function ensureDialog() {
    let dialog = document.getElementById("shared-static-view-dialog");
    if (dialog) return dialog;
    ensureStyles();
    dialog = document.createElement("dialog");
    dialog.id = "shared-static-view-dialog";
    dialog.className = "shared-view-dialog";
    dialog.innerHTML = `
      <div class="shared-view-head">
        <div class="shared-view-head-row">
          <div><strong id="shared-view-title">Vista condivisa</strong><div id="shared-view-feedback" class="shared-view-feedback"></div></div>
          <button id="shared-view-close" class="btn" type="button">CHIUDI</button>
        </div>
        <div class="shared-view-controls">
          <input id="shared-view-key" aria-label="Data o mese">
          <button id="shared-view-refresh" class="btn" type="button">AGGIORNA</button>
          <button id="shared-view-publish" class="btn btn-primary" type="button">PUBBLICA VERSIONE ATTUALE</button>
        </div>
      </div>
      <div id="shared-view-body" class="shared-view-body"></div>`;
    document.body.appendChild(dialog);

    dialog.querySelector("#shared-view-close").addEventListener("click", () => dialog.close());
    dialog.querySelector("#shared-view-refresh").addEventListener("click", () => {
      const key = dialog.querySelector("#shared-view-key").value.trim();
      openView(activeType, key);
    });
    dialog.querySelector("#shared-view-publish").addEventListener("click", publishCurrent);
    dialog.addEventListener("close", () => {
      if (typeof unsubscribe === "function") unsubscribe();
      unsubscribe = null;
    });
    return dialog;
  }

  function stringifySafe(value) {
    try { return JSON.stringify(value, null, 2); } catch (_) { return String(value ?? ""); }
  }

  function renderPayload(documentValue) {
    const body = document.getElementById("shared-view-body");
    if (!body) return;
    if (!documentValue?.payload) {
      body.innerHTML = `<div class="shared-view-empty">Nessuna vista condivisa pubblicata per questo periodo.</div>`;
      return;
    }
    const updated = documentValue.updatedAt?.toDate?.() || documentValue.updatedAt || "";
    const updatedText = updated instanceof Date ? updated.toLocaleString("it-IT") : String(updated || "");
    const payload = documentValue.payload;
    const items = Array.isArray(payload) ? payload : (Array.isArray(payload?.items) ? payload.items : null);
    let content = "";
    if (items) {
      content = `<div class="shared-view-list">${items.map((item, index) => `
        <article class="shared-view-card"><h4>${escapeHtml(item?.nome || item?.name || item?.commessa || item?.operatore || `Elemento ${index + 1}`)}</h4>
        <pre class="shared-view-json">${escapeHtml(stringifySafe(item))}</pre></article>`).join("")}</div>`;
    } else {
      content = `<pre class="shared-view-json">${escapeHtml(stringifySafe(payload))}</pre>`;
    }
    body.innerHTML = `<div class="shared-view-meta">Versione ${escapeHtml(documentValue.version || 1)}${updatedText ? ` · aggiornata ${escapeHtml(updatedText)}` : ""}${documentValue.authorName ? ` · ${escapeHtml(documentValue.authorName)}` : ""}</div>${content}`;
  }

  function setFeedback(message, error = false) {
    const node = document.getElementById("shared-view-feedback");
    if (!node) return;
    node.textContent = message || "";
    node.style.color = error ? "#b00020" : "#146c2e";
  }

  function openView(type, key) {
    const api = window.HeraSharedStaticViews;
    const dialog = ensureDialog();
    activeType = type;
    activeKey = key || (type === "squadre" ? todayKey() : currentMonthKey());
    const input = dialog.querySelector("#shared-view-key");
    input.type = type === "squadre" ? "date" : "month";
    input.value = activeKey;
    dialog.querySelector("#shared-view-title").textContent = type === "squadre" ? "Vista squadre condivisa" : "Vista calendario condivisa";
    dialog.querySelector("#shared-view-publish").hidden = !isAdmin();
    setFeedback("");
    renderPayload(api?.readLocal?.(type, activeKey));
    if (typeof unsubscribe === "function") unsubscribe();
    unsubscribe = api?.subscribe?.(type, activeKey, renderPayload, (error) => setFeedback(error?.message || "Errore lettura vista condivisa", true)) || null;
    if (!dialog.open) dialog.showModal();
  }

  async function publishCurrent() {
    if (!isAdmin()) return setFeedback("Solo un amministratore può pubblicare.", true);
    const api = window.HeraSharedStaticViews;
    const key = document.getElementById("shared-view-key")?.value?.trim() || activeKey;
    setFeedback("Pubblicazione in corso…");
    try {
      const result = activeType === "squadre"
        ? await api.publishSquadre(key)
        : await api.publishCalendar(key);
      setFeedback(result?.skipped ? "La versione condivisa era già aggiornata." : "Vista condivisa pubblicata su tutti i dispositivi.");
    } catch (error) {
      setFeedback(error?.message || "Pubblicazione non riuscita.", true);
    }
  }

  function addMenuButton(afterId, id, label, type) {
    if (document.getElementById(id)) return true;
    const anchor = document.getElementById(afterId);
    if (!anchor?.parentNode) return false;
    const button = document.createElement("button");
    button.id = id;
    button.type = "button";
    button.className = "btn menu-title-btn";
    button.innerHTML = `<span class="menu-item-icon" aria-hidden="true">🖼️</span>${label}`;
    button.addEventListener("click", () => openView(type));
    anchor.insertAdjacentElement("afterend", button);
    return true;
  }

  function install() {
    const a = addMenuButton("open-panel-squadre", "open-shared-squadre-view", "Vista squadre condivisa", "squadre");
    const b = addMenuButton("open-hours-btn", "open-shared-calendar-view", "Vista calendario condivisa", "calendario");
    return a && b;
  }

  let attempts = 0;
  const timer = setInterval(() => {
    attempts += 1;
    if (install() || attempts >= 80) clearInterval(timer);
  }, 250);

  window.HeraSharedStaticViewsUi = { openSquadre: (date) => openView("squadre", date), openCalendar: (month) => openView("calendario", month) };
})();
