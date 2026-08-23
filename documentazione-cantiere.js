(() => {
  "use strict";

  if (window.HeraCantiereDocuments?.installed) return;

  const COMMESSA_ID = "lavori-occasionali";
  const SOURCE = "documentazione-cantiere";
  const MAX_FILE_SIZE = 15 * 1024 * 1024;
  const RETRY_MS = 300;
  const MAX_RETRIES = 80;
  const state = {
    observer: null,
    decorating: false,
    retries: 0,
    overlay: null,
    counts: new Map(),
    docs: new Map(),
    navigationBypass: new WeakSet()
  };

  const text = (value) => String(value ?? "").trim();
  const normalize = (value) => text(value).replace(/\s+/g, " ").toLocaleUpperCase("it-IT");
  const esc = (value) => text(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

  function isAdmin() {
    try { return typeof canManageData === "function" && canManageData(); }
    catch (_) { return false; }
  }

  function currentUserData() {
    let user = null;
    try { user = currentUser || firebase.auth?.().currentUser || null; }
    catch (_) { try { user = firebase.auth?.().currentUser || null; } catch (_) {} }
    return {
      uid: text(user?.uid),
      email: text(user?.email),
      name: typeof getOperatorDisplayName === "function"
        ? text(getOperatorDisplayName())
        : text(user?.displayName || user?.email || "Operatore")
    };
  }

  function feedback(message, isError = false) {
    if (typeof showToast === "function") {
      try { showToast(message, isError ? "error" : "success"); return; } catch (_) {}
    }
    if (isError) alert(message);
  }

  function collectionName() {
    try { return typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse"; }
    catch (_) { return "commesse"; }
  }

  function serverNow() {
    return firebase.firestore.FieldValue.serverTimestamp();
  }

  function plantId(plant) {
    return text(plant?.id || plant?.docId || plant?.impiantoId || plant?.["ID SAP"] || plant?.idSap);
  }

  function plantName(plant) {
    return text(plant?.denominazione || plant?.["Denominazione Impianto"] || plant?.nome || plant?.impianto || "Cantiere");
  }

  function plantCommessaId(plant) {
    return text(plant?.commessaId || plant?.parentCommessaId || plant?.commessa?.id || window.selectedCommessaId || window.currentCommessaId);
  }

  function isOccasionalPlant(plant) {
    const id = plantCommessaId(plant).toLowerCase();
    return plant?.lavoroOccasionale === true || plant?.multiCantiere === true || id === COMMESSA_ID;
  }

  function keyForPlant(plant) {
    const id = plantId(plant);
    return id ? `${COMMESSA_ID}::${id}` : "";
  }

  function plantRef(plant) {
    const id = plantId(plant);
    if (!id) throw new Error("Cantiere non identificato.");
    return db.collection(collectionName()).doc(COMMESSA_ID).collection("impianti").doc(id);
  }

  function availablePlants() {
    const result = new Map();
    try {
      if (Array.isArray(currentImpianti)) {
        currentImpianti.filter(isOccasionalPlant).forEach((plant) => {
          const key = keyForPlant(plant);
          if (key) result.set(key, plant);
        });
      }
    } catch (_) {}
    try {
      if (impiantiByCommessaId instanceof Map) {
        const cached = impiantiByCommessaId.get(COMMESSA_ID);
        if (Array.isArray(cached)) cached.forEach((plant) => {
          const normalizedPlant = plant?.commessaId ? plant : { ...plant, commessaId: COMMESSA_ID };
          const key = keyForPlant(normalizedPlant);
          if (key) result.set(key, normalizedPlant);
        });
      }
    } catch (_) {}
    return Array.from(result.values());
  }

  function storageService() {
    if (!firebase?.apps?.length || typeof firebase.storage !== "function") {
      throw new Error("Firebase Storage non disponibile.");
    }
    return firebase.storage();
  }

  function safeFileName(name) {
    const raw = text(name || `documento-${Date.now()}`)
      .replace(/[\\/:*?"<>|#%{}\[\]]+/g, "-")
      .replace(/\s+/g, " ")
      .trim();
    return raw.length <= 110 ? raw : `${raw.slice(0, 88)}-${Date.now()}`;
  }

  function categoryIcon(category, mime) {
    if (text(mime).startsWith("image/")) return "🖼️";
    if (text(mime).includes("pdf")) return "📄";
    const icons = { PREVENTIVO: "💶", PLANIMETRIA: "🗺️", SICUREZZA: "🦺", ORDINE_LAVORO: "📝", AUTORIZZAZIONE: "✅", VERBALE: "📋", CONSUNTIVO: "🧾", ALTRO: "📎" };
    return icons[text(category)] || "📎";
  }

  function categoryLabel(category) {
    const labels = { PREVENTIVO: "Preventivo", FOTO_PRIMA: "Foto prima lavori", FOTO_DOPO: "Foto dopo lavori", PLANIMETRIA: "Planimetria", SICUREZZA: "POS / Sicurezza", ORDINE_LAVORO: "Ordine di lavoro", AUTORIZZAZIONE: "Autorizzazione", VERBALE: "Verbale", CONSUNTIVO: "Consuntivo / Fattura", ALTRO: "Altro" };
    return labels[text(category)] || "Documento";
  }

  function formatSize(bytes) {
    const n = Number(bytes || 0);
    if (!n) return "";
    if (n < 1024) return `${n} B`;
    if (n < 1024 * 1024) return `${Math.round(n / 1024)} KB`;
    return `${(n / (1024 * 1024)).toFixed(1)} MB`;
  }

  async function loadLegacyDocs(plant) {
    try {
      const snap = await plantRef(plant).collection("documentiPdf").get();
      return snap.docs.map((doc) => {
        const data = doc.data() || {};
        return {
          id: doc.id,
          legacy: true,
          ownerUserId: text(data.createdBy),
          title: text(data.title || data.fileName || "Documento PDF"),
          fileName: text(data.fileName),
          mimeType: text(data.mimeType || "application/pdf"),
          fileSize: Number(data.fileSize || 0),
          fileUrl: text(data.downloadUrl),
          storagePath: text(data.storagePath),
          category: "ALTRO",
          note: "",
          importantBeforeNavigation: data.showBeforeNavigation !== false,
          createdByName: text(data.createdByName),
          createdAt: data.createdAt || null
        };
      }).filter((item) => item.fileUrl);
    } catch (error) {
      console.warn("Documentazione legacy non disponibile:", error);
      return [];
    }
  }

  async function loadCloudDocs(plant) {
    const key = keyForPlant(plant);
    if (!key) return [];
    try {
      const snap = await db.collection("documents").where("impiantoKey", "==", key).get();
      return snap.docs.map((doc) => ({ id: doc.id, ...doc.data(), legacy: false }))
        .filter((item) => item.source === SOURCE && text(item.commessaId) === COMMESSA_ID && text(item.fileUrl));
    } catch (error) {
      console.warn("Documentazione cantiere non disponibile:", error);
      return [];
    }
  }

  async function loadDocs(plant, force = false) {
    const key = keyForPlant(plant);
    if (!key) return [];
    if (!force && state.docs.has(key)) return state.docs.get(key);
    const [cloud, legacy] = await Promise.all([loadCloudDocs(plant), loadLegacyDocs(plant)]);
    const all = [...cloud, ...legacy].sort((a, b) => {
      const ta = a.createdAt?.toMillis?.() || a.createdAt?.seconds * 1000 || 0;
      const tb = b.createdAt?.toMillis?.() || b.createdAt?.seconds * 1000 || 0;
      return tb - ta;
    });
    state.docs.set(key, all);
    state.counts.set(key, all.length);
    return all;
  }

  async function uploadDocument(plant, file, values) {
    if (!isAdmin()) throw new Error("Operazione riservata all'amministratore.");
    const user = currentUserData();
    if (!user.uid) throw new Error("Accedi all'app prima di caricare documenti.");
    if (!(file instanceof Blob) || !file.size) throw new Error("Seleziona un file valido.");
    if (file.size > MAX_FILE_SIZE) throw new Error(`${file.name || "Il file"} supera il limite di 15 MB.`);

    const docRef = db.collection("documents").doc();
    const fileName = safeFileName(file.name);
    const path = `documents/${user.uid}/${docRef.id}/${fileName}`;
    const objectRef = storageService().ref(path);
    let uploaded = false;
    try {
      await objectRef.put(file, {
        contentType: text(file.type) || "application/octet-stream",
        customMetadata: {
          source: SOURCE,
          commessaId: COMMESSA_ID,
          impiantoId: plantId(plant),
          documentId: docRef.id,
          ownerUserId: user.uid
        }
      });
      uploaded = true;
      const fileUrl = await objectRef.getDownloadURL();
      await docRef.set({
        id: docRef.id,
        source: SOURCE,
        ownerUserId: user.uid,
        createdBy: user.uid,
        createdByEmail: user.email,
        createdByName: user.name,
        visibility: "global",
        sharedToAll: true,
        sharedUserIds: [],
        commessaIds: [COMMESSA_ID],
        commessaId: COMMESSA_ID,
        impiantoId: plantId(plant),
        impiantoKey: keyForPlant(plant),
        impiantoName: plantName(plant),
        title: text(values.title || file.name || fileName),
        category: text(values.category || "ALTRO"),
        note: text(values.note),
        importantBeforeNavigation: values.importantBeforeNavigation === true,
        showBeforeNavigation: true,
        fileName,
        fileUrl,
        storagePath: path,
        mimeType: text(file.type) || "application/octet-stream",
        fileSize: Number(file.size),
        uploadStatus: "completed",
        versionHistoryEnabled: false,
        createdAt: serverNow(),
        updatedAt: serverNow()
      });
      state.docs.delete(keyForPlant(plant));
      state.counts.delete(keyForPlant(plant));
      return docRef.id;
    } catch (error) {
      if (uploaded) await objectRef.delete().catch(() => null);
      throw error;
    }
  }

  async function deleteDocument(plant, item) {
    if (item.legacy) {
      feedback("Il PDF storico resta protetto. Puoi sostituirlo caricando il documento aggiornato.", true);
      return false;
    }
    const user = currentUserData();
    if (!isAdmin() || text(item.ownerUserId) !== user.uid) {
      feedback("Puoi eliminare soltanto i documenti caricati dal tuo account amministratore.", true);
      return false;
    }
    if (!confirm(`Eliminare “${item.title || item.fileName || "Documento"}”?`)) return false;
    try {
      if (text(item.storagePath)) await storageService().ref(item.storagePath).delete().catch(() => null);
      await db.collection("documents").doc(item.id).delete();
      state.docs.delete(keyForPlant(plant));
      state.counts.delete(keyForPlant(plant));
      feedback("Documento eliminato.");
      return true;
    } catch (error) {
      console.error("Eliminazione documento cantiere fallita:", error);
      feedback("Non sono riuscito a eliminare il documento.", true);
      return false;
    }
  }

  function closeOverlay() {
    if (!state.overlay) return;
    state.overlay.remove();
    state.overlay = null;
    document.documentElement.style.overflow = "";
  }

  function shell(plant, title = "📁 Documentazione cantiere") {
    closeOverlay();
    const overlay = document.createElement("section");
    overlay.className = "cantiere-doc-overlay";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.innerHTML = `<div class="cantiere-doc-panel"><header class="cantiere-doc-head"><div class="cantiere-doc-head-main"><h2>${esc(title)}</h2><p>${esc(plantName(plant))}</p></div><button type="button" class="cantiere-doc-close" aria-label="Chiudi">✕</button></header><main class="cantiere-doc-body"></main></div>`;
    overlay.querySelector(".cantiere-doc-close").addEventListener("click", closeOverlay);
    overlay.addEventListener("click", (event) => { if (event.target === overlay) closeOverlay(); });
    document.body.appendChild(overlay);
    document.documentElement.style.overflow = "hidden";
    state.overlay = overlay;
    return overlay.querySelector(".cantiere-doc-body");
  }

  function openUrl(url) {
    if (!url) return;
    window.open(url, "_blank", "noopener,noreferrer");
  }

  function renderItems(plant, host, items) {
    if (!items.length) {
      host.innerHTML = '<div class="cantiere-doc-empty">📂 Nessuna documentazione presente per questo cantiere.</div>';
      return;
    }
    const user = currentUserData();
    host.innerHTML = `<div class="cantiere-doc-list">${items.map((item) => {
      const canDelete = isAdmin() && !item.legacy && text(item.ownerUserId) === user.uid;
      const meta = [categoryLabel(item.category), formatSize(item.fileSize), text(item.createdByName)].filter(Boolean).join(" • ");
      return `<article class="cantiere-doc-item" data-doc-id="${esc(item.id)}"><div class="cantiere-doc-icon">${categoryIcon(item.category, item.mimeType)}</div><div class="cantiere-doc-info"><div class="cantiere-doc-title">${esc(item.title || item.fileName || "Documento")}</div><div class="cantiere-doc-meta">${esc(meta)}</div>${item.note ? `<div class="cantiere-doc-desc">${esc(item.note)}</div>` : ""}${item.importantBeforeNavigation ? '<span class="cantiere-doc-important">⚠️ DA LEGGERE PRIMA DEL LAVORO</span>' : ""}</div><div class="cantiere-doc-actions"><button type="button" class="cantiere-doc-open" data-open-doc="${esc(item.id)}">APRI</button>${canDelete ? `<button type="button" class="cantiere-doc-delete" data-delete-doc="${esc(item.id)}" title="Elimina">🗑️</button>` : ""}</div></article>`;
    }).join("")}</div>`;
    host.querySelectorAll("[data-open-doc]").forEach((button) => button.addEventListener("click", () => {
      const item = items.find((entry) => text(entry.id) === text(button.dataset.openDoc));
      if (item) openUrl(item.fileUrl || item.downloadUrl);
    }));
    host.querySelectorAll("[data-delete-doc]").forEach((button) => button.addEventListener("click", async () => {
      const item = items.find((entry) => text(entry.id) === text(button.dataset.deleteDoc));
      if (item && await deleteDocument(plant, item)) await openDocuments(plant, true);
    }));
  }

  function uploadForm(plant, host, onSaved) {
    const form = document.createElement("div");
    form.className = "cantiere-doc-form";
    form.innerHTML = `<label>Tipo documento</label><select data-doc-category><option value="PREVENTIVO">Preventivo</option><option value="FOTO_PRIMA">Foto prima lavori</option><option value="FOTO_DOPO">Foto dopo lavori</option><option value="PLANIMETRIA">Planimetria</option><option value="SICUREZZA">POS / Sicurezza</option><option value="ORDINE_LAVORO">Ordine di lavoro</option><option value="AUTORIZZAZIONE">Autorizzazione</option><option value="VERBALE">Verbale</option><option value="CONSUNTIVO">Consuntivo / Fattura</option><option value="ALTRO">Altro</option></select><label>Titolo</label><input type="text" data-doc-title placeholder="Es. Preventivo potatura 2026"><label>Nota</label><textarea data-doc-note placeholder="Informazioni utili per la squadra"></textarea><label class="cantiere-doc-check"><input type="checkbox" data-doc-important> ⚠️ Da leggere prima di iniziare il lavoro</label><label>File (max 15 MB)</label><input type="file" data-doc-file accept="application/pdf,image/*,.doc,.docx,.xls,.xlsx,.odt,.ods" required><div class="cantiere-doc-form-actions"><button type="button" class="cantiere-doc-save">CARICA</button><button type="button" class="cantiere-doc-cancel">ANNULLA</button></div>`;
    host.prepend(form);
    form.querySelector(".cantiere-doc-cancel").addEventListener("click", () => form.remove());
    form.querySelector(".cantiere-doc-save").addEventListener("click", async (event) => {
      const button = event.currentTarget;
      const file = form.querySelector("[data-doc-file]").files?.[0];
      if (!file) return feedback("Seleziona un file.", true);
      button.disabled = true;
      button.textContent = "CARICO…";
      try {
        await uploadDocument(plant, file, {
          category: form.querySelector("[data-doc-category]").value,
          title: form.querySelector("[data-doc-title]").value,
          note: form.querySelector("[data-doc-note]").value,
          importantBeforeNavigation: form.querySelector("[data-doc-important]").checked
        });
        feedback("Documentazione cantiere caricata.");
        form.remove();
        if (typeof onSaved === "function") await onSaved();
      } catch (error) {
        console.error("Upload documentazione cantiere fallito:", error);
        feedback(`Caricamento non riuscito: ${text(error?.message || error)}`, true);
      } finally {
        if (button.isConnected) { button.disabled = false; button.textContent = "CARICA"; }
      }
    });
  }

  async function openDocuments(plant, force = false) {
    const body = shell(plant);
    body.innerHTML = '<div class="cantiere-doc-empty">⏳ Carico documentazione…</div>';
    const items = await loadDocs(plant, force);
    body.innerHTML = "";
    if (isAdmin()) {
      const toolbar = document.createElement("div");
      toolbar.className = "cantiere-doc-toolbar";
      toolbar.innerHTML = '<button type="button" class="cantiere-doc-upload">➕ AGGIUNGI DOCUMENTO</button><span class="cantiere-doc-note">PDF, foto e documenti • max 15 MB</span>';
      body.appendChild(toolbar);
      toolbar.querySelector(".cantiere-doc-upload").addEventListener("click", () => {
        if (body.querySelector(".cantiere-doc-form")) return;
        uploadForm(plant, body, async () => openDocuments(plant, true));
      });
    }
    const listHost = document.createElement("div");
    body.appendChild(listHost);
    renderItems(plant, listHost, items);
    void decorateCards(true);
  }

  async function showNavigationPrompt(plant, button, items) {
    const body = shell(plant, "📁 Documentazione disponibile");
    const important = items.some((item) => item.importantBeforeNavigation === true);
    body.innerHTML = `<div class="cantiere-doc-nav-card"><div class="cantiere-doc-nav-alert">${important ? "⚠️ Questo cantiere contiene documenti indicati da leggere prima del lavoro." : "Sono presenti documenti per questo cantiere."}<br><strong>${items.length} ${items.length === 1 ? "documento disponibile" : "documenti disponibili"}.</strong></div><div class="cantiere-doc-nav-actions"><button type="button" class="cantiere-doc-nav-view">📁 VISUALIZZA DOCUMENTI</button><button type="button" class="cantiere-doc-nav-continue">🧭 CONTINUA A NAVIGARE</button></div></div>`;
    body.querySelector(".cantiere-doc-nav-view").addEventListener("click", () => openDocuments(plant));
    body.querySelector(".cantiere-doc-nav-continue").addEventListener("click", () => {
      closeOverlay();
      state.navigationBypass.add(button);
      button.click();
    });
  }

  function findPlantForElement(element) {
    if (!(element instanceof Element)) return null;
    const plants = availablePlants();
    let node = element;
    for (let depth = 0; node && node !== document.body && depth < 10; depth += 1, node = node.parentElement) {
      const nodeText = normalize(node.textContent);
      for (const plant of plants) {
        const name = normalize(plantName(plant));
        if (name && nodeText.includes(name)) return plant;
      }
    }
    return null;
  }

  function isNavigateButton(button) {
    return button instanceof HTMLButtonElement && normalize(button.textContent).includes("NAVIGA");
  }

  function isGearButton(button) {
    if (!(button instanceof HTMLButtonElement)) return false;
    const label = normalize([button.textContent, button.title, button.getAttribute("aria-label")].filter(Boolean).join(" "));
    return label.includes("⚙") || label.includes("INGRAN") || label.includes("GESTIONE");
  }

  function cardForPlant(plant) {
    const name = normalize(plantName(plant));
    if (!name) return null;
    const buttons = Array.from(document.querySelectorAll("button"));
    const anchors = buttons.filter((button) => normalize(button.textContent) === "FATTO" || isNavigateButton(button));
    for (const anchor of anchors) {
      let node = anchor.parentElement;
      for (let depth = 0; node && node !== document.body && depth < 9; depth += 1, node = node.parentElement) {
        if (normalize(node.textContent).includes(name)) return node;
      }
    }
    return null;
  }

  function actionRow(card) {
    const nav = Array.from(card.querySelectorAll("button")).find(isNavigateButton);
    if (!nav) return null;
    let row = nav.parentElement;
    for (let i = 0; row && row !== card && i < 3; i += 1) {
      if (row.querySelectorAll("button").length >= 2) return row;
      row = row.parentElement;
    }
    return nav.parentElement;
  }

  async function decoratePlant(plant, force = false) {
    const card = cardForPlant(plant);
    const key = keyForPlant(plant);
    if (!card || !key) return;
    card.dataset.cantiereDocKey = key;

    if (isAdmin() && !card.querySelector("[data-cantiere-doc-admin]")) {
      const gear = Array.from(card.querySelectorAll("button")).find(isGearButton);
      if (gear) {
        const admin = document.createElement("button");
        admin.type = "button";
        admin.className = "cantiere-doc-admin";
        admin.dataset.cantiereDocAdmin = key;
        admin.innerHTML = "📁<br>DOCUMENTAZIONE";
        admin.title = "Documentazione cantiere";
        admin.addEventListener("click", (event) => { event.preventDefault(); event.stopPropagation(); void openDocuments(plant); });
        gear.insertAdjacentElement("afterend", admin);
      }
    }

    let count = state.counts.get(key);
    if (force || typeof count !== "number") {
      const items = await loadDocs(plant, force);
      count = items.length;
    }
    const existing = card.querySelector("[data-cantiere-doc-read]");
    if (!count) { if (existing) existing.remove(); return; }
    if (existing) {
      const badge = existing.querySelector(".cantiere-doc-count");
      if (badge) badge.textContent = String(count);
      return;
    }
    const wrap = document.createElement("div");
    wrap.className = "cantiere-doc-read-wrap";
    wrap.dataset.cantiereDocRead = key;
    wrap.innerHTML = `<button type="button" class="cantiere-doc-read">📁 LEGGI DOCUMENTAZIONE <span class="cantiere-doc-count">${count}</span></button>`;
    wrap.querySelector("button").addEventListener("click", (event) => { event.preventDefault(); event.stopPropagation(); void openDocuments(plant); });
    const row = actionRow(card);
    if (row) row.insertAdjacentElement("afterend", wrap);
    else card.appendChild(wrap);
  }

  async function decorateCards(force = false) {
    if (state.decorating) return;
    state.decorating = true;
    try {
      for (const plant of availablePlants()) await decoratePlant(plant, force);
    } finally {
      state.decorating = false;
    }
  }

  function installNavigationGuard() {
    document.addEventListener("click", async (event) => {
      const button = event.target?.closest?.("button");
      if (!isNavigateButton(button)) return;
      const plant = findPlantForElement(button);
      if (!plant || !isOccasionalPlant(plant)) return;
      if (state.navigationBypass.has(button)) {
        state.navigationBypass.delete(button);
        return;
      }
      event.preventDefault();
      event.stopImmediatePropagation();
      const items = await loadDocs(plant);
      if (!items.length) {
        state.navigationBypass.add(button);
        button.click();
        return;
      }
      await showNavigationPrompt(plant, button, items);
    }, true);
  }

  function installObserver() {
    if (state.observer) return;
    let timer = 0;
    state.observer = new MutationObserver(() => {
      clearTimeout(timer);
      timer = setTimeout(() => void decorateCards(false), 120);
    });
    state.observer.observe(document.body, { childList: true, subtree: true });
  }

  function init() {
    if (!window.firebase?.apps?.length || !window.db) {
      if (state.retries++ < MAX_RETRIES) return setTimeout(init, RETRY_MS);
      console.warn("Documentazione cantiere: Firebase non disponibile.");
      return;
    }
    installNavigationGuard();
    installObserver();
    void decorateCards(false);
  }

  window.HeraCantiereDocuments = {
    installed: true,
    open: openDocuments,
    refresh: (plant) => loadDocs(plant, true)
  };
  window.HeraOccasionalPdfStorage = window.HeraCantiereDocuments;

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init, { once: true });
  else init();
})();