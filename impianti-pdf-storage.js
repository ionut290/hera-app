(() => {
  "use strict";

  if (window.HeraImpiantiPdfStorage?.installed) return;

  const MAX_FILE_SIZE = 15 * 1024 * 1024;
  const state = {
    observer: null,
    refreshing: false,
    navigationBypass: new WeakSet(),
    viewer: null
  };

  const text = (value) => String(value ?? "").trim();
  const normalize = (value) => text(value).replace(/\s+/g, " ").toLocaleUpperCase("it-IT");
  const collectionName = () => typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";
  const serverNow = () => firebase.firestore.FieldValue.serverTimestamp();

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

  function escapeHtml(value) {
    return text(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function safeFileName(name) {
    const cleaned = text(name || "documento.pdf")
      .replace(/[\\/:*?"<>|#%{}\[\]]+/g, "-")
      .replace(/\s+/g, " ")
      .trim();
    const pdf = cleaned.toLowerCase().endsWith(".pdf") ? cleaned : `${cleaned}.pdf`;
    return pdf.length <= 110 ? pdf : `${pdf.slice(0, 92)}-${Date.now()}.pdf`;
  }

  function formatSize(bytes) {
    const value = Number(bytes || 0);
    if (value < 1024) return `${value} B`;
    if (value < 1024 * 1024) return `${Math.round(value / 1024)} KB`;
    return `${(value / (1024 * 1024)).toFixed(1)} MB`;
  }

  function isPdf(file) {
    return file instanceof Blob && (
      text(file.type).toLowerCase() === "application/pdf"
      || text(file.name).toLowerCase().endsWith(".pdf")
    );
  }

  function feedback(message, isError = false) {
    if (typeof showToast === "function") {
      try { showToast(message, isError ? "error" : "success"); return; } catch (_) {}
    }
    if (isError) alert(message);
  }

  function plantId(plant) {
    return text(plant?.id || plant?.docId || plant?.impiantoId || plant?.["ID SAP"] || plant?.idSap);
  }

  function plantName(plant) {
    return text(
      plant?.denominazione
      || plant?.["Denominazione Impianto"]
      || plant?.nome
      || plant?.impianto
      || "Impianto"
    );
  }

  function selectedCommessaId() {
    const direct = text(window.selectedCommessaId || window.currentCommessaId);
    if (direct) return direct;
    const selectors = [
      document.getElementById("commessa-select"),
      document.getElementById("squadra-commessa"),
      document.getElementById("hours-table-commessa-select")
    ];
    for (const select of selectors) {
      const value = text(select?.value);
      if (value) return value;
    }
    return "";
  }

  function commessaIdForPlant(plant) {
    return text(
      plant?.commessaId
      || plant?.parentCommessaId
      || plant?.commessa?.id
      || selectedCommessaId()
    );
  }

  function identity(plant) {
    const commessaId = commessaIdForPlant(plant);
    const id = plantId(plant);
    if (!commessaId || !id) return null;
    return { commessaId, plantId: id, key: `${commessaId}::${id}` };
  }

  function plantRef(plant) {
    const id = identity(plant);
    if (!id) throw new Error("Impianto o commessa non identificati.");
    return db.collection(collectionName()).doc(id.commessaId).collection("impianti").doc(id.plantId);
  }

  function storageService() {
    if (!firebase?.apps?.length || typeof firebase.storage !== "function") {
      throw new Error("Firebase Storage non disponibile.");
    }
    return firebase.storage();
  }

  function availablePlants() {
    const result = new Map();
    const currentCommessa = selectedCommessaId();

    // SOLA LETTURA: questo modulo non modifica mai currentImpianti o impiantiByCommessaId.
    try {
      if (Array.isArray(currentImpianti)) {
        currentImpianti.forEach((plant) => {
          const id = identity(plant);
          if (id) result.set(id.key, plant);
        });
      }
    } catch (_) {}

    try {
      if (currentCommessa && impiantiByCommessaId instanceof Map) {
        const cached = impiantiByCommessaId.get(currentCommessa);
        if (Array.isArray(cached)) {
          cached.forEach((plant) => {
            const id = identity({ ...plant, commessaId: plant?.commessaId || currentCommessa });
            if (id) result.set(id.key, plant?.commessaId ? plant : { ...plant, commessaId: currentCommessa });
          });
        }
      }
    } catch (_) {}

    return Array.from(result.values());
  }

  async function uploadOnePdf(plant, file) {
    const id = identity(plant);
    if (!id) throw new Error("Impianto o commessa non identificati.");
    const user = currentUserData();
    if (!user.uid) throw new Error("Accedi all'app prima di allegare un PDF.");
    if (!isPdf(file)) throw new Error(`${file?.name || "Il file"} non è un PDF.`);
    if (!file.size) throw new Error(`${file.name} è vuoto.`);
    if (file.size > MAX_FILE_SIZE) throw new Error(`${file.name} supera il limite di 15 MB.`);

    const docRef = plantRef(plant).collection("documentiPdf").doc();
    const fileName = safeFileName(file.name);
    const path = `impianti-pdf/${id.commessaId}/${id.plantId}/${user.uid}/${docRef.id}/${fileName}`;
    const objectRef = storageService().ref(path);
    let uploaded = false;

    try {
      await objectRef.put(file, {
        contentType: "application/pdf",
        customMetadata: {
          commessaId: id.commessaId,
          plantId: id.plantId,
          documentId: docRef.id,
          ownerUserId: user.uid
        }
      });
      uploaded = true;
      const downloadUrl = await objectRef.getDownloadURL();
      await docRef.set({
        id: docRef.id,
        commessaId: id.commessaId,
        plantId: id.plantId,
        plantName: plantName(plant),
        title: text(file.name) || fileName,
        fileName,
        mimeType: "application/pdf",
        fileSize: Number(file.size || 0),
        storagePath: path,
        downloadUrl,
        visibility: "commessa",
        sharedToAll: true,
        showBeforeNavigation: true,
        includeInWhatsapp: false,
        createdBy: user.uid,
        createdByEmail: user.email,
        createdByName: user.name,
        createdAt: serverNow(),
        updatedAt: serverNow()
      });
      return { id: docRef.id, downloadUrl, storagePath: path };
    } catch (error) {
      if (uploaded) await objectRef.delete().catch(() => null);
      throw error;
    }
  }

  async function loadPdfs(plant) {
    const snap = await plantRef(plant).collection("documentiPdf").get();
    return snap.docs
      .map((doc) => ({ id: doc.id, ...doc.data() }))
      .filter((item) => item.mimeType === "application/pdf" && item.downloadUrl)
      .sort((a, b) => (a.createdAt?.toMillis?.() || 0) - (b.createdAt?.toMillis?.() || 0));
  }

  async function deletePdf(plant, item) {
    const user = currentUserData();
    let admin = false;
    try { admin = typeof canManageData === "function" && canManageData(); } catch (_) {}
    if (!admin && text(item.createdBy) !== user.uid) {
      feedback("Puoi eliminare soltanto i PDF caricati da te.", true);
      return false;
    }
    if (!confirm(`Eliminare il PDF “${item.title || item.fileName || "Documento"}”?`)) return false;
    if (text(item.storagePath)) await storageService().ref(item.storagePath).delete().catch(() => null);
    await plantRef(plant).collection("documentiPdf").doc(item.id).delete();
    feedback("PDF eliminato.");
    return true;
  }

  function choosePdfs(plant, host, trigger) {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "application/pdf,.pdf";
    input.multiple = true;
    input.hidden = true;
    document.body.appendChild(input);

    input.addEventListener("change", async () => {
      const files = Array.from(input.files || []);
      input.remove();
      if (!files.length) return;
      const invalid = files.find((file) => !isPdf(file));
      if (invalid) return feedback("Puoi allegare soltanto file PDF.", true);
      const oversized = files.find((file) => file.size > MAX_FILE_SIZE);
      if (oversized) return feedback(`${oversized.name} supera il limite di 15 MB.`, true);

      const oldHtml = trigger?.innerHTML || "📎 ALLEGA PDF";
      if (trigger) {
        trigger.disabled = true;
        trigger.innerHTML = "⏳ CARICO PDF…";
      }
      try {
        for (const file of files) await uploadOnePdf(plant, file);
        feedback(files.length === 1 ? "PDF allegato all'impianto." : `${files.length} PDF allegati all'impianto.`);
        if (host) await renderPdfList(plant, host);
      } catch (error) {
        console.error("Upload PDF impianto fallito:", error);
        const code = text(error?.code);
        const message = text(error?.message || error);
        feedback(`Caricamento PDF non riuscito${code ? ` (${code})` : ""}: ${message || "errore sconosciuto"}`, true);
      } finally {
        if (trigger) {
          trigger.disabled = false;
          trigger.innerHTML = oldHtml;
        }
      }
    }, { once: true });
    input.addEventListener("cancel", () => input.remove(), { once: true });
    input.click();
  }

  async function renderPdfList(plant, host) {
    if (!host?.isConnected) return;
    host.innerHTML = '<small class="muted">Carico PDF…</small>';
    try {
      const items = await loadPdfs(plant);
      if (!items.length) {
        host.innerHTML = '<small class="muted">Nessun PDF allegato.</small>';
        return;
      }
      const user = currentUserData();
      let admin = false;
      try { admin = typeof canManageData === "function" && canManageData(); } catch (_) {}
      host.innerHTML = items.map((item) => {
        const canDelete = admin || text(item.createdBy) === user.uid;
        return `<div class="impianto-pdf-item" data-pdf-id="${escapeHtml(item.id)}">
          <button type="button" class="impianto-pdf-open" data-open-pdf="${escapeHtml(item.id)}"><span>📄</span><span><strong>${escapeHtml(item.title || item.fileName || "Documento PDF")}</strong><small>${escapeHtml(formatSize(item.fileSize))}${item.createdByName ? ` • ${escapeHtml(item.createdByName)}` : ""}</small></span></button>
          ${canDelete ? `<button type="button" class="impianto-pdf-delete" data-delete-pdf="${escapeHtml(item.id)}" title="Elimina PDF">🗑️</button>` : ""}
        </div>`;
      }).join("");

      host.querySelectorAll("[data-open-pdf]").forEach((button) => {
        button.addEventListener("click", () => {
          const item = items.find((entry) => entry.id === button.dataset.openPdf);
          if (item?.downloadUrl) window.open(item.downloadUrl, "_blank", "noopener,noreferrer");
        });
      });
      host.querySelectorAll("[data-delete-pdf]").forEach((button) => {
        button.addEventListener("click", async () => {
          const item = items.find((entry) => entry.id === button.dataset.deletePdf);
          if (item && await deletePdf(plant, item)) await renderPdfList(plant, host);
        });
      });
    } catch (error) {
      console.warn("Lettura PDF impianto non riuscita:", error);
      host.innerHTML = '<small class="muted">PDF non disponibili in questo momento.</small>';
    }
  }

  function isGearButton(button) {
    if (!(button instanceof Element)) return false;
    const label = normalize([
      button.textContent,
      button.getAttribute("aria-label"),
      button.getAttribute("title"),
      button.dataset?.action
    ].filter(Boolean).join(" "));
    return label.includes("⚙")
      || label.includes("INGRAN")
      || label.includes("IMPOSTAZ")
      || label.includes("GESTIONE");
  }

  function findPlantForElement(element) {
    if (!(element instanceof Element)) return null;
    const candidates = availablePlants();
    let node = element;
    for (let depth = 0; node && node !== document.body && depth < 10; depth += 1, node = node.parentElement) {
      const nodeText = normalize(node.textContent);
      for (const plant of candidates) {
        const name = normalize(plantName(plant));
        const id = identity(plant);
        if (!id || !name) continue;
        if (node.querySelector?.(`[data-impianto-pdf-key="${CSS.escape(id.key)}"]`) || nodeText.includes(name)) return plant;
      }
    }
    return null;
  }

  function findGearPlacement(plant) {
    const name = normalize(plantName(plant));
    const id = identity(plant);
    if (!name || !id) return null;
    const gears = Array.from(document.querySelectorAll("button, [role='button']")).filter(isGearButton);
    for (const gear of gears) {
      let node = gear;
      for (let depth = 0; node && node !== document.body && depth < 10; depth += 1, node = node.parentElement) {
        if (!normalize(node.textContent).includes(name)) continue;
        return { card: node, gear, anchor: gear.parentElement || gear };
      }
    }
    return null;
  }

  function createPdfSection(plant) {
    const id = identity(plant);
    if (!id) return null;
    const section = document.createElement("section");
    section.className = "impianto-pdf-section";
    section.dataset.impiantoPdfKey = id.key;
    section.innerHTML = `<div class="impianto-pdf-head"><button type="button" class="btn impianto-pdf-add">📎 ALLEGA PDF</button><small class="muted">Mostrato prima di NAVIGA • max 15 MB</small></div><div class="impianto-pdf-list"></div>`;
    const host = section.querySelector(".impianto-pdf-list");
    const add = section.querySelector(".impianto-pdf-add");
    add.addEventListener("click", () => choosePdfs(plant, host, add));
    void renderPdfList(plant, host);
    return section;
  }

  function decorateCards() {
    availablePlants().forEach((plant) => {
      const id = identity(plant);
      if (!id) return;
      const placement = findGearPlacement(plant);
      if (!placement?.anchor?.isConnected) return;
      let section = document.querySelector(`[data-impianto-pdf-key="${CSS.escape(id.key)}"]`);
      if (!section) section = createPdfSection(plant);
      if (!section) return;
      if (section.previousElementSibling !== placement.anchor) {
        placement.anchor.insertAdjacentElement("afterend", section);
      }
    });
  }

  function closeViewer() {
    if (!state.viewer) return;
    state.viewer.remove();
    state.viewer = null;
    document.documentElement.classList.remove("impianto-pdf-viewer-open");
  }

  function openNavigationViewer(plant, items, originalButton) {
    closeViewer();
    let index = 0;
    const overlay = document.createElement("section");
    overlay.className = "impianto-pdf-navigation-viewer";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.innerHTML = `
      <header class="impianto-pdf-navigation-head">
        <div><strong>📄 Documento impianto</strong><small class="impianto-pdf-navigation-title"></small></div>
        <div class="impianto-pdf-navigation-counter"></div>
        <button type="button" class="impianto-pdf-navigation-close" aria-label="Chiudi">✕</button>
      </header>
      <main class="impianto-pdf-navigation-body"><iframe title="Documento PDF dell'impianto"></iframe></main>
      <footer class="impianto-pdf-navigation-actions">
        <button type="button" class="btn impianto-pdf-navigation-prev">← PDF PRECEDENTE</button>
        <button type="button" class="btn impianto-pdf-navigation-next">PDF SUCCESSIVO →</button>
        <button type="button" class="btn btn-primary impianto-pdf-navigation-continue">CONTINUA NAVIGAZIONE</button>
      </footer>`;

    const frame = overlay.querySelector("iframe");
    const title = overlay.querySelector(".impianto-pdf-navigation-title");
    const counter = overlay.querySelector(".impianto-pdf-navigation-counter");
    const prev = overlay.querySelector(".impianto-pdf-navigation-prev");
    const next = overlay.querySelector(".impianto-pdf-navigation-next");

    function render() {
      const item = items[index];
      frame.src = item.downloadUrl;
      title.textContent = `${plantName(plant)} • ${item.title || item.fileName || "Documento PDF"}`;
      counter.textContent = `${index + 1} / ${items.length}`;
      prev.hidden = items.length < 2;
      next.hidden = items.length < 2;
      prev.disabled = index === 0;
      next.disabled = index === items.length - 1;
    }

    prev.addEventListener("click", () => { if (index > 0) { index -= 1; render(); } });
    next.addEventListener("click", () => { if (index < items.length - 1) { index += 1; render(); } });
    overlay.querySelector(".impianto-pdf-navigation-close").addEventListener("click", closeViewer);
    overlay.querySelector(".impianto-pdf-navigation-continue").addEventListener("click", () => {
      closeViewer();
      state.navigationBypass.add(originalButton);
      originalButton.click();
      queueMicrotask(() => state.navigationBypass.delete(originalButton));
    });

    document.body.appendChild(overlay);
    state.viewer = overlay;
    document.documentElement.classList.add("impianto-pdf-viewer-open");
    render();
  }

  async function handleNavigationClick(event) {
    const button = event.target?.closest?.("button, a");
    if (!button || state.navigationBypass.has(button)) return;
    const label = normalize(button.textContent || button.getAttribute("aria-label") || button.getAttribute("title") || "");
    if (!label.startsWith("NAVIGA")) return;

    const plant = findPlantForElement(button);
    if (!plant) return;

    event.preventDefault();
    event.stopImmediatePropagation();
    const wasDisabled = button.disabled;
    if ("disabled" in button) button.disabled = true;
    try {
      const items = await loadPdfs(plant);
      if (!items.length) {
        state.navigationBypass.add(button);
        if ("disabled" in button) button.disabled = wasDisabled;
        button.click();
        queueMicrotask(() => state.navigationBypass.delete(button));
        return;
      }
      if ("disabled" in button) button.disabled = wasDisabled;
      openNavigationViewer(plant, items, button);
    } catch (error) {
      console.error("Controllo PDF prima di NAVIGA fallito:", error);
      if ("disabled" in button) button.disabled = wasDisabled;
      feedback("Non riesco a caricare i PDF dell'impianto. Riprova prima di avviare la navigazione.", true);
    }
  }

  function installStyles() {
    if (document.getElementById("impianti-pdf-storage-style")) return;
    const style = document.createElement("style");
    style.id = "impianti-pdf-storage-style";
    style.textContent = `
      .impianto-pdf-section{width:100%;box-sizing:border-box;margin-top:8px;padding:10px;border:1px solid #dbe4ef;border-radius:14px;background:#f8fafc}
      .impianto-pdf-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap}.impianto-pdf-add{font-weight:800}
      .impianto-pdf-list{display:grid;gap:7px;margin-top:8px}.impianto-pdf-item{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:7px;align-items:center}
      .impianto-pdf-open{display:grid;grid-template-columns:30px minmax(0,1fr);gap:8px;align-items:center;width:100%;padding:9px 10px;border:1px solid #dbe4ef;border-radius:12px;background:#fff;text-align:left;color:inherit}
      .impianto-pdf-open>span:first-child{font-size:22px}.impianto-pdf-open strong,.impianto-pdf-open small{display:block;overflow-wrap:anywhere}.impianto-pdf-open small{margin-top:2px;color:#64748b}
      .impianto-pdf-delete{width:42px;height:42px;border:1px solid #fecaca;border-radius:12px;background:#fff7f7;font-size:18px}
      html.impianto-pdf-viewer-open,html.impianto-pdf-viewer-open body{overflow:hidden}
      .impianto-pdf-navigation-viewer{position:fixed;inset:0;z-index:2147483600;background:#eef2f7;display:grid;grid-template-rows:auto minmax(0,1fr) auto}
      .impianto-pdf-navigation-head{display:grid;grid-template-columns:minmax(0,1fr) auto auto;gap:12px;align-items:center;padding:max(12px,env(safe-area-inset-top)) 14px 12px;background:#fff;border-bottom:1px solid #dbe4ef}
      .impianto-pdf-navigation-head strong,.impianto-pdf-navigation-head small{display:block}.impianto-pdf-navigation-head small{margin-top:3px;color:#64748b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
      .impianto-pdf-navigation-counter{font-weight:800;white-space:nowrap}.impianto-pdf-navigation-close{width:44px;height:44px;border:1px solid #dbe4ef;border-radius:12px;background:#fff;font-size:20px}
      .impianto-pdf-navigation-body{min-height:0;background:#d8dee8}.impianto-pdf-navigation-body iframe{display:block;width:100%;height:100%;border:0;background:#fff}
      .impianto-pdf-navigation-actions{display:grid;grid-template-columns:auto auto minmax(220px,1fr);gap:8px;padding:10px 12px max(10px,env(safe-area-inset-bottom));background:#fff;border-top:1px solid #dbe4ef}
      .impianto-pdf-navigation-continue{min-height:48px;font-weight:900}
      @media(max-width:700px){.impianto-pdf-navigation-actions{grid-template-columns:1fr 1fr}.impianto-pdf-navigation-continue{grid-column:1/-1}.impianto-pdf-navigation-head{grid-template-columns:minmax(0,1fr) auto}.impianto-pdf-navigation-counter{grid-row:2;grid-column:1}.impianto-pdf-navigation-close{grid-row:1/3;grid-column:2}}
    `;
    document.head.appendChild(style);
  }

  function refresh() {
    if (state.refreshing) return;
    state.refreshing = true;
    try {
      installStyles();
      decorateCards();
    } finally {
      state.refreshing = false;
    }
  }

  state.observer = new MutationObserver(() => queueMicrotask(refresh));
  state.observer.observe(document.documentElement, { childList: true, subtree: true });
  document.addEventListener("click", handleNavigationClick, true);
  document.addEventListener("change", (event) => {
    if (["commessa-select", "squadra-commessa", "hours-table-commessa-select"].includes(event.target?.id)) refresh();
  });

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", refresh, { once: true });
  else refresh();

  window.HeraImpiantiPdfStorage = {
    installed: true,
    version: "1.0.0",
    refresh,
    uploadOnePdf,
    choosePdfs,
    loadPdfs,
    openNavigationViewer
  };
})();
