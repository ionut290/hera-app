(() => {
  "use strict";
  if (window.HeraOccasionalPdfStorage?.installed) return;

  const COMMESSA_ID = "lavori-occasionali";
  const MAX_FILE_SIZE = 15 * 1024 * 1024;
  const RETRY_MS = 250;
  const MAX_RETRIES = 80;
  const state = {
    observer: null,
    refreshing: false,
    popupRetries: 0,
    popupBridgeInstalled: false,
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
    return text(plant?.id || plant?.docId || plant?.impiantoId);
  }

  function plantName(plant) {
    return text(plant?.denominazione || plant?.["Denominazione Impianto"] || plant?.nome || "Cantiere occasionale");
  }

  function isOccasionalPlant(plant) {
    const commessaId = text(
      plant?.commessaId || plant?.parentCommessaId || plant?.commessa?.id || window.selectedCommessaId
    ).toLowerCase();
    return plant?.lavoroOccasionale === true || plant?.multiCantiere === true || commessaId === COMMESSA_ID;
  }

  function plantRef(plant) {
    const id = plantId(plant);
    if (!id) throw new Error("Cantiere non identificato.");
    return db.collection(collectionName()).doc(COMMESSA_ID).collection("impianti").doc(id);
  }

  function storageService() {
    if (!firebase?.apps?.length || typeof firebase.storage !== "function") {
      throw new Error("Firebase Storage non disponibile.");
    }
    return firebase.storage();
  }

  async function uploadOnePdf(plant, file) {
    const user = currentUserData();
    if (!user.uid) throw new Error("Accedi all'app prima di allegare un PDF.");
    if (!isPdf(file)) throw new Error(`${file?.name || "Il file"} non è un PDF.`);
    if (!file.size) throw new Error(`${file.name} è vuoto.`);
    if (file.size > MAX_FILE_SIZE) throw new Error(`${file.name} supera il limite di 15 MB.`);

    const ref = plantRef(plant).collection("documentiPdf").doc();
    const fileName = safeFileName(file.name);
    const path = `lavori-occasionali/${plantId(plant)}/${user.uid}/${ref.id}/${fileName}`;
    const objectRef = storageService().ref(path);
    let uploaded = false;
    try {
      await objectRef.put(file, {
        contentType: "application/pdf",
        customMetadata: {
          commessaId: COMMESSA_ID,
          plantId: plantId(plant),
          documentId: ref.id,
          ownerUserId: user.uid
        }
      });
      uploaded = true;
      const downloadUrl = await objectRef.getDownloadURL();
      await ref.set({
        id: ref.id,
        commessaId: COMMESSA_ID,
        plantId: plantId(plant),
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
      return { id: ref.id, downloadUrl, storagePath: path };
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
    if (text(item.storagePath)) await storageService().ref(item.storagePath).delete();
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

      const oldHtml = trigger?.innerHTML || trigger?.textContent || "📎 ALLEGA PDF";
      if (trigger) {
        trigger.disabled = true;
        trigger.innerHTML = "⏳ CARICO PDF…";
      }
      try {
        for (const file of files) await uploadOnePdf(plant, file);
        feedback(files.length === 1 ? "PDF allegato al cantiere." : `${files.length} PDF allegati al cantiere.`);
        if (host) await renderPdfList(plant, host);
      } catch (error) {
        console.error("Upload PDF lavoro occasionale fallito:", error);
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
    host.innerHTML = '<small class="muted">Carico documenti PDF…</small>';
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
        return `<div class="occasional-pdf-item" data-pdf-id="${escapeHtml(item.id)}">
          <button type="button" class="occasional-pdf-open" data-open-pdf="${escapeHtml(item.id)}"><span>📄</span><span><strong>${escapeHtml(item.title || item.fileName || "Documento PDF")}</strong><small>${escapeHtml(formatSize(item.fileSize))}${item.createdByName ? ` • ${escapeHtml(item.createdByName)}` : ""}</small></span></button>
          ${canDelete ? `<button type="button" class="occasional-pdf-delete" data-delete-pdf="${escapeHtml(item.id)}" title="Elimina PDF">🗑️</button>` : ""}
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
      console.warn("Lettura PDF cantiere non riuscita:", error);
      host.innerHTML = '<small class="muted">PDF non disponibili in questo momento.</small>';
    }
  }

  function plants() {
    const all = [];
    try { if (Array.isArray(currentImpianti)) all.push(...currentImpianti); } catch (_) {}
    try {
      const list = impiantiByCommessaId instanceof Map ? impiantiByCommessaId.get(COMMESSA_ID) : null;
      if (Array.isArray(list)) all.push(...list);
    } catch (_) {}
    return [...new Map(all
      .filter((plant) => isOccasionalPlant(plant))
      .map((plant) => [plantId(plant) || plantName(plant), plant])).values()];
  }

  function findPlantForButton(button) {
    if (!(button instanceof Element)) return null;
    const candidates = plants();
    let node = button;
    for (let depth = 0; node && node !== document.body && depth < 10; depth += 1, node = node.parentElement) {
      for (const plant of candidates) {
        const id = plantId(plant);
        const name = normalize(plantName(plant));
        if (!id || !name) continue;
        const section = node.querySelector?.(`[data-occasional-pdf-plant="${CSS.escape(id)}"]`);
        if (section || normalize(node.textContent).includes(name)) return plant;
      }
    }
    return null;
  }

  function isGearButton(button) {
    if (!(button instanceof Element)) return false;
    const label = normalize([
      button.textContent,
      button.getAttribute("aria-label"),
      button.getAttribute("title")
    ].filter(Boolean).join(" "));
    return label.includes("⚙") || label.includes("INGRAN") || label.includes("IMPOSTAZ") || label.includes("GESTIONE");
  }

  function findGearAnchorForPlant(plant) {
    const name = normalize(plantName(plant));
    if (!name) return null;
    const gears = Array.from(document.querySelectorAll("button")).filter(isGearButton);
    for (const gear of gears) {
      let node = gear;
      for (let depth = 0; node && node !== document.body && depth < 10; depth += 1, node = node.parentElement) {
        if (!normalize(node.textContent).includes(name)) continue;
        const row = gear.parentElement || gear;
        return { card: node, gear, anchor: row };
      }
    }
    return null;
  }

  function createPdfSection(plant) {
    const section = document.createElement("section");
    section.className = "occasional-pdf-section occasional-pdf-under-gear";
    section.dataset.occasionalPdfPlant = plantId(plant);
    section.innerHTML = `<div class="occasional-pdf-head"><button type="button" class="btn occasional-pdf-add">📎 ALLEGA PDF</button><small class="muted">Il PDF verrà mostrato prima di NAVIGA • max 15 MB</small></div><div class="occasional-pdf-list"></div>`;
    const host = section.querySelector(".occasional-pdf-list");
    const add = section.querySelector(".occasional-pdf-add");
    add.addEventListener("click", () => choosePdfs(plant, host, add));
    void renderPdfList(plant, host);
    return section;
  }

  function decorateCards() {
    plants().forEach((plant) => {
      const id = plantId(plant);
      if (!id) return;
      const placement = findGearAnchorForPlant(plant);
      if (!placement?.anchor?.isConnected) return;

      let section = document.querySelector(`[data-occasional-pdf-plant="${CSS.escape(id)}"]`);
      if (!section) section = createPdfSection(plant);

      const expectedPrevious = placement.anchor;
      if (section.previousElementSibling !== expectedPrevious) {
        expectedPrevious.insertAdjacentElement("afterend", section);
      }
    });
  }

  function closeViewer() {
    if (!state.viewer) return;
    state.viewer.remove();
    state.viewer = null;
    document.documentElement.classList.remove("occasional-pdf-viewer-open");
  }

  function openNavigationViewer(plant, items, originalButton) {
    closeViewer();
    let index = 0;
    const overlay = document.createElement("section");
    overlay.className = "occasional-pdf-navigation-viewer";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.innerHTML = `
      <header class="occasional-pdf-navigation-head">
        <div><strong>📄 Documento cantiere</strong><small class="occasional-pdf-navigation-title"></small></div>
        <div class="occasional-pdf-navigation-counter"></div>
        <button type="button" class="occasional-pdf-navigation-close" aria-label="Chiudi">✕</button>
      </header>
      <main class="occasional-pdf-navigation-body"><iframe title="Documento PDF del cantiere"></iframe></main>
      <footer class="occasional-pdf-navigation-actions">
        <button type="button" class="btn occasional-pdf-navigation-prev">← PDF PRECEDENTE</button>
        <button type="button" class="btn occasional-pdf-navigation-next">PDF SUCCESSIVO →</button>
        <button type="button" class="btn btn-primary occasional-pdf-navigation-continue">CONTINUA NAVIGAZIONE</button>
      </footer>`;

    const frame = overlay.querySelector("iframe");
    const title = overlay.querySelector(".occasional-pdf-navigation-title");
    const counter = overlay.querySelector(".occasional-pdf-navigation-counter");
    const prev = overlay.querySelector(".occasional-pdf-navigation-prev");
    const next = overlay.querySelector(".occasional-pdf-navigation-next");

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
    overlay.querySelector(".occasional-pdf-navigation-close").addEventListener("click", closeViewer);
    overlay.querySelector(".occasional-pdf-navigation-continue").addEventListener("click", () => {
      closeViewer();
      state.navigationBypass.add(originalButton);
      originalButton.click();
      queueMicrotask(() => state.navigationBypass.delete(originalButton));
    });

    document.body.appendChild(overlay);
    state.viewer = overlay;
    document.documentElement.classList.add("occasional-pdf-viewer-open");
    render();
  }

  async function handleNavigationClick(event) {
    const button = event.target?.closest?.("button, a");
    if (!button) return;
    if (state.navigationBypass.has(button)) return;
    const label = normalize(button.textContent || button.getAttribute("aria-label") || "");
    if (!label.startsWith("NAVIGA")) return;

    const plant = findPlantForButton(button);
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
      feedback("Non riesco a caricare il PDF del cantiere. Riprova prima di avviare la navigazione.", true);
    }
  }

  function removePdfFromWhatsappPopup(chooser, plant) {
    if (!chooser || !isOccasionalPlant(plant)) return;
    chooser.querySelector('[data-photo-source="pdf"], .whazzup-photo-pdf-btn')?.remove();
    chooser.querySelector("[data-shared-pdf-list]")?.remove();
    const heading = chooser.querySelector("#whazzup-photo-source-title");
    if (heading) heading.textContent = "Come vuoi aggiungere le foto?";
    const description = heading?.nextElementSibling;
    if (description?.tagName === "P") {
      description.textContent = "Il PDF del cantiere non viene inviato con WhatsApp: viene mostrato prima di NAVIGA.";
    }
  }

  function installPopupBridge() {
    if (state.popupBridgeInstalled) return true;
    if (!window.HeraWhazzupPdfV2 || typeof window.openWhazzupPhotoSourceChooser !== "function") return false;
    const original = window.openWhazzupPhotoSourceChooser;
    if (original.__occasionalPdfNavigationBridge === true) {
      state.popupBridgeInstalled = true;
      return true;
    }
    function openWhazzupPhotoSourceChooserWithoutOccasionalPdf(plant) {
      const result = original.apply(this, arguments);
      const clean = () => {
        const choosers = Array.from(document.querySelectorAll(".whazzup-photo-source-chooser"));
        removePdfFromWhatsappPopup(choosers[choosers.length - 1], plant);
      };
      queueMicrotask(clean);
      window.setTimeout(clean, 50);
      return result;
    }
    openWhazzupPhotoSourceChooserWithoutOccasionalPdf.__occasionalPdfNavigationBridge = true;
    window.openWhazzupPhotoSourceChooser = openWhazzupPhotoSourceChooserWithoutOccasionalPdf;
    state.popupBridgeInstalled = true;
    return true;
  }

  function retryPopupBridge() {
    if (installPopupBridge()) return;
    state.popupRetries += 1;
    if (state.popupRetries < MAX_RETRIES) window.setTimeout(retryPopupBridge, RETRY_MS);
    else console.warn("Bridge PDF Lavori occasionali non installato: popup foto non disponibile.");
  }

  function installStyles() {
    if (document.getElementById("occasional-pdf-storage-style")) return;
    const style = document.createElement("style");
    style.id = "occasional-pdf-storage-style";
    style.textContent = `
      .occasional-pdf-section{margin-top:10px;padding:10px;border:1px solid #dbe4ef;border-radius:14px;background:#f8fafc}
      .occasional-pdf-under-gear{width:100%;box-sizing:border-box}
      .occasional-pdf-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap}.occasional-pdf-head .occasional-pdf-add{font-weight:800}
      .occasional-pdf-list{display:grid;gap:7px;margin-top:8px}.occasional-pdf-item{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:7px;align-items:center}
      .occasional-pdf-open{display:grid;grid-template-columns:30px minmax(0,1fr);gap:8px;align-items:center;width:100%;padding:9px 10px;border:1px solid #dbe4ef;border-radius:12px;background:#fff;text-align:left;color:inherit}
      .occasional-pdf-open>span:first-child{font-size:22px}.occasional-pdf-open strong,.occasional-pdf-open small{display:block;overflow-wrap:anywhere}.occasional-pdf-open small{margin-top:2px;color:#64748b}
      .occasional-pdf-delete{width:42px;height:42px;border:1px solid #fecaca;border-radius:12px;background:#fff7f7;font-size:18px}
      html.occasional-pdf-viewer-open,html.occasional-pdf-viewer-open body{overflow:hidden}
      .occasional-pdf-navigation-viewer{position:fixed;inset:0;z-index:2147483600;background:#eef2f7;display:grid;grid-template-rows:auto minmax(0,1fr) auto}
      .occasional-pdf-navigation-head{display:grid;grid-template-columns:minmax(0,1fr) auto auto;gap:12px;align-items:center;padding:max(12px,env(safe-area-inset-top)) 14px 12px;background:#fff;border-bottom:1px solid #dbe4ef}
      .occasional-pdf-navigation-head strong,.occasional-pdf-navigation-head small{display:block}.occasional-pdf-navigation-head small{margin-top:3px;color:#64748b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
      .occasional-pdf-navigation-counter{font-weight:800;white-space:nowrap}.occasional-pdf-navigation-close{width:44px;height:44px;border:1px solid #dbe4ef;border-radius:12px;background:#fff;font-size:20px}
      .occasional-pdf-navigation-body{min-height:0;background:#d8dee8}.occasional-pdf-navigation-body iframe{display:block;width:100%;height:100%;border:0;background:#fff}
      .occasional-pdf-navigation-actions{display:grid;grid-template-columns:auto auto minmax(220px,1fr);gap:8px;padding:10px 12px max(10px,env(safe-area-inset-bottom));background:#fff;border-top:1px solid #dbe4ef}
      .occasional-pdf-navigation-continue{min-height:48px;font-weight:900}
      @media(max-width:700px){.occasional-pdf-navigation-actions{grid-template-columns:1fr 1fr}.occasional-pdf-navigation-continue{grid-column:1/-1}.occasional-pdf-navigation-head{grid-template-columns:minmax(0,1fr) auto}.occasional-pdf-navigation-counter{grid-row:2;grid-column:1}.occasional-pdf-navigation-close{grid-row:1/3;grid-column:2}}
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
  document.addEventListener("change", (event) => {
    if (event.target?.id === "squadra-commessa") refresh();
  });
  document.addEventListener("click", handleNavigationClick, true);

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", refresh, { once: true });
  else refresh();

  window.HeraOccasionalPdfStorage = {
    installed: true,
    version: "2.1.0",
    refresh,
    isOccasionalPlant,
    uploadOnePdf,
    choosePdfs,
    loadPdfs,
    openNavigationViewer
  };
  retryPopupBridge();
})();