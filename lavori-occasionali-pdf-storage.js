(() => {
  "use strict";
  if (window.HeraOccasionalPdfStorage?.installed) return;

  const COMMESSA_ID = "lavori-occasionali";
  const MAX_FILE_SIZE = 15 * 1024 * 1024;
  const POPUP_RETRY_MS = 250;
  const POPUP_MAX_RETRIES = 80;
  const state = { observer: null, refreshing: false, popupRetries: 0, popupBridgeInstalled: false };

  const text = (value) => String(value ?? "").trim();
  const normalize = (value) => text(value).replace(/\s+/g, " ").toLocaleUpperCase("it-IT");
  const collectionName = () => typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";
  const currentUserData = () => {
    let user = null;
    try { user = currentUser || firebase.auth?.().currentUser || null; } catch (_) { user = firebase.auth?.().currentUser || null; }
    return {
      uid: text(user?.uid),
      email: text(user?.email),
      name: typeof getOperatorDisplayName === "function" ? text(getOperatorDisplayName()) : text(user?.displayName || user?.email || "Operatore")
    };
  };
  const serverNow = () => firebase.firestore.FieldValue.serverTimestamp();

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
    return text(plant?.denominazione || plant?.nome || "Cantiere occasionale");
  }

  function isOccasionalPlant(plant) {
    const commessaId = text(plant?.commessaId || plant?.parentCommessaId || plant?.commessa?.id || window.selectedCommessaId).toLowerCase();
    return plant?.lavoroOccasionale === true || plant?.multiCantiere === true || commessaId === COMMESSA_ID;
  }

  function plantRef(plant) {
    const id = plantId(plant);
    if (!id) throw new Error("Cantiere non identificato.");
    return db.collection(collectionName()).doc(COMMESSA_ID).collection("impianti").doc(id);
  }

  function storageService() {
    if (!firebase?.apps?.length || typeof firebase.storage !== "function") throw new Error("Firebase Storage non disponibile.");
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
      .filter((item) => item.mimeType === "application/pdf")
      .sort((a, b) => (b.createdAt?.toMillis?.() || 0) - (a.createdAt?.toMillis?.() || 0));
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
      const oldHtml = trigger?.innerHTML || trigger?.textContent || "📎 PDF";
      if (trigger) { trigger.disabled = true; trigger.innerHTML = "⏳ CARICO PDF…"; }
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
        if (trigger) { trigger.disabled = false; trigger.innerHTML = oldHtml; }
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
      .filter((plant) => plant?.lavoroOccasionale === true)
      .map((plant) => [plantId(plant) || plantName(plant), plant])).values()];
  }

  function decorateCards() {
    plants().forEach((plant) => {
      const id = plantId(plant);
      const name = normalize(plantName(plant));
      if (!id || !name) return;
      document.querySelectorAll("button").forEach((button) => {
        if (normalize(button.textContent) !== "FATTO") return;
        let card = button.parentElement;
        while (card && card !== document.body && !normalize(card.textContent).includes(name)) card = card.parentElement;
        if (!card || card === document.body || card.querySelector(`[data-occasional-pdf-plant="${CSS.escape(id)}"]`)) return;

        const section = document.createElement("section");
        section.className = "occasional-pdf-section";
        section.dataset.occasionalPdfPlant = id;
        section.innerHTML = `<div class="occasional-pdf-head"><button type="button" class="btn occasional-pdf-add">📎 ALLEGA PDF</button><small class="muted">PDF condivisi con gli utenti dell’app • max 15 MB</small></div><div class="occasional-pdf-list"></div>`;
        const hours = card.querySelector(".occasional-hours-actions");
        if (hours) hours.insertAdjacentElement("afterend", section);
        else button.parentElement?.insertAdjacentElement("afterend", section);
        const host = section.querySelector(".occasional-pdf-list");
        const add = section.querySelector(".occasional-pdf-add");
        add.addEventListener("click", () => choosePdfs(plant, host, add));
        void renderPdfList(plant, host);
      });
    });
  }

  function installStyles() {
    if (document.getElementById("occasional-pdf-storage-style")) return;
    const style = document.createElement("style");
    style.id = "occasional-pdf-storage-style";
    style.textContent = `
      .occasional-pdf-section{margin-top:10px;padding:10px;border:1px solid #dbe4ef;border-radius:14px;background:#f8fafc}
      .occasional-pdf-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap}.occasional-pdf-head .occasional-pdf-add{font-weight:800}
      .occasional-pdf-list{display:grid;gap:7px;margin-top:8px}.occasional-pdf-item{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:7px;align-items:center}
      .occasional-pdf-open{display:grid;grid-template-columns:30px minmax(0,1fr);gap:8px;align-items:center;width:100%;padding:9px 10px;border:1px solid #dbe4ef;border-radius:12px;background:#fff;text-align:left;color:inherit}
      .occasional-pdf-open>span:first-child{font-size:22px}.occasional-pdf-open strong,.occasional-pdf-open small{display:block;overflow-wrap:anywhere}.occasional-pdf-open small{margin-top:2px;color:#64748b}
      .occasional-pdf-delete{width:42px;height:42px;border:1px solid #fecaca;border-radius:12px;background:#fff7f7;font-size:18px}
    `;
    document.head.appendChild(style);
  }

  function enhanceOccasionalPopup(chooser, plant) {
    if (!chooser || !isOccasionalPlant(plant) || chooser.dataset.occasionalPdfStorage === "1") return;
    chooser.dataset.occasionalPdfStorage = "1";
    const button = chooser.querySelector('[data-photo-source="pdf"], .whazzup-photo-pdf-btn');
    if (!button) return;

    button.innerHTML = '<span aria-hidden="true">📄</span><strong>Allega PDF</strong><small>Firebase Storage • collegato al cantiere • senza scadenza</small>';
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopImmediatePropagation();
      const section = document.querySelector(`[data-occasional-pdf-plant="${CSS.escape(plantId(plant))}"]`);
      const host = section?.querySelector(".occasional-pdf-list") || null;
      choosePdfs(plant, host, button);
    }, true);

    chooser.querySelector("[data-shared-pdf-list]")?.remove();
    const heading = chooser.querySelector("#whazzup-photo-source-title");
    if (heading) heading.textContent = "Come vuoi aggiungere gli allegati?";
    const description = heading?.nextElementSibling;
    if (description?.tagName === "P") description.textContent = "Le foto restano nel flusso attuale. I PDF vengono salvati nel Firebase Storage del cantiere occasionale.";
  }

  function installPopupBridge() {
    if (state.popupBridgeInstalled) return true;
    if (!window.HeraWhazzupPdfV2 || typeof window.openWhazzupPhotoSourceChooser !== "function") return false;
    const original = window.openWhazzupPhotoSourceChooser;
    if (original.__occasionalPdfStorageBridge === true) {
      state.popupBridgeInstalled = true;
      return true;
    }

    function openWhazzupPhotoSourceChooserWithOccasionalStorage(plant) {
      const result = original.apply(this, arguments);
      queueMicrotask(() => {
        const choosers = Array.from(document.querySelectorAll(".whazzup-photo-source-chooser"));
        enhanceOccasionalPopup(choosers[choosers.length - 1], plant);
      });
      return result;
    }
    openWhazzupPhotoSourceChooserWithOccasionalStorage.__occasionalPdfStorageBridge = true;
    window.openWhazzupPhotoSourceChooser = openWhazzupPhotoSourceChooserWithOccasionalStorage;
    state.popupBridgeInstalled = true;
    return true;
  }

  function retryPopupBridge() {
    if (installPopupBridge()) return;
    state.popupRetries += 1;
    if (state.popupRetries < POPUP_MAX_RETRIES) window.setTimeout(retryPopupBridge, POPUP_RETRY_MS);
    else console.warn("Bridge PDF Lavori occasionali non installato: popup allegati non disponibile.");
  }

  function refresh() {
    if (state.refreshing) return;
    state.refreshing = true;
    try { installStyles(); decorateCards(); }
    finally { state.refreshing = false; }
  }

  state.observer = new MutationObserver(() => queueMicrotask(refresh));
  state.observer.observe(document.documentElement, { childList: true, subtree: true });
  document.addEventListener("change", (event) => { if (event.target?.id === "squadra-commessa") refresh(); });
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", refresh, { once: true });
  else refresh();

  window.HeraOccasionalPdfStorage = {
    installed: true,
    version: "1.1.0",
    refresh,
    isOccasionalPlant,
    uploadOnePdf,
    choosePdfs,
    loadPdfs
  };
  retryPopupBridge();
})();
