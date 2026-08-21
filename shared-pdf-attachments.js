(function installWhazzupPdfDriveV2(global) {
  "use strict";

  const COLLECTION = "documents";
  const SOURCE = "whazzup-pdf-drive-v2";
  const MAX_FILE_SIZE = 15 * 1024 * 1024;
  const MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
  const FUNCTION_NAME = "uploadWhazzupPdfToDrive";
  const REGION = "europe-west1";
  const RETRY_MS = 250;
  const MAX_RETRIES = 80;
  let retries = 0;
  let installed = false;

  function text(value) {
    return String(value == null ? "" : value).trim();
  }

  function escapeHtml(value) {
    return text(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function safeFileName(name) {
    const raw = text(name || "documento.pdf")
      .replace(/[\\/:*?\"<>|#%{}\[\]]+/g, "-")
      .replace(/\s+/g, " ")
      .trim();
    const pdf = raw.toLowerCase().endsWith(".pdf") ? raw : `${raw}.pdf`;
    return pdf.length <= 110 ? pdf : `${pdf.slice(0, 96)}-${Date.now()}.pdf`;
  }

  function getServices() {
    const firebase = global.firebase;
    if (!firebase?.apps?.length) return null;
    if (typeof firebase.firestore !== "function" || typeof firebase.functions !== "function") return null;
    const user = global.currentUser || firebase.auth?.().currentUser || null;
    return {
      firebase,
      db: global.db || firebase.firestore(),
      user,
      functions: firebase.app().functions(REGION)
    };
  }

  function getImpiantoKey(impianto) {
    if (typeof global.getWhazzupPhotoKey === "function") {
      try {
        const key = text(global.getWhazzupPhotoKey(impianto));
        if (key) return key;
      } catch (_) {}
    }
    return text(
      impianto?.id || impianto?.impiantoId || impianto?.physicalPlantId || impianto?.migrationSourceId
      || impianto?.idSap || impianto?.["ID SAP"] || impianto?.denominazione || impianto?.["Denominazione Impianto"]
    );
  }

  function getCommessaId(impianto) {
    return text(impianto?.commessaId || impianto?.parentCommessaId || impianto?.commessa?.id || global.selectedCommessaId);
  }

  function getCommessaName(impianto) {
    const direct = text(impianto?.commessaName || impianto?.commessaNome || impianto?.commessa?.nome);
    if (direct) return direct;
    const id = getCommessaId(impianto);
    try {
      const item = global.commesseById?.get?.(id);
      const name = text(item?.nome || item?.name);
      if (name) return name;
    } catch (_) {}
    return id || "Generale";
  }

  function isPdf(file) {
    return file instanceof Blob && (
      String(file.type || "").toLowerCase() === "application/pdf"
      || String(file.name || "").toLowerCase().endsWith(".pdf")
    );
  }

  function formatSize(bytes) {
    const value = Number(bytes || 0);
    if (value < 1024) return `${value} B`;
    if (value < 1024 * 1024) return `${Math.round(value / 1024)} KB`;
    return `${(value / (1024 * 1024)).toFixed(1)} MB`;
  }

  function feedback(message, isError = false) {
    if (typeof global.showToast === "function") {
      try {
        global.showToast(message, isError ? "error" : "success");
        return;
      } catch (_) {}
    }
    if (isError) global.alert?.(message);
  }

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(reader.error || new Error("Impossibile leggere il PDF."));
      reader.onload = () => {
        const value = String(reader.result || "");
        resolve(value.includes(",") ? value.slice(value.indexOf(",") + 1) : value);
      };
      reader.readAsDataURL(file);
    });
  }

  async function uploadPdf(impianto, file) {
    const services = getServices();
    if (!services?.user?.uid) throw new Error("Accedi all'app prima di caricare un PDF.");
    if (!isPdf(file)) throw new Error("Puoi allegare soltanto file PDF.");
    if (!file.size) throw new Error("Il PDF selezionato è vuoto.");
    if (file.size > MAX_FILE_SIZE) throw new Error("Il PDF supera il limite massimo di 15 MB.");

    const impiantoKey = getImpiantoKey(impianto);
    if (!impiantoKey) throw new Error("Non riesco a identificare l'impianto.");

    const fileName = safeFileName(file.name);
    const base64 = await fileToBase64(file);
    const callable = services.functions.httpsCallable(FUNCTION_NAME);
    const response = await callable({
      base64,
      fileName,
      mimeType: "application/pdf",
      commessaId: getCommessaId(impianto),
      commessaName: getCommessaName(impianto)
    });
    const result = response?.data || {};
    const fileId = text(result.fileId);
    const fileUrl = text(result.fileUrl);
    if (!fileId || !fileUrl) throw new Error("Il server non ha restituito il PDF caricato.");

    const docRef = services.db.collection(COLLECTION).doc();
    const expiresAtDate = new Date(Date.now() + MAX_AGE_MS);
    const now = services.firebase.firestore.FieldValue.serverTimestamp();
    await docRef.set({
      id: docRef.id,
      title: text(file.name) || fileName,
      fileName,
      fileUrl,
      driveFileId: fileId,
      storageProvider: "drive",
      storagePath: `drive:${fileId}`,
      mimeType: "application/pdf",
      fileSize: Number(file.size || 0),
      ownerUserId: services.user.uid,
      createdBy: services.user.uid,
      createdByEmail: text(services.user.email),
      createdByName: text(services.user.displayName || services.user.email || "Operatore"),
      visibility: "global",
      sharedToAll: true,
      commessaId: getCommessaId(impianto),
      commessaIds: getCommessaId(impianto) ? [getCommessaId(impianto)] : [],
      commessaName: getCommessaName(impianto),
      impiantoKey,
      impiantoId: text(impianto?.id || impianto?.impiantoId),
      impiantoName: text(impianto?.denominazione || impianto?.["Denominazione Impianto"] || impianto?.nome),
      source: SOURCE,
      category: "WHAZZUP PDF",
      uploadStatus: "completed",
      createdAt: now,
      updatedAt: now,
      expiresAt: services.firebase.firestore.Timestamp.fromDate(expiresAtDate),
      expiresAtIso: expiresAtDate.toISOString(),
      autoDeleteAfterDays: 30
    });

    return { id: docRef.id, fileUrl, fileId, fileName, expiresAt: expiresAtDate };
  }

  async function loadSharedPdfs(impianto) {
    const services = getServices();
    const impiantoKey = getImpiantoKey(impianto);
    if (!services?.user?.uid || !impiantoKey) return [];
    const snapshot = await services.db.collection(COLLECTION).where("impiantoKey", "==", impiantoKey).get();
    const now = Date.now();
    return snapshot.docs
      .map((doc) => ({ id: doc.id, ...doc.data() }))
      .filter((item) => item.source === SOURCE && item.uploadStatus === "completed")
      .filter((item) => {
        const expiry = item.expiresAt?.toMillis?.() || Date.parse(item.expiresAtIso || "") || 0;
        return !expiry || expiry > now;
      })
      .sort((a, b) => (b.createdAt?.toMillis?.() || 0) - (a.createdAt?.toMillis?.() || 0));
  }

  async function openOrShare(item) {
    const url = text(item?.fileUrl);
    if (!url) return;
    if (navigator.share) {
      try {
        await navigator.share({ title: item.title || item.fileName || "PDF", url });
        return;
      } catch (error) {
        if (error?.name === "AbortError") return;
      }
    }
    global.open(url, "_blank", "noopener,noreferrer");
  }

  async function renderSharedPdfs(chooser, impianto) {
    if (!chooser?.isConnected) return;
    let host = chooser.querySelector("[data-shared-pdf-list]");
    if (!host) {
      host = document.createElement("section");
      host.dataset.sharedPdfList = "1";
      host.className = "whazzup-shared-pdf-list";
      const card = chooser.querySelector(".whazzup-photo-source-card") || chooser.firstElementChild;
      card?.appendChild(host);
    }
    if (!host) return;
    host.innerHTML = '<p class="muted whazzup-shared-pdf-loading">Carico i PDF condivisi…</p>';
    try {
      const items = await loadSharedPdfs(impianto);
      if (!items.length) {
        host.innerHTML = '<p class="muted whazzup-shared-pdf-empty">Nessun PDF condiviso per questo impianto.</p>';
        return;
      }
      host.innerHTML = `<div class="whazzup-shared-pdf-heading"><strong>PDF condivisi</strong><small>${items.length} disponibili</small></div>${items.map((item) => {
        const expiry = item.expiresAt?.toDate?.();
        const expiryText = expiry instanceof Date && !Number.isNaN(expiry.getTime()) ? ` • fino al ${expiry.toLocaleDateString("it-IT")}` : "";
        return `<button type="button" class="whazzup-shared-pdf-item" data-shared-pdf-id="${escapeHtml(item.id)}"><span aria-hidden="true">📄</span><span><strong>${escapeHtml(item.title || item.fileName || "Documento PDF")}</strong><small>${formatSize(item.fileSize)}${expiryText}</small></span></button>`;
      }).join("")}`;
      host.querySelectorAll("[data-shared-pdf-id]").forEach((button) => {
        button.addEventListener("click", () => {
          const item = items.find((entry) => entry.id === button.dataset.sharedPdfId);
          if (item) void openOrShare(item);
        });
      });
    } catch (error) {
      console.warn("Lettura PDF condivisi non disponibile:", error);
      host.innerHTML = '<p class="muted whazzup-shared-pdf-empty">PDF condivisi non disponibili in questo momento.</p>';
    }
  }

  function choosePdfs(impianto, chooser) {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "application/pdf,.pdf";
    input.multiple = true;
    input.hidden = true;
    input.addEventListener("change", async () => {
      const files = Array.from(input.files || []);
      input.remove();
      if (!files.length) return;
      const invalid = files.find((file) => !isPdf(file));
      if (invalid) return feedback("Puoi allegare soltanto file PDF.", true);
      const oversized = files.find((file) => file.size > MAX_FILE_SIZE);
      if (oversized) return feedback(`Il file ${oversized.name} supera 15 MB.`, true);

      const button = chooser?.querySelector('[data-photo-source="pdf"]');
      const original = button?.innerHTML || "";
      if (button) {
        button.disabled = true;
        button.innerHTML = '<span aria-hidden="true">⏳</span><strong>Caricamento PDF…</strong><small>Salvataggio su Drive centrale</small>';
      }
      try {
        for (const file of files) await uploadPdf(impianto, file);
        feedback(files.length === 1 ? "PDF caricato e condiviso con tutti." : `${files.length} PDF caricati e condivisi con tutti.`);
        await renderSharedPdfs(chooser, impianto);
      } catch (error) {
        console.error("Upload PDF Whazzup V2 fallito:", error);
        const code = text(error?.code);
        const message = text(error?.message || error);
        feedback(`Caricamento PDF non riuscito${code ? ` (${code})` : ""}: ${message || "errore sconosciuto"}`, true);
      } finally {
        if (button) {
          button.disabled = false;
          button.innerHTML = original;
        }
      }
    }, { once: true });
    input.addEventListener("cancel", () => input.remove(), { once: true });
    document.body.appendChild(input);
    input.click();
  }

  function installStyles() {
    if (document.getElementById("whazzup-pdf-drive-v2-styles")) return;
    const style = document.createElement("style");
    style.id = "whazzup-pdf-drive-v2-styles";
    style.textContent = `
      .whazzup-photo-source-actions .whazzup-photo-pdf-btn{display:grid;grid-template-columns:56px 1fr;grid-template-rows:auto auto;column-gap:12px;align-items:center;text-align:left;width:100%;min-height:84px;border:1px solid rgba(13,148,136,.28);background:rgba(240,253,250,.94)}
      .whazzup-photo-source-actions .whazzup-photo-pdf-btn>span{grid-row:1/3;font-size:30px;text-align:center}.whazzup-photo-source-actions .whazzup-photo-pdf-btn strong{font-size:1.02rem}.whazzup-photo-source-actions .whazzup-photo-pdf-btn small{opacity:.72}
      .whazzup-shared-pdf-list{margin-top:14px;border-top:1px solid rgba(15,23,42,.1);padding-top:12px}.whazzup-shared-pdf-heading{display:flex;align-items:baseline;justify-content:space-between;gap:10px;margin-bottom:8px}.whazzup-shared-pdf-heading small{opacity:.62}
      .whazzup-shared-pdf-item{display:grid;grid-template-columns:36px 1fr;gap:8px;align-items:center;width:100%;border:1px solid rgba(15,23,42,.1);border-radius:12px;background:#fff;padding:10px;margin:7px 0;text-align:left}.whazzup-shared-pdf-item>span:first-child{font-size:24px}.whazzup-shared-pdf-item strong,.whazzup-shared-pdf-item small{display:block}.whazzup-shared-pdf-item small{margin-top:2px;opacity:.65}.whazzup-shared-pdf-loading,.whazzup-shared-pdf-empty{margin:4px 0}
    `;
    document.head.appendChild(style);
  }

  function enhanceChooser(chooser, impianto) {
    if (!chooser || chooser.dataset.pdfDriveV2 === "1") return;
    chooser.dataset.pdfDriveV2 = "1";
    const actions = chooser.querySelector(".whazzup-photo-source-actions");
    if (!actions) return;
    const pdfButton = document.createElement("button");
    pdfButton.type = "button";
    pdfButton.className = "btn whazzup-photo-pdf-btn";
    pdfButton.dataset.photoSource = "pdf";
    pdfButton.innerHTML = '<span aria-hidden="true">📄</span><strong>Aggiungi PDF</strong><small>Condiviso con tutti per 30 giorni</small>';
    pdfButton.addEventListener("click", () => choosePdfs(impianto, chooser));
    actions.appendChild(pdfButton);

    const heading = chooser.querySelector("#whazzup-photo-source-title");
    if (heading) heading.textContent = "Come vuoi aggiungere gli allegati?";
    const description = heading?.nextElementSibling;
    if (description?.tagName === "P") description.textContent = "Le foto restano sul dispositivo. I PDF vengono salvati sul Drive centrale e condivisi con tutti per 30 giorni.";
    void renderSharedPdfs(chooser, impianto);
  }

  function tryInstall() {
    if (installed) return true;
    const original = global.openWhazzupPhotoSourceChooser;
    if (typeof original !== "function") return false;
    installStyles();
    global.openWhazzupPhotoSourceChooser = function openWhazzupPhotoSourceChooserWithPdfV2(impianto) {
      const result = original.apply(this, arguments);
      queueMicrotask(() => {
        const choosers = Array.from(document.querySelectorAll(".whazzup-photo-source-chooser"));
        enhanceChooser(choosers[choosers.length - 1], impianto);
      });
      return result;
    };
    installed = true;
    global.HeraWhazzupPdfV2 = Object.freeze({ uploadPdf, loadSharedPdfs, source: SOURCE });
    return true;
  }

  function retryInstall() {
    if (tryInstall()) return;
    retries += 1;
    if (retries < MAX_RETRIES) global.setTimeout(retryInstall, RETRY_MS);
    else console.warn("PDF Whazzup V2 non installato: selettore foto non disponibile.");
  }

  retryInstall();
})(globalThis);
