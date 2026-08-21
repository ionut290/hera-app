(function installSharedWhazzupPdfAttachments(global) {
  "use strict";

  const COLLECTION = "documents";
  const SOURCE = "whazzup-impianto-pdf";
  const MAX_FILE_SIZE = 15 * 1024 * 1024;
  const MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
  const DRIVE_CALLABLE_URL = "https://europe-west1-hera-app-6cd2b.cloudfunctions.net/uploadWhazzupPdfToDrive";
  const WRAP_RETRY_MS = 250;
  const WRAP_MAX_RETRIES = 80;
  let retryCount = 0;
  let installed = false;

  function normalize(value) {
    return String(value == null ? "" : value).trim();
  }

  function safeFileName(name) {
    const normalized = normalize(name || "documento.pdf")
      .replace(/[\\/:*?"<>|#%{}\[\]]+/g, "-")
      .replace(/\s+/g, " ")
      .trim();
    const shortened = normalized.length > 105 ? `${normalized.slice(0, 95)}-${Date.now()}.pdf` : normalized;
    return shortened.toLowerCase().endsWith(".pdf") ? shortened : `${shortened}.pdf`;
  }

  function getImpiantoKey(impianto) {
    if (typeof global.getWhazzupPhotoKey === "function") {
      try {
        const key = normalize(global.getWhazzupPhotoKey(impianto));
        if (key) return key;
      } catch (_) {}
    }
    return normalize(
      impianto?.id || impianto?.impiantoId || impianto?.physicalPlantId || impianto?.migrationSourceId
      || impianto?.idSap || impianto?.["ID SAP"] || impianto?.denominazione || impianto?.["Denominazione Impianto"]
    );
  }

  function getCommessaId(impianto) {
    return normalize(impianto?.commessaId || impianto?.parentCommessaId || impianto?.commessa?.id || global.selectedCommessaId);
  }

  function getCommessaName(impianto) {
    const direct = normalize(impianto?.commessaName || impianto?.commessaNome || impianto?.commessa?.nome);
    if (direct) return direct;
    const id = getCommessaId(impianto);
    try {
      const item = global.commesseById?.get?.(id);
      const name = normalize(item?.nome || item?.name);
      if (name) return name;
    } catch (_) {}
    return id || "Generale";
  }

  function getFirebaseServices() {
    const firebase = global.firebase;
    if (!firebase || !firebase.apps?.length || typeof firebase.firestore !== "function") return null;
    const db = global.db || firebase.firestore();
    const storage = typeof firebase.storage === "function" ? firebase.storage() : null;
    const user = global.currentUser || firebase.auth?.().currentUser || null;
    return { firebase, db, storage, user };
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

  function showFeedback(message, isError = false) {
    if (typeof global.showToast === "function") {
      try {
        global.showToast(message, isError ? "error" : "success");
        return;
      } catch (_) {}
    }
    if (isError) global.alert?.(message);
  }

  function shouldUseDriveFallback(error) {
    const text = `${error?.code || ""} ${error?.message || ""} ${error || ""}`.toLowerCase();
    return text.includes("storage/unknown") || text.includes("404") || text.includes("not found") || text.includes("bucket");
  }

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(reader.error || new Error("Lettura PDF non riuscita."));
      reader.onload = () => {
        const value = String(reader.result || "");
        resolve(value.includes(",") ? value.slice(value.indexOf(",") + 1) : value);
      };
      reader.readAsDataURL(file);
    });
  }

  async function uploadToDriveFallback(impianto, file, services, docRef, fileName) {
    const token = await services.user.getIdToken();
    const base64 = await fileToBase64(file);
    const response = await fetch(DRIVE_CALLABLE_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      },
      body: JSON.stringify({
        data: {
          base64,
          fileName,
          mimeType: "application/pdf",
          commessaId: getCommessaId(impianto),
          commessaName: getCommessaName(impianto)
        }
      })
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || payload?.error) {
      const message = payload?.error?.message || `Fallback Drive non disponibile (${response.status}).`;
      throw new Error(message);
    }
    const result = payload.result || payload.data || {};
    if (!result.fileId || !result.fileUrl) throw new Error("Drive non ha restituito il PDF caricato.");
    return {
      provider: "drive",
      fileUrl: result.fileUrl,
      storagePath: `drive:${result.fileId}`,
      driveFileId: result.fileId
    };
  }

  async function persistDocumentMetadata(impianto, file, services, docRef, fileName, uploadResult, expiresAtDate) {
    const { firebase, user } = services;
    const commessaId = getCommessaId(impianto);
    const now = firebase.firestore.FieldValue.serverTimestamp();
    await docRef.set({
      id: docRef.id,
      title: normalize(file.name) || fileName,
      fileName,
      fileUrl: uploadResult.fileUrl,
      storagePath: uploadResult.storagePath,
      storageProvider: uploadResult.provider,
      driveFileId: uploadResult.driveFileId || "",
      mimeType: "application/pdf",
      fileSize: Number(file.size),
      ownerUserId: user.uid,
      createdBy: user.uid,
      createdByEmail: normalize(user.email),
      createdByName: normalize(user.displayName || user.email || "Operatore"),
      visibility: "global",
      sharedToAll: false,
      sharedUserIds: [],
      commessaIds: commessaId ? [commessaId] : [],
      commessaId: commessaId || "",
      impiantoKey: getImpiantoKey(impianto),
      impiantoId: normalize(impianto?.id || impianto?.impiantoId),
      impiantoName: normalize(impianto?.denominazione || impianto?.["Denominazione Impianto"] || impianto?.nome),
      source: SOURCE,
      category: "WHAZZUP PDF",
      uploadStatus: "completed",
      versionHistoryEnabled: false,
      createdAt: now,
      updatedAt: now,
      expiresAt: firebase.firestore.Timestamp.fromDate(expiresAtDate),
      expiresAtIso: expiresAtDate.toISOString(),
      autoDeleteAfterDays: 30
    });
  }

  async function uploadPdf(impianto, file) {
    const services = getFirebaseServices();
    if (!services?.user?.uid) throw new Error("Accedi all'app prima di caricare un PDF.");
    if (!isPdf(file)) throw new Error("Puoi allegare soltanto file PDF.");
    if (!file.size) throw new Error("Il PDF selezionato è vuoto.");
    if (file.size > MAX_FILE_SIZE) throw new Error("Il PDF supera il limite massimo di 15 MB.");

    const impiantoKey = getImpiantoKey(impianto);
    if (!impiantoKey) throw new Error("Non riesco a identificare l'impianto per questo PDF.");

    const docRef = services.db.collection(COLLECTION).doc();
    const fileName = safeFileName(file.name);
    const expiresAtDate = new Date(Date.now() + MAX_AGE_MS);
    let uploadResult = null;
    let storageRef = null;

    if (services.storage) {
      const storagePath = `documents/${services.user.uid}/${docRef.id}/${fileName}`;
      storageRef = services.storage.ref().child(storagePath);
      try {
        await storageRef.put(file, {
          contentType: "application/pdf",
          customMetadata: { source: SOURCE, documentId: docRef.id, ownerUserId: services.user.uid, impiantoKey }
        });
        uploadResult = {
          provider: "storage",
          fileUrl: await storageRef.getDownloadURL(),
          storagePath,
          driveFileId: ""
        };
      } catch (error) {
        if (!shouldUseDriveFallback(error)) throw error;
        console.warn("Firebase Storage non disponibile per PDF: uso Drive centrale.", error);
      }
    }

    if (!uploadResult) {
      uploadResult = await uploadToDriveFallback(impianto, file, services, docRef, fileName);
    }

    try {
      await persistDocumentMetadata(impianto, file, services, docRef, fileName, uploadResult, expiresAtDate);
    } catch (error) {
      if (uploadResult.provider === "storage" && storageRef) await storageRef.delete().catch(() => null);
      throw error;
    }

    return { id: docRef.id, fileName, fileUrl: uploadResult.fileUrl, fileSize: file.size, expiresAt: expiresAtDate, provider: uploadResult.provider };
  }

  async function pickPdfs(impianto, chooser) {
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
      if (invalid) return showFeedback("Puoi allegare soltanto file PDF.", true);
      const oversized = files.find((file) => file.size > MAX_FILE_SIZE);
      if (oversized) return showFeedback(`Il file ${oversized.name} supera 15 MB.`, true);

      const button = chooser?.querySelector('[data-photo-source="pdf"]');
      const original = button?.innerHTML || "";
      if (button) {
        button.disabled = true;
        button.innerHTML = "<span aria-hidden=\"true\">⏳</span><strong>Caricamento PDF...</strong><small>Salvataggio condiviso in corso</small>";
      }
      try {
        for (const file of files) await uploadPdf(impianto, file);
        showFeedback(files.length === 1 ? "PDF condiviso con tutti gli utenti." : `${files.length} PDF condivisi con tutti gli utenti.`);
        await renderSharedPdfs(chooser, impianto);
      } catch (error) {
        console.error("Caricamento PDF Whazzup non riuscito:", error);
        showFeedback(`Caricamento PDF non riuscito: ${error?.message || error}`, true);
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

  async function loadSharedPdfs(impianto) {
    const services = getFirebaseServices();
    const impiantoKey = getImpiantoKey(impianto);
    if (!services?.user?.uid || !impiantoKey) return [];
    const snapshot = await services.db.collection(COLLECTION).where("impiantoKey", "==", impiantoKey).get();
    const now = Date.now();
    return snapshot.docs
      .map((doc) => ({ id: doc.id, ...doc.data() }))
      .filter((item) => item.source === SOURCE && item.uploadStatus === "completed")
      .filter((item) => {
        const expires = item.expiresAt?.toMillis?.() || Date.parse(item.expiresAtIso || "") || 0;
        return !expires || expires > now;
      })
      .sort((a, b) => (b.createdAt?.toMillis?.() || 0) - (a.createdAt?.toMillis?.() || 0));
  }

  async function shareOrOpenPdf(item) {
    const url = normalize(item.fileUrl);
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
      host.innerHTML = `<div class="whazzup-shared-pdf-heading"><strong>PDF condivisi</strong><small>${items.length} disponibili per tutti</small></div>${items.map((item) => {
        const title = normalize(item.title || item.fileName || "Documento PDF");
        const size = formatSize(item.fileSize);
        const expiry = item.expiresAt?.toDate?.();
        const expiryText = expiry instanceof Date && !Number.isNaN(expiry.getTime()) ? ` • fino al ${expiry.toLocaleDateString("it-IT")}` : "";
        return `<button type="button" class="whazzup-shared-pdf-item" data-shared-pdf-id="${item.id}"><span aria-hidden="true">📄</span><span><strong>${title.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}</strong><small>${size}${expiryText}</small></span></button>`;
      }).join("")}`;
      host.querySelectorAll("[data-shared-pdf-id]").forEach((button) => {
        button.addEventListener("click", () => {
          const item = items.find((entry) => entry.id === button.dataset.sharedPdfId);
          if (item) void shareOrOpenPdf(item);
        });
      });
    } catch (error) {
      console.warn("Lettura PDF condivisi Whazzup non disponibile:", error);
      host.innerHTML = '<p class="muted whazzup-shared-pdf-empty">PDF condivisi non disponibili in questo momento.</p>';
    }
  }

  function enhanceChooser(chooser, impianto) {
    if (!chooser || chooser.dataset.sharedPdfEnhanced === "1") return;
    chooser.dataset.sharedPdfEnhanced = "1";
    const actions = chooser.querySelector(".whazzup-photo-source-actions");
    if (!actions) return;
    const pdfButton = document.createElement("button");
    pdfButton.type = "button";
    pdfButton.className = "btn whazzup-photo-pdf-btn";
    pdfButton.dataset.photoSource = "pdf";
    pdfButton.innerHTML = '<span aria-hidden="true">📄</span><strong>Aggiungi PDF</strong><small>Documento condiviso con tutti</small>';
    pdfButton.addEventListener("click", () => void pickPdfs(impianto, chooser));
    actions.appendChild(pdfButton);

    const heading = chooser.querySelector("#whazzup-photo-source-title");
    if (heading) heading.textContent = "Come vuoi aggiungere gli allegati?";
    const description = heading?.nextElementSibling;
    if (description?.tagName === "P") description.textContent = "Le foto restano sul dispositivo; i PDF vengono condivisi con tutti gli utenti per 30 giorni.";
    void renderSharedPdfs(chooser, impianto);
  }

  function installStyles() {
    if (document.getElementById("whazzup-shared-pdf-styles")) return;
    const style = document.createElement("style");
    style.id = "whazzup-shared-pdf-styles";
    style.textContent = `
      .whazzup-photo-source-actions .whazzup-photo-pdf-btn{display:grid;grid-template-columns:56px 1fr;grid-template-rows:auto auto;column-gap:12px;align-items:center;text-align:left;width:100%;min-height:84px;border:1px solid rgba(13,148,136,.28);background:rgba(240,253,250,.94)}
      .whazzup-photo-source-actions .whazzup-photo-pdf-btn>span{grid-row:1/3;font-size:30px;text-align:center}.whazzup-photo-source-actions .whazzup-photo-pdf-btn strong{font-size:1.02rem}.whazzup-photo-source-actions .whazzup-photo-pdf-btn small{opacity:.72}
      .whazzup-shared-pdf-list{margin-top:14px;border-top:1px solid rgba(15,23,42,.1);padding-top:12px}.whazzup-shared-pdf-heading{display:flex;align-items:baseline;justify-content:space-between;gap:10px;margin-bottom:8px}.whazzup-shared-pdf-heading small{opacity:.62}
      .whazzup-shared-pdf-item{display:grid;grid-template-columns:36px 1fr;gap:8px;align-items:center;width:100%;border:1px solid rgba(15,23,42,.1);border-radius:12px;background:#fff;padding:10px;margin:7px 0;text-align:left}.whazzup-shared-pdf-item>span:first-child{font-size:24px}.whazzup-shared-pdf-item strong,.whazzup-shared-pdf-item small{display:block}.whazzup-shared-pdf-item small{margin-top:2px;opacity:.65}.whazzup-shared-pdf-loading,.whazzup-shared-pdf-empty{margin:4px 0}
    `;
    document.head.appendChild(style);
  }

  function tryInstall() {
    if (installed) return true;
    const original = global.openWhazzupPhotoSourceChooser;
    if (typeof original !== "function") return false;
    installStyles();
    global.openWhazzupPhotoSourceChooser = function openWhazzupPhotoSourceChooserWithPdf(impianto) {
      const result = original.apply(this, arguments);
      queueMicrotask(() => {
        const choosers = Array.from(document.querySelectorAll(".whazzup-photo-source-chooser"));
        enhanceChooser(choosers[choosers.length - 1], impianto);
      });
      return result;
    };
    installed = true;
    global.HeraSharedWhazzupPdf = Object.freeze({ uploadPdf, loadSharedPdfs, source: SOURCE, maxAgeMs: MAX_AGE_MS });
    return true;
  }

  function retryInstall() {
    if (tryInstall()) return;
    retryCount += 1;
    if (retryCount < WRAP_MAX_RETRIES) global.setTimeout(retryInstall, WRAP_RETRY_MS);
    else console.warn("Modulo PDF Whazzup non installato: selettore foto non disponibile.");
  }

  retryInstall();
})(globalThis);
