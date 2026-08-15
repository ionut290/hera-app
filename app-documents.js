"use strict";
(function installVargaDocumentsModule(global) {
  if (global.VargaDocumentsModule) return;
  const api = {};
  function openPrivateDocsPage() {
    if (!currentUser) {
      alert("Devi fare login per usare i documenti.");
      return;
    }
    window.location.hash = "documenti";
    applyRoute();
    window.HeraDocuments?.activate?.();
    closeSideMenu();
  }
  api.openPrivateDocsPage = openPrivateDocsPage;
  function openPrivateDocsUploadPage() {
    openPrivateDocsPage();
    applyPrivateDocPreset("pin");
    setTimeout(() => {
      ui.privateDocsForm?.scrollIntoView({ behavior: "smooth", block: "start" });
      ui.privateDocsName?.focus();
    }, 50);
  }
  api.openPrivateDocsUploadPage = openPrivateDocsUploadPage;
  function closePrivateDocsPage() {
    window.HeraDocuments?.deactivate?.();
    window.location.hash = "";
    applyRoute();
  }
  api.closePrivateDocsPage = closePrivateDocsPage;
  function stopPosDocumentsSubscription() {
    if (unsubscribePosDocuments) {
      unsubscribePosDocuments();
      unsubscribePosDocuments = null;
    }
    posDocuments = [];
  }
  api.stopPosDocumentsSubscription = stopPosDocumentsSubscription;
  function subscribePosDocuments() {
    stopPosDocumentsSubscription();
    const query = canManageData()
      ? db.collection("posDocuments")
      : db.collection("posDocuments").where("active", "==", true);
    unsubscribePosDocuments = query.onSnapshot((snapshot) => {
      posDocuments = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      renderPosDocuments();
    }, (error) => {
      console.error("Errore caricamento documenti POS:", error);
      if (ui.posDocumentsList) ui.posDocumentsList.innerHTML = "<p class='muted'>Impossibile caricare i documenti POS.</p>";
    });
  }
  api.subscribePosDocuments = subscribePosDocuments;
  function getFilteredPosDocuments() {
    const canManage = canManageData();
    const search = String(ui.posSearch?.value || "").trim().toLowerCase();
    return posDocuments
      .filter((doc) => canManage || doc.active === true)
      .filter((doc) => {
        if (!search) return true;
        return [doc.title, doc.description, doc.category]
          .some((value) => String(value || "").toLowerCase().includes(search));
      })
      .sort((a, b) => {
        const categoryCompare = String(a.category || "Altro").localeCompare(String(b.category || "Altro"), "it");
        if (categoryCompare !== 0) return categoryCompare;
        const orderCompare = Number(a.order || 0) - Number(b.order || 0);
        if (orderCompare !== 0) return orderCompare;
        return String(a.title || "").localeCompare(String(b.title || ""), "it");
      });
  }
  api.getFilteredPosDocuments = getFilteredPosDocuments;
  function renderPosDocuments() {
    if (!ui.posDocumentsList) return;
    const canManage = canManageData();
    updateDriveConnectVisibility();
    ui.openPosBtn?.classList.remove("hidden");
    if (ui.openPosBtn) ui.openPosBtn.disabled = false;
    ui.posAdminCard?.classList.toggle("hidden", !canManage);
    const documents = getFilteredPosDocuments();
    if (!documents.length) {
      ui.posDocumentsList.innerHTML = "<p class='muted'>Nessun documento disponibile.</p>";
      return;
    }
    ui.posDocumentsList.innerHTML = "";
    const grouped = new Map();
    documents.forEach((doc) => {
      const category = String(doc.category || "Altro").trim() || "Altro";
      if (!grouped.has(category)) grouped.set(category, []);
      grouped.get(category).push(doc);
    });
    grouped.forEach((items, category) => {
      const group = document.createElement("section");
      group.className = "pos-category-group";
      group.innerHTML = `<h3>📁 ${escapeHTML(category)}</h3>`;
      const grid = document.createElement("div");
      grid.className = "pos-document-grid";
      items.forEach((doc) => grid.appendChild(createPosDocumentCard(doc, canManage)));
      group.appendChild(grid);
      ui.posDocumentsList.appendChild(group);
    });
  }
  api.renderPosDocuments = renderPosDocuments;
  function createPosDocumentCard(doc, canManage) {
    const card = document.createElement("article");
    card.className = "pos-document-card";
    if (doc.active === false) card.classList.add("is-inactive");
    const title = document.createElement("h4");
    title.textContent = doc.title || "Documento senza titolo";
    const description = document.createElement("p");
    description.className = "muted";
    description.textContent = doc.description || "Nessuna descrizione.";
    const actions = document.createElement("div");
    actions.className = "item-actions pos-document-actions";
    const driveUrl = String(doc.driveUrl || "").trim();
    if (driveUrl) {
      const link = document.createElement("a");
      link.className = "btn pos-open-link";
      link.href = driveUrl;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = "Apri documento";
      actions.appendChild(link);
    } else {
      const unavailable = document.createElement("p");
      unavailable.className = "muted pos-unavailable";
      unavailable.textContent = "Documento non disponibile.";
      actions.appendChild(unavailable);
    }
    if (canManage) {
      const editBtn = createButton("Modifica", () => openPosDocumentForm(doc));
      const deleteBtn = createButton("Elimina", () => deletePosDocument(doc));
      actions.appendChild(editBtn);
      actions.appendChild(deleteBtn);
      const meta = document.createElement("p");
      meta.className = "muted pos-admin-meta";
      meta.textContent = `Ordine: ${Number(doc.order || 0)} • ${doc.active === false ? "Non attivo" : "Attivo"}`;
      card.append(title, description, meta, actions);
      return card;
    }
    card.append(title, description, actions);
    return card;
  }
  api.createPosDocumentCard = createPosDocumentCard;
  function openPosDocumentForm(doc = null) {
    if (!canManageData()) return;
    ui.posDocumentForm?.classList.remove("hidden");
    if (ui.posAddToggleBtn) ui.posAddToggleBtn.textContent = doc ? "Modifica documento" : "➕ Aggiungi documento";
    if (ui.posDocumentId) ui.posDocumentId.value = doc?.id || "";
    if (ui.posTitle) ui.posTitle.value = doc?.title || "";
    if (ui.posDescription) ui.posDescription.value = doc?.description || "";
    if (ui.posDriveUrl) ui.posDriveUrl.value = doc?.driveUrl || "";
    if (ui.posCategory) ui.posCategory.value = doc?.category || POS_DEFAULT_CATEGORIES[0];
    if (ui.posOrder) ui.posOrder.value = Number(doc?.order || 0);
    if (ui.posActive) ui.posActive.checked = doc?.active !== false;
    ui.posTitle?.focus();
  }
  api.openPosDocumentForm = openPosDocumentForm;
  function closePosDocumentForm() {
    ui.posDocumentForm?.reset();
    if (ui.posDocumentId) ui.posDocumentId.value = "";
    if (ui.posActive) ui.posActive.checked = true;
    if (ui.posAddToggleBtn) ui.posAddToggleBtn.textContent = "➕ Aggiungi documento";
    ui.posDocumentForm?.classList.add("hidden");
    if (ui.posFeedback) ui.posFeedback.textContent = "";
  }
  api.closePosDocumentForm = closePosDocumentForm;
  async function savePosDocument(event) {
    event.preventDefault();
    if (!canManageData()) {
      alert("Solo l'admin può salvare documenti POS.");
      return;
    }
    const id = String(ui.posDocumentId?.value || "").trim();
    const now = firebase.firestore.FieldValue.serverTimestamp();
    const payload = {
      title: String(ui.posTitle?.value || "").trim(),
      description: String(ui.posDescription?.value || "").trim(),
      driveUrl: String(ui.posDriveUrl?.value || "").trim(),
      category: String(ui.posCategory?.value || "").trim() || "Altro",
      order: Number(ui.posOrder?.value || 0),
      active: Boolean(ui.posActive?.checked),
      updatedAt: now
    };
    if (!payload.title) {
      alert("Inserisci il titolo documento.");
      return;
    }
    if (id) {
      await db.collection("posDocuments").doc(id).set(payload, { merge: true });
    } else {
      await db.collection("posDocuments").add({
        ...payload,
        createdAt: now,
        createdBy: currentUser?.email || ""
      });
    }
    if (ui.posFeedback) ui.posFeedback.textContent = "Documento salvato.";
    closePosDocumentForm();
  }
  api.savePosDocument = savePosDocument;
  async function deletePosDocument(doc) {
    if (!canManageData()) {
      alert("Solo l'admin può eliminare documenti POS.");
      return;
    }
    const ok = window.confirm(`Eliminare il documento "${doc.title || "senza titolo"}"?`);
    if (!ok) return;
    await db.collection("posDocuments").doc(doc.id).delete();
  }
  api.deletePosDocument = deletePosDocument;
  function subscribePrivateDocs() {
    // The document archive owns both the new privacy-scoped queries and the
    // historical per-user subscription (documents.js). Keeping the former broad
    // renderer active here would overwrite the tabbed archive with the legacy list.
    return undefined;
  }
  api.subscribePrivateDocs = subscribePrivateDocs;
  function stopPrivateDocsSubscription() {
    if (unsubscribePrivateDocs) {
      unsubscribePrivateDocs();
      unsubscribePrivateDocs = null;
    }
  }
  api.stopPrivateDocsSubscription = stopPrivateDocsSubscription;
  function getPrivateDocsDriveToken() {
    return String(localStorage.getItem("googleDriveAccessToken") || "").trim();
  }
  api.getPrivateDocsDriveToken = getPrivateDocsDriveToken;
  async function getOrCreatePrivateDocsFolder(token, uid) {
    const query = [
      "name='Hera App - Documenti privati'",
      "mimeType='application/vnd.google-apps.folder'",
      "trashed=false"
    ].join(" and ");
    const rootSearch = await driveApiFetchWithToken(token, `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name)&pageSize=1`, { method: "GET" });
    let rootFolderId = rootSearch?.files?.[0]?.id || "";
    if (!rootFolderId) {
      const createdRoot = await driveApiFetchWithToken(token, "https://www.googleapis.com/drive/v3/files", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: "Hera App - Documenti privati", mimeType: "application/vnd.google-apps.folder" })
      });
      rootFolderId = createdRoot.id;
    }
  
    const userQuery = [
      `name='${String(uid || "").replace(/'/g, "\\'")}'`,
      "mimeType='application/vnd.google-apps.folder'",
      "trashed=false",
      `'${rootFolderId}' in parents`
    ].join(" and ");
    const userSearch = await driveApiFetchWithToken(token, `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(userQuery)}&fields=files(id,name)&pageSize=1`, { method: "GET" });
    const existingUserFolder = userSearch?.files?.[0]?.id || "";
    if (existingUserFolder) return existingUserFolder;
  
    const createdUserFolder = await driveApiFetchWithToken(token, "https://www.googleapis.com/drive/v3/files", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: uid, mimeType: "application/vnd.google-apps.folder", parents: [rootFolderId] })
    });
    return createdUserFolder.id;
  }
  api.getOrCreatePrivateDocsFolder = getOrCreatePrivateDocsFolder;
  async function uploadPrivateDocumentToDrive(file, uid) {
    const token = getPrivateDocsDriveToken();
    if (!token) {
      throw new Error("Google Drive non autorizzato. Rifai il login Google prima di usare il salvataggio Drive.");
    }
    const folderId = await getOrCreatePrivateDocsFolder(token, uid);
    const metadata = {
      name: file.name || "documento",
      parents: [folderId]
    };
    const boundary = "hera-private-doc-upload";
    const body = [
      `--${boundary}`,
      "Content-Type: application/json; charset=UTF-8",
      "",
      JSON.stringify(metadata),
      `--${boundary}`,
      `Content-Type: ${file.type || "application/octet-stream"}`,
      "",
      file,
      `--${boundary}--`
    ];
    const payload = new Blob(body, { type: `multipart/related; boundary=${boundary}` });
    const uploaded = await driveApiFetchWithToken(token, "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink", {
      method: "POST",
      headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
      body: payload
    });
    return {
      driveFileId: uploaded.id || "",
      driveFileName: uploaded.name || file.name || "documento",
      driveWebViewLink: uploaded.webViewLink || `https://drive.google.com/file/d/${uploaded.id}/view`
    };
  }
  api.uploadPrivateDocumentToDrive = uploadPrivateDocumentToDrive;
  async function savePrivateDocument(event) {
    event.preventDefault();
    if (!currentUser) return;
    try {
      const name = String(ui.privateDocsName.value || "").trim();
      const note = String(ui.privateDocsNote.value || "").trim();
      const file = ui.privateDocsFile.files?.[0] || ui.privateDocsCamera.files?.[0] || null;
      if (!name) {
        ui.privateDocsFeedback.textContent = "La denominazione è obbligatoria.";
        return;
      }
      let fileDataUrl = "";
      let fileName = "";
      let fileType = "";
      let fileSize = 0;
      let driveFileId = "";
      let driveWebViewLink = "";
      const useDriveUpload = Boolean(file);
      if (file) {
        fileSize = Number(file.size || 0);
        fileName = file.name || "documento";
        fileType = file.type || "application/octet-stream";
        if (useDriveUpload) {
          ui.privateDocsFeedback.textContent = "Caricamento sul cloud centralizzato...";
          const upload = await uploadBlobToDrive(file, fileName, fileType, driveReportsFolderId, { driveType: "DOCUMENTI", commessaName: "Documenti" });
          driveFileId = upload.fileId;
          driveWebViewLink = upload.webViewLink;
        } else {
          fileDataUrl = await readFileAsDataUrl(file);
        }
      }
      await db.collection("privateDocuments").doc(currentUser.uid).collection("items").add({
        name,
        note,
        fileName,
        fileType,
        fileSize,
        fileDataUrl,
        driveFileId,
        driveWebViewLink,
        storageMode: driveFileId ? "drive" : "firestore",
        createdAt: firebase.firestore.FieldValue.serverTimestamp(),
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      });
      ui.privateDocsForm.reset();
      ui.privateDocsFeedback.textContent = "Documento personale salvato.";
    } catch (error) {
      console.error("Salvataggio documento personale non riuscito:", error);
      ui.privateDocsFeedback.textContent = error?.message || "Errore durante il salvataggio del documento.";
    }
  }
  api.savePrivateDocument = savePrivateDocument;
  async function deletePrivateDocument(docId) {
    if (!currentUser || !docId) return;
    const ok = window.confirm("Eliminare questo documento personale?");
    if (!ok) return;
    await db.collection("privateDocuments").doc(currentUser.uid).collection("items").doc(docId).delete();
  }
  api.deletePrivateDocument = deletePrivateDocument;
  function renderPrivateDocsList() {
    if (!ui.privateDocsList) return;
    if (!currentUser) {
      ui.privateDocsList.innerHTML = "<p class='muted'>Fai login per usare i documenti.</p>";
      return;
    }
    if (!privateDocsRecords.length) {
      ui.privateDocsList.innerHTML = "<p class='muted'>Nessun documento personale salvato.</p>";
      return;
    }
    ui.privateDocsList.innerHTML = "";
    privateDocsRecords.forEach((item) => {
      const row = document.createElement("div");
      row.className = "simple-list-item stacked";
      const createdAt = item.createdAt?.toDate ? item.createdAt.toDate().toLocaleString("it-IT") : "-";
      row.innerHTML = `
        <div>
          <strong>${escapeHTML(item.name || "Documento")}</strong>
          <p class="muted">${escapeHTML(item.note || "-")}</p>
          <p class="muted">Data inserimento: ${escapeHTML(createdAt)}</p>
        </div>
      `;
      const actions = document.createElement("div");
      actions.className = "actions-row";
      if (item.driveWebViewLink) {
        actions.appendChild(createButton("Apri su Drive", () => window.open(item.driveWebViewLink, "_blank")));
      } else if (item.fileDataUrl) {
        actions.appendChild(createButton("Apri allegato", () => window.open(item.fileDataUrl, "_blank")));
      }
      actions.appendChild(createButton("Elimina", () => deletePrivateDocument(item.id)));
      row.appendChild(actions);
      ui.privateDocsList.appendChild(row);
    });
  }
  api.renderPrivateDocsList = renderPrivateDocsList;
  function openDocumentLink(value) {
    const raw = String(value || "").trim();
    if (!raw) return;
    const normalized = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`;
    window.open(normalized, "_blank");
  }
  api.openDocumentLink = openDocumentLink;
  function buildDocumentViewerUrl(rawUrl = "") {
    const url = String(rawUrl || "").trim();
    if (!url) return "";
    if (/docs\.google\.com\/spreadsheets/i.test(url)) return url;
    if (/drive\.google\.com/i.test(url)) {
      return `https://docs.google.com/gview?embedded=1&url=${encodeURIComponent(url)}`;
    }
    return `https://docs.google.com/gview?embedded=1&url=${encodeURIComponent(url)}`;
  }
  api.buildDocumentViewerUrl = buildDocumentViewerUrl;
  function openNotificationDocumentViewer(rawUrl, title = "Documento") {
    const viewerUrl = buildDocumentViewerUrl(rawUrl);
    if (!viewerUrl) return;
    if (ui.notificationDocViewerTitle) ui.notificationDocViewerTitle.textContent = title;
    if (ui.notificationDocViewerFrame) ui.notificationDocViewerFrame.src = viewerUrl;
    ui.notificationDocViewerModal?.classList.remove("hidden");
    ui.notificationDocViewerModal?.setAttribute("aria-hidden", "false");
  }
  api.openNotificationDocumentViewer = openNotificationDocumentViewer;
  function closeNotificationDocumentViewer() {
    if (ui.notificationDocViewerFrame) ui.notificationDocViewerFrame.src = "";
    ui.notificationDocViewerModal?.classList.add("hidden");
    ui.notificationDocViewerModal?.setAttribute("aria-hidden", "true");
  }
  api.closeNotificationDocumentViewer = closeNotificationDocumentViewer;
  Object.assign(global, api);
  global.VargaDocumentsModule = Object.freeze({ ...api });
})(window);
