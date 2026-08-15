(() => {
  "use strict";

  const button = document.getElementById("identity-card-btn");
  const viewer = document.getElementById("identity-card-viewer");
  const viewerBody = document.getElementById("identity-card-viewer-body");
  const closeButton = document.getElementById("identity-card-close-btn");
  const replaceButton = document.getElementById("identity-card-replace-btn");
  const identityShareButton = document.getElementById("identity-card-share-btn");
  const identityStatus = document.getElementById("identity-card-status");
  const businessPreview = document.getElementById("business-card-preview");
  const businessPhoto = document.getElementById("business-card-photo");
  const businessName = document.getElementById("business-card-name");
  const businessRole = document.getElementById("business-card-role");
  const businessCompany = document.getElementById("business-card-company");
  const businessContacts = document.getElementById("business-card-contacts");
  const businessBio = document.getElementById("business-card-bio");
  const businessForm = document.getElementById("business-card-form");
  const businessEditButton = document.getElementById("business-card-edit-btn");
  const businessCancelButton = document.getElementById("business-card-cancel-btn");
  const businessShareButton = document.getElementById("business-card-share-btn");
  const businessSaveButton = document.getElementById("business-card-save-btn");
  const businessFeedback = document.getElementById("business-card-feedback");
  const businessInputs = {
    role: document.getElementById("business-card-role-input"),
    company: document.getElementById("business-card-company-input"),
    phone: document.getElementById("business-card-phone-input"),
    email: document.getElementById("business-card-email-input"),
    address: document.getElementById("business-card-address-input"),
    website: document.getElementById("business-card-website-input"),
    bio: document.getElementById("business-card-bio-input")
  };
  const pinButton = document.getElementById("fuel-pin-btn");
  const pinViewer = document.getElementById("fuel-pin-viewer");
  const pinClose = document.getElementById("fuel-pin-close-btn");
  const pinValue = document.getElementById("fuel-pin-value");
  const pinForm = document.getElementById("fuel-pin-form");
  const q8DriverCodeInput = document.getElementById("fuel-q8-driver-code-input");
  const q8PinInput = document.getElementById("fuel-q8-pin-input");
  const eniliveDriverCodeInput = document.getElementById("fuel-enilive-driver-code-input");
  const enilivePinInput = document.getElementById("fuel-enilive-pin-input");
  const pinCopy = document.getElementById("fuel-pin-copy-btn");
  const pinSave = document.getElementById("fuel-pin-save-btn");
  const pinFeedback = document.getElementById("fuel-pin-feedback");
  if (!button || !viewer || !viewerBody) return;

  let currentUser = null;
  let identityCard = null;
  let fuelPinDocument = null;
  let businessCardDocument = null;
  let unsubscribe = null;
  let bodyOverflowBeforeViewer = "";
  let fullscreenImage = null;
  let imageScale = 1;
  let imageOffset = { x: 0, y: 0 };
  let dragStart = null;

  const normalizedText = (item) => `${item?.name || ""} ${item?.note || ""}`.toLocaleLowerCase("it-IT");
  const isIdentityCard = (item) => {
    const text = normalizedText(item);
    return text.includes("tessera") && (text.includes("riconoscimento") || text.includes("tesserino"));
  };
  const isFuelPin = (item) => normalizedText(item).includes("pin carburante");
  const isBusinessCard = (item) => normalizedText(item).includes("carta da visita");

  const emptyBusinessCard = () => ({
    role: "", company: "", phone: "", email: "", address: "", website: "", bio: ""
  });

  const readBusinessCard = () => {
    const note = String(businessCardDocument?.note || "").trim();
    if (!note) return emptyBusinessCard();
    try {
      return { ...emptyBusinessCard(), ...JSON.parse(note) };
    } catch {
      return emptyBusinessCard();
    }
  };

  const fallbackAvatar = () => {
    const name = String(currentUser?.displayName || currentUser?.email || "U").trim();
    const initials = name.split(/\s+/).slice(0, 2).map((part) => part[0] || "").join("").toUpperCase() || "U";
    return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(
      `<svg xmlns="http://www.w3.org/2000/svg" width="180" height="180"><rect width="180" height="180" rx="90" fill="#dbeafe"/><text x="90" y="108" text-anchor="middle" font-family="Arial,sans-serif" font-size="64" font-weight="700" fill="#1d4ed8">${initials}</text></svg>`
    )}`;
  };

  const setBusinessFeedback = (message) => {
    if (businessFeedback) businessFeedback.textContent = message;
  };

  const addBusinessContact = (icon, value) => {
    if (!businessContacts || !value) return;
    const item = document.createElement("span");
    item.className = "business-card-contact";
    item.textContent = `${icon} ${value}`;
    businessContacts.appendChild(item);
  };

  const renderBusinessCard = () => {
    const data = readBusinessCard();
    const name = String(currentUser?.displayName || currentUser?.email || "Utente");
    if (businessPhoto) {
      businessPhoto.src = currentUser?.photoURL || fallbackAvatar();
      businessPhoto.onerror = () => { businessPhoto.src = fallbackAvatar(); };
    }
    if (businessName) businessName.textContent = name;
    if (businessRole) businessRole.textContent = data.role || "Aggiungi la tua qualifica";
    if (businessCompany) businessCompany.textContent = data.company || "Aggiungi l’azienda";
    if (businessContacts) {
      businessContacts.innerHTML = "";
      addBusinessContact("☎", data.phone);
      addBusinessContact("✉", data.email || currentUser?.email || "");
      addBusinessContact("⌖", data.address);
      addBusinessContact("🌐", data.website);
    }
    if (businessBio) {
      businessBio.textContent = data.bio || "";
      businessBio.hidden = !data.bio;
    }
    Object.entries(businessInputs).forEach(([key, input]) => {
      if (input) input.value = data[key] || (key === "email" ? currentUser?.email || "" : "");
    });
    if (businessEditButton) businessEditButton.textContent = businessCardDocument ? "Modifica" : "Compila";
  };

  const toggleBusinessForm = (visible) => {
    businessForm?.classList.toggle("hidden", !visible);
    if (visible) {
      renderBusinessCard();
      businessInputs.role?.focus();
    }
    setBusinessFeedback("");
  };

  const updateButtons = () => {
    button.disabled = !currentUser;
    pinButton && (pinButton.disabled = !currentUser);
    button.classList.toggle("has-card", Boolean(identityCard));
    pinButton?.classList.toggle("has-pin", Boolean(fuelPinDocument?.note));
    button.title = identityCard ? "Mostra il tesserino a schermo intero" : "Inserisci il tesserino di riconoscimento";
    button.setAttribute("aria-label", button.title);
  };

  const closeViewer = () => {
    closeFullscreenImage(false);
    viewer.classList.add("hidden");
    viewer.setAttribute("aria-hidden", "true");
    viewerBody.innerHTML = "";
    toggleBusinessForm(false);
    document.body.style.overflow = bodyOverflowBeforeViewer;
  };

  const applyImageTransform = () => {
    const image = fullscreenImage?.querySelector("img");
    if (image) image.style.transform = `translate3d(${imageOffset.x}px, ${imageOffset.y}px, 0) scale(${imageScale})`;
    const status = fullscreenImage?.querySelector(".identity-fullscreen-zoom-status");
    if (status) status.textContent = `${Math.round(imageScale * 100)}%`;
  };

  function closeFullscreenImage(useHistory = true) {
    if (!fullscreenImage) return;
    fullscreenImage.remove();
    fullscreenImage = null;
    viewer.classList.remove("identity-image-open");
    if (useHistory && history.state?.identityCardFullscreen) history.back();
  }

  const openFullscreenImage = (source, alt) => {
    if (fullscreenImage) return;
    imageScale = 1; imageOffset = { x: 0, y: 0 };
    fullscreenImage = document.createElement("section");
    fullscreenImage.className = "identity-image-fullscreen";
    fullscreenImage.setAttribute("role", "dialog");
    fullscreenImage.setAttribute("aria-modal", "true");
    fullscreenImage.setAttribute("aria-label", "Tessera a schermo intero");
    fullscreenImage.innerHTML = `<header><span class="identity-fullscreen-zoom-status">100%</span><div><button type="button" data-zoom="out" aria-label="Riduci">−</button><button type="button" data-zoom="reset" aria-label="Ripristina zoom">1:1</button><button type="button" data-zoom="in" aria-label="Ingrandisci">＋</button><button type="button" data-close aria-label="Chiudi schermo intero">✕</button></div></header><div class="identity-image-pan"><img alt=""></div>`;
    const image = fullscreenImage.querySelector("img"); image.src = source; image.alt = alt;
    const pan = fullscreenImage.querySelector(".identity-image-pan");
    fullscreenImage.querySelector("[data-close]").addEventListener("click", () => closeFullscreenImage());
    fullscreenImage.addEventListener("click", (event) => {
      const action = event.target.closest("[data-zoom]")?.dataset.zoom;
      if (!action) return;
      if (action === "reset") { imageScale = 1; imageOffset = { x: 0, y: 0 }; }
      else imageScale = Math.max(1, Math.min(5, imageScale + (action === "in" ? .5 : -.5)));
      applyImageTransform();
    });
    pan.addEventListener("wheel", (event) => { event.preventDefault(); imageScale = Math.max(1, Math.min(5, imageScale + (event.deltaY < 0 ? .25 : -.25))); applyImageTransform(); }, { passive: false });
    pan.addEventListener("pointerdown", (event) => { dragStart = { x: event.clientX - imageOffset.x, y: event.clientY - imageOffset.y }; pan.setPointerCapture(event.pointerId); });
    pan.addEventListener("pointermove", (event) => { if (!dragStart || imageScale === 1) return; imageOffset = { x: event.clientX - dragStart.x, y: event.clientY - dragStart.y }; applyImageTransform(); });
    pan.addEventListener("pointerup", () => { dragStart = null; });
    viewer.appendChild(fullscreenImage); viewer.classList.add("identity-image-open");
    history.pushState({ ...(history.state || {}), identityCardFullscreen: true }, "");
    fullscreenImage.querySelector("[data-close]").focus();
  };

  const closePinViewer = () => {
    pinViewer?.classList.add("hidden");
    pinViewer?.setAttribute("aria-hidden", "true");
    [q8DriverCodeInput, q8PinInput, eniliveDriverCodeInput, enilivePinInput].forEach((input) => {
      if (input) input.value = "";
    });
    document.body.style.overflow = "";
  };

  const openUpload = () => {
    closeViewer();
    window.location.hash = "documenti";
    window.dispatchEvent(new HashChangeEvent("hashchange"));
    window.setTimeout(() => {
      document.getElementById("private-docs-preset-tessera-btn")?.click();
      document.getElementById("private-docs-form")?.scrollIntoView({ behavior: "smooth", block: "start" });
      document.getElementById("private-docs-file")?.focus();
    }, 100);
  };

  const drivePreviewUrl = (item) => {
    const storedId = String(item?.driveFileId || "").trim();
    if (storedId) return `https://drive.google.com/file/d/${encodeURIComponent(storedId)}/preview`;
    const url = String(item?.driveWebViewLink || "").trim();
    const fileId = url.match(/\/d\/([^/?#]+)/)?.[1] || new URLSearchParams(url.split("?")[1] || "").get("id") || "";
    return fileId ? `https://drive.google.com/file/d/${encodeURIComponent(fileId)}/preview` : url;
  };

  const renderIdentityCard = () => {
    viewerBody.innerHTML = "";
    identityStatus?.classList.toggle("ready", Boolean(identityCard));
    if (identityStatus) identityStatus.textContent = identityCard ? "Disponibile" : "Non inserita";
    if (identityShareButton) identityShareButton.disabled = !identityCard;
    if (!identityCard) {
      const placeholder = document.createElement("button");
      placeholder.className = "btn";
      placeholder.type = "button";
      placeholder.textContent = "＋ Inserisci la tessera di riconoscimento";
      placeholder.addEventListener("click", openUpload);
      viewerBody.appendChild(placeholder);
      return;
    }
    const fileType = String(identityCard.fileType || "").toLowerCase();
    if (identityCard.fileDataUrl && fileType.startsWith("image/")) {
      const image = document.createElement("img");
      image.src = identityCard.fileDataUrl;
      image.alt = "Tessera di riconoscimento";
      const enlarge = document.createElement("button");
      enlarge.type = "button";
      enlarge.className = "identity-card-enlarge";
      enlarge.setAttribute("aria-label", "Apri la tessera a schermo intero con zoom");
      enlarge.appendChild(image);
      enlarge.addEventListener("click", () => openFullscreenImage(image.src, image.alt));
      viewerBody.appendChild(enlarge);
      return;
    }
    const source = identityCard.fileDataUrl || drivePreviewUrl(identityCard);
    if (!source) return;
    const frame = document.createElement("iframe");
    frame.src = source;
    frame.title = "Tessera di riconoscimento";
    frame.allow = "fullscreen";
    viewerBody.appendChild(frame);
  };

  const openViewer = () => {
    if (!currentUser) return window.alert("Devi fare login per usare il tesserino di riconoscimento.");
    renderIdentityCard();
    renderBusinessCard();
    viewer.classList.remove("hidden");
    viewer.setAttribute("aria-hidden", "false");
    bodyOverflowBeforeViewer = document.body.style.overflow;
    document.body.style.overflow = "hidden";
  };

  const saveBusinessCard = async (event) => {
    event.preventDefault();
    if (!currentUser || !window.firebase?.firestore) return;
    const data = Object.fromEntries(
      Object.entries(businessInputs).map(([key, input]) => [key, String(input?.value || "").trim()])
    );
    if (!data.role && !data.company && !data.phone) {
      return setBusinessFeedback("Inserisci almeno qualifica, azienda o telefono.");
    }
    businessSaveButton && (businessSaveButton.disabled = true);
    setBusinessFeedback("Salvataggio in corso...");
    try {
      const documentData = {
        name: "Carta da visita",
        note: JSON.stringify(data),
        fileName: "", fileType: "", fileSize: 0, fileDataUrl: "",
        driveFileId: "", driveWebViewLink: "",
        storageMode: "private-firestore",
        ownerUid: currentUser.uid,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      };
      const items = firebase.firestore().collection("privateDocuments").doc(currentUser.uid).collection("items");
      if (businessCardDocument?.id) {
        await items.doc(businessCardDocument.id).set(documentData, { merge: true });
      } else {
        const created = await items.add({ ...documentData, createdAt: firebase.firestore.FieldValue.serverTimestamp() });
        businessCardDocument = { id: created.id, ...documentData };
      }
      renderBusinessCard();
      toggleBusinessForm(false);
    } catch (error) {
      console.error("Salvataggio carta da visita non riuscito:", error);
      setBusinessFeedback("Salvataggio non riuscito. Verifica la connessione e riprova.");
    } finally {
      businessSaveButton && (businessSaveButton.disabled = false);
    }
  };

  const shareFileOrText = async ({ file, title, text, url }) => {
    if (navigator.share) {
      const payload = { title, text };
      if (file && navigator.canShare?.({ files: [file] })) payload.files = [file];
      else if (url) payload.url = url;
      await navigator.share(payload);
      return true;
    }
    const value = [text, url].filter(Boolean).join("\n");
    await navigator.clipboard.writeText(value);
    window.alert("Condivisione non disponibile: informazioni copiate.");
    return true;
  };

  const shareIdentityCard = async () => {
    if (!identityCard) return;
    identityShareButton && (identityShareButton.disabled = true);
    try {
      let file = null;
      if (identityCard.fileDataUrl) {
        const blob = await fetch(identityCard.fileDataUrl).then((response) => response.blob());
        const extension = blob.type.includes("pdf") ? "pdf" : (blob.type.split("/")[1] || "jpg");
        file = new File([blob], `tessera-riconoscimento.${extension}`, { type: blob.type });
      }
      const url = file ? "" : String(identityCard.driveWebViewLink || drivePreviewUrl(identityCard) || "");
      await shareFileOrText({
        file,
        title: "Tessera di riconoscimento",
        text: `Tessera di riconoscimento di ${currentUser?.displayName || currentUser?.email || "utente"}`,
        url
      });
    } catch (error) {
      if (error?.name !== "AbortError") window.alert("Non è stato possibile condividere la tessera.");
    } finally {
      identityShareButton && (identityShareButton.disabled = false);
    }
  };

  const shareBusinessCard = async () => {
    const data = readBusinessCard();
    if (!businessCardDocument) {
      toggleBusinessForm(true);
      return setBusinessFeedback("Compila e salva prima la carta da visita.");
    }
    businessShareButton && (businessShareButton.disabled = true);
    try {
      let file = null;
      if (!window.html2canvas && window.HeraHeavyLibs?.ensure) {
        try { await window.HeraHeavyLibs.ensure("html2canvas"); } catch (_) {}
      }
      if (window.html2canvas && businessPreview) {
        const canvas = await window.html2canvas(businessPreview, {
          scale: Math.min(3, window.devicePixelRatio || 2),
          backgroundColor: "#ffffff",
          useCORS: true
        });
        const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png", .96));
        if (blob) file = new File([blob], "carta-da-visita.png", { type: "image/png" });
      }
      const text = [
        currentUser?.displayName || currentUser?.email || "Utente",
        data.role,
        data.company,
        data.phone ? `Tel: ${data.phone}` : "",
        data.email ? `Email: ${data.email}` : "",
        data.website
      ].filter(Boolean).join("\n");
      await shareFileOrText({ file, title: "Carta da visita", text });
    } catch (error) {
      if (error?.name !== "AbortError") window.alert("Non è stato possibile condividere la carta da visita.");
    } finally {
      businessShareButton && (businessShareButton.disabled = false);
    }
  };

  const showPinFeedback = (message) => { if (pinFeedback) pinFeedback.textContent = message; };
  const setFuelPinDisplayMode = (hasSavedCredentials) => {
    pinViewer?.classList.toggle("fuel-pin-summary-only", hasSavedCredentials);
    if (pinForm) pinForm.hidden = hasSavedCredentials;
    if (pinFeedback) pinFeedback.hidden = hasSavedCredentials;
  };
  const emptyFuelCredentials = () => ({ q8DriverCode: "", q8Pin: "", eniliveDriverCode: "", enilivePin: "" });
  const readFuelCredentials = () => {
    const note = String(fuelPinDocument?.note || "").trim();
    if (!note) return emptyFuelCredentials();
    try {
      const parsed = JSON.parse(note);
      return { ...emptyFuelCredentials(), ...parsed };
    } catch {
      return { ...emptyFuelCredentials(), q8Pin: note };
    }
  };
  const renderFuelCredentials = (credentials) => {
    if (!pinValue) return;
    const providers = [
      ["Q8", credentials.q8DriverCode, credentials.q8Pin],
      ["Enilive", credentials.eniliveDriverCode, credentials.enilivePin]
    ];
    if (!providers.some(([, code, pin]) => code || pin)) {
      pinValue.textContent = "Dati non disponibili";
      return;
    }
    pinValue.innerHTML = providers.map(([name, code, pin]) =>
      `<section class="fuel-credential-provider"><strong>${name}</strong><span>Codice autista: ${code || "—"}</span><span>PIN: ${pin || "—"}</span></section>`
    ).join("");
  };
  const fillFuelInputs = (credentials) => {
    if (q8DriverCodeInput) q8DriverCodeInput.value = credentials.q8DriverCode || "";
    if (q8PinInput) q8PinInput.value = credentials.q8Pin || "";
    if (eniliveDriverCodeInput) eniliveDriverCodeInput.value = credentials.eniliveDriverCode || "";
    if (enilivePinInput) enilivePinInput.value = credentials.enilivePin || "";
  };
  const openPinViewer = () => {
    if (!currentUser) return window.alert("Devi fare login per visualizzare il PIN carburante.");
    const credentials = readFuelCredentials();
    renderFuelCredentials(credentials);
    fillFuelInputs(credentials);
    const available = Object.values(credentials).some(Boolean);
    setFuelPinDisplayMode(available);
    showPinFeedback(available ? "Dati carburante personali disponibili." : "Inserisci i dati Q8 e/o Enilive.");
    pinViewer?.classList.remove("hidden");
    pinViewer?.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";
  };

  const saveFuelPin = async (event) => {
    event.preventDefault();
    const credentials = {
      q8DriverCode: String(q8DriverCodeInput?.value || "").trim(),
      q8Pin: String(q8PinInput?.value || "").trim(),
      eniliveDriverCode: String(eniliveDriverCodeInput?.value || "").trim(),
      enilivePin: String(enilivePinInput?.value || "").trim()
    };
    const hasCompleteProvider = (credentials.q8DriverCode && credentials.q8Pin)
      || (credentials.eniliveDriverCode && credentials.enilivePin);
    if (!currentUser || !hasCompleteProvider) {
      return showPinFeedback("Completa codice autista e PIN per Q8 e/o Enilive.");
    }
    const value = JSON.stringify(credentials);
    pinSave && (pinSave.disabled = true);
    showPinFeedback("Salvataggio...");
    try {
      const data = {
        name: "PIN carburante",
        note: value,
        fileName: "", fileType: "", fileSize: 0, fileDataUrl: "",
        driveFileId: "", driveWebViewLink: "",
        storageMode: "private-firestore",
        ownerUid: currentUser.uid,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      };
      const items = firebase.firestore().collection("privateDocuments").doc(currentUser.uid).collection("items");
      if (fuelPinDocument?.id) await items.doc(fuelPinDocument.id).set(data, { merge: true });
      else await items.add({ ...data, createdAt: firebase.firestore.FieldValue.serverTimestamp() });
      renderFuelCredentials(credentials);
      setFuelPinDisplayMode(true);
      showPinFeedback("Dati carburante salvati correttamente.");
    } catch (error) {
      console.error("Salvataggio PIN carburante non riuscito:", error);
      showPinFeedback("Salvataggio non riuscito. Verifica i permessi Firebase.");
    } finally {
      pinSave && (pinSave.disabled = false);
    }
  };

  const copyFuelPin = async () => {
    const credentials = readFuelCredentials();
    if (!fuelPinDocument?.note) {
      credentials.q8DriverCode = String(q8DriverCodeInput?.value || "").trim();
      credentials.q8Pin = String(q8PinInput?.value || "").trim();
      credentials.eniliveDriverCode = String(eniliveDriverCodeInput?.value || "").trim();
      credentials.enilivePin = String(enilivePinInput?.value || "").trim();
    }
    const rows = [];
    if (credentials.q8DriverCode || credentials.q8Pin) rows.push(`Q8 — Codice autista: ${credentials.q8DriverCode || "—"} — PIN: ${credentials.q8Pin || "—"}`);
    if (credentials.eniliveDriverCode || credentials.enilivePin) rows.push(`Enilive — Codice autista: ${credentials.eniliveDriverCode || "—"} — PIN: ${credentials.enilivePin || "—"}`);
    const value = rows.join("\n");
    if (!value) return showPinFeedback("Nessun dato carburante da copiare.");
    try {
      await navigator.clipboard.writeText(value);
      showPinFeedback("PIN copiato.");
    } catch {
      const fallback = document.createElement("textarea");
      fallback.value = value;
      document.body.appendChild(fallback);
      fallback.select();
      document.execCommand("copy");
      fallback.remove();
      showPinFeedback("PIN copiato.");
    }
  };

  const subscribe = (user) => {
    if (unsubscribe) unsubscribe();
    unsubscribe = null;
    identityCard = null;
    fuelPinDocument = null;
    businessCardDocument = null;
    currentUser = user || null;
    updateButtons();
    if (!user || !window.firebase?.firestore) return;
    unsubscribe = firebase.firestore().collection("privateDocuments").doc(user.uid).collection("items")
      .orderBy("createdAt", "desc")
      .onSnapshot((snapshot) => {
        const items = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        identityCard = items.find(isIdentityCard) || null;
        fuelPinDocument = items.find(isFuelPin) || null;
        businessCardDocument = items.find(isBusinessCard) || null;
        if (!viewer.classList.contains("hidden")) {
          renderIdentityCard();
          renderBusinessCard();
        }
        if (!pinViewer?.classList.contains("hidden")) {
          const credentials = readFuelCredentials();
          renderFuelCredentials(credentials);
          fillFuelInputs(credentials);
          setFuelPinDisplayMode(Object.values(credentials).some(Boolean));
        }
        updateButtons();
      }, (error) => {
        console.error("Caricamento dati personali non riuscito:", error);
        identityCard = null;
        fuelPinDocument = null;
        businessCardDocument = null;
        updateButtons();
      });
  };

  button.addEventListener("click", openViewer);
  closeButton?.addEventListener("click", closeViewer);
  replaceButton?.addEventListener("click", openUpload);
  identityShareButton?.addEventListener("click", shareIdentityCard);
  businessEditButton?.addEventListener("click", () => toggleBusinessForm(true));
  businessCancelButton?.addEventListener("click", () => toggleBusinessForm(false));
  businessForm?.addEventListener("submit", saveBusinessCard);
  businessShareButton?.addEventListener("click", shareBusinessCard);
  pinButton?.addEventListener("click", openPinViewer);
  pinClose?.addEventListener("click", closePinViewer);
  pinForm?.addEventListener("submit", saveFuelPin);
  pinCopy?.addEventListener("click", copyFuelPin);
  viewer.addEventListener("click", (event) => { if (event.target === viewer) closeViewer(); });
  pinViewer?.addEventListener("click", (event) => { if (event.target === pinViewer) closePinViewer(); });
  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    if (fullscreenImage) closeFullscreenImage();
    else if (!viewer.classList.contains("hidden")) closeViewer();
    if (!pinViewer?.classList.contains("hidden")) closePinViewer();
  });
  window.addEventListener("popstate", () => {
    if (fullscreenImage) closeFullscreenImage(false);
  });
  window.addEventListener("orientationchange", () => {
    imageOffset = { x: 0, y: 0 };
    applyImageTransform();
  });

  updateButtons();
  if (window.firebase?.auth) firebase.auth().onAuthStateChanged(subscribe);
})();
