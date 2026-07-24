(() => {
  "use strict";

  const button = document.getElementById("identity-card-btn");
  const viewer = document.getElementById("identity-card-viewer");
  const viewerBody = document.getElementById("identity-card-viewer-body");
  const closeButton = document.getElementById("identity-card-close-btn");
  const replaceButton = document.getElementById("identity-card-replace-btn");
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
  let unsubscribe = null;

  const normalizedText = (item) => `${item?.name || ""} ${item?.note || ""}`.toLocaleLowerCase("it-IT");
  const isIdentityCard = (item) => {
    const text = normalizedText(item);
    return text.includes("tessera") && (text.includes("riconoscimento") || text.includes("tesserino"));
  };
  const isFuelPin = (item) => normalizedText(item).includes("pin carburante");

  const updateButtons = () => {
    button.disabled = !currentUser;
    pinButton && (pinButton.disabled = !currentUser);
    button.classList.toggle("has-card", Boolean(identityCard));
    pinButton?.classList.toggle("has-pin", Boolean(fuelPinDocument?.note));
    button.title = identityCard ? "Mostra il tesserino a schermo intero" : "Inserisci il tesserino di riconoscimento";
    button.setAttribute("aria-label", button.title);
  };

  const closeViewer = () => {
    viewer.classList.add("hidden");
    viewer.setAttribute("aria-hidden", "true");
    viewerBody.innerHTML = "";
    document.body.style.overflow = "";
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

  const openViewer = () => {
    if (!currentUser) return window.alert("Devi fare login per usare il tesserino di riconoscimento.");
    if (!identityCard) return openUpload();
    viewerBody.innerHTML = "";
    const fileType = String(identityCard.fileType || "").toLowerCase();
    if (identityCard.fileDataUrl && fileType.startsWith("image/")) {
      const image = document.createElement("img");
      image.src = identityCard.fileDataUrl;
      image.alt = "Tesserino di riconoscimento";
      viewerBody.appendChild(image);
    } else {
      const source = identityCard.fileDataUrl || drivePreviewUrl(identityCard);
      if (!source) return openUpload();
      const frame = document.createElement("iframe");
      frame.src = source;
      frame.title = "Tesserino di riconoscimento";
      frame.allow = "fullscreen";
      viewerBody.appendChild(frame);
    }
    viewer.classList.remove("hidden");
    viewer.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";
  };

  const showPinFeedback = (message) => { if (pinFeedback) pinFeedback.textContent = message; };
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
    currentUser = user || null;
    updateButtons();
    if (!user || !window.firebase?.firestore) return;
    unsubscribe = firebase.firestore().collection("privateDocuments").doc(user.uid).collection("items")
      .orderBy("createdAt", "desc")
      .onSnapshot((snapshot) => {
        const items = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        identityCard = items.find(isIdentityCard) || null;
        fuelPinDocument = items.find(isFuelPin) || null;
        if (!pinViewer?.classList.contains("hidden")) {
          const credentials = readFuelCredentials();
          renderFuelCredentials(credentials);
          fillFuelInputs(credentials);
        }
        updateButtons();
      }, (error) => {
        console.error("Caricamento dati personali non riuscito:", error);
        identityCard = null;
        fuelPinDocument = null;
        updateButtons();
      });
  };

  button.addEventListener("click", openViewer);
  closeButton?.addEventListener("click", closeViewer);
  replaceButton?.addEventListener("click", openUpload);
  pinButton?.addEventListener("click", openPinViewer);
  pinClose?.addEventListener("click", closePinViewer);
  pinForm?.addEventListener("submit", saveFuelPin);
  pinCopy?.addEventListener("click", copyFuelPin);
  viewer.addEventListener("click", (event) => { if (event.target === viewer) closeViewer(); });
  pinViewer?.addEventListener("click", (event) => { if (event.target === pinViewer) closePinViewer(); });
  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    if (!viewer.classList.contains("hidden")) closeViewer();
    if (!pinViewer?.classList.contains("hidden")) closePinViewer();
  });

  updateButtons();
  if (window.firebase?.auth) firebase.auth().onAuthStateChanged(subscribe);
})();
