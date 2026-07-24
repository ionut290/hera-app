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
  const pinInput = document.getElementById("fuel-pin-input");
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
    if (pinInput) pinInput.value = "";
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
  const openPinViewer = () => {
    if (!currentUser) return window.alert("Devi fare login per visualizzare il PIN carburante.");
    const value = String(fuelPinDocument?.note || "").trim();
    if (pinValue) pinValue.textContent = value || "Dati non disponibili";
    if (pinInput) pinInput.value = value;
    showPinFeedback(value ? "PIN personale disponibile." : "Inserisci il tuo PIN carburante.");
    pinViewer?.classList.remove("hidden");
    pinViewer?.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";
  };

  const saveFuelPin = async (event) => {
    event.preventDefault();
    const value = String(pinInput?.value || "").trim();
    if (!currentUser || !value) return showPinFeedback("Inserisci un PIN valido.");
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
      showPinFeedback("PIN salvato correttamente.");
    } catch (error) {
      console.error("Salvataggio PIN carburante non riuscito:", error);
      showPinFeedback("Salvataggio non riuscito. Verifica i permessi Firebase.");
    } finally {
      pinSave && (pinSave.disabled = false);
    }
  };

  const copyFuelPin = async () => {
    const value = String(fuelPinDocument?.note || pinInput?.value || "").trim();
    if (!value) return showPinFeedback("Nessun PIN da copiare.");
    try {
      await navigator.clipboard.writeText(value);
      showPinFeedback("PIN copiato.");
    } catch {
      pinInput?.select();
      document.execCommand("copy");
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
          const value = String(fuelPinDocument?.note || "").trim();
          pinValue.textContent = value || "Dati non disponibili";
          pinInput.value = value;
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
