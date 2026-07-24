(() => {
  "use strict";

  const button = document.getElementById("identity-card-btn");
  const viewer = document.getElementById("identity-card-viewer");
  const viewerBody = document.getElementById("identity-card-viewer-body");
  const closeButton = document.getElementById("identity-card-close-btn");
  const replaceButton = document.getElementById("identity-card-replace-btn");
  if (!button || !viewer || !viewerBody) return;

  let currentUser = null;
  let identityCard = null;
  let unsubscribe = null;

  const isIdentityCard = (item) => {
    const text = `${item?.name || ""} ${item?.note || ""}`.toLocaleLowerCase("it-IT");
    return text.includes("tessera") && (text.includes("riconoscimento") || text.includes("tesserino"));
  };

  const updateButton = () => {
    button.disabled = !currentUser;
    button.classList.toggle("has-card", Boolean(identityCard));
    button.title = identityCard
      ? "Mostra il tesserino a schermo intero"
      : "Inserisci il tesserino di riconoscimento";
    button.setAttribute("aria-label", button.title);
  };

  const closeViewer = () => {
    viewer.classList.add("hidden");
    viewer.setAttribute("aria-hidden", "true");
    viewerBody.innerHTML = "";
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
    const fileId = url.match(/\/d\/([^/?#]+)/)?.[1]
      || new URLSearchParams(url.split("?")[1] || "").get("id")
      || "";
    return fileId ? `https://drive.google.com/file/d/${encodeURIComponent(fileId)}/preview` : url;
  };

  const openViewer = () => {
    if (!currentUser) {
      window.alert("Devi fare login per usare il tesserino di riconoscimento.");
      return;
    }
    if (!identityCard) {
      openUpload();
      return;
    }

    viewerBody.innerHTML = "";
    const fileType = String(identityCard.fileType || "").toLowerCase();
    if (identityCard.fileDataUrl && fileType.startsWith("image/")) {
      const image = document.createElement("img");
      image.src = identityCard.fileDataUrl;
      image.alt = "Tesserino di riconoscimento";
      viewerBody.appendChild(image);
    } else {
      const source = identityCard.fileDataUrl || drivePreviewUrl(identityCard);
      if (!source) {
        openUpload();
        return;
      }
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

  const subscribe = (user) => {
    if (unsubscribe) unsubscribe();
    unsubscribe = null;
    identityCard = null;
    currentUser = user || null;
    updateButton();
    if (!user || !window.firebase?.firestore) return;

    unsubscribe = firebase.firestore()
      .collection("privateDocuments")
      .doc(user.uid)
      .collection("items")
      .orderBy("createdAt", "desc")
      .onSnapshot((snapshot) => {
        identityCard = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })).find(isIdentityCard) || null;
        updateButton();
      }, (error) => {
        console.error("Caricamento tesserino non riuscito:", error);
        identityCard = null;
        updateButton();
      });
  };

  button.addEventListener("click", openViewer);
  closeButton?.addEventListener("click", closeViewer);
  replaceButton?.addEventListener("click", openUpload);
  viewer.addEventListener("click", (event) => {
    if (event.target === viewer) closeViewer();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !viewer.classList.contains("hidden")) closeViewer();
  });

  updateButton();
  if (window.firebase?.auth) firebase.auth().onAuthStateChanged(subscribe);
})();
