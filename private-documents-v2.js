(() => {
  "use strict";

  const form = document.getElementById("private-docs-form");
  const nameInput = document.getElementById("private-docs-name");
  const noteInput = document.getElementById("private-docs-note");
  const fileInput = document.getElementById("private-docs-file");
  const cameraInput = document.getElementById("private-docs-camera");
  const saveButton = document.getElementById("private-docs-save-btn");
  const feedback = document.getElementById("private-docs-feedback");
  if (!form || !window.firebase?.firestore) return;

  const MAX_DATA_URL_LENGTH = 720000;
  const MAX_PDF_BYTES = 500000;

  const showFeedback = (message, type = "") => {
    if (!feedback) return;
    feedback.textContent = message;
    feedback.classList.toggle("private-docs-feedback-success", type === "success");
    feedback.classList.toggle("private-docs-feedback-error", type === "error");
  };

  const readAsDataUrl = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Non riesco a leggere il file selezionato."));
    reader.readAsDataURL(file);
  });

  const loadImage = (dataUrl) => new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("La foto selezionata non è valida."));
    image.src = dataUrl;
  });

  async function optimizeImage(file) {
    const originalDataUrl = await readAsDataUrl(file);
    const image = await loadImage(originalDataUrl);
    const longestSide = Math.max(image.naturalWidth, image.naturalHeight);
    const scale = Math.min(1, 1600 / Math.max(1, longestSide));
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
    const context = canvas.getContext("2d", { alpha: false });
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);

    let quality = 0.86;
    let dataUrl = canvas.toDataURL("image/jpeg", quality);
    while (dataUrl.length > MAX_DATA_URL_LENGTH && quality > 0.42) {
      quality -= 0.08;
      dataUrl = canvas.toDataURL("image/jpeg", quality);
    }
    if (dataUrl.length > MAX_DATA_URL_LENGTH) {
      throw new Error("La foto è ancora troppo grande. Ritagliala e riprova.");
    }
    return {
      dataUrl,
      fileName: String(file.name || "documento.jpg").replace(/\.[^.]+$/, "") + ".jpg",
      fileType: "image/jpeg",
      fileSize: Math.round((dataUrl.length * 3) / 4)
    };
  }

  async function prepareFile(file) {
    if (!file) {
      return { dataUrl: "", fileName: "", fileType: "", fileSize: 0 };
    }
    const fileType = String(file.type || "").toLowerCase();
    if (fileType.startsWith("image/")) return optimizeImage(file);
    if (fileType !== "application/pdf") {
      throw new Error("Puoi allegare soltanto immagini oppure PDF.");
    }
    if (Number(file.size || 0) > MAX_PDF_BYTES) {
      throw new Error("Il PDF è troppo grande. Il limite per i documenti personali è 500 KB.");
    }
    return {
      dataUrl: await readAsDataUrl(file),
      fileName: file.name || "documento.pdf",
      fileType,
      fileSize: Number(file.size || 0)
    };
  }

  const clearOtherFileInput = (source, other) => {
    source?.addEventListener("change", () => {
      if (source.files?.length && other) other.value = "";
      showFeedback(source.files?.[0] ? `File selezionato: ${source.files[0].name}` : "");
    });
  };

  clearOtherFileInput(fileInput, cameraInput);
  clearOtherFileInput(cameraInput, fileInput);

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    event.stopImmediatePropagation();

    const user = firebase.auth().currentUser;
    if (!user) {
      showFeedback("Devi effettuare il login prima di salvare.", "error");
      return;
    }

    const name = String(nameInput?.value || "").trim();
    const note = String(noteInput?.value || "").trim();
    const file = fileInput?.files?.[0] || cameraInput?.files?.[0] || null;
    if (!name) {
      showFeedback("Inserisci la denominazione del documento.", "error");
      nameInput?.focus();
      return;
    }

    saveButton.disabled = true;
    showFeedback(file ? "Preparazione e salvataggio del documento..." : "Salvataggio...");

    try {
      const prepared = await prepareFile(file);
      await firebase.firestore()
        .collection("privateDocuments")
        .doc(user.uid)
        .collection("items")
        .add({
          name,
          note,
          fileName: prepared.fileName,
          fileType: prepared.fileType,
          fileSize: prepared.fileSize,
          fileDataUrl: prepared.dataUrl,
          driveFileId: "",
          driveWebViewLink: "",
          storageMode: "private-firestore",
          ownerUid: user.uid,
          createdAt: firebase.firestore.FieldValue.serverTimestamp(),
          updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        });

      form.reset();
      showFeedback("Documento personale salvato correttamente.", "success");
    } catch (error) {
      console.error("Salvataggio documento personale non riuscito:", error);
      const code = String(error?.code || "");
      if (code.includes("permission-denied")) {
        showFeedback("Salvataggio non autorizzato. Verifica i permessi Firebase dei documenti personali.", "error");
      } else {
        showFeedback(error?.message || "Salvataggio non riuscito. Riprova.", "error");
      }
    } finally {
      saveButton.disabled = false;
    }
  }, true);
})();
