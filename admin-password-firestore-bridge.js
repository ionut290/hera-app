(() => {
  "use strict";

  const VERSION = "1.0.0";
  const REQUEST_COLLECTION = "adminPasswordRequests";
  const REQUEST_TIMEOUT_MS = 45000;
  let busy = false;
  let activeProfile = null;

  function isManager() {
    try {
      return typeof canManageData === "function" && canManageData();
    } catch (_) {
      return false;
    }
  }

  function profiles() {
    if (typeof platformUsers === "undefined" || !Array.isArray(platformUsers)) return [];
    return platformUsers.map((item) => ({
      uid: String(item?.uid || item?.id || "").trim(),
      email: String(item?.email || "").trim().toLowerCase(),
      displayName: String(
        item?.displayName ||
        item?.nomeCompleto ||
        item?.fullName ||
        [item?.nome || item?.firstName, item?.cognome || item?.lastName].filter(Boolean).join(" ") ||
        item?.email ||
        "Utente"
      ).trim()
    })).filter((item) => item.uid && item.email);
  }

  function findProfile(uid) {
    const normalizedUid = String(uid || "").trim();
    return profiles().find((item) => item.uid === normalizedUid) || null;
  }

  function ensureDialog() {
    let dialog = document.getElementById("admin-password-firestore-dialog");
    if (dialog) return dialog;

    dialog = document.createElement("dialog");
    dialog.id = "admin-password-firestore-dialog";
    dialog.className = "biometric-dialog admin-password-dialog";
    dialog.innerHTML = `
      <form id="admin-password-firestore-form" method="dialog">
        <h2>Cambia password utente</h2>
        <p id="admin-password-firestore-user" class="muted"></p>
        <section class="admin-password-section">
          <p>L'amministratore genera una nuova password temporanea e la comunica direttamente all'utente.</p>
          <p><strong>Nessuna email viene inviata.</strong> La password precedente smette di funzionare. Al prossimo accesso l'utente dovrà scegliere una nuova password personale prima di continuare.</p>
          <button id="admin-password-firestore-generate" class="btn btn-primary" type="button">GENERA NUOVA PASSWORD</button>
        </section>
        <section id="admin-password-firestore-result" class="admin-password-result hidden" aria-live="polite">
          <strong>Nuova password temporanea</strong>
          <div class="auth-password-row">
            <input id="admin-password-firestore-value" type="text" readonly autocomplete="off">
            <button id="admin-password-firestore-copy" class="btn" type="button">COPIA</button>
          </div>
          <p class="muted">Comunica questa password all'utente. La password viene restituita cifrata dal server e non viene salvata in chiaro in Firestore.</p>
        </section>
        <p id="admin-password-firestore-feedback" class="muted" role="status" aria-live="polite"></p>
        <div class="actions-row">
          <button id="admin-password-firestore-close" class="btn" type="button">CHIUDI</button>
        </div>
      </form>`;
    document.body.appendChild(dialog);

    dialog.addEventListener("cancel", (event) => {
      event.preventDefault();
      closeDialog();
    });
    dialog.querySelector("#admin-password-firestore-close")?.addEventListener("click", closeDialog);
    dialog.querySelector("#admin-password-firestore-generate")?.addEventListener("click", () => void generateTemporaryPassword());
    dialog.querySelector("#admin-password-firestore-copy")?.addEventListener("click", copyTemporaryPassword);
    return dialog;
  }

  function setFeedback(message) {
    const node = ensureDialog().querySelector("#admin-password-firestore-feedback");
    if (node) node.textContent = String(message || "");
  }

  function setBusy(value) {
    busy = Boolean(value);
    const dialog = ensureDialog();
    dialog.querySelectorAll("button").forEach((button) => {
      if (button.id !== "admin-password-firestore-close") button.disabled = busy;
    });
  }

  function openDialog(profile) {
    if (!isManager() || !profile?.uid || !profile?.email) return;
    const currentEmail = String(firebase.auth()?.currentUser?.email || "").trim().toLowerCase();
    if (profile.email === currentEmail) {
      window.alert("Per sicurezza modifica la password del tuo account amministratore direttamente dal tuo account.");
      return;
    }

    activeProfile = profile;
    const dialog = ensureDialog();
    dialog.querySelector("#admin-password-firestore-user").textContent = `${profile.displayName} • ${profile.email}`;
    dialog.querySelector("#admin-password-firestore-value").value = "";
    dialog.querySelector("#admin-password-firestore-result").classList.add("hidden");
    setFeedback("");
    if (!dialog.open) dialog.showModal();
  }

  function closeDialog() {
    const dialog = document.getElementById("admin-password-firestore-dialog");
    if (dialog?.open) dialog.close();
    activeProfile = null;
    busy = false;
  }

  function requireFirebase() {
    if (!window.firebase || typeof firebase.auth !== "function" || typeof firebase.firestore !== "function") {
      throw new Error("Firebase non disponibile.");
    }
    const currentUser = firebase.auth().currentUser;
    if (!currentUser?.uid || !currentUser?.email) {
      throw new Error("Sessione amministratore non valida. Esci e accedi di nuovo.");
    }
    return currentUser;
  }

  async function createEncryptionKeyPair() {
    if (!window.crypto?.subtle) {
      throw new Error("Il browser non supporta il canale sicuro per la password.");
    }
    const keyPair = await window.crypto.subtle.generateKey(
      {
        name: "RSA-OAEP",
        modulusLength: 2048,
        publicExponent: new Uint8Array([1, 0, 1]),
        hash: "SHA-256"
      },
      true,
      ["encrypt", "decrypt"]
    );
    const publicKeyJwk = await window.crypto.subtle.exportKey("jwk", keyPair.publicKey);
    return { privateKey: keyPair.privateKey, publicKeyJwk };
  }

  function base64ToBytes(value) {
    const binary = window.atob(String(value || ""));
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }

  async function decryptTemporaryPassword(encryptedValue, privateKey) {
    if (!encryptedValue) throw new Error("Il server non ha restituito la password cifrata.");
    const decrypted = await window.crypto.subtle.decrypt(
      { name: "RSA-OAEP" },
      privateKey,
      base64ToBytes(encryptedValue)
    );
    return new TextDecoder().decode(decrypted);
  }

  function waitForCompletion(ref) {
    return new Promise((resolve, reject) => {
      let unsubscribe = null;
      const timeout = window.setTimeout(() => {
        try { unsubscribe?.(); } catch (_) {}
        reject(new Error("Il server sta impiegando troppo tempo. Riprova tra poco."));
      }, REQUEST_TIMEOUT_MS);

      unsubscribe = ref.onSnapshot(
        (snapshot) => {
          if (!snapshot.exists) return;
          const data = snapshot.data() || {};
          if (!["completed", "failed"].includes(data.status)) return;
          window.clearTimeout(timeout);
          try { unsubscribe?.(); } catch (_) {}
          resolve(data);
        },
        (error) => {
          window.clearTimeout(timeout);
          try { unsubscribe?.(); } catch (_) {}
          reject(error);
        }
      );
    });
  }

  async function generateTemporaryPassword() {
    if (!isManager() || !activeProfile || busy) return;
    const confirmed = window.confirm(
      `Cambiare la password di ${activeProfile.displayName}?\n\n` +
      "La password precedente smetterà di funzionare. Verrà generata una password temporanea da comunicare all'utente. Al prossimo accesso dovrà sceglierne una nuova personale."
    );
    if (!confirmed) return;

    setBusy(true);
    setFeedback("Generazione nuova password...");
    let requestRef = null;
    try {
      const currentUser = requireFirebase();
      const { privateKey, publicKeyJwk } = await createEncryptionKeyPair();
      const db = firebase.firestore();
      requestRef = db.collection(REQUEST_COLLECTION).doc();

      await requestRef.set({
        requestId: requestRef.id,
        targetUid: activeProfile.uid,
        targetEmail: activeProfile.email,
        requestedByUid: currentUser.uid,
        requestedByEmail: String(currentUser.email || "").trim().toLowerCase(),
        publicKeyJwk,
        status: "pending",
        createdAt: firebase.firestore.FieldValue.serverTimestamp(),
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      });

      const result = await waitForCompletion(requestRef);
      if (result.status === "failed") {
        const error = new Error(result.errorMessage || "Cambio password non riuscito.");
        error.code = result.errorCode || "internal";
        throw error;
      }
      if (result.targetUid && String(result.targetUid) !== String(activeProfile.uid)) {
        throw new Error("Il server ha aggiornato un account diverso da quello selezionato.");
      }

      const temporaryPassword = await decryptTemporaryPassword(
        result.encryptedTemporaryPassword,
        privateKey
      );
      if (!temporaryPassword) throw new Error("Password temporanea vuota.");

      const dialog = ensureDialog();
      dialog.querySelector("#admin-password-firestore-value").value = temporaryPassword;
      dialog.querySelector("#admin-password-firestore-result").classList.remove("hidden");
      setFeedback("Password cambiata correttamente. Comunicala all'utente: al prossimo accesso dovrà crearne una nuova personale.");

      try { await requestRef.delete(); } catch (_) {}
      requestRef = null;
    } catch (error) {
      console.error("Gestione password via Firestore fallita:", error);
      const code = String(error?.code || "").toLowerCase();
      if (code.includes("permission-denied")) setFeedback("Operazione non autorizzata: accedi con un account amministratore.");
      else if (code.includes("not-found")) setFeedback("Account Firebase dell'utente non trovato.");
      else if (code.includes("unavailable") || code.includes("network")) setFeedback("Connessione non disponibile. Riprova quando sei online.");
      else setFeedback(error?.message || "Cambio password non riuscito.");
      if (requestRef) {
        try { await requestRef.delete(); } catch (_) {}
      }
    } finally {
      setBusy(false);
    }
  }

  async function copyTemporaryPassword() {
    const input = ensureDialog().querySelector("#admin-password-firestore-value");
    const value = String(input?.value || "");
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
      setFeedback("Password temporanea copiata.");
    } catch (_) {
      input?.select();
      setFeedback("Seleziona la password e copiala manualmente.");
    }
  }

  function interceptPasswordButton(event) {
    const button = event.target?.closest?.("[data-admin-manage-password]");
    if (!button) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    event.stopPropagation();

    const profile = findProfile(button.dataset.adminManagePassword);
    if (!profile) {
      window.alert("Profilo utente non trovato. Aggiorna Gestione utenti e riprova.");
      return;
    }
    openDialog(profile);
  }

  function initialize() {
    document.addEventListener("click", interceptPasswordButton, true);
  }

  window.HeraAdminPasswordFirestoreBridge = {
    installed: true,
    version: VERSION,
    open: openDialog
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initialize, { once: true });
  else initialize();
})();
