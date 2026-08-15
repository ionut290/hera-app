(() => {
  "use strict";

  const VERSION = "1.0.0";
  const REQUEST_TIMEOUT_MS = 45000;
  const SEARCH_LIST_IDS = [
    "pending-users-list",
    "admin-users-list",
    "user-ban-list",
    "user-permissions-list"
  ];
  let busy = false;
  let activeProfile = null;
  let observer = null;
  let panelActivated = false;
  let profilesCache = [];
  let profilesPromise = null;

  function isManager() {
    try {
      return typeof canManageData === "function" && canManageData();
    } catch (_) {
      return false;
    }
  }

  function normalize(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  }

  function mapProfile(item) {
    const displayName = String(
      item?.displayName ||
      item?.nomeCompleto ||
      item?.fullName ||
      [item?.nome || item?.firstName, item?.cognome || item?.lastName].filter(Boolean).join(" ") ||
      item?.email ||
      "Utente"
    ).trim();
    return {
      uid: String(item?.uid || item?.id || "").trim(),
      email: String(item?.email || "").trim().toLowerCase(),
      displayName,
      searchText: normalize(`${displayName} ${item?.email || ""} ${item?.uid || item?.id || ""}`)
    };
  }

  function profilesFromApp() {
    try {
      if (typeof platformUsers === "undefined" || !Array.isArray(platformUsers)) return [];
      return platformUsers.map(mapProfile).filter((item) => item.uid && item.email);
    } catch (_) {
      return [];
    }
  }

  async function loadProfiles() {
    const appProfiles = profilesFromApp();
    if (appProfiles.length) {
      profilesCache = appProfiles;
      return appProfiles;
    }
    if (profilesCache.length) return profilesCache;
    if (profilesPromise) return profilesPromise;
    if (!window.firebase || typeof firebase.firestore !== "function") return [];

    profilesPromise = firebase.firestore().collection("platformUsers").get()
      .then((snapshot) => snapshot.docs
        .map((doc) => mapProfile({ id: doc.id, uid: doc.id, ...(doc.data() || {}) }))
        .filter((item) => item.uid && item.email))
      .then((items) => {
        profilesCache = items;
        return items;
      })
      .catch((error) => {
        console.warn("Caricamento profili per Gestione utenti non riuscito:", error);
        return [];
      })
      .finally(() => {
        profilesPromise = null;
      });
    return profilesPromise;
  }

  function profileForCard(card, profiles) {
    if (!card) return null;
    const dataText = normalize(Array.from(card.querySelectorAll(
      "[data-user-id],[data-uid],[data-user],[data-ban-user],[data-user-ban],[data-user-permission]"
    )).map((node) => Object.values(node.dataset || {}).join(" ")).join(" "));
    const text = normalize(card.textContent || "");
    const combined = `${text} ${dataText}`.trim();

    const byEmail = profiles.find((profile) => profile.email && combined.includes(normalize(profile.email)));
    if (byEmail) return byEmail;
    const byUid = profiles.find((profile) => profile.uid && dataText.includes(normalize(profile.uid)));
    if (byUid) return byUid;

    return [...profiles]
      .sort((a, b) => b.displayName.length - a.displayName.length)
      .find((profile) => {
        const name = normalize(profile.displayName);
        return name.length >= 3 && combined.includes(name);
      }) || null;
  }

  function ensureSearch() {
    if (!isManager()) return null;
    const panel = document.getElementById("panel-utenti");
    if (!panel) return null;
    let input = panel.querySelector("#user-management-search-input");
    if (input) return input;

    const toolbar = document.createElement("div");
    toolbar.className = "personale-search-toolbar user-management-search-toolbar";
    toolbar.innerHTML = `
      <div class="personale-search-input-wrap">
        <span aria-hidden="true">🔍</span>
        <input id="user-management-search-input" type="text" inputmode="search" enterkeyhint="search" placeholder="Cerca utente per nome o email" autocomplete="off" autocapitalize="none" spellcheck="false">
      </div>
      <button id="user-management-search-clear" class="btn btn-small" type="button">Azzera</button>
      <p id="user-management-search-feedback" class="muted" role="status" aria-live="polite"></p>`;
    panel.insertBefore(toolbar, panel.firstChild);

    input = toolbar.querySelector("#user-management-search-input");
    input?.addEventListener("input", applySearch);
    toolbar.querySelector("#user-management-search-clear")?.addEventListener("click", () => {
      if (!input) return;
      input.value = "";
      applySearch();
      input.focus();
    });
    return input;
  }

  function applySearch() {
    const input = document.getElementById("user-management-search-input");
    const feedback = document.getElementById("user-management-search-feedback");
    const query = normalize(input?.value || "");
    let visible = 0;

    SEARCH_LIST_IDS.forEach((id) => {
      const list = document.getElementById(id);
      if (!list) return;
      Array.from(list.children).forEach((item) => {
        if (!(item instanceof HTMLElement)) return;
        const matches = !query || normalize(item.textContent || "").includes(query);
        item.hidden = !matches;
        if (matches) visible += 1;
      });
    });

    if (!feedback) return;
    if (!query) feedback.textContent = "";
    else if (visible) feedback.textContent = `${visible} risultato${visible === 1 ? "" : "i"} trovato${visible === 1 ? "" : "i"}.`;
    else feedback.textContent = "Nessun utente trovato.";
  }

  function ensureDialog() {
    let dialog = document.getElementById("admin-password-private-dialog");
    if (dialog) return dialog;

    dialog = document.createElement("dialog");
    dialog.id = "admin-password-private-dialog";
    dialog.className = "biometric-dialog admin-password-dialog";
    dialog.innerHTML = `
      <form id="admin-password-private-form">
        <h2>Cambia password utente</h2>
        <p id="admin-password-private-user" class="muted"></p>
        <section class="admin-password-section">
          <p>L'amministratore genera una nuova password temporanea e la comunica direttamente all'utente.</p>
          <p><strong>Nessuna email viene inviata.</strong> La password precedente smette di funzionare. Al prossimo accesso l'utente dovrà scegliere una nuova password personale prima di continuare.</p>
          <button id="admin-password-private-generate" class="btn btn-primary" type="button">GENERA NUOVA PASSWORD</button>
        </section>
        <section id="admin-password-private-result" class="admin-password-result hidden" aria-live="polite">
          <strong>Nuova password temporanea</strong>
          <div class="auth-password-row">
            <input id="admin-password-private-value" type="text" readonly autocomplete="off">
            <button id="admin-password-private-copy" class="btn" type="button">COPIA</button>
          </div>
          <p class="muted">Comunica questa password all'utente. La password viene restituita cifrata dal server e non viene salvata in chiaro in Firestore.</p>
        </section>
        <p id="admin-password-private-feedback" class="muted" role="status" aria-live="polite"></p>
        <div class="actions-row">
          <button id="admin-password-private-close" class="btn" type="button">CHIUDI</button>
        </div>
      </form>`;
    document.body.appendChild(dialog);

    dialog.addEventListener("cancel", (event) => {
      event.preventDefault();
      closeDialog();
    });
    dialog.querySelector("#admin-password-private-close")?.addEventListener("click", closeDialog);
    dialog.querySelector("#admin-password-private-generate")?.addEventListener("click", () => void generateTemporaryPassword());
    dialog.querySelector("#admin-password-private-copy")?.addEventListener("click", copyTemporaryPassword);
    return dialog;
  }

  function setFeedback(message) {
    const node = ensureDialog().querySelector("#admin-password-private-feedback");
    if (node) node.textContent = String(message || "");
  }

  function setBusy(value) {
    busy = Boolean(value);
    const dialog = ensureDialog();
    dialog.querySelectorAll("button").forEach((button) => {
      if (button.id !== "admin-password-private-close") button.disabled = busy;
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
    dialog.querySelector("#admin-password-private-user").textContent = `${profile.displayName} • ${profile.email}`;
    dialog.querySelector("#admin-password-private-value").value = "";
    dialog.querySelector("#admin-password-private-result").classList.add("hidden");
    setFeedback("");
    if (!dialog.open) dialog.showModal();
  }

  function closeDialog() {
    const dialog = document.getElementById("admin-password-private-dialog");
    if (dialog?.open) dialog.close();
    activeProfile = null;
    busy = false;
  }

  function requireFirebaseUser() {
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
      const currentUser = requireFirebaseUser();
      const { privateKey, publicKeyJwk } = await createEncryptionKeyPair();
      const db = firebase.firestore();
      requestRef = db
        .collection("privateDocuments")
        .doc(currentUser.uid)
        .collection("adminPasswordRequests")
        .doc();

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
      dialog.querySelector("#admin-password-private-value").value = temporaryPassword;
      dialog.querySelector("#admin-password-private-result").classList.remove("hidden");
      setFeedback("Password cambiata correttamente. Comunicala all'utente: al prossimo accesso dovrà crearne una nuova personale.");

      try { await requestRef.delete(); } catch (_) {}
      requestRef = null;
    } catch (error) {
      console.error("Gestione password tramite richiesta privata Firestore fallita:", error);
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
    const input = ensureDialog().querySelector("#admin-password-private-value");
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

  function createPasswordButton(profile) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn";
    button.dataset.adminPrivatePassword = profile.uid;
    button.textContent = "CAMBIA PASSWORD";
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openDialog(profile);
    });
    return button;
  }

  async function enhanceUserCards() {
    if (!isManager()) return;
    const profiles = await loadProfiles();
    if (!profiles.length) return;
    const currentEmail = String(firebase.auth()?.currentUser?.email || "").trim().toLowerCase();

    ["user-ban-list", "pending-users-list"].forEach((listId) => {
      const list = document.getElementById(listId);
      if (!list) return;
      Array.from(list.children).forEach((card) => {
        if (!(card instanceof HTMLElement)) return;
        if (card.querySelector("[data-admin-private-password]")) return;
        const profile = profileForCard(card, profiles);
        if (!profile || profile.email === currentEmail) return;
        card.appendChild(createPasswordButton(profile));
      });
    });
  }

  function interceptLegacyPasswordButton(event) {
    const legacy = event.target?.closest?.("[data-admin-manage-password]");
    if (!legacy) return;
    const uid = String(legacy.dataset.adminManagePassword || "").trim();
    const profile = profilesCache.find((item) => item.uid === uid) || profilesFromApp().find((item) => item.uid === uid);
    if (!profile) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    event.stopPropagation();
    openDialog(profile);
  }

  async function enhance() {
    if (!panelActivated || !isManager()) return;
    ensureSearch();
    await enhanceUserCards();
    applySearch();
  }

  function installObserver() {
    if (observer || !panelActivated) return;
    const panel = document.getElementById("panel-utenti");
    if (!panel) return;
    let scheduled = false;
    observer = new MutationObserver(() => {
      if (scheduled) return;
      scheduled = true;
      window.setTimeout(() => {
        scheduled = false;
        void enhance();
      }, 80);
    });
    observer.observe(panel, { childList: true, subtree: true });
  }

  function activateUserPanel() {
    panelActivated = true;
    installObserver();
    void enhance();
  }

  function initialize() {
    document.addEventListener("click", interceptLegacyPasswordButton, true);
    document.addEventListener("click", (event) => {
      if (event.target?.closest?.("#open-panel-utenti")) {
        window.setTimeout(activateUserPanel, 120);
      }
    }, true);
    const panel = document.getElementById("panel-utenti");
    if (panel && !panel.classList.contains("hidden")) activateUserPanel();
  }

  window.HeraAdminUserAccessTools = {
    installed: true,
    version: VERSION,
    refresh: enhance,
    activate: activateUserPanel,
    search: applySearch,
    openPassword: openDialog
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initialize, { once: true });
  else initialize();
})();
