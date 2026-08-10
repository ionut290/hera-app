(() => {
  "use strict";

  const SCRIPT_VERSION = "3.3.0";
  const SEARCH_LIST_IDS = [
    "pending-users-list",
    "admin-users-list",
    "user-ban-list",
    "user-permissions-list"
  ];
  let busy = false;
  let observer = null;
  let activeProfile = null;

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

  function currentProfiles() {
    if (typeof platformUsers === "undefined" || !Array.isArray(platformUsers)) return [];
    return platformUsers.map((item) => {
      const displayName = String(
        item?.displayName ||
        item?.nomeCompleto ||
        item?.fullName ||
        [item?.nome || item?.firstName, item?.cognome || item?.lastName].filter(Boolean).join(" ") ||
        item?.email ||
        "Utente"
      ).trim();
      return {
        id: String(item?.id || item?.uid || ""),
        uid: String(item?.uid || item?.id || ""),
        email: String(item?.email || "").trim().toLowerCase(),
        displayName,
        searchText: normalize(`${displayName} ${item?.email || ""} ${item?.uid || item?.id || ""}`)
      };
    }).filter((item) => item.uid && item.email);
  }

  function getCallable() {
    if (!window.firebase || typeof firebase.app !== "function") {
      throw new Error("Firebase Functions non disponibile.");
    }
    const app = firebase.app();
    if (!app || typeof app.functions !== "function") {
      throw new Error("Firebase Functions non disponibile.");
    }
    const functionsClient = app.functions("europe-west1");
    if (!functionsClient || typeof functionsClient.httpsCallable !== "function") {
      throw new Error("Firebase Functions non disponibile.");
    }
    return functionsClient.httpsCallable("adminSetUserPassword");
  }

  function ensureDialog() {
    let dialog = document.getElementById("admin-password-dialog");
    if (dialog) return dialog;

    dialog = document.createElement("dialog");
    dialog.id = "admin-password-dialog";
    dialog.className = "biometric-dialog admin-password-dialog";
    dialog.innerHTML = `
      <form id="admin-password-form" method="dialog">
        <h2>Cambia password utente</h2>
        <p id="admin-password-user" class="muted"></p>
        <section class="admin-password-section">
          <p>L'amministratore genera una nuova password temporanea e la comunica direttamente all'utente.</p>
          <p><strong>Nessuna email viene inviata.</strong> La password precedente smette di funzionare. Al prossimo accesso l'utente dovrà scegliere una nuova password personale prima di continuare.</p>
          <button id="admin-generate-temp-password" class="btn btn-primary" type="button">GENERA NUOVA PASSWORD</button>
        </section>
        <section id="admin-password-result" class="admin-password-result hidden" aria-live="polite">
          <strong>Nuova password temporanea</strong>
          <div class="auth-password-row">
            <input id="admin-password-result-value" type="text" readonly autocomplete="off">
            <button id="admin-copy-temp-password" class="btn" type="button">COPIA</button>
          </div>
          <p class="muted">Comunica questa password all'utente. Non viene salvata in chiaro in Firestore e viene mostrata qui solo dopo la generazione.</p>
        </section>
        <p id="admin-password-feedback" class="muted" role="status" aria-live="polite"></p>
        <div class="actions-row">
          <button id="admin-password-close" class="btn" type="button">CHIUDI</button>
        </div>
      </form>`;
    document.body.appendChild(dialog);

    dialog.addEventListener("cancel", (event) => {
      event.preventDefault();
      closeDialog();
    });
    dialog.querySelector("#admin-password-close")?.addEventListener("click", closeDialog);
    dialog.querySelector("#admin-generate-temp-password")?.addEventListener("click", () => void generateTemporaryPassword());
    dialog.querySelector("#admin-copy-temp-password")?.addEventListener("click", copyTemporaryPassword);
    return dialog;
  }

  function setFeedback(message) {
    const node = ensureDialog().querySelector("#admin-password-feedback");
    if (node) node.textContent = message || "";
  }

  function setBusy(value) {
    busy = value;
    const dialog = ensureDialog();
    dialog.querySelectorAll("button").forEach((button) => {
      if (button.id !== "admin-password-close") button.disabled = value;
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
    dialog.querySelector("#admin-password-user").textContent = `${profile.displayName} • ${profile.email}`;
    dialog.querySelector("#admin-password-result-value").value = "";
    dialog.querySelector("#admin-password-result").classList.add("hidden");
    setFeedback("");
    if (!dialog.open) dialog.showModal();
  }

  function closeDialog() {
    const dialog = document.getElementById("admin-password-dialog");
    if (dialog?.open) dialog.close();
    activeProfile = null;
    busy = false;
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
    try {
      const callable = getCallable();
      const response = await callable({
        uid: activeProfile.uid,
        email: activeProfile.email
      });
      const result = response?.data || null;
      if (!result?.success || !result?.temporaryPassword) {
        throw new Error("Il server non ha restituito la nuova password temporanea.");
      }
      if (result.uid && String(result.uid) !== String(activeProfile.uid)) {
        throw new Error("Il server ha aggiornato un account diverso da quello selezionato.");
      }
      const dialog = ensureDialog();
      dialog.querySelector("#admin-password-result-value").value = result.temporaryPassword;
      dialog.querySelector("#admin-password-result").classList.remove("hidden");
      setFeedback("Password cambiata correttamente. Comunica la password temporanea all'utente: al prossimo accesso dovrà crearne una nuova personale.");
    } catch (error) {
      console.error("Gestione password amministratore fallita:", error);
      const code = String(error?.code || "").toLowerCase();
      if (code.includes("permission-denied")) setFeedback("Operazione non autorizzata: serve un account amministratore.");
      else if (code.includes("not-found")) setFeedback("Account Firebase dell'utente non trovato.");
      else if (code.includes("unavailable") || code.includes("network")) setFeedback("Connessione non disponibile. Riprova quando sei online.");
      else if (code.includes("internal")) setFeedback("Il server non è riuscito a cambiare la password. Riprova tra poco.");
      else setFeedback(error?.message || "Cambio password non riuscito.");
    } finally {
      setBusy(false);
    }
  }

  async function copyTemporaryPassword() {
    const value = String(ensureDialog().querySelector("#admin-password-result-value")?.value || "");
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
      setFeedback("Password temporanea copiata.");
    } catch (_) {
      const input = ensureDialog().querySelector("#admin-password-result-value");
      input?.select();
      setFeedback("Seleziona la password e copiala manualmente.");
    }
  }

  function profileByLinkedPerson(personId) {
    if (typeof personaleRecords === "undefined" || !Array.isArray(personaleRecords)) return null;
    const person = personaleRecords.find((item) => String(item?.id || "") === String(personId || ""));
    if (!person) return null;
    const uid = String(person.linkedUserId || "");
    const email = String(person.linkedUserEmail || person.email || "").trim().toLowerCase();
    return currentProfiles().find((profile) => (uid && profile.uid === uid) || (email && profile.email === email)) || null;
  }

  function profileForCard(card, profiles = currentProfiles()) {
    if (!card) return null;
    const text = normalize(card.textContent || "");
    const dataText = normalize(Array.from(card.querySelectorAll("[data-user-id],[data-uid],[data-user],[data-ban-user],[data-user-ban]"))
      .map((node) => Object.values(node.dataset || {}).join(" "))
      .join(" "));
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

  function createPasswordButton(profile, label = "CAMBIA PASSWORD") {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn";
    button.dataset.adminManagePassword = profile.uid;
    button.textContent = label;
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openDialog(profile);
    });
    return button;
  }

  function enhanceRegistryModal(root = document) {
    if (!isManager()) return;
    root.querySelectorAll?.(".registry-account-actions").forEach((actions) => {
      if (actions.querySelector("[data-admin-manage-password]")) return;
      const modal = actions.closest("#registry-modal");
      const personId = modal?.querySelector("[data-link-user]")?.dataset?.linkUser ||
        modal?.querySelector("[data-unlink-user]")?.dataset?.unlinkUser || "";
      if (!personId) return;
      const profile = profileByLinkedPerson(personId);
      if (!profile) return;
      actions.appendChild(createPasswordButton(profile, "Cambia password"));
    });
  }

  function enhancePendingCards(root = document) {
    if (!isManager()) return;
    const profiles = currentProfiles();
    root.querySelectorAll?.(".pending-user-card").forEach((card) => {
      if (card.querySelector("[data-admin-manage-password]")) return;
      const profile = profileForCard(card, profiles);
      if (!profile) return;
      const actions = card.querySelector(".actions-row") || card;
      actions.appendChild(createPasswordButton(profile));
    });
  }

  function enhanceBanCards(root = document) {
    if (!isManager()) return;
    const list = root.querySelector?.("#user-ban-list");
    if (!list) return;
    const profiles = currentProfiles();
    const currentEmail = String(firebase.auth()?.currentUser?.email || "").trim().toLowerCase();

    Array.from(list.children).forEach((card) => {
      if (!(card instanceof HTMLElement)) return;
      if (card.querySelector("[data-admin-manage-password]")) return;
      const profile = profileForCard(card, profiles);
      if (!profile || profile.email === currentEmail) return;
      card.appendChild(createPasswordButton(profile));
    });
  }

  function ensureUserSearch(root = document) {
    if (!isManager()) return null;
    const panel = root.querySelector?.("#panel-utenti") || document.getElementById("panel-utenti");
    if (!panel) return null;
    let input = panel.querySelector("#user-management-search-input");
    if (input) return input;

    const toolbar = document.createElement("div");
    toolbar.className = "personale-search-toolbar user-management-search-toolbar";
    toolbar.innerHTML = `
      <label class="personale-search-input-wrap" for="user-management-search-input">
        <span aria-hidden="true">🔍</span>
        <input id="user-management-search-input" type="search" placeholder="Cerca utente per nome o email" autocomplete="off">
      </label>
      <button id="user-management-search-clear" class="btn btn-small" type="button">Azzera</button>
      <p id="user-management-search-feedback" class="muted" role="status" aria-live="polite"></p>`;

    panel.insertBefore(toolbar, panel.firstChild);
    input = toolbar.querySelector("#user-management-search-input");
    input?.addEventListener("input", applyUserSearch);
    toolbar.querySelector("#user-management-search-clear")?.addEventListener("click", () => {
      if (!input) return;
      input.value = "";
      applyUserSearch();
      input.focus();
    });
    return input;
  }

  function applyUserSearch() {
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

  function enhance() {
    ensureUserSearch(document);
    enhanceRegistryModal(document);
    enhancePendingCards(document);
    enhanceBanCards(document);
    applyUserSearch();
  }

  function installObserver() {
    if (observer || !document.body) return;
    observer = new MutationObserver(enhance);
    observer.observe(document.body, { childList: true, subtree: true });
  }

  function initialize() {
    enhance();
    installObserver();
    document.addEventListener("click", (event) => {
      if (event.target?.closest?.("#open-panel-utenti, #open-panel-personale")) {
        window.setTimeout(enhance, 150);
      }
    }, true);
  }

  window.HeraAdminPasswordManager = {
    installed: true,
    version: SCRIPT_VERSION,
    open: openDialog,
    refresh: enhance,
    search: applyUserSearch
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initialize, { once: true });
  else initialize();
})();
