(() => {
  "use strict";

  const SCRIPT_VERSION = "3.0.0";
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

  function currentProfiles() {
    if (typeof platformUsers === "undefined" || !Array.isArray(platformUsers)) return [];
    return platformUsers.map((item) => ({
      id: String(item?.id || item?.uid || ""),
      uid: String(item?.uid || item?.id || ""),
      email: String(item?.email || "").trim().toLowerCase(),
      displayName: String(item?.displayName || item?.nomeCompleto || item?.fullName || item?.email || "Utente")
    })).filter((item) => item.uid && item.email);
  }

  function getCallable() {
    if (!window.firebase || typeof firebase.functions !== "function") {
      throw new Error("Firebase Functions non disponibile.");
    }
    return firebase.functions("europe-west1").httpsCallable("adminSetUserPassword");
  }

  function ensureDialog() {
    let dialog = document.getElementById("admin-password-dialog");
    if (dialog) return dialog;

    dialog = document.createElement("dialog");
    dialog.id = "admin-password-dialog";
    dialog.className = "biometric-dialog admin-password-dialog";
    dialog.innerHTML = `
      <form id="admin-password-form" method="dialog">
        <h2>Password temporanea utente</h2>
        <p id="admin-password-user" class="muted"></p>
        <section class="admin-password-section">
          <p>L'amministratore genera una password temporanea e la comunica direttamente all'utente.</p>
          <p><strong>Nessuna email viene inviata.</strong> L'utente accede dall'app con la password temporanea e deve scegliere subito una nuova password personale prima di poter continuare.</p>
          <button id="admin-generate-temp-password" class="btn btn-primary" type="button">GENERA PASSWORD TEMPORANEA</button>
        </section>
        <section id="admin-password-result" class="admin-password-result hidden" aria-live="polite">
          <strong>Password temporanea</strong>
          <div class="auth-password-row">
            <input id="admin-password-result-value" type="text" readonly autocomplete="off">
            <button id="admin-copy-temp-password" class="btn" type="button">COPIA</button>
          </div>
          <p class="muted">Comunica questa password all'utente. Non viene salvata in Firestore e viene mostrata qui solo dopo la generazione.</p>
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
      `Generare una password temporanea per ${activeProfile.displayName}?\n\n` +
      "La password precedente smetterà di funzionare. Nessuna email verrà inviata. Al prossimo accesso l'utente dovrà scegliere una nuova password dentro l'app."
    );
    if (!confirmed) return;

    setBusy(true);
    setFeedback("Generazione password temporanea...");
    try {
      const callable = getCallable();
      const response = await callable({
        uid: activeProfile.uid,
        email: activeProfile.email,
        mode: "temporary"
      });
      const result = response?.data || null;
      if (!result?.success || !result.temporaryPassword) {
        throw new Error("Il server non ha restituito la password temporanea.");
      }
      const dialog = ensureDialog();
      dialog.querySelector("#admin-password-result-value").value = result.temporaryPassword;
      dialog.querySelector("#admin-password-result").classList.remove("hidden");
      setFeedback("Password temporanea creata. Comunicala all'utente: il cambio successivo avverrà direttamente nell'app.");
    } catch (error) {
      console.error("Gestione password amministratore fallita:", error);
      const code = String(error?.code || "").toLowerCase();
      if (code.includes("permission-denied")) setFeedback("Operazione non autorizzata: serve un account amministratore.");
      else if (code.includes("not-found")) setFeedback("Account Firebase dell'utente non trovato.");
      else if (code.includes("unavailable") || code.includes("network")) setFeedback("Connessione non disponibile. Riprova quando sei online.");
      else setFeedback(error?.message || "Creazione password temporanea non riuscita.");
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
      const button = document.createElement("button");
      button.type = "button";
      button.className = "btn";
      button.dataset.adminManagePassword = profile.uid;
      button.textContent = "Password temporanea";
      button.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        openDialog(profile);
      });
      actions.appendChild(button);
    });
  }

  function enhancePendingCards(root = document) {
    if (!isManager()) return;
    const profiles = currentProfiles();
    root.querySelectorAll?.(".pending-user-card").forEach((card) => {
      if (card.querySelector("[data-admin-manage-password]")) return;
      const emailText = Array.from(card.querySelectorAll("p,dd"))
        .map((node) => String(node.textContent || "").trim().toLowerCase())
        .find((value) => value.includes("@"));
      const profile = profiles.find((item) => item.email === emailText);
      if (!profile) return;
      const actions = card.querySelector(".actions-row") || card;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "btn";
      button.dataset.adminManagePassword = profile.uid;
      button.textContent = "PASSWORD TEMPORANEA";
      button.addEventListener("click", () => openDialog(profile));
      actions.appendChild(button);
    });
  }

  function enhance() {
    enhanceRegistryModal(document);
    enhancePendingCards(document);
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
    refresh: enhance
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initialize, { once: true });
  else initialize();
})();
