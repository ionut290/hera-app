(() => {
  "use strict";

  const MIN_PASSWORD_LENGTH = 10;
  const SCRIPT_VERSION = "1.0.0";
  let busy = false;
  let observer = null;

  const esc = (value) => String(value ?? "").replace(/[&<>'"]/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;"
  }[char]));

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

  async function markPasswordChangeRequired(profile) {
    if (!window.firebase || typeof firebase.firestore !== "function") return;
    await firebase.firestore().collection("platformUsers").doc(profile.uid).set({
      mustChangePassword: true,
      passwordResetRequestedAt: firebase.firestore.FieldValue.serverTimestamp(),
      passwordResetRequestedBy: firebase.auth().currentUser?.uid || "",
      passwordResetRequestedByEmail: firebase.auth().currentUser?.email || "",
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  }

  function getContinueUrl() {
    const base = "https://creative-syrniki-dddbae.netlify.app/";
    const url = new URL(base);
    url.searchParams.set("passwordReset", "complete");
    return url.toString();
  }

  async function sendResetEmail(profile) {
    if (!window.firebase || typeof firebase.auth !== "function") {
      throw new Error("Firebase Authentication non disponibile.");
    }
    const auth = firebase.auth();
    auth.languageCode = "it";
    try {
      await auth.sendPasswordResetEmail(profile.email, {
        url: getContinueUrl(),
        handleCodeInApp: false
      });
    } catch (error) {
      const code = String(error?.code || "").toLowerCase();
      if (!["auth/unauthorized-continue-uri", "auth/invalid-continue-uri"].includes(code)) throw error;
      await auth.sendPasswordResetEmail(profile.email);
    }
  }

  async function requestPasswordReset(profile) {
    if (!isManager() || busy || !profile?.email) return;
    const confirmed = window.confirm(
      `Reimpostare la password di ${profile.displayName}?\n\n` +
      `Verrà inviata un'email a ${profile.email} con il link sicuro per scegliere la nuova password.`
    );
    if (!confirmed) return;

    busy = true;
    try {
      await markPasswordChangeRequired(profile);
      await sendResetEmail(profile);
      alert(`Email di cambio password inviata a ${profile.email}.`);
    } catch (error) {
      console.error("Reset password amministratore fallito:", error);
      const code = String(error?.code || "").toLowerCase();
      if (code === "auth/network-request-failed") alert("Connessione non disponibile. Riprova quando sei online.");
      else if (code === "auth/too-many-requests") alert("Troppe richieste di reset. Attendi qualche minuto e riprova.");
      else alert(error?.message || "Non è stato possibile avviare il cambio password.");
    } finally {
      busy = false;
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
      if (actions.querySelector("[data-admin-reset-password]")) return;
      const modal = actions.closest("#registry-modal");
      const editor = modal?.querySelector("#registry-editor");
      const personId = modal?.querySelector("[data-link-user]")?.dataset?.linkUser ||
        modal?.querySelector("[data-unlink-user]")?.dataset?.unlinkUser || "";
      if (!personId) return;
      const profile = profileByLinkedPerson(personId);
      if (!profile) return;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "btn";
      button.dataset.adminResetPassword = profile.uid;
      button.dataset.adminResetEmail = profile.email;
      button.textContent = "Cambia password";
      button.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        void requestPasswordReset(profile);
      });
      actions.appendChild(button);
    });
  }

  function enhancePendingCards(root = document) {
    if (!isManager()) return;
    const profiles = currentProfiles();
    root.querySelectorAll?.(".pending-user-card").forEach((card) => {
      if (card.querySelector("[data-admin-reset-password]")) return;
      const emailText = Array.from(card.querySelectorAll("p,dd")).map((node) => String(node.textContent || "").trim().toLowerCase())
        .find((value) => value.includes("@"));
      const profile = profiles.find((item) => item.email === emailText);
      if (!profile) return;
      const actions = card.querySelector(".actions-row") || card;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "btn";
      button.dataset.adminResetPassword = profile.uid;
      button.textContent = "CAMBIA PASSWORD";
      button.addEventListener("click", () => void requestPasswordReset(profile));
      actions.appendChild(button);
    });
  }

  function enhance() {
    enhanceRegistryModal(document);
    enhancePendingCards(document);
  }

  function installObserver() {
    if (observer || !document.body) return;
    observer = new MutationObserver(() => enhance());
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
    requestPasswordReset,
    refresh: enhance,
    minPasswordLength: MIN_PASSWORD_LENGTH
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initialize, { once: true });
  else initialize();
})();
