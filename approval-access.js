(() => {
  "use strict";
  const PHONE = "393892352575";
  const WAIT_MS = 60000;
  const REGION = "europe-west1";
  const APP_URL = "https://creative-syrniki-dddbae.netlify.app";
  const ANDROID_URL = "https://play.google.com/store/apps/details?id=it.vargacantieri.hera";
  const PENDING = new Set(["in_attesa", "richiesta_inviata"]);
  const ACTIVE = new Set(["attivo", "active", "approved", "autorizzato", "abilitato"]);
  let db;
  let auth;
  let user;
  let profile;
  let unsubscribe;
  let cooldownTimer;
  let pendingUsersById = new Map();

  const el = (id) => document.getElementById(id);
  const serverTime = () => firebase.firestore.FieldValue.serverTimestamp();
  const normalizeStatus = (value) => String(value || "").trim().toLowerCase();
  const hasExplicitStatus = (data) => Boolean(
    data && (
      Object.prototype.hasOwnProperty.call(data, "statoAccount") ||
      Object.prototype.hasOwnProperty.call(data, "accountStatus") ||
      Object.prototype.hasOwnProperty.call(data, "banned")
    )
  );
  const statusOf = (data) => {
    if (!data) return "in_attesa";
    if (data.banned === true) return "bloccato";
    const explicit = normalizeStatus(data.statoAccount || data.accountStatus);
    if (explicit) return explicit;
    // I profili creati dal nuovo flusso sono gli unici che devono attendere
    // l'approvazione. Tutti i documenti storici già presenti in Firestore,
    // privi dei nuovi campi, restano autorizzati senza alcuna migrazione.
    if (data.profileMigratedByEmail === false) return "in_attesa";
    return "attivo";
  };
  const isAllowedStatus = (status) => ACTIVE.has(normalizeStatus(status));
  const splitName = (firebaseUser) => {
    const parts = String(firebaseUser.displayName || "").trim().split(/\s+/).filter(Boolean);
    return { nome: parts[0] || "", cognome: parts.slice(1).join(" "), nomeCompleto: parts.join(" ") || firebaseUser.email || "Utente" };
  };
  const dateValue = (value) => value?.toDate ? value.toDate() : value ? new Date(value) : null;
  const formatDateTime = (value) => {
    const date = dateValue(value);
    return date && !Number.isNaN(date.getTime()) ? date.toLocaleString("it-IT") : "—";
  };
  const escape = (value) => String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char]));
  const firstValue = (source, keys) => keys.map((key) => String(source?.[key] || "").trim()).find(Boolean) || "";

  function createRequestId() {
    if (window.crypto?.randomUUID) return window.crypto.randomUUID();
    return `${Date.now()}_${Math.random().toString(36).slice(2, 12)}`;
  }

  function approvalMessage(target, administratorName) {
    const userName = firstValue(target, ["nomeCompleto", "displayName", "fullName"]) || target?.email || "";
    const userEmail = String(target?.email || "").trim();
    return [
      "✅ ACCESSO A VARGA CANTIERI APPROVATO",
      "",
      `Ciao ${userName},`,
      "",
      `l’amministratore ${administratorName} ha accettato la tua richiesta. Ora puoi accedere all’app Varga Cantieri.`,
      "",
      "A COSA SERVE L’APP",
      "• consultare le commesse e gli impianti di lavoro;",
      "• vedere squadre, attività e informazioni operative;",
      "• aprire la navigazione verso gli impianti;",
      "• consultare documenti, comunicazioni e aggiornamenti autorizzati.",
      "",
      "COME ACCEDERE",
      `Email: ${userEmail}`,
      "Password: usa la password che hai scelto durante la registrazione.",
      "Non comunicare la password ad altre persone.",
      "",
      "INSTALLAZIONE SU ANDROID",
      `1. Apri Google Play: ${ANDROID_URL}`,
      "2. Premi Installa.",
      "3. Apri Varga Cantieri e accedi con l’email e la password scelte.",
      "",
      "INSTALLAZIONE SU IPHONE",
      `1. Apri con Safari: ${APP_URL}`,
      "2. Premi Condividi (quadrato con freccia verso l’alto).",
      "3. Premi Aggiungi alla schermata Home e poi Aggiungi.",
      "4. Apri l’icona Varga Cantieri e accedi con l’email e la password scelte.",
      "",
      `Accesso web: ${APP_URL}`,
      "",
      "Benvenuto e buon lavoro!"
    ].join("\n");
  }

  function approvalPhone(target) {
    return firstValue(target, ["whatsapp", "whatsappPhone", "telefono", "cellulare", "phone", "phoneNumber"]);
  }

  function normalizePhone(value) {
    let digits = String(value || "").replace(/\D/g, "");
    if (digits.startsWith("00")) digits = digits.slice(2);
    if (/^3\d{8,10}$/.test(digits)) digits = `39${digits}`;
    return digits;
  }

  function ensureApprovalResultDialog() {
    let dialog = el("access-approval-result-dialog");
    if (dialog) return dialog;
    dialog = document.createElement("section");
    dialog.id = "access-approval-result-dialog";
    dialog.className = "access-approval-result-dialog hidden";
    dialog.setAttribute("role", "dialog");
    dialog.setAttribute("aria-modal", "true");
    dialog.setAttribute("aria-labelledby", "access-approval-result-title");
    dialog.innerHTML = `
      <div class="access-approval-result-card">
        <div class="access-approval-result-icon" aria-hidden="true">✅</div>
        <h2 id="access-approval-result-title">Utente sbloccato</h2>
        <p id="access-approval-result-email-status" class="access-approval-result-status"></p>
        <p class="muted">Puoi inviare allo stesso utente anche il medesimo messaggio tramite WhatsApp.</p>
        <label class="access-approval-phone-label">Numero WhatsApp (facoltativo)
          <input id="access-approval-result-phone" type="tel" inputmode="tel" autocomplete="tel" placeholder="Es. 393331234567">
        </label>
        <p class="muted">Se lasci il numero vuoto, WhatsApp si apre e puoi scegliere il contatto.</p>
        <div class="access-approval-result-actions">
          <button id="access-approval-result-email" class="btn btn-primary" type="button">INVIA EMAIL</button>
          <button id="access-approval-result-whatsapp" class="btn access-whatsapp" type="button">INVIA LO STESSO MESSAGGIO SU WHATSAPP</button>
          <button id="access-approval-result-copy" class="btn" type="button">COPIA MESSAGGIO</button>
          <button id="access-approval-result-close" class="btn" type="button">CHIUDI</button>
        </div>
      </div>`;
    document.body.appendChild(dialog);
    el("access-approval-result-close")?.addEventListener("click", () => dialog.classList.add("hidden"));
    el("access-approval-result-email")?.addEventListener("click", () => {
      const message = String(dialog.dataset.message || "");
      const email = String(dialog.dataset.email || "").trim();
      if (!message || !email) return alert("Email o messaggio di approvazione non disponibile.");
      window.location.href = `mailto:${encodeURIComponent(email)}?subject=${encodeURIComponent("✅ Accesso a Varga Cantieri approvato")}&body=${encodeURIComponent(message)}`;
    });
    el("access-approval-result-whatsapp")?.addEventListener("click", () => {
      const message = String(dialog.dataset.message || "");
      if (!message) return alert("Messaggio di approvazione non disponibile.");
      const phone = normalizePhone(el("access-approval-result-phone")?.value || "");
      window.open(`https://wa.me/${phone}?text=${encodeURIComponent(message)}`, "_blank", "noopener,noreferrer");
    });
    el("access-approval-result-copy")?.addEventListener("click", async () => {
      const message = String(dialog.dataset.message || "");
      try {
        await navigator.clipboard.writeText(message);
        alert("Messaggio copiato.");
      } catch (_) {
        window.prompt("Copia il messaggio:", message);
      }
    });
    return dialog;
  }

  function showApprovalResult(result) {
    const dialog = ensureApprovalResultDialog();
    dialog.dataset.message = String(result.message || "");
    dialog.dataset.email = String(result.userEmail || "");
    el("access-approval-result-title").textContent = `${result.userName || "Utente"} è stato sbloccato`;
    el("access-approval-result-email-status").textContent = result.emailSent
      ? `Email di conferma inviata automaticamente a ${result.userEmail}.`
      : `Utente sbloccato. ${result.emailError || "Premi INVIA EMAIL oppure usa WhatsApp per avvisarlo."}`;
    el("access-approval-result-email-status").classList.toggle("is-warning", !result.emailSent);
    el("access-approval-result-phone").value = String(result.phone || "");
    dialog.classList.remove("hidden");
  }

  function stop() {
    if (unsubscribe) unsubscribe();
    unsubscribe = null;
    clearInterval(cooldownTimer);
  }

  function hide() {
    const screen = el("access-approval-screen");
    document.body.classList.remove("access-approval-locked");
    screen?.classList.add("hidden");
  }

  function show(state, message = "") {
    const screen = el("access-approval-screen");
    if (!screen) return;
    document.body.classList.add("access-approval-locked");
    screen.classList.remove("hidden");
    el("access-approval-loading").classList.toggle("hidden", state !== "checking" && state !== "error");
    el("access-approval-content").classList.toggle("hidden", state === "checking" || state === "error");
    el("access-approval-error").classList.toggle("hidden", state !== "error");
    if (state === "error") el("access-approval-error-message").textContent = message || "Impossibile verificare l’autorizzazione. Controlla la connessione e riprova.";
    const rejected = state === "rifiutato" || state === "bloccato";
    el("access-approval-title").textContent = rejected ? "ACCESSO NON AUTORIZZATO" : "ACCESSO IN ATTESA DI APPROVAZIONE";
    el("access-approval-description").textContent = rejected
      ? "La tua richiesta di accesso non è stata approvata. Per maggiori informazioni contatta l’amministratore."
      : "Il tuo account è stato registrato correttamente, ma deve essere autorizzato da un amministratore. Premi il pulsante qui sotto per inviare la richiesta di accesso tramite WhatsApp.";
    el("access-approval-status").textContent = rejected ? (state === "bloccato" ? "Account bloccato" : "Accesso rifiutato") : "In attesa di approvazione";
    el("access-request-whatsapp").textContent = rejected ? "CONTATTA L’AMMINISTRATORE SU WHATSAPP" : "RICHIEDI ACCESSO SU WHATSAPP";
    el("access-check-approval").classList.toggle("hidden", rejected);
    if (profile) {
      el("access-approval-name").textContent = profile.nomeCompleto || profile.displayName || user?.displayName || "—";
      el("access-approval-email").textContent = profile.email || user?.email || "—";
    }
  }

  async function ensureProfile(firebaseUser) {
    const ref = db.collection("platformUsers").doc(firebaseUser.uid);
    const bootstrap = window.HeraPlatformProfileBootstrap;
    const bootstrapAgeMs = Date.now() - Number(bootstrap?.loadedAt || 0);
    if (
      bootstrap?.uid === firebaseUser.uid
      && bootstrap.exists === true
      && bootstrapAgeMs >= 0
      && bootstrapAgeMs <= 2000
    ) {
      window.HeraPlatformProfileBootstrap = null;
      return bootstrap.data && typeof bootstrap.data === "object" ? bootstrap.data : {};
    }
    return db.runTransaction(async (transaction) => {
      const snapshot = await transaction.get(ref);
      if (snapshot.exists) {
        // Non scrivere sui profili storici: la scrittura automatica dei nuovi
        // campi può essere negata dalle regole Firestore e bloccare l'accesso.
        // statusOf() gestisce in memoria la compatibilità con gli utenti esistenti.
        return snapshot.data() || {};
      }
      const names = splitName(firebaseUser);
      const created = {
        uid: firebaseUser.uid, email: firebaseUser.email || "", displayName: names.nomeCompleto,
        photoURL: firebaseUser.photoURL || "",
        providerId: Array.isArray(firebaseUser.providerData) ? String(firebaseUser.providerData.find((provider) => provider?.providerId === "google.com")?.providerId || "") : "",
        emailVerified: firebaseUser.emailVerified === true,
        ...names, statoAccount: "in_attesa", accountStatus: "in_attesa", ruolo: "user", role: "user",
        primoAccessoAt: serverTime(), firstLoginAt: serverTime(), approvatoAt: null, approvedAt: null,
        approvatoDa: null, approvedBy: null, numeroRichieste: 0, requestCount: 0
      };
      transaction.set(ref, created);
      return created;
    });
  }

  function listen() {
    stop();
    unsubscribe = db.collection("platformUsers").doc(user.uid).onSnapshot((snapshot) => {
      if (!snapshot.exists) return;
      profile = { id: snapshot.id, ...(snapshot.data() || {}) };
      const status = statusOf(profile);
      if (isAllowedStatus(status)) {
        stop();
        el("access-approval-feedback").textContent = "Accesso autorizzato";
        setTimeout(() => window.location.reload(), 700);
      } else show(status);
    }, () => show("error"));
  }

  async function verify(firebaseUser) {
    db = firebase.firestore(); auth = firebase.auth(); user = firebaseUser;
    hide();
    try {
      profile = { id: user.uid, ...(await ensureProfile(user)) };
      const status = statusOf(profile);
      const allowed = isAllowedStatus(status);
      if (!allowed) { show(status); listen(); }
      else hide();
      return { allowed, profile, status };
    } catch (error) {
      console.error("Verifica autorizzazione fallita", error);
      show("error");
      return { allowed: false, error };
    }
  }

  function whatsappText() {
    const now = new Date();
    const date = now.toLocaleDateString("it-IT", { day: "2-digit", month: "2-digit", year: "numeric" });
    const time = now.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" });
    return `🔐 RICHIESTA ACCESSO VARGA CANTIERI\n\nUn nuovo utente richiede l’autorizzazione per accedere all’app.\n\n👤 Nome e cognome: ${profile?.nomeCompleto || profile?.displayName || user.displayName || ""}\n📧 Email: ${profile?.email || user.email || ""}\n🆔 ID utente: ${user.uid}\n📅 Data richiesta: ${date}\n🕐 Ora richiesta: ${time}\n\nL’utente è attualmente in attesa di approvazione.\n\nAccedi nell’app come amministratore e apri:\n\nMenu → Gestione utenti → Utenti in attesa`;
  }

  async function requestWhatsapp() {
    const button = el("access-request-whatsapp");
    if (button.disabled) return;
    try {
      if (PENDING.has(statusOf(profile))) {
        await db.collection("platformUsers").doc(user.uid).update({
          statoAccount: "richiesta_inviata", accountStatus: "richiesta_inviata",
          richiestaWhatsappAt: serverTime(), whatsappRequestedAt: serverTime(),
          ultimaRichiestaAt: serverTime(), lastRequestAt: serverTime(),
          numeroRichieste: firebase.firestore.FieldValue.increment(1), requestCount: firebase.firestore.FieldValue.increment(1)
        });
      }
      window.open(`https://wa.me/${PHONE}?text=${encodeURIComponent(whatsappText())}`, "_blank", "noopener,noreferrer");
      el("access-approval-feedback").textContent = "Richiesta preparata. Completa l’invio su WhatsApp.";
      let remaining = WAIT_MS / 1000;
      button.disabled = true;
      button.textContent = `RIPROVA TRA ${remaining}s`;
      cooldownTimer = setInterval(() => {
        remaining -= 1;
        button.textContent = remaining > 0 ? `RIPROVA TRA ${remaining}s` : "RICHIEDI ACCESSO SU WHATSAPP";
        if (remaining <= 0) { clearInterval(cooldownTimer); button.disabled = false; }
      }, 1000);
    } catch (error) { show("error"); }
  }

  async function refresh() {
    show("checking");
    try {
      const snapshot = await db.collection("platformUsers").doc(user.uid).get({ source: "server" });
      profile = snapshot.data();
      const status = statusOf(profile);
      if (isAllowedStatus(status)) hide();
      else show(status);
    } catch (_) { show("error"); }
  }

  async function approve(uid, button) {
    const admin = firebase.auth().currentUser;
    if (!admin) throw new Error("Sessione amministratore non disponibile. Esci e accedi di nuovo.");
    const target = pendingUsersById.get(String(uid)) || { uid, id: uid, email: "" };
    const administratorName = firstValue(profile, ["nomeCompleto", "displayName"])
      || admin.displayName
      || admin.email
      || "Amministratore";
    const requestId = createRequestId();
    button.disabled = true;
    button.textContent = "SBLOCCO E INVIO EMAIL…";

    try {
      const callable = firebase.app().functions(REGION).httpsCallable("approveUserAccess");
      const response = await callable({
        targetUid: uid,
        requestId,
        administratorName
      });
      if (!response?.data?.approved) throw new Error("Il backend non ha confermato lo sblocco.");
      showApprovalResult(response.data);
      return;
    } catch (cloudError) {
      const code = String(cloudError?.code || "");
      if (["functions/unauthenticated", "functions/permission-denied", "functions/invalid-argument", "functions/not-found"].includes(code)) {
        throw cloudError;
      }
      console.warn("Callable approvazione non disponibile; applico il salvataggio Firestore compatibile.", cloudError);
    }

    const ref = db.collection("platformUsers").doc(uid);
    const audit = db.collection("userAccessAudit").doc(`approval_${uid}_${requestId}`);
    const patch = {
      statoAccount: "attivo", accountStatus: "attivo", role: "user", ruolo: "user",
      approvatoAt: serverTime(), approvedAt: serverTime(), approvatoDa: admin.uid,
      approvatoDaEmail: admin.email || "", approvatoDaNome: administratorName,
      approvedBy: admin.email || admin.uid, banned: false
    };
    const batch = db.batch();
    batch.update(ref, patch);
    batch.set(audit, {
      userId: uid, action: "approvazione", reason: "", administratorUid: admin.uid,
      administratorEmail: admin.email || "", administratorName, requestId, createdAt: serverTime()
    }, { merge: true });
    await batch.commit();
    showApprovalResult({
      approved: true,
      emailSent: false,
      emailError: "Il servizio email non era raggiungibile: invia lo stesso avviso tramite WhatsApp.",
      userName: firstValue(target, ["nomeCompleto", "displayName", "fullName"]) || target.email || "Utente",
      userEmail: target.email || "",
      phone: approvalPhone(target),
      administratorName,
      message: approvalMessage(target, administratorName),
      requestId
    });
  }

  async function decide(uid, decision, button) {
    if (!window.confirm(decision === "attivo" ? "Autorizzare questo utente?\n\nL’utente potrà accedere alle normali funzioni e ai dati dell’app." : "Rifiutare l’accesso a questo utente?")) return;
    if (decision === "attivo") {
      try {
        await approve(uid, button);
      } catch (error) {
        console.error("Sblocco utente non riuscito", error);
        alert(`Sblocco non riuscito: ${error?.message || "controlla la connessione e riprova."}`);
        button.disabled = false;
        button.textContent = "SBLOCCA UTENTE";
      }
      return;
    }
    const reason = decision === "rifiutato" ? (window.prompt("Motivazione facoltativa:", "") || "") : "";
    const admin = firebase.auth().currentUser;
    if (!admin) throw new Error("Sessione amministratore non disponibile. Esci e accedi di nuovo.");
    const ref = db.collection("platformUsers").doc(uid);
    const audit = db.collection("userAccessAudit").doc();
    const patch = decision === "attivo"
      ? { statoAccount: "attivo", accountStatus: "attivo", role: "user", ruolo: "user", approvatoAt: serverTime(), approvedAt: serverTime(), approvatoDa: admin.uid, approvatoDaEmail: admin.email || "", approvedBy: admin.email || admin.uid, banned: false }
      : { statoAccount: "rifiutato", accountStatus: "rifiutato", rifiutatoAt: serverTime(), rejectedAt: serverTime(), rifiutatoDa: admin.uid, rejectedBy: admin.email || admin.uid, motivoRifiuto: reason };
    const batch = db.batch();
    batch.update(ref, patch);
    batch.set(audit, { userId: uid, action: decision === "attivo" ? "approvazione" : "rifiuto", reason, administratorUid: admin.uid, administratorEmail: admin.email || "", createdAt: serverTime() });
    await batch.commit();
    alert("Accesso rifiutato.");
  }

  function renderAdmin(users, isAdmin) {
    const container = el("pending-users-list");
    if (!container) return;
    const pending = isAdmin ? users.filter((item) => PENDING.has(statusOf(item))).sort((a, b) => (dateValue(b.ultimaRichiestaAt || b.primoAccessoAt)?.getTime() || 0) - (dateValue(a.ultimaRichiestaAt || a.primoAccessoAt)?.getTime() || 0)) : [];
    pendingUsersById = new Map(pending.map((item) => [String(item.id || item.uid || ""), item]));
    el("pending-users-section").classList.toggle("hidden", !isAdmin);
    const badge = el("pending-users-badge");
    badge.textContent = pending.length;
    badge.classList.toggle("hidden", !pending.length);
    const menuBadge = el("pending-users-menu-badge");
    menuBadge.textContent = pending.length;
    menuBadge.classList.toggle("hidden", !pending.length);
    container.innerHTML = pending.length ? pending.map((item) => `<article class="pending-user-card"><h4>${escape(item.nomeCompleto || item.displayName)}</h4><p>${escape(item.email)}</p><dl><dt>ID utente</dt><dd>${escape(item.uid || item.id)}</dd><dt>Primo accesso</dt><dd>${formatDateTime(item.primoAccessoAt || item.firstLoginAt)}</dd><dt>Richiesta WhatsApp</dt><dd>${formatDateTime(item.ultimaRichiestaAt || item.richiestaWhatsappAt)}</dd><dt>Stato</dt><dd>${escape(statusOf(item))}</dd><dt>Richieste inviate</dt><dd>${Number(item.numeroRichieste || item.requestCount || 0)}</dd></dl><div class="actions-row"><button class="btn btn-primary" type="button" data-approve-user="${escape(item.id)}">SBLOCCA UTENTE</button><button class="btn btn-danger" type="button" data-reject-user="${escape(item.id)}">RIFIUTA ACCESSO</button></div></article>`).join("") : "<p class='muted'>Nessun utente in attesa.</p>";
  }

  document.addEventListener("click", (event) => {
    const approve = event.target.closest("[data-approve-user]");
    const reject = event.target.closest("[data-reject-user]");
    if (approve) void decide(approve.dataset.approveUser, "attivo", approve);
    if (reject) void decide(reject.dataset.rejectUser, "rifiutato", reject).catch((error) => {
      console.error("Rifiuto accesso non riuscito", error);
      alert(`Rifiuto non riuscito: ${error?.message || "controlla la connessione e riprova."}`);
    });
  });
  el("access-request-whatsapp")?.addEventListener("click", requestWhatsapp);
  el("access-check-approval")?.addEventListener("click", refresh);
  el("access-approval-retry")?.addEventListener("click", refresh);
  el("access-approval-logout")?.addEventListener("click", async () => { stop(); await auth.signOut(); });

  window.HeraAccessApproval = { verify, renderAdmin, statusOf };
})();
