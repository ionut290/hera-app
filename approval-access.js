(() => {
  "use strict";
  const PHONE = "393892352575";
  const WAIT_MS = 60000;
  const PENDING = new Set(["in_attesa", "richiesta_inviata"]);
  const ACTIVE = new Set(["attivo", "active", "approved", "autorizzato", "abilitato"]);
  let db;
  let auth;
  let user;
  let profile;
  let unsubscribe;
  let cooldownTimer;

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

  async function decide(uid, decision) {
    if (!window.confirm(decision === "attivo" ? "Autorizzare questo utente?\n\nL’utente potrà accedere alle normali funzioni e ai dati dell’app." : "Rifiutare l’accesso a questo utente?")) return;
    const reason = decision === "rifiutato" ? (window.prompt("Motivazione facoltativa:", "") || "") : "";
    const admin = auth.currentUser;
    const ref = db.collection("platformUsers").doc(uid);
    const audit = db.collection("userAccessAudit").doc();
    const patch = decision === "attivo"
      ? { statoAccount: "attivo", accountStatus: "attivo", role: "user", ruolo: "user", approvatoAt: serverTime(), approvedAt: serverTime(), approvatoDa: admin.uid, approvatoDaEmail: admin.email || "", approvedBy: admin.email || admin.uid, banned: false }
      : { statoAccount: "rifiutato", accountStatus: "rifiutato", rifiutatoAt: serverTime(), rejectedAt: serverTime(), rifiutatoDa: admin.uid, rejectedBy: admin.email || admin.uid, motivoRifiuto: reason };
    const batch = db.batch();
    batch.update(ref, patch);
    batch.set(audit, { userId: uid, action: decision === "attivo" ? "approvazione" : "rifiuto", reason, administratorUid: admin.uid, administratorEmail: admin.email || "", createdAt: serverTime() });
    await batch.commit();
    alert(decision === "attivo" ? "Utente sbloccato correttamente." : "Accesso rifiutato.");
  }

  function renderAdmin(users, isAdmin) {
    const container = el("pending-users-list");
    if (!container) return;
    const pending = isAdmin ? users.filter((item) => PENDING.has(statusOf(item))).sort((a, b) => (dateValue(b.ultimaRichiestaAt || b.primoAccessoAt)?.getTime() || 0) - (dateValue(a.ultimaRichiestaAt || a.primoAccessoAt)?.getTime() || 0)) : [];
    el("pending-users-section").classList.toggle("hidden", !isAdmin);
    const badge = el("pending-users-badge");
    badge.textContent = pending.length;
    badge.classList.toggle("hidden", !pending.length);
    const menuBadge = el("pending-users-menu-badge");
    menuBadge.textContent = pending.length;
    menuBadge.classList.toggle("hidden", !pending.length);
    container.innerHTML = pending.length ? pending.map((item) => `<article class="pending-user-card"><h4>${escape(item.nomeCompleto || item.displayName)}</h4><p>${escape(item.email)}</p><dl><dt>ID utente</dt><dd>${escape(item.uid || item.id)}</dd><dt>Primo accesso</dt><dd>${formatDateTime(item.primoAccessoAt || item.firstLoginAt)}</dd><dt>Richiesta WhatsApp</dt><dd>${formatDateTime(item.ultimaRichiestaAt || item.richiestaWhatsappAt)}</dd><dt>Stato</dt><dd>${escape(statusOf(item))}</dd><dt>Richieste inviate</dt><dd>${Number(item.numeroRichieste || item.requestCount || 0)}</dd></dl><div class="actions-row"><button class="btn btn-primary" data-approve-user="${escape(item.id)}">SBLOCCA UTENTE</button><button class="btn btn-danger" data-reject-user="${escape(item.id)}">RIFIUTA ACCESSO</button></div></article>`).join("") : "<p class='muted'>Nessun utente in attesa.</p>";
  }

  document.addEventListener("click", (event) => {
    const approve = event.target.closest("[data-approve-user]");
    const reject = event.target.closest("[data-reject-user]");
    if (approve) void decide(approve.dataset.approveUser, "attivo");
    if (reject) void decide(reject.dataset.rejectUser, "rifiutato");
  });
  el("access-request-whatsapp")?.addEventListener("click", requestWhatsapp);
  el("access-check-approval")?.addEventListener("click", refresh);
  el("access-approval-retry")?.addEventListener("click", refresh);
  el("access-approval-logout")?.addEventListener("click", async () => { stop(); await auth.signOut(); });

  window.HeraAccessApproval = { verify, renderAdmin, statusOf };
})();