(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.HeraManagementCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const PERSONNEL_COLUMNS = ["ID_OPERATORE","CODICE_OPERATORE","NOME","COGNOME","CODICE_FISCALE","DATA_NASCITA","LUOGO_NASCITA","TELEFONO","EMAIL","RUOLO","MANSIONE","LIVELLO","TIPO_CONTRATTO","DATA_ASSUNZIONE","DATA_FINE_CONTRATTO","MATRICOLA","AZIENDA","SEDE","RESPONSABILE","STATO","SQUADRA","COMMESSE","PATENTE","ABILITAZIONI","DATA_SCADENZA_PATENTE","DATA_SCADENZA_VISITA_MEDICA","DATA_SCADENZA_CONTRATTO","CONTATTO_EMERGENZA","TELEFONO_EMERGENZA","NOTE","ATTIVO","EMAIL_ACCESSO_APP","RUOLO_APP","CONSENTI_ACCESSO_APP","COMMESSE_ABILITATE","MODALITA_AGGIORNAMENTO_COMMESSE","MANTIENI_ABILITAZIONI_ESISTENTI","LINKED_USER_ID","LINKED_USER_EMAIL","PROFILO_COLLEGATO","FONTE_FOTO_PROFILO"];
  const VEHICLE_COLUMNS = ["ID_MEZZO","CODICE_MEZZO","TARGA","CATEGORIA","MARCA","MODELLO","ANNO_IMMATRICOLAZIONE","NUMERO_TELAIO","ALIMENTAZIONE","CILINDRATA","POTENZA","CHILOMETRI","ORE_MOTORE","PORTATA","PESO","NUMERO_POSTI","AZIENDA_PROPRIETARIA","STATO","SQUADRA_ASSEGNATA","OPERATORE_RESPONSABILE","COMMESSE","LUOGO_DEPOSITO","DATA_REVISIONE","DATA_SCADENZA_ASSICURAZIONE","DATA_SCADENZA_BOLLO","DATA_PROSSIMA_MANUTENZIONE","DATA_SCADENZA_TARATURA","DATA_INIZIO_NOLEGGIO","DATA_FINE_NOLEGGIO","FORNITORE_NOLEGGIO","NUMERO_POLIZZA","COMPAGNIA_ASSICURATIVA","NOTE","ATTIVO","FOTO_URL","DOCUMENTI_URL"];
  const PERSON_STATUSES = ["Attivo","Assente","Ferie","Malattia","Permesso","Infortunio","Riposo","Sospeso","Non disponibile","Cessato"];
  const VEHICLE_STATUSES = ["Disponibile","Assegnato","In uso","In manutenzione","Guasto","Fuori servizio","Noleggiato","Restituito","Dismesso"];
  const VEHICLE_CATEGORIES = ["Camion","Furgone","Autovettura","Trattore grande","Trattore piccolo","Escavatore","Pala","Minipala","Decespugliatore","Motosega","Soffiatore","Trincia","Spazzatrice","Piattaforma","Rimorchio","Attrezzatura","Altro"];
  const MODES = ["AGGIUNGI","SOSTITUISCI","RIMUOVI","NON_MODIFICARE"];
  const clean = (value) => String(value == null ? "" : value).trim().replace(/\s+/g, " ");
  const key = (value) => clean(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  const unique = (values) => [...new Set((values || []).map(clean).filter(Boolean))];
  const splitIds = (value) => unique(String(value || "").split(";").map(clean));
  const legacyEnabledIds = (person, resolve) => {
    if (Array.isArray(person?.enabledCommessaIds)) return unique(person.enabledCommessaIds);
    const raw = [person?.authorizedCommessaIds, person?.assignedCommessaIds, person?.commessaIds, person?.commesseAbilitate].flatMap((v) => Array.isArray(v) ? v : splitIds(v));
    return unique(raw.map((v) => resolve ? resolve(v) : v).filter(Boolean));
  };
  function applyEnabledMode(current, requested, mode) {
    const before = unique(current), values = unique(requested), normalizedMode = key(mode) || "NON_MODIFICARE";
    if (normalizedMode === "NON_MODIFICARE") return before;
    if (!MODES.includes(normalizedMode)) throw new Error("Modalità di aggiornamento commesse non valida.");
    if (normalizedMode === "AGGIUNGI") return unique([...before, ...values]);
    if (normalizedMode === "RIMUOVI") return before.filter((id) => !values.includes(id));
    return values;
  }
  function identify(row, records, type) {
    const fields = type === "personale" ? [["ID_OPERATORE","id"],["CODICE_OPERATORE","codiceOperatore"],["EMAIL_ACCESSO_APP","emailAccessoApp"]] : [["ID_MEZZO","id"],["CODICE_MEZZO","codiceMezzo"],["TARGA","targa"]];
    for (const [column, field] of fields) {
      const value = key(row[column]);
      if (!value) continue;
      const matches = records.filter((record) => key(record[field] || (field === "codiceMezzo" ? record.nId : field === "emailAccessoApp" ? record.email : "")) === value);
      if (matches.length > 1) return { duplicate: true, field: column };
      if (matches.length === 1) return { record: matches[0], field: column };
    }
    return {};
  }
  function validateRow(row, type, existing = [], resolveCommessa = () => "") {
    const errors = [], warnings = [], isPerson = type === "personale";
    const match = identify(row, existing, type);
    if (match.duplicate) errors.push(`Identificativo duplicato: ${match.field}.`);
    const required = isPerson ? ["NOME","COGNOME","STATO"] : ["CODICE_MEZZO","CATEGORIA","STATO"];
    if (!match.record) required.forEach((field) => { if (!clean(row[field])) errors.push(`Campo obbligatorio mancante: ${field}.`); });
    const statuses = isPerson ? PERSON_STATUSES : VEHICLE_STATUSES;
    if (clean(row.STATO) && !statuses.some((v) => key(v) === key(row.STATO))) errors.push("Stato non riconosciuto.");
    if (!isPerson && clean(row.CATEGORIA) && !VEHICLE_CATEGORIES.some((v) => key(v) === key(row.CATEGORIA))) errors.push("Categoria non riconosciuta.");
    if (isPerson && clean(row.EMAIL_ACCESSO_APP) && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean(row.EMAIL_ACCESSO_APP))) errors.push("Email di accesso non valida.");
    if (isPerson && clean(row.MODALITA_AGGIORNAMENTO_COMMESSE) && !MODES.includes(key(row.MODALITA_AGGIORNAMENTO_COMMESSE))) errors.push("Modalità di aggiornamento commesse non valida.");
    const requested = splitIds(row[isPerson ? "COMMESSE_ABILITATE" : "COMMESSE"]);
    const ids = requested.map(resolveCommessa);
    requested.forEach((value, index) => { if (!ids[index]) errors.push(`La commessa “${value}” non esiste.`); });
    const mode = isPerson ? (key(row.MODALITA_AGGIORNAMENTO_COMMESSE) || "NON_MODIFICARE") : "NON_MODIFICARE";
    const current = isPerson && match.record ? legacyEnabledIds(match.record, resolveCommessa) : [];
    const finalIds = isPerson ? applyEnabledMode(current, ids.filter(Boolean), mode) : [];
    const added = finalIds.filter((id) => !current.includes(id)), removed = current.filter((id) => !finalIds.includes(id));
    if (removed.length) warnings.push(`Saranno rimosse ${removed.length} abilitazioni.`);
    return { row, existing: match.record || null, errors, warnings, current, added, removed, finalIds, action: errors.length ? "Errore" : match.record ? "Aggiorna" : "Nuovo" };
  }
  function normalizeAppUser(user = {}) {
    const providerData = Array.isArray(user.providerData) ? user.providerData : [];
    const google = providerData.find((item) => item?.providerId === "google.com") || {};
    return {
      uid: clean(user.uid || user.id),
      displayName: clean(user.displayName || google.displayName),
      email: clean(user.email || google.email).toLowerCase(),
      photoURL: clean(user.photoURL || google.photoURL),
      providerId: clean(user.providerId || google.providerId || user.linkedAuthProvider || "google.com"),
      emailVerified: user.emailVerified === true,
      role: clean(user.role || user.ruolo || "operatore"),
      accountStatus: clean(user.statoAccount || user.accountStatus),
      firstLoginAt: user.firstLoginAt || user.createdAt || null,
      lastLoginAt: user.lastLoginAt || user.lastSeenAt || null
    };
  }
  // Il nome Google è solo una proposta modificabile: con più parole non si
  // presume quale sia il cognome, ma si presenta una suddivisione iniziale.
  function proposeDisplayName(displayName) {
    const parts = clean(displayName).split(" ").filter(Boolean);
    return { nome: parts.shift() || "", cognome: parts.join(" "), ambiguous: parts.length > 1 };
  }
  function findUserLinkConflicts(user, personnel = [], currentOperatorId = "") {
    const profile = normalizeAppUser(user);
    const linked = personnel.filter((p) => p.id !== currentOperatorId && clean(p.linkedUserId) === profile.uid);
    const emailMatches = personnel.filter((p) => p.id !== currentOperatorId && profile.email && [p.email, p.emailAccessoApp, p.linkedUserEmail].some((v) => clean(v).toLowerCase() === profile.email));
    return { linked, emailMatches, duplicateUid: linked.length > 0, ambiguousEmail: emailMatches.length > 1 };
  }
  function buildGoogleLinkPatch(operator = {}, user = {}, options = {}) {
    const profile = normalizeAppUser(user);
    if (!profile.uid) throw new Error("L’utente selezionato non possiede un UID valido.");
    const proposed = proposeDisplayName(profile.displayName);
    const manualPhoto = clean(operator.profilePhotoSource).toLowerCase() === "manual";
    const patch = {
      linkedUserId: profile.uid,
      linkedUserEmail: profile.email,
      linkedAuthProvider: profile.providerId,
      profileSource: "google",
      googleDisplayName: profile.displayName,
      emailVerified: profile.emailVerified
    };
    if (!clean(operator.nome)) patch.nome = clean(options.nome ?? proposed.nome);
    if (!clean(operator.cognome)) patch.cognome = clean(options.cognome ?? proposed.cognome);
    if (!clean(operator.email)) patch.email = profile.email;
    if (!clean(operator.emailAccessoApp)) patch.emailAccessoApp = profile.email;
    if (!manualPhoto && profile.photoURL) {
      patch.photoURL = profile.photoURL;
      patch.profilePhotoSource = "google";
    }
    return patch;
  }
  function buildGoogleLoginSyncPatch(operator = {}, user = {}) {
    const profile = normalizeAppUser(user);
    const patch = { emailVerified: profile.emailVerified, googleDisplayName: profile.displayName, linkedAuthProvider: profile.providerId };
    if (clean(operator.profilePhotoSource).toLowerCase() === "google" && profile.photoURL) patch.photoURL = profile.photoURL;
    return patch;
  }
  return { PERSONNEL_COLUMNS, VEHICLE_COLUMNS, PERSON_STATUSES, VEHICLE_STATUSES, VEHICLE_CATEGORIES, MODES, clean, key, unique, splitIds, legacyEnabledIds, applyEnabledMode, identify, validateRow, normalizeAppUser, proposeDisplayName, findUserLinkConflicts, buildGoogleLinkPatch, buildGoogleLoginSyncPatch };
});
