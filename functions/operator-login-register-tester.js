"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");

const FIREBASE_WEB_API_KEY = "AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k";
const INTERNAL_LOGIN_DOMAIN = "operatori.vargacantieri.app";
const ACCESS_REQUESTS_COLLECTION = "accessRequests";
const ADMIN_EMAIL_FALLBACK = "ionut29019@gmail.com";

function text(value) {
  return String(value ?? "").trim();
}

function normalizeUsername(value) {
  return text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, ".")
    .replace(/^\.+|\.+$/g, "")
    .replace(/\.{2,}/g, ".");
}

function normalizeName(value) {
  return text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function validEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ""));
}

function firstValue(data, keys) {
  for (const key of keys) {
    const value = text(data?.[key]);
    if (value) return value;
  }
  return "";
}

function getFirstName(data) {
  return firstValue(data, ["nome", "name", "firstName", "NOME"]);
}

function getLastName(data) {
  return firstValue(data, ["cognome", "surname", "lastName", "COGNOME"]);
}

function getDisplayName(data) {
  return firstValue(data, ["nomeCompleto", "fullName", "displayName", "nominativo", "operatore"])
    || [getFirstName(data), getLastName(data)].filter(Boolean).join(" ")
    || "Operatore";
}

function getPersonnelEmail(data) {
  return text(
    data?.linkedUserEmail ||
    data?.LINKED_USER_EMAIL ||
    data?.emailAccessoApp ||
    data?.EMAIL_ACCESSO_APP ||
    ""
  ).toLowerCase();
}

function getPersonnelUid(data) {
  return text(data?.linkedUserId || data?.LINKED_USER_ID || "");
}

function randomPassword(length = 14) {
  const alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%";
  let out = "";
  for (let i = 0; i < length; i += 1) out += alphabet[Math.floor(Math.random() * alphabet.length)];
  return out;
}

async function requireAdmin(context) {
  const email = text(context?.auth?.token?.email).toLowerCase();
  if (!context?.auth?.uid) throw new functions.https.HttpsError("unauthenticated", "Accesso amministratore richiesto.");
  if (email === ADMIN_EMAIL_FALLBACK) return true;
  const profile = await admin.firestore().collection("platformUsers").doc(context.auth.uid).get();
  const data = profile.exists ? (profile.data() || {}) : {};
  if (data.isAdmin === true || data.admin === true || ["admin", "administrator", "amministratore"].includes(text(data.role || data.ruolo).toLowerCase())) return true;
  throw new functions.https.HttpsError("permission-denied", "Funzione riservata all’amministratore.");
}

async function loginWithUsername(data) {
  const username = normalizeUsername(data?.username);
  const password = String(data?.password || "");

  if (!username || username.length > 100 || password.length < 6 || password.length > 256) {
    throw new functions.https.HttpsError("invalid-argument", "Username o password non validi.");
  }

  const db = admin.firestore();
  const matches = await db.collection("personale")
    .where("loginUsername", "==", username)
    .limit(2)
    .get();

  if (matches.size !== 1) {
    throw new functions.https.HttpsError("unauthenticated", "Username o password non corretti.");
  }

  const personnel = matches.docs[0].data() || {};
  const email = getPersonnelEmail(personnel);
  const expectedUid = getPersonnelUid(personnel);

  if (!validEmail(email)) {
    throw new functions.https.HttpsError("unauthenticated", "Username o password non corretti.");
  }

  let response;
  try {
    response = await fetch(
      `https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${encodeURIComponent(FIREBASE_WEB_API_KEY)}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, password, returnSecureToken: true })
      }
    );
  } catch (error) {
    console.error("Login username: Identity Toolkit non raggiungibile.", {
      code: error?.code || "",
      message: error?.message || ""
    });
    throw new functions.https.HttpsError("unavailable", "Servizio di accesso temporaneamente non disponibile.");
  }

  const payload = await response.json().catch(() => ({}));
  if (!response.ok || !payload?.localId) {
    throw new functions.https.HttpsError("unauthenticated", "Username o password non corretti.");
  }

  if (expectedUid && payload.localId !== expectedUid) {
    throw new functions.https.HttpsError("failed-precondition", "Collegamento account non valido. Contatta l’amministratore.");
  }

  const customToken = await admin.auth().createCustomToken(payload.localId);
  return { token: customToken };
}

async function findPersonnelByName(firstName, lastName) {
  const wantedFirst = normalizeName(firstName);
  const wantedLast = normalizeName(lastName);
  const snapshot = await admin.firestore().collection("personale").get();
  return snapshot.docs.filter((doc) => {
    const data = doc.data() || {};
    const directFirst = normalizeName(getFirstName(data));
    const directLast = normalizeName(getLastName(data));
    if (directFirst && directLast) return directFirst === wantedFirst && directLast === wantedLast;
    return normalizeName(getDisplayName(data)) === normalizeName(`${firstName} ${lastName}`);
  });
}

async function createAccessRequest(firstName, lastName, personnelId = "", reason = "not-found") {
  const db = admin.firestore();
  const duplicate = await db.collection(ACCESS_REQUESTS_COLLECTION)
    .where("normalizedFullName", "==", normalizeName(`${firstName} ${lastName}`))
    .where("status", "==", "pending")
    .limit(1)
    .get();
  if (!duplicate.empty) return duplicate.docs[0].id;
  const ref = await db.collection(ACCESS_REQUESTS_COLLECTION).add({
    firstName: text(firstName).slice(0, 80),
    lastName: text(lastName).slice(0, 80),
    displayName: `${text(firstName)} ${text(lastName)}`.trim(),
    normalizedFullName: normalizeName(`${firstName} ${lastName}`),
    personnelId: text(personnelId),
    reason,
    status: "pending",
    createdAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  });
  return ref.id;
}

async function requestAccessLookup(data) {
  const firstName = text(data?.firstName).slice(0, 80);
  const lastName = text(data?.lastName).slice(0, 80);
  if (!firstName || !lastName) throw new functions.https.HttpsError("invalid-argument", "Nome e cognome sono obbligatori.");

  const matches = await findPersonnelByName(firstName, lastName);
  if (matches.length === 1) {
    const personnel = matches[0].data() || {};
    const username = normalizeUsername(personnel.loginUsername || personnel.LOGIN_USERNAME || "");
    if (username) {
      return {
        status: "found",
        found: true,
        username,
        displayName: getDisplayName(personnel)
      };
    }
    const requestId = await createAccessRequest(firstName, lastName, matches[0].id, "existing-needs-account");
    return { status: "pending", found: true, requestId, existingPersonnel: true };
  }

  const reason = matches.length > 1 ? "ambiguous-name" : "not-found";
  const requestId = await createAccessRequest(firstName, lastName, "", reason);
  return { status: "pending", found: false, requestId };
}

async function uniqueUsername(base, excludePersonnelId = "") {
  const normalizedBase = normalizeUsername(base) || "operatore";
  const db = admin.firestore();
  let candidate = normalizedBase;
  for (let suffix = 1; suffix < 1000; suffix += 1) {
    const matches = await db.collection("personale").where("loginUsername", "==", candidate).limit(2).get();
    const conflict = matches.docs.some((doc) => doc.id !== excludePersonnelId);
    if (!conflict) return candidate;
    candidate = `${normalizedBase}.${suffix + 1}`;
  }
  throw new functions.https.HttpsError("resource-exhausted", "Impossibile generare uno username univoco.");
}

async function provisionPersonnel(personnelDoc, firstName, lastName) {
  const db = admin.firestore();
  const existing = personnelDoc?.exists ? (personnelDoc.data() || {}) : {};
  const effectiveFirst = getFirstName(existing) || firstName;
  const effectiveLast = getLastName(existing) || lastName;
  const displayName = getDisplayName(existing) !== "Operatore" ? getDisplayName(existing) : `${effectiveFirst} ${effectiveLast}`.trim();
  const username = await uniqueUsername(existing.loginUsername || displayName, personnelDoc?.id || "");
  const loginEmail = validEmail(getPersonnelEmail(existing)) ? getPersonnelEmail(existing) : `${username}@${INTERNAL_LOGIN_DOMAIN}`;
  const temporaryPassword = randomPassword();
  const linkedUid = getPersonnelUid(existing);
  let user = null;

  if (linkedUid) {
    try { user = await admin.auth().getUser(linkedUid); } catch (error) { if (error?.code !== "auth/user-not-found") throw error; }
  }
  if (!user) {
    try { user = await admin.auth().getUserByEmail(loginEmail); } catch (error) { if (error?.code !== "auth/user-not-found") throw error; }
  }
  if (user) {
    user = await admin.auth().updateUser(user.uid, { password: temporaryPassword, displayName, disabled: false });
  } else {
    user = await admin.auth().createUser({ email: loginEmail, password: temporaryPassword, displayName, emailVerified: true, disabled: false });
  }

  let ref = personnelDoc?.ref || null;
  if (!ref) ref = db.collection("personale").doc();
  await ref.set({
    nome: effectiveFirst,
    cognome: effectiveLast,
    nomeCompleto: displayName,
    displayName,
    linkedUserId: user.uid,
    LINKED_USER_ID: user.uid,
    linkedUserEmail: loginEmail,
    LINKED_USER_EMAIL: loginEmail,
    emailAccessoApp: loginEmail,
    EMAIL_ACCESSO_APP: loginEmail,
    loginUsername: username,
    generatedInternalLogin: loginEmail.endsWith(`@${INTERNAL_LOGIN_DOMAIN}`),
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });

  await db.collection("platformUsers").doc(user.uid).set({
    uid: user.uid,
    email: loginEmail,
    displayName,
    firstName: effectiveFirst,
    lastName: effectiveLast,
    loginUsername: username,
    statoAccount: "attivo",
    accountStatus: "attivo",
    role: "user",
    ruolo: "user",
    isAdmin: false,
    admin: false,
    mustChangePassword: false,
    generatedInternalLogin: loginEmail.endsWith(`@${INTERNAL_LOGIN_DOMAIN}`),
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });

  return { personnelId: ref.id, uid: user.uid, username, email: loginEmail, temporaryPassword, displayName };
}

async function listAccessRequests(_data, context) {
  await requireAdmin(context);
  const snapshot = await admin.firestore().collection(ACCESS_REQUESTS_COLLECTION)
    .where("status", "==", "pending")
    .limit(50)
    .get();
  return {
    requests: snapshot.docs.map((doc) => ({ id: doc.id, ...(doc.data() || {}) }))
  };
}

async function approveAccessRequest(data, context) {
  await requireAdmin(context);
  const requestId = text(data?.requestId);
  if (!requestId) throw new functions.https.HttpsError("invalid-argument", "Richiesta non valida.");
  const db = admin.firestore();
  const requestRef = db.collection(ACCESS_REQUESTS_COLLECTION).doc(requestId);
  const requestSnap = await requestRef.get();
  if (!requestSnap.exists) throw new functions.https.HttpsError("not-found", "Richiesta non trovata.");
  const request = requestSnap.data() || {};
  if (request.status !== "pending") throw new functions.https.HttpsError("failed-precondition", "Richiesta già gestita.");

  let personnelDoc = null;
  if (text(request.personnelId)) {
    const snap = await db.collection("personale").doc(text(request.personnelId)).get();
    if (snap.exists) personnelDoc = snap;
  }
  const credentials = await provisionPersonnel(personnelDoc, request.firstName, request.lastName);
  await requestRef.set({
    status: "approved",
    approvedAt: admin.firestore.FieldValue.serverTimestamp(),
    approvedBy: context.auth.uid,
    approvedByEmail: text(context.auth.token.email).toLowerCase(),
    personnelId: credentials.personnelId,
    username: credentials.username,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  return { approved: true, credentials };
}

async function rejectAccessRequest(data, context) {
  await requireAdmin(context);
  const requestId = text(data?.requestId);
  if (!requestId) throw new functions.https.HttpsError("invalid-argument", "Richiesta non valida.");
  const ref = admin.firestore().collection(ACCESS_REQUESTS_COLLECTION).doc(requestId);
  const snap = await ref.get();
  if (!snap.exists) throw new functions.https.HttpsError("not-found", "Richiesta non trovata.");
  if ((snap.data() || {}).status !== "pending") throw new functions.https.HttpsError("failed-precondition", "Richiesta già gestita.");
  await ref.set({
    status: "rejected",
    rejectedAt: admin.firestore.FieldValue.serverTimestamp(),
    rejectedBy: context.auth.uid,
    rejectedByEmail: text(context.auth.token.email).toLowerCase(),
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  return { rejected: true };
}

async function registerNewUser(data) {
  const email = String(data?.email || "").trim().toLowerCase();
  const password = String(data?.temporaryPassword || "");
  const firstName = String(data?.firstName || "").trim().slice(0, 80);
  const lastName = String(data?.lastName || "").trim().slice(0, 80);
  const displayName = [firstName, lastName].filter(Boolean).join(" ");

  if (!validEmail(email)) throw new functions.https.HttpsError("invalid-argument", "Indirizzo email non valido.");
  if (!firstName || !lastName) throw new functions.https.HttpsError("invalid-argument", "Nome e cognome sono obbligatori.");
  if (password.length < 10) throw new functions.https.HttpsError("invalid-argument", "La password deve contenere almeno 10 caratteri.");

  try {
    await admin.auth().getUserByEmail(email);
    throw new functions.https.HttpsError("already-exists", "Account già esistente.");
  } catch (error) {
    if (error instanceof functions.https.HttpsError) throw error;
    if (error.code !== "auth/user-not-found") throw new functions.https.HttpsError("internal", "Impossibile verificare l’indirizzo email. Riprova tra poco.");
  }

  let user = null;
  try {
    user = await admin.auth().createUser({ email, password, displayName, emailVerified: false, disabled: false });
    await admin.firestore().collection("platformUsers").doc(user.uid).set({
      uid: user.uid,
      email,
      displayName,
      firstName,
      lastName,
      role: "user",
      ruolo: "user",
      isAdmin: false,
      admin: false,
      banned: false,
      mustChangePassword: false,
      selfRegistered: true,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });
    return { created: true };
  } catch (error) {
    if (error?.code === "auth/email-already-exists") throw new functions.https.HttpsError("already-exists", "Account già esistente.");
    if (["auth/invalid-password", "auth/weak-password"].includes(error?.code)) throw new functions.https.HttpsError("invalid-argument", "La password deve contenere almeno 10 caratteri.");
    if (user?.uid) {
      try { await admin.auth().deleteUser(user.uid); } catch (_) {}
    }
    throw new functions.https.HttpsError("internal", "Creazione account non riuscita. Riprova tra poco.");
  }
}

exports.registerTester = functions.region("europe-west1").https.onCall(async (data, context) => {
  const action = String(data?.action || "");
  if (action === "loginUsername") return loginWithUsername(data);
  if (action === "requestAccessLookup") return requestAccessLookup(data);
  if (action === "listAccessRequests") return listAccessRequests(data, context);
  if (action === "approveAccessRequest") return approveAccessRequest(data, context);
  if (action === "rejectAccessRequest") return rejectAccessRequest(data, context);
  return registerNewUser(data);
});
