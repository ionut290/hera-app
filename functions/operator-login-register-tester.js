"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");

const FIREBASE_WEB_API_KEY = "AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k";

function normalizeUsername(value) {
  return String(value || "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, ".")
    .replace(/^\.+|\.+$/g, "")
    .replace(/\.{2,}/g, ".");
}

function validEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ""));
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
  const email = String(
    personnel.linkedUserEmail ||
    personnel.LINKED_USER_EMAIL ||
    personnel.emailAccessoApp ||
    personnel.EMAIL_ACCESSO_APP ||
    ""
  ).trim().toLowerCase();
  const expectedUid = String(personnel.linkedUserId || personnel.LINKED_USER_ID || "").trim();

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
    console.error("Login username: UID non coerente con il personale.", {
      username,
      expectedUid,
      authenticatedUid: payload.localId
    });
    throw new functions.https.HttpsError(
      "failed-precondition",
      "Collegamento account non valido. Contatta l’amministratore."
    );
  }

  const customToken = await admin.auth().createCustomToken(payload.localId);
  return { token: customToken };
}

async function registerNewUser(data) {
  const email = String(data?.email || "").trim().toLowerCase();
  const password = String(data?.temporaryPassword || "");
  const firstName = String(data?.firstName || "").trim().slice(0, 80);
  const lastName = String(data?.lastName || "").trim().slice(0, 80);
  const displayName = [firstName, lastName].filter(Boolean).join(" ");

  if (!validEmail(email)) {
    throw new functions.https.HttpsError("invalid-argument", "Indirizzo email non valido.");
  }
  if (!firstName || !lastName) {
    throw new functions.https.HttpsError("invalid-argument", "Nome e cognome sono obbligatori.");
  }
  if (password.length < 10) {
    throw new functions.https.HttpsError(
      "invalid-argument",
      "La password deve contenere almeno 10 caratteri."
    );
  }

  try {
    await admin.auth().getUserByEmail(email);
    throw new functions.https.HttpsError("already-exists", "Account già esistente.");
  } catch (error) {
    if (error instanceof functions.https.HttpsError) throw error;
    if (error.code !== "auth/user-not-found") {
      console.error("Verifica account per registrazione fallita.", {
        code: error?.code || "",
        message: error?.message || ""
      });
      throw new functions.https.HttpsError(
        "internal",
        "Impossibile verificare l’indirizzo email. Riprova tra poco."
      );
    }
  }

  let user = null;
  try {
    user = await admin.auth().createUser({
      email,
      password,
      displayName,
      emailVerified: false,
      disabled: false
    });
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
    if (error?.code === "auth/email-already-exists") {
      throw new functions.https.HttpsError("already-exists", "Account già esistente.");
    }
    if (["auth/invalid-password", "auth/weak-password"].includes(error?.code)) {
      throw new functions.https.HttpsError(
        "invalid-argument",
        "La password deve contenere almeno 10 caratteri."
      );
    }

    if (user?.uid) {
      try {
        await admin.auth().deleteUser(user.uid);
      } catch (cleanupError) {
        console.error("Pulizia account incompleto fallita.", {
          uid: user.uid,
          code: cleanupError?.code || "",
          message: cleanupError?.message || ""
        });
      }
    }

    console.error("Creazione account fallita.", {
      code: error?.code || "",
      message: error?.message || ""
    });
    throw new functions.https.HttpsError(
      "internal",
      "Creazione account non riuscita. Riprova tra poco."
    );
  }
}

exports.registerTester = functions.region("europe-west1").https.onCall(async (data) => {
  if (String(data?.action || "") === "loginUsername") {
    return loginWithUsername(data);
  }
  return registerNewUser(data);
});
