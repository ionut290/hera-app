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

exports.loginWithOperatorUsername = functions
  .region("europe-west1")
  .https.onCall(async (data) => {
    const username = normalizeUsername(data?.username);
    const password = String(data?.password || "");

    if (!username || username.length > 100 || password.length < 6 || password.length > 256) {
      throw new functions.https.HttpsError(
        "invalid-argument",
        "Username o password non validi."
      );
    }

    const db = admin.firestore();
    const matches = await db
      .collection("personale")
      .where("loginUsername", "==", username)
      .limit(2)
      .get();

    if (matches.size !== 1) {
      throw new functions.https.HttpsError(
        "unauthenticated",
        "Username o password non corretti."
      );
    }

    const personnel = matches.docs[0].data() || {};
    const email = String(
      personnel.linkedUserEmail ||
      personnel.LINKED_USER_EMAIL ||
      personnel.emailAccessoApp ||
      personnel.EMAIL_ACCESSO_APP ||
      ""
    ).trim().toLowerCase();
    const expectedUid = String(
      personnel.linkedUserId || personnel.LINKED_USER_ID || ""
    ).trim();

    if (!validEmail(email)) {
      throw new functions.https.HttpsError(
        "unauthenticated",
        "Username o password non corretti."
      );
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
      throw new functions.https.HttpsError(
        "unavailable",
        "Servizio di accesso temporaneamente non disponibile."
      );
    }

    const payload = await response.json().catch(() => ({}));
    if (!response.ok || !payload?.localId) {
      throw new functions.https.HttpsError(
        "unauthenticated",
        "Username o password non corretti."
      );
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
  });
