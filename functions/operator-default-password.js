"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");

const DEFAULT_OPERATOR_PASSWORD = "12345678";
const ADMIN_EMAIL_FALLBACK = "ionut29019@gmail.com";

function text(value) {
  return String(value ?? "").trim();
}

async function requireAdmin(context) {
  const email = text(context?.auth?.token?.email).toLowerCase();
  if (!context?.auth?.uid) {
    throw new functions.https.HttpsError("unauthenticated", "Accesso amministratore richiesto.");
  }
  if (email === ADMIN_EMAIL_FALLBACK) return;

  const snap = await admin.firestore().collection("platformUsers").doc(context.auth.uid).get();
  const profile = snap.exists ? (snap.data() || {}) : {};
  const role = text(profile.role || profile.ruolo).toLowerCase();
  if (profile.isAdmin === true || profile.admin === true || ["admin", "administrator", "amministratore"].includes(role)) return;

  throw new functions.https.HttpsError("permission-denied", "Funzione riservata all’amministratore.");
}

exports.setOperatorDefaultPassword = functions.region("europe-west1").https.onCall(async (data, context) => {
  await requireAdmin(context);
  const uid = text(data?.uid);
  if (!uid) {
    throw new functions.https.HttpsError("invalid-argument", "UID operatore mancante.");
  }

  const user = await admin.auth().updateUser(uid, {
    password: DEFAULT_OPERATOR_PASSWORD,
    disabled: false
  });

  await admin.firestore().collection("platformUsers").doc(uid).set({
    defaultOperatorPasswordProvisioned: true,
    defaultOperatorPasswordProvisionedAt: admin.firestore.FieldValue.serverTimestamp(),
    mustChangePassword: false,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });

  return {
    ok: true,
    uid: user.uid,
    password: DEFAULT_OPERATOR_PASSWORD
  };
});
