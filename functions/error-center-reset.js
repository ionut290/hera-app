"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");

const REGION = "europe-west1";
const ADMIN_EMAIL = "ionut29019@gmail.com";
const GROUPS_COLLECTION = "appErrorGroups";
const SUMMARY_COLLECTION = "systemCounters";
const SUMMARY_DOCUMENT = "errorCenterSummary";
const PUBLIC_CALLABLE_OPTIONS = Object.freeze({ invoker: "public" });
const DELETE_BATCH_SIZE = 400;

function db() {
  return admin.firestore();
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

async function isAdminContext(context) {
  if (!context.auth?.uid) return false;
  const email = normalizeEmail(context.auth.token?.email);
  if (email === ADMIN_EMAIL || context.auth.token?.admin === true || context.auth.token?.isAdmin === true) return true;
  try {
    const snapshot = await db().collection("appConfig").doc("adminUsers").get();
    const emails = Array.isArray(snapshot.data()?.emails) ? snapshot.data().emails.map(normalizeEmail) : [];
    return emails.includes(email);
  } catch (_) {
    return false;
  }
}

async function requireAdmin(context) {
  if (!context.auth?.uid) throw new functions.https.HttpsError("unauthenticated", "Accesso richiesto.");
  if (!(await isAdminContext(context))) throw new functions.https.HttpsError("permission-denied", "Funzione riservata all'amministratore.");
}

async function deleteAllErrorGroups() {
  let deleted = 0;
  while (true) {
    const snapshot = await db().collection(GROUPS_COLLECTION).limit(DELETE_BATCH_SIZE).get();
    if (snapshot.empty) break;
    const batch = db().batch();
    snapshot.docs.forEach((doc) => batch.delete(doc.ref));
    await batch.commit();
    deleted += snapshot.size;
    if (snapshot.size < DELETE_BATCH_SIZE) break;
  }
  return deleted;
}

exports.resetErrorCenter = functions.region(REGION).runWith(PUBLIC_CALLABLE_OPTIONS).https.onCall(async (_data, context) => {
  await requireAdmin(context);

  const deletedGroups = await deleteAllErrorGroups();
  await db().collection(SUMMARY_COLLECTION).doc(SUMMARY_DOCUMENT).set({
    totalEvents: 0,
    totalGroups: 0,
    unseenAlerts: 0,
    lastErrorAt: null,
    lastGroupId: "",
    lastSeverity: "info",
    lastTitle: "",
    resetAt: admin.firestore.FieldValue.serverTimestamp(),
    resetByUid: context.auth.uid,
    resetByEmail: normalizeEmail(context.auth.token?.email),
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: false });

  return { reset: true, deletedGroups };
});
