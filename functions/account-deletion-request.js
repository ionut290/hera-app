"use strict";

// STAGED, NON-DESTRUCTIVE account deletion request intake.
// Must NOT be exported from functions/main.js or deployed without the documented
// owner/privacy approval and a verified manual fulfillment process.
const admin = require("firebase-admin");
const { onCall, HttpsError } = require("firebase-functions/v2/https");

const MAX_AUTH_AGE_SECONDS = 600;
const COLLECTION = "accountDeletionRequests";
const REVIEW_STATUSES = new Set(["pending", "in_review", "completed", "rejected"]);

function identity(request, requireFresh) {
  const uid = request?.auth?.uid;
  if (typeof uid !== "string" || !/^[A-Za-z0-9_-]{1,128}$/.test(uid)) {
    throw new HttpsError("unauthenticated", "Accedi per richiedere la cancellazione.");
  }
  if (requireFresh) {
    const authTime = Number(request.auth?.token?.auth_time);
    const now = Math.floor(Date.now() / 1000);
    if (!Number.isFinite(authTime) || authTime > now + 60 ||
        now - authTime > MAX_AUTH_AGE_SECONDS) {
      throw new HttpsError("failed-precondition",
        "Per sicurezza esci e accedi nuovamente prima di inviare la richiesta.");
    }
  }
  return uid;
}

exports.requestOwnVargaAccountDeletion = onCall({
  region: "europe-west1",
  invoker: "public", // Firebase bearer token is mandatory inside the handler
  maxInstances: 3,
  timeoutSeconds: 30
}, async (request) => {
  const uid = identity(request, true);
  const data = request.data;
  if (!data || typeof data !== "object" || Array.isArray(data) ||
      Object.keys(data).length !== 1 ||
      data.confirmation !== "RICHIEDO ELIMINAZIONE") {
    throw new HttpsError("invalid-argument", "Conferma esplicita mancante o richiesta non valida.");
  }

  // The path is ALWAYS derived from verified Auth UID. No user-supplied UID,
  // address, business identifier, file path, or private employee information.
  const ref = admin.firestore().collection(COLLECTION).doc(uid);
  try {
    const result = await admin.firestore().runTransaction(async (tx) => {
      const snapshot = await tx.get(ref);
      if (snapshot.exists) {
        // Never reset a completed, rejected or in-review request to pending.
        const status = snapshot.data()?.status;
        if (!REVIEW_STATUSES.has(status)) {
          throw new HttpsError("failed-precondition",
            "Richiesta esistente in stato non riconosciuto. Contatta l'assistenza.");
        }
        return { requested: true, status, alreadyExists: true };
      }
      tx.create(ref, {
        uid,
        status: "pending",
        submittedAt: admin.firestore.FieldValue.serverTimestamp(),
        origin: "authenticated_app",
        // No email, name, personnel documents, or deletion from this endpoint.
        reviewBeforeErasureRequired: true
      });
      return { requested: true, status: "pending", alreadyExists: false };
    });
    return result;
  } catch (error) {
    if (error instanceof HttpsError) throw error;
    console.error("Account deletion request intake failed", { code: error?.code || "unknown" });
    throw new HttpsError("internal", "Richiesta non confermata dal server. Riprova.");
  }
});

exports.getOwnVargaAccountDeletionStatus = onCall({
  region: "europe-west1",
  invoker: "public",
  maxInstances: 3,
  timeoutSeconds: 20
}, async (request) => {
  const uid = identity(request, false);
  if (request.data != null && (typeof request.data !== "object" ||
      Array.isArray(request.data) || Object.keys(request.data).length !== 0)) {
    throw new HttpsError("invalid-argument", "Richiesta non valida.");
  }
  try {
    const snapshot = await admin.firestore().collection(COLLECTION).doc(uid).get();
    if (!snapshot.exists) return { requested: false };
    const status = snapshot.data()?.status;
    if (!REVIEW_STATUSES.has(status)) {
      throw new HttpsError("failed-precondition", "Stato richiesta non riconosciuto.");
    }
    return { requested: true, status };
  } catch (error) {
    if (error instanceof HttpsError) throw error;
    console.error("Account deletion request status failed", { code: error?.code || "unknown" });
    throw new HttpsError("internal", "Stato non disponibile. Riprova.");
  }
});
