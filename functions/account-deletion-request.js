"use strict";

// STAGED, NON-DESTRUCTIVE account deletion request intake.
// Must NOT be exported from functions/main.js or deployed without the documented
// owner/privacy approval and a verified manual fulfillment process.
const admin = require("firebase-admin");
const { onCall, HttpsError } = require("firebase-functions/v2/https");

const MAX_AUTH_AGE_SECONDS = 600;
const COLLECTION = "accountDeletionRequests";
const REVIEW_STATUSES = new Set(["pending", "in_review", "completed", "rejected"]);
// Only these fixed, reviewed codes may be disclosed to the account owner.
// Never return free-text operator notes, business data or arbitrary fields.
const SAFE_REJECTION_REASONS = new Set([
  "identity_not_verified",
  "retention_review_required",
  "request_needs_clarification"
]);

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

// Firebase verifies callable tokens, but a signed token can remain usable for a
// short period after the user is disabled or deleted. Never accept a request
// or expose its status unless the same UID still has an active Auth record.
// This remains a STAGED, NON-DESTRUCTIVE service; it never revokes tokens.
async function requireActiveAccount(uid, token) {
  let user;
  try {
    user = await admin.auth().getUser(uid);
  } catch (error) {
    if (error?.code === "auth/user-not-found") {
      throw new HttpsError("unauthenticated", "Sessione non valida. Accedi nuovamente.");
    }
    console.error("Deletion request Auth verification failed", { code: error?.code || "unknown" });
    throw new HttpsError("internal", "Verifica dell'account non disponibile. Riprova.");
  }
  if (!user || user.disabled) {
    throw new HttpsError("unauthenticated", "Sessione non valida. Accedi nuovamente.");
  }

  // Firebase callable ID-token validation alone does not check revocation.
  // Even a still-valid token must not submit or READ a deletion request after
  // a password reset, explicit sign-out of all devices or token revocation.
  // The live Auth record has the timestamp after which tokens are valid.
  // Fail closed if either timestamp is missing or unparseable.
  const authTime = Number(token?.auth_time);
  const validAfterText = user.tokensValidAfterTime;
  const validAfter = typeof validAfterText === "string" ? Date.parse(validAfterText) : NaN;
  const now = Math.floor(Date.now() / 1000);
  if (!Number.isSafeInteger(authTime) || authTime <= 0 || authTime > now + 60) {
    throw new HttpsError("unauthenticated", "Sessione scaduta. Accedi nuovamente.");
  }
  if (!Number.isFinite(validAfter)) {
    throw new HttpsError("internal", "Verifica della sessione non disponibile. Riprova.");
  }
  if (authTime * 1000 < validAfter) {
    throw new HttpsError("unauthenticated", "Sessione revocata. Accedi nuovamente.");
  }
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
  await requireActiveAccount(uid, request.auth?.token);

  // The path is ALWAYS derived from verified Auth UID. No user-supplied UID,
  // address, business identifier, file path, or private employee information.
  const ref = admin.firestore().collection(COLLECTION).doc(uid);
  try {
    const result = await admin.firestore().runTransaction(async (tx) => {
      const snapshot = await tx.get(ref);
      if (snapshot.exists) {
        // Never reset a completed, rejected or in-review request to pending.
        const previous = snapshot.data() || {};
        const status = previous.status;
        if (!REVIEW_STATUSES.has(status)) {
          throw new HttpsError("failed-precondition",
            "Richiesta esistente in stato non riconosciuto. Contatta l'assistenza.");
        }
        if (status === "rejected") {
          // Rejection must not permanently prevent the account owner from
          // asking again. Preserve any previous reviewer notes/disposition
          // instead of erasing the document, and audit each renewed request.
          // A fresh Firebase Auth sign-in was already required above.
          const count = previous.reopenCount === undefined ? 0 : previous.reopenCount;
          if (!Number.isSafeInteger(count) || count < 0 || count >= 10000) {
            throw new HttpsError("failed-precondition",
              "Cronologia richieste non valida. Contatta l'assistenza.");
          }
          const now = admin.firestore.FieldValue.serverTimestamp();
          tx.update(ref, {
            status: "pending",
            submittedAt: now,
            reopenedAt: now,
            reopenCount: count + 1,
            reviewBeforeErasureRequired: true
          });
          return { requested: true, status: "pending", alreadyExists: true, reopened: true };
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
  await requireActiveAccount(uid, request.auth?.token);
  try {
    const snapshot = await admin.firestore().collection(COLLECTION).doc(uid).get();
    if (!snapshot.exists) return { requested: false };
    const status = snapshot.data()?.status;
    if (!REVIEW_STATUSES.has(status)) {
      throw new HttpsError("failed-precondition", "Stato richiesta non riconosciuto.");
    }
    const response = { requested: true, status };
    if (status === "rejected") {
      const reason = snapshot.data()?.rejectionReasonCode;
      if (SAFE_REJECTION_REASONS.has(reason)) {
        response.reasonCode = reason;
      }
    }
    return response;
  } catch (error) {
    if (error instanceof HttpsError) throw error;
    console.error("Account deletion request status failed", { code: error?.code || "unknown" });
    throw new HttpsError("internal", "Stato non disponibile. Riprova.");
  }
});
