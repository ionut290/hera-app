"use strict";

/**
 * STAGED, NOT DEPLOYED, privacy-request operator queue. Review does not erase
 * accounts. It deliberately cannot mark a request "completed": only a
 * separately verified deletion/retention process may do that.
 *
 * Do not import into functions/main.js or any routine production workflow.
 * Requires BOTH dedicated reviewer and admin custom claims on an active Auth
 * user record. A stale sign-in token or revoked live claim always fails closed.
 */
const admin = require("firebase-admin");
const { onCall, HttpsError } = require("firebase-functions/v2/https");

const REGION = "europe-west1";
const COLLECTION = "accountDeletionRequests";
const MAX_AUTH_AGE_SECONDS = 600;
const MAX_QUEUE_RESULTS = 25;
const REASONS = new Set([
  "identity_not_verified",
  "retention_review_required",
  "request_needs_clarification"
]);

function validUid(uid) {
  return typeof uid === "string" && /^[A-Za-z0-9_-]{1,128}$/.test(uid);
}
function assertObject(data, fields) {
  if (!data || typeof data !== "object" || Array.isArray(data) ||
      Object.keys(data).length !== fields.length ||
      fields.some(key => !Object.prototype.hasOwnProperty.call(data, key))) {
    throw new HttpsError("invalid-argument", "Richiesta di revisione non valida.");
  }
}
async function reviewer(request) {
  const uid = request?.auth?.uid;
  if (!validUid(uid) ||
      request.auth?.token?.privacy_deletion_reviewer !== true ||
      request.auth?.token?.admin !== true) {
    throw new HttpsError("permission-denied", "Operatore privacy non autorizzato.");
  }
  const authTime = Number(request.auth?.token?.auth_time);
  const now = Math.floor(Date.now() / 1000);
  if (!Number.isFinite(authTime) || authTime > now + 60 ||
      now - authTime > MAX_AUTH_AGE_SECONDS) {
    throw new HttpsError("failed-precondition",
      "Accedi nuovamente prima di gestire una richiesta di cancellazione.");
  }
  let live;
  try {
    live = await admin.auth().getUser(uid);
  } catch (_) {
    throw new HttpsError("permission-denied", "Operatore non verificato.");
  }
  if (live.disabled ||
      live.customClaims?.privacy_deletion_reviewer !== true ||
      live.customClaims?.admin !== true) {
    throw new HttpsError("permission-denied", "Autorizzazione di revisione revocata.");
  }

  // A valid signed token can still circulate after an administrator revokes
  // refresh tokens (or the account password changes). Authorization must
  // check the live Firebase Auth revocation timestamp, not just custom claims.
  // Admin SDK omits tokensValidAfterTime on accounts never revoked.
  const validSinceText = live.tokensValidAfterTime;
  if (validSinceText != null) {
    const validSince = typeof validSinceText === "string"
      ? Date.parse(validSinceText) : NaN;
    if (!Number.isFinite(validSince)) {
      throw new HttpsError("permission-denied",
        "Verifica revoca autorizzazione non disponibile.");
    }
    if (authTime * 1000 < validSince) {
      throw new HttpsError("permission-denied",
        "Sessione operatore revocata. Accedi nuovamente.");
    }
  }
  return uid;
}

exports.listVargaAccountDeletionReviewQueue = onCall({
  region: REGION, invoker: "public", timeoutSeconds: 20, maxInstances: 2
}, async request => {
  await reviewer(request);
  if (request.data != null) assertObject(request.data, []);
  try {
    const docs = await admin.firestore().collection(COLLECTION)
      .where("status", "in", ["pending", "in_review"])
      .limit(MAX_QUEUE_RESULTS).get();
    return {
      requests: docs.docs.map(doc => ({
        uid: doc.id,
        status: doc.data()?.status,
        reopened: Number.isSafeInteger(doc.data()?.reopenCount) &&
          doc.data().reopenCount > 0
      })).filter(item => validUid(item.uid) &&
        (item.status === "pending" || item.status === "in_review")),
      truncated: docs.size === MAX_QUEUE_RESULTS
    };
  } catch (error) {
    console.error("Deletion review queue read failed", { code: error?.code || "unknown" });
    throw new HttpsError("internal", "Coda richieste non disponibile.");
  }
});

exports.reviewVargaAccountDeletionRequest = onCall({
  region: REGION, invoker: "public", timeoutSeconds: 30, maxInstances: 2
}, async request => {
  const operatorUid = await reviewer(request);
  const data = request.data;
  const action = data?.action;
  if (action !== "start_review" && action !== "reject") {
    throw new HttpsError("invalid-argument",
      "Azione non valida: la revisione non può completare la cancellazione.");
  }
  assertObject(data, action === "reject"
    ? ["uid", "action", "reasonCode"] : ["uid", "action"]);
  const uid = data.uid;
  if (!validUid(uid) || uid === operatorUid) {
    throw new HttpsError("invalid-argument", "Identificativo richiesta non valido.");
  }
  if (action === "reject" && !REASONS.has(data.reasonCode)) {
    throw new HttpsError("invalid-argument", "Motivazione di revisione non valida.");
  }
  const ref = admin.firestore().collection(COLLECTION).doc(uid);
  try {
    return await admin.firestore().runTransaction(async tx => {
      const snapshot = await tx.get(ref);
      if (!snapshot.exists || snapshot.data()?.uid !== uid) {
        throw new HttpsError("not-found", "Richiesta non trovata.");
      }
      const current = snapshot.data().status;
      if (current === "completed" || current === "rejected") {
        throw new HttpsError("failed-precondition",
          "La richiesta è già chiusa e non può essere modificata qui.");
      }
      if (action === "start_review" && current !== "pending" ||
          action === "reject" && current !== "in_review") {
        throw new HttpsError("failed-precondition",
          "Transizione di stato non consentita. Aggiorna la coda.");
      }
      const next = action === "start_review" ? "in_review" : "rejected";
      const patch = {
        status: next,
        reviewedByUid: operatorUid,
        reviewedAt: admin.firestore.FieldValue.serverTimestamp()
      };
      if (action === "reject") patch.rejectionReasonCode = data.reasonCode;
      tx.update(ref, patch);
      return { updated: true, status: next };
    });
  } catch (error) {
    if (error instanceof HttpsError) throw error;
    console.error("Deletion review decision failed", { code: error?.code || "unknown" });
    throw new HttpsError("internal", "Revisione non registrata.");
  }
});
