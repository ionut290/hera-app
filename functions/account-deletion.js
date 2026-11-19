"use strict";

// Self-service, authenticated account deletion.
// Business records shared with a company are not deleted by this action.
const admin = require("firebase-admin");
const { onCall, HttpsError } = require("firebase-functions/v2/https");

const REGION = "europe-west1";
// Explicit default Storage bucket: functions/index.js initializes Admin without storageBucket.
const STORAGE_BUCKET = "hera-app-6cd2b.firebasestorage.app";
const MAX_AUTH_AGE_SECONDS = 600;

function validUserId(value) {
  return typeof value === "string" && /^[a-zA-Z0-9_-]{1,128}$/.test(value);
}

async function recursiveDeleteIfPresent(db, path) {
  await db.recursiveDelete(db.doc(path));
}

async function deleteMatchedDocuments(db, collectionName, field, uid) {
  const snapshot = await db.collection(collectionName).where(field, "==", uid).get();
  for (const doc of snapshot.docs) {
    await db.recursiveDelete(doc.ref);
  }
}

async function referencedByRetainedDocument(db, storagePath, personalIds) {
  for (const field of ["storagePath", "thumbnailPath"]) {
    const matches = await db.collection("documents").where(field, "==", storagePath).get();
    if (matches.docs.some((doc) => !personalIds.has(doc.id))) return true;
  }
  return false;
}

function safePersonalStoragePath(path, uid) {
  const prefix = "documents/" + uid + "/";
  // Never delete a company/shared attachment or another user's objects.
  return typeof path === "string" && path.startsWith(prefix)
    && path.length > prefix.length && !path.includes("..");
}

exports.deleteOwnVargaAccount = onCall({
  region: REGION,
  invoker: "public", // Firebase bearer auth is enforced below
  timeoutSeconds: 300,
  maxInstances: 3
}, async (request) => {
  const auth = request.auth;
  const uid = auth && auth.uid;
  if (!validUserId(uid)) {
    throw new HttpsError("unauthenticated", "Accedi per eliminare il tuo account.");
  }
  if (request.data?.confirmation !== "ELIMINA") {
    throw new HttpsError("invalid-argument", "Conferma esplicita mancante.");
  }
  const authTime = Number(auth.token?.auth_time);
  const now = Math.floor(Date.now() / 1000);
  if (!Number.isFinite(authTime) || authTime > now + 60 || now - authTime > MAX_AUTH_AGE_SECONDS) {
    throw new HttpsError("failed-precondition", "Per sicurezza esci, accedi nuovamente e ripeti l'eliminazione.");
  }

  // Deletion is scoped exclusively to request.auth.uid; client-supplied UID,
  // email, document IDs, and storage paths cannot expand its authority.
  const db = admin.firestore();
  try {
    await admin.auth().getUser(uid);

    // Personal documents only. Shared/global operational documents must remain
    // available to the business and are not removed by account deletion.
    const owned = await db.collection("documents").where("ownerUserId", "==", uid).get();
    // Legacy documents with no visibility are shown as personal in documents.js.
    // Include them so account deletion cannot silently orphan personal data.
    const personal = owned.docs.filter((doc) => !doc.data()?.visibility || doc.data()?.visibility === "personal");
    const paths = [];
    for (const doc of personal) {
      const file = doc.data()?.storagePath;
      const thumbnail = doc.data()?.thumbnailPath;
      if (!safePersonalStoragePath(file, uid) ||
          (thumbnail && !safePersonalStoragePath(thumbnail, uid))) {
        throw new HttpsError("failed-precondition",
          "Un documento personale usa un percorso di archivio non previsto. Contatta l'assistenza per completare l'eliminazione.");
      }
      paths.push(file);
      if (thumbnail) paths.push(thumbnail);
    }

    // Validate the deletion plan without changing user data or Apple credentials.
    // Clients request this first, before revoking an Apple authorization code.
    if (request.data?.mode === "preflight") {
      return { ready: true, personalDocuments: personal.length };
    }

    // Do not destroy business files, shared geodata, timesheets, commesse or
    // records belonging to other people. Remove personal files before metadata.
    if (paths.length) {
      const bucket = admin.storage().bucket(STORAGE_BUCKET);
      const personalIds = new Set(personal.map((doc) => doc.id));
      for (const path of new Set(paths)) {
        // Retain objects also referenced by company/shared document metadata.
        if (await referencedByRetainedDocument(db, path, personalIds)) continue;
        await bucket.file(path).delete({ ignoreNotFound: true });
      }
    }
    for (const doc of personal) await db.recursiveDelete(doc.ref);

    const personalQueries = [
      ["documentReminderAcknowledgements", "userId"],
      ["alertReadConfirmations", "utente"],
      ["userAlertAcknowledgements", "userId"],
      ["userAlerts", "userId"]
    ];
    for (const [collection, field] of personalQueries) {
      await deleteMatchedDocuments(db, collection, field, uid);
    }

    // Employee registry records may have a legitimate business retention purpose,
    // but must not remain linked to a deleted Firebase Authentication identity.
    // Remove ONLY auth-derived fields, not personnel/contract/working-time data.
    const linkedPersonnel = await db.collection("personale").where("linkedUserId", "==", uid).get();
    const deleteField = admin.firestore.FieldValue.delete();
    for (const doc of linkedPersonnel.docs) {
      const person = doc.data() || {};
      const unlink = {
        linkedUserId: deleteField,
        linkedAuthProvider: deleteField,
        googleLastLoginAt: deleteField,
        googleDisplayName: deleteField,
        emailVerified: deleteField
      };
      // A Google-sourced portrait is associated with the deleted login.
      // Preserve an operator's separately managed business photo.
      if (person.profilePhotoSource === "google") {
        unlink.photoURL = deleteField;
        unlink.profilePhotoSource = deleteField;
      }
      await doc.ref.update(unlink);
    }

    // Personal private records and authorization links are always owned by UID.
    for (const collection of [
      "privateDocuments", "operatorPositions", "userOperatorLinks", "platformUsers"
    ]) {
      await recursiveDeleteIfPresent(db, collection + "/" + uid);
    }

    // Authentication is removed last: if cleanup fails the account remains
    // usable so the request can be retried; never fake deletion success.
    await admin.auth().deleteUser(uid);
    return { deleted: true };
  } catch (error) {
    if (error instanceof HttpsError) throw error;
    console.error("Account deletion failed:", { code: error?.code || "unknown" });
    throw new HttpsError("internal",
      "Eliminazione non completata. L'account non è stato confermato come cancellato: riprova o contatta l'assistenza.");
  }
});