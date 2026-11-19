"use strict";

// Self-service, authenticated account deletion.
// Business records shared with a company are not deleted by this action.
const admin = require("firebase-admin");
const { randomUUID } = require("node:crypto");
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

// Remove an account's identity from shared central notifications without
// deleting operational messages or other recipients' delivery/read state.
// Use server-side lookups so the client cannot choose another account to scrub.
async function scrubCentralNotifications(db, uid) {
  const deleteValue = admin.firestore.FieldValue.delete();
  const refs = new Map();
  for (const [field, operator, value] of [
    ["actorId", "==", uid],
    ["recipientUserIds", "array-contains", uid],
    ["readBy." + uid, "==", true],
    ["readBy." + uid, "==", false],
    ["deletedBy." + uid, "==", true],
    ["deletedBy." + uid, "==", false]
  ]) {
    const snapshot = await db.collection("notifications").where(field, operator, value).get();
    for (const doc of snapshot.docs) refs.set(doc.ref.path, doc);
  }
  for (const doc of refs.values()) {
    const notification = doc.data() || {};
    const patch = {};
    if (notification.actorId === uid) {
      patch.actorId = deleteValue;
      patch.actorName = "Account eliminato";
    }
    if (Array.isArray(notification.recipientUserIds) &&
        notification.recipientUserIds.includes(uid)) {
      patch.recipientUserIds = admin.firestore.FieldValue.arrayRemove(uid);
    }
    if (notification.readBy &&
        Object.prototype.hasOwnProperty.call(notification.readBy, uid)) {
      patch["readBy." + uid] = deleteValue;
    }
    if (notification.deletedBy &&
        Object.prototype.hasOwnProperty.call(notification.deletedBy, uid)) {
      patch["deletedBy." + uid] = deleteValue;
    }
    if (Object.keys(patch).length) await doc.ref.update(patch);
  }
}

// Preserve work messages, but unlink the deleted user's author identity.
// Never edit the recipient, message body or job metadata: they can belong to
// other staff or be necessary for a still-active business conversation.
async function unlinkSharedAuthor(db, collectionName, identityField, uid, nameField, emailField) {
  const records = await db.collection(collectionName).where(identityField, "==", uid).get();
  const deleted = admin.firestore.FieldValue.delete();
  for (const doc of records.docs) {
    const data = doc.data() || {};
    const patch = { [identityField]: deleted };
    if (Object.prototype.hasOwnProperty.call(data, emailField)) patch[emailField] = deleted;
    if (Object.prototype.hasOwnProperty.call(data, nameField)) patch[nameField] = "Account eliminato";
    await doc.ref.update(patch);
  }
}


/**
 * A deleted account must not remain the explicit addressee of retained chat.
 * Keep the sender and the message intact: both are needed by other participants.
 * This only matches Firebase UIDs stored as recipientId. Legacy email aliases,
 * content text and metadata require separate, contextual review. Using an
 * unlinkable non-empty tombstone preserves private-message routing.
 */
async function unlinkDeletedChatRecipient(db, uid) {
  const messages = await db.collection("chatMessages").where("recipientId", "==", uid).get();
  for (const message of messages.docs) {
    // A missing recipientId can turn a private conversation into a broadcast
    // in legacy clients. Replace it with a non-matching, unlinkable tombstone.
    // Never store the deleted UID or account email in this value.
    await message.ref.update({ recipientId: "deleted-recipient-" + randomUUID() });
  }
}


/**
 * Minimise legacy notification email queue records linked EXACTLY to this
 * authenticated account. The queue uses scalar "to" and "createdBy" fields.
 * Never remove another recipient or guess addresses from message text.
 * Already-delivered external emails cannot be recalled by Firebase.
 */
async function minimizeOwnedEmailQueue(db, uid, accountEmail) {
  const aliases = [uid];
  if (typeof accountEmail === "string" && accountEmail.trim()) {
    aliases.push(accountEmail.trim());
  }
  const refs = new Map();
  for (const value of aliases) {
    for (const field of ["to", "createdBy"]) {
      const matches = await db.collection("notificationEmailQueue").where(field, "==", value).get();
      for (const doc of matches.docs) refs.set(doc.ref.path, doc);
    }
  }
  const erased = admin.firestore.FieldValue.delete();
  for (const doc of refs.values()) {
    const data = doc.data() || {};
    const patch = {};
    if (aliases.includes(data.to)) {
      // Cancel ONLY an exact-recipient message. Do not delete shared queues.
      patch.to = erased;
      if (Object.prototype.hasOwnProperty.call(data, "html")) patch.html = erased;
      if (Object.prototype.hasOwnProperty.call(data, "text")) patch.text = erased;
      if (data.status === "pending") patch.status = "cancelled_account_deleted";
    }
    if (aliases.includes(data.createdBy)) patch.createdBy = "Account eliminato";
    if (Object.keys(patch).length) await doc.ref.update(patch);
  }
}

/**
 * Operational security logs can remain for their documented retention period,
 * but no longer need the account's Firebase UID, email, name or user agent.
 * Only exact UID/email matches are scrubbed, not arbitrary detail strings.
 */
async function minimizeOwnActivityLogs(db, uid, accountEmail) {
  const refs = new Map();
  const matches = [["userId", uid]];
  if (typeof accountEmail === "string" && accountEmail.trim()) {
    matches.push(["userEmail", accountEmail.trim()]);
  }
  for (const [field, value] of matches) {
    const docs = await db.collection("activityLogs").where(field, "==", value).get();
    for (const doc of docs.docs) refs.set(doc.ref.path, doc);
  }
  const erased = admin.firestore.FieldValue.delete();
  for (const doc of refs.values()) {
    const data = doc.data() || {};
    const patch = {};
    if (data.userId === uid) patch.userId = erased;
    if (data.userEmail === accountEmail && accountEmail) patch.userEmail = erased;
    if (data.userId === uid || (data.userEmail === accountEmail && accountEmail)) {
      if (Object.prototype.hasOwnProperty.call(data, "userName")) patch.userName = "Account eliminato";
      if (Object.prototype.hasOwnProperty.call(data, "deviceInfo")) patch.deviceInfo = erased;
    }
    if (Object.keys(patch).length) await doc.ref.update(patch);
  }
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
    const ownAuthRecord = await admin.auth().getUser(uid);

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
      ["userAlerts", "userId"],
      // Temporary password-reset challenges must not outlive a deleted login.
      ["passwordRecoveryChallenges", "uid"]
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

    // Shared business notifications must not keep an obsolete Firebase UID.
    await scrubCentralNotifications(db, uid);

    // The operator approved minimising identity in retained work communications.
    // Only rows linked by exact author/recipient UID are touched. We retain
    // message text, enterprise metadata and identities of all other users.
    // TTL/retention and legacy email aliases are separate controls.
    await unlinkSharedAuthor(db, "chatMessages", "senderId", uid, "senderName", "senderEmail");
    await unlinkDeletedChatRecipient(db, uid);
    await unlinkSharedAuthor(db, "appNotifications", "createdByUid", uid, "createdByName", "createdByEmail");

    // Queue and logs use a different, legacy identity schema. Only the
    // verified Auth email of the deleting user may be used for exact matches.
    await minimizeOwnedEmailQueue(db, uid, ownAuthRecord.email || "");
    await minimizeOwnActivityLogs(db, uid, ownAuthRecord.email || "");

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