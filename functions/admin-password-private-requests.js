"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");
const { onDocumentCreated } = require("firebase-functions/v2/firestore");
const crypto = require("crypto");

const ADMIN_EMAIL = "ionut29019@gmail.com";

async function getAdminEmails(db) {
  const snapshot = await db.collection("appConfig").doc("adminUsers").get();
  const configured = snapshot.exists && Array.isArray(snapshot.data().emails)
    ? snapshot.data().emails
    : [];
  return new Set(
    [ADMIN_EMAIL, ...configured]
      .map((email) => String(email || "").trim().toLowerCase())
      .filter(Boolean)
  );
}

async function resolveAdministrator(userId, requestedByEmail, db) {
  let requester;
  try {
    requester = await admin.auth().getUser(String(userId || "").trim());
  } catch (_) {
    throw new functions.https.HttpsError("unauthenticated", "Account amministratore non valido.");
  }
  const email = String(requester.email || "").trim().toLowerCase();
  const claimedEmail = String(requestedByEmail || "").trim().toLowerCase();
  const adminEmails = await getAdminEmails(db);
  if (!email || !adminEmails.has(email) || (claimedEmail && claimedEmail !== email)) {
    throw new functions.https.HttpsError(
      "permission-denied",
      "Solo un amministratore può modificare le password degli utenti."
    );
  }
  return { uid: requester.uid, email };
}

async function resolveTargetUser(data) {
  const uid = String(data?.targetUid || "").trim();
  const email = String(data?.targetEmail || "").trim().toLowerCase();
  if (uid) {
    try {
      return await admin.auth().getUser(uid);
    } catch (error) {
      if (!["auth/user-not-found", "auth/invalid-uid"].includes(error?.code)) throw error;
    }
  }
  if (email) {
    try {
      return await admin.auth().getUserByEmail(email);
    } catch (error) {
      if (error?.code !== "auth/user-not-found") throw error;
    }
  }
  throw new functions.https.HttpsError("not-found", "Utente Firebase non trovato.");
}

function generateTemporaryPassword() {
  const upper = "ABCDEFGHJKLMNPQRSTUVWXYZ";
  const lower = "abcdefghijkmnopqrstuvwxyz";
  const digits = "23456789";
  const symbols = "!@#$%";
  const all = upper + lower + digits + symbols;
  const chars = [
    upper[crypto.randomInt(upper.length)],
    lower[crypto.randomInt(lower.length)],
    digits[crypto.randomInt(digits.length)],
    symbols[crypto.randomInt(symbols.length)]
  ];
  while (chars.length < 16) chars.push(all[crypto.randomInt(all.length)]);
  for (let index = chars.length - 1; index > 0; index -= 1) {
    const swapIndex = crypto.randomInt(index + 1);
    [chars[index], chars[swapIndex]] = [chars[swapIndex], chars[index]];
  }
  return chars.join("");
}

function encryptTemporaryPassword(temporaryPassword, publicKeyJwk) {
  if (!publicKeyJwk || typeof publicKeyJwk !== "object" || publicKeyJwk.kty !== "RSA") {
    throw new functions.https.HttpsError("invalid-argument", "Chiave di sicurezza richiesta non valida.");
  }
  try {
    const publicKey = crypto.createPublicKey({ key: publicKeyJwk, format: "jwk" });
    return crypto.publicEncrypt(
      {
        key: publicKey,
        oaepHash: "sha256",
        padding: crypto.constants.RSA_PKCS1_OAEP_PADDING
      },
      Buffer.from(temporaryPassword, "utf8")
    ).toString("base64");
  } catch (error) {
    console.error("Cifratura password temporanea fallita", {
      code: error?.code || "",
      message: error?.message || ""
    });
    throw new functions.https.HttpsError("invalid-argument", "Canale sicuro password non disponibile.");
  }
}

async function applyTemporaryPassword(db, target, administrator, temporaryPassword, requestedEmail) {
  await admin.auth().updateUser(target.uid, {
    password: temporaryPassword,
    disabled: false,
    emailVerified: true
  });
  await admin.auth().revokeRefreshTokens(target.uid);

  const now = admin.firestore.FieldValue.serverTimestamp();
  await db.collection("platformUsers").doc(target.uid).set({
    uid: target.uid,
    email: target.email || String(requestedEmail || "").trim().toLowerCase(),
    emailVerified: true,
    mustChangePassword: true,
    temporaryPasswordIssuedAt: now,
    passwordManagedByAdminAt: now,
    passwordManagedByAdminUid: administrator.uid,
    passwordManagedByAdminEmail: administrator.email,
    passwordManagementMode: "temporary-private-firestore",
    updatedAt: now
  }, { merge: true });

  await db.collection("userAccessAudit").add({
    userId: target.uid,
    userEmail: target.email || "",
    action: "temporary-password-issued",
    forceChange: true,
    emailFlowRequired: false,
    administratorUid: administrator.uid,
    administratorEmail: administrator.email,
    channel: "private-firestore-trigger",
    createdAt: now
  });
}

exports.processAdminPasswordPrivateRequest = onDocumentCreated(
  {
    document: "privateDocuments/{userId}/adminPasswordRequests/{requestId}",
    region: "europe-west1"
  },
  async (event) => {
    const snapshot = event.data;
    if (!snapshot) return null;

    const requestRef = snapshot.ref;
    const data = snapshot.data() || {};
    const db = admin.firestore();
    const deleteField = admin.firestore.FieldValue.delete();

    try {
      if (data.requestedByUid && String(data.requestedByUid) !== String(event.params.userId)) {
        throw new functions.https.HttpsError("permission-denied", "Richiesta amministratore non valida.");
      }

      const administrator = await resolveAdministrator(
        event.params.userId,
        data.requestedByEmail,
        db
      );
      const target = await resolveTargetUser(data);
      if (target.uid === administrator.uid) {
        throw new functions.https.HttpsError(
          "failed-precondition",
          "Per sicurezza modifica la tua password amministratore dal tuo account."
        );
      }

      const temporaryPassword = generateTemporaryPassword();
      const encryptedTemporaryPassword = encryptTemporaryPassword(
        temporaryPassword,
        data.publicKeyJwk
      );

      await requestRef.set({
        status: "processing",
        encryptedTemporaryPassword,
        publicKeyJwk: deleteField,
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      }, { merge: true });

      await applyTemporaryPassword(
        db,
        target,
        administrator,
        temporaryPassword,
        data.targetEmail
      );

      await requestRef.set({
        status: "completed",
        targetUid: target.uid,
        targetEmail: target.email || String(data.targetEmail || "").trim().toLowerCase(),
        mustChangePassword: true,
        completedAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
      return null;
    } catch (error) {
      console.error("Richiesta privata Firestore cambio password fallita", {
        requestId: event.params.requestId,
        administratorUid: event.params.userId,
        targetUid: String(data.targetUid || ""),
        targetEmail: String(data.targetEmail || ""),
        code: error?.code || "",
        message: error?.message || ""
      });

      const code = String(error?.code || "internal").replace(/^functions\//, "");
      const safeMessage = error instanceof functions.https.HttpsError
        ? error.message
        : "Non è stato possibile creare la password temporanea.";
      await requestRef.set({
        status: "failed",
        errorCode: code,
        errorMessage: safeMessage,
        publicKeyJwk: deleteField,
        encryptedTemporaryPassword: deleteField,
        completedAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
      return null;
    }
  }
);
