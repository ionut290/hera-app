"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");
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

async function assertAdmin(context, db) {
  const email = String(context.auth?.token?.email || "").trim().toLowerCase();
  if (!context.auth || !email) {
    throw new functions.https.HttpsError("unauthenticated", "Login richiesto.");
  }
  const adminEmails = await getAdminEmails(db);
  if (!adminEmails.has(email)) {
    throw new functions.https.HttpsError(
      "permission-denied",
      "Solo un amministratore può modificare le password degli utenti."
    );
  }
  return { uid: context.auth.uid, email };
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

async function resolveTargetUser(data) {
  const uid = String(data?.uid || "").trim();
  const email = String(data?.email || "").trim().toLowerCase();
  if (uid) {
    try {
      return await admin.auth().getUser(uid);
    } catch (error) {
      if (error?.code !== "auth/user-not-found") throw error;
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

exports.adminSetUserPassword = functions
  .region("europe-west1")
  .https.onCall(async (data, context) => {
    const db = admin.firestore();
    const administrator = await assertAdmin(context, db);
    const target = await resolveTargetUser(data);

    if (target.uid === administrator.uid) {
      throw new functions.https.HttpsError(
        "failed-precondition",
        "Per sicurezza modifica la tua password amministratore dal tuo account."
      );
    }

    const temporaryPassword = generateTemporaryPassword();

    try {
      await admin.auth().updateUser(target.uid, {
        password: temporaryPassword,
        disabled: false,
        // L'account viene considerato verificato dall'amministratore: il flusso
        // temporaneo deve funzionare completamente dentro l'app, senza email.
        emailVerified: true
      });
      await admin.auth().revokeRefreshTokens(target.uid);

      const now = admin.firestore.FieldValue.serverTimestamp();
      await db.collection("platformUsers").doc(target.uid).set({
        uid: target.uid,
        email: target.email || String(data?.email || "").trim().toLowerCase(),
        emailVerified: true,
        mustChangePassword: true,
        temporaryPasswordIssuedAt: now,
        passwordManagedByAdminAt: now,
        passwordManagedByAdminUid: administrator.uid,
        passwordManagedByAdminEmail: administrator.email,
        passwordManagementMode: "temporary",
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
        createdAt: now
      });

      return {
        success: true,
        uid: target.uid,
        email: target.email || "",
        mode: "temporary",
        mustChangePassword: true,
        temporaryPassword
      };
    } catch (error) {
      console.error("Creazione password temporanea amministratore fallita", {
        targetUid: target.uid,
        code: error?.code || "",
        message: error?.message || ""
      });
      if (error instanceof functions.https.HttpsError) throw error;
      throw new functions.https.HttpsError(
        "internal",
        "Non è stato possibile creare la password temporanea. Riprova."
      );
    }
  });
