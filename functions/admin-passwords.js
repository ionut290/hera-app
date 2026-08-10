"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");
const crypto = require("crypto");

const ADMIN_EMAIL = "ionut29019@gmail.com";
const MIN_PASSWORD_LENGTH = 10;

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
    const mode = String(data?.mode || "temporary").trim().toLowerCase();
    if (!new Set(["temporary", "custom"]).has(mode)) {
      throw new functions.https.HttpsError("invalid-argument", "Modalità password non valida.");
    }

    const target = await resolveTargetUser(data);
    if (target.uid === administrator.uid) {
      throw new functions.https.HttpsError(
        "failed-precondition",
        "Per sicurezza modifica la tua password amministratore dal tuo account."
      );
    }

    let nextPassword = "";
    let mustChangePassword = false;
    if (mode === "temporary") {
      nextPassword = generateTemporaryPassword();
      mustChangePassword = true;
    } else {
      nextPassword = String(data?.password || "");
      if (nextPassword.length < MIN_PASSWORD_LENGTH) {
        throw new functions.https.HttpsError(
          "invalid-argument",
          `La password deve contenere almeno ${MIN_PASSWORD_LENGTH} caratteri.`
        );
      }
    }

    try {
      await admin.auth().updateUser(target.uid, {
        password: nextPassword,
        disabled: false
      });
      await admin.auth().revokeRefreshTokens(target.uid);

      const now = admin.firestore.FieldValue.serverTimestamp();
      const profilePatch = {
        uid: target.uid,
        email: target.email || String(data?.email || "").trim().toLowerCase(),
        mustChangePassword,
        passwordManagedByAdminAt: now,
        passwordManagedByAdminUid: administrator.uid,
        passwordManagedByAdminEmail: administrator.email,
        passwordManagementMode: mode,
        updatedAt: now
      };
      if (mode === "temporary") {
        profilePatch.temporaryPasswordIssuedAt = now;
      } else {
        profilePatch.passwordChangedAt = now;
      }

      await db.collection("platformUsers").doc(target.uid).set(profilePatch, { merge: true });
      await db.collection("userAccessAudit").add({
        userId: target.uid,
        userEmail: target.email || "",
        action: mode === "temporary" ? "temporary-password-issued" : "password-set-by-admin",
        forceChange: mustChangePassword,
        administratorUid: administrator.uid,
        administratorEmail: administrator.email,
        createdAt: now
      });

      return {
        success: true,
        uid: target.uid,
        email: target.email || "",
        mode,
        mustChangePassword,
        temporaryPassword: mode === "temporary" ? nextPassword : null
      };
    } catch (error) {
      console.error("Cambio password amministratore fallito", {
        targetUid: target.uid,
        mode,
        code: error?.code || "",
        message: error?.message || ""
      });
      if (error instanceof functions.https.HttpsError) throw error;
      if (["auth/invalid-password", "auth/weak-password"].includes(error?.code)) {
        throw new functions.https.HttpsError("invalid-argument", "Password non valida.");
      }
      throw new functions.https.HttpsError(
        "internal",
        "Non è stato possibile modificare la password. Riprova."
      );
    }
  });
