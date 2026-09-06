"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");
const crypto = require("crypto");
const {
  MIN_CODE_LENGTH,
  MIN_PASSWORD_LENGTH,
  isRecoveryCodeCandidate,
  isValidNewPassword,
  createSalt,
  hashRecoveryCode,
  secureEqual,
  opaqueId
} = require("./password-recovery-code-core");

const REGION = "europe-west1";
const ADMIN_EMAIL = "ionut29019@gmail.com";
const CONFIG_PATH = "securityPrivate/passwordRecoveryCode";
const ATTEMPT_COLLECTION = "passwordRecoveryAttempts";
const CHALLENGE_COLLECTION = "passwordRecoveryChallenges";
const AUDIT_COLLECTION = "userAccessAudit";
const ACTIVE_STATUSES = new Set(["attivo", "active", "approved", "autorizzato", "abilitato"]);
const MAX_FAILURES = 5;
const FAILURE_WINDOW_MS = 15 * 60 * 1000;
const LOCK_MS = 30 * 60 * 1000;
const CHALLENGE_MS = 10 * 60 * 1000;

function callable(handler) {
  return functions.region(REGION).runWith({ timeoutSeconds: 30, memory: "256MB" }).https.onCall(handler);
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function timestampMs(value) {
  if (value && typeof value.toMillis === "function") return value.toMillis();
  const number = Number(value || 0);
  return Number.isFinite(number) ? number : 0;
}

function requestIp(context) {
  const forwarded = String(context?.rawRequest?.headers?.["x-forwarded-for"] || "").split(",")[0].trim();
  return forwarded || String(context?.rawRequest?.ip || "unknown");
}

async function requireAdministrator(context, db) {
  if (!context.auth?.uid) {
    throw new functions.https.HttpsError("unauthenticated", "Accedi come amministratore e riprova.");
  }
  const requester = await admin.auth().getUser(context.auth.uid);
  const email = normalizeEmail(requester.email || context.auth.token?.email);
  const configured = await db.collection("appConfig").doc("adminUsers").get();
  const adminEmails = new Set([
    ADMIN_EMAIL,
    ...(configured.exists && Array.isArray(configured.data()?.emails) ? configured.data().emails : [])
  ].map(normalizeEmail).filter(Boolean));
  const adminClaim = context.auth.token?.admin === true || context.auth.token?.isAdmin === true;
  if (!email || (!adminClaim && !adminEmails.has(email))) {
    throw new functions.https.HttpsError("permission-denied", "Funzione riservata all’amministratore.");
  }
  return { uid: requester.uid, email };
}

function attemptRefs(db, email, context) {
  return [
    db.collection(ATTEMPT_COLLECTION).doc(opaqueId(`email:${email}`)),
    db.collection(ATTEMPT_COLLECTION).doc(opaqueId(`ip:${requestIp(context)}`))
  ];
}

async function verifyCodeWithRateLimit(db, email, code, context) {
  const configRef = db.doc(CONFIG_PATH);
  const refs = attemptRefs(db, email, context);
  const nowMs = Date.now();
  return db.runTransaction(async (transaction) => {
    const [configSnapshot, ...attemptSnapshots] = await Promise.all([
      transaction.get(configRef),
      ...refs.map((ref) => transaction.get(ref))
    ]);
    for (const snapshot of attemptSnapshots) {
      if (timestampMs(snapshot.data()?.lockedUntil) > nowMs) {
        throw new functions.https.HttpsError("resource-exhausted", "Troppi tentativi. Attendi 30 minuti e riprova.");
      }
    }
    const config = configSnapshot.data() || {};
    const valid = config.enabled === true
      && config.salt
      && config.codeHash
      && secureEqual(hashRecoveryCode(code, config.salt), config.codeHash);
    refs.forEach((ref, index) => {
      const previous = attemptSnapshots[index].data() || {};
      const previousWindow = timestampMs(previous.windowStartedAt);
      const sameWindow = previousWindow > 0 && nowMs - previousWindow < FAILURE_WINDOW_MS;
      const failCount = valid ? 0 : (sameWindow ? Number(previous.failCount || 0) : 0) + 1;
      transaction.set(ref, {
        failCount,
        windowStartedAt: admin.firestore.Timestamp.fromMillis(sameWindow ? previousWindow : nowMs),
        lockedUntil: admin.firestore.Timestamp.fromMillis(!valid && failCount >= MAX_FAILURES ? nowMs + LOCK_MS : 0),
        lastAttemptAt: admin.firestore.FieldValue.serverTimestamp(),
        lastResult: valid ? "accepted" : "rejected"
      }, { merge: true });
    });
    return valid;
  });
}

function profileIsAdministrator(profile, email) {
  const role = String(profile?.role || profile?.ruolo || "").trim().toLowerCase();
  return email === ADMIN_EMAIL || profile?.isAdmin === true || profile?.admin === true || /admin|amministratore/.test(role);
}

async function emailIsConfiguredAdministrator(db, email) {
  if (email === ADMIN_EMAIL) return true;
  const snapshot = await db.collection("appConfig").doc("adminUsers").get();
  const emails = snapshot.exists && Array.isArray(snapshot.data()?.emails) ? snapshot.data().emails : [];
  return emails.map(normalizeEmail).includes(email);
}

function profileIsActive(profile) {
  if (!profile || profile.banned === true) return false;
  const status = String(profile.statoAccount || profile.accountStatus || "").trim().toLowerCase();
  if (!status) return profile.profileMigratedByEmail !== false;
  return ACTIVE_STATUSES.has(status);
}

exports.setPasswordRecoveryCode = callable(async (data, context) => {
  const db = admin.firestore();
  const administrator = await requireAdministrator(context, db);
  const code = String(data?.code || "").trim();
  if (!isRecoveryCodeCandidate(code)) {
    throw new functions.https.HttpsError(
      "invalid-argument",
      `Il codice deve iniziare con REC-, contenere almeno ${MIN_CODE_LENGTH} caratteri e non avere spazi.`
    );
  }
  const salt = createSalt();
  await db.doc(CONFIG_PATH).set({
    enabled: true,
    codeHash: hashRecoveryCode(code, salt),
    salt,
    version: 1,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedByUid: administrator.uid,
    updatedByEmail: administrator.email
  });
  await db.collection(AUDIT_COLLECTION).add({
    action: "password-recovery-code-configured",
    administratorUid: administrator.uid,
    administratorEmail: administrator.email,
    createdAt: admin.firestore.FieldValue.serverTimestamp()
  });
  return { configured: true };
});

exports.getPasswordRecoveryCodeStatus = callable(async (_data, context) => {
  const db = admin.firestore();
  await requireAdministrator(context, db);
  const snapshot = await db.doc(CONFIG_PATH).get();
  const data = snapshot.data() || {};
  return {
    configured: snapshot.exists && data.enabled === true,
    updatedAt: data.updatedAt?.toMillis ? data.updatedAt.toMillis() : null,
    updatedByEmail: normalizeEmail(data.updatedByEmail)
  };
});

exports.startPasswordRecoveryWithCode = callable(async (data, context) => {
  const db = admin.firestore();
  const email = normalizeEmail(data?.email);
  const code = String(data?.code || "").trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !isRecoveryCodeCandidate(code)) {
    throw new functions.https.HttpsError("invalid-argument", "Email o codice di recupero non valido.");
  }
  if (!(await verifyCodeWithRateLimit(db, email, code, context))) {
    throw new functions.https.HttpsError("permission-denied", "Codice di recupero non valido.");
  }

  let user;
  try {
    user = await admin.auth().getUserByEmail(email);
  } catch (error) {
    if (error?.code === "auth/user-not-found") {
      throw new functions.https.HttpsError("permission-denied", "Account non disponibile per il recupero con codice.");
    }
    throw error;
  }
  const profileSnapshot = await db.collection("platformUsers").doc(user.uid).get();
  const profile = profileSnapshot.data() || null;
  if (!profileIsActive(profile)) {
    throw new functions.https.HttpsError("permission-denied", "Account non attivo. Contatta l’amministratore.");
  }
  if (profileIsAdministrator(profile, email) || await emailIsConfiguredAdministrator(db, email)) {
    throw new functions.https.HttpsError("permission-denied", "Per gli amministratori è obbligatorio il recupero tramite email.");
  }

  const challengeToken = crypto.randomBytes(32).toString("base64url");
  const challengeId = opaqueId(challengeToken);
  await db.collection(CHALLENGE_COLLECTION).doc(challengeId).set({
    uid: user.uid,
    email,
    state: "issued",
    createdAt: admin.firestore.FieldValue.serverTimestamp(),
    expiresAt: admin.firestore.Timestamp.fromMillis(Date.now() + CHALLENGE_MS)
  });
  return { allowed: true, challengeToken, expiresInSeconds: CHALLENGE_MS / 1000 };
});

exports.completePasswordRecoveryWithCode = callable(async (data) => {
  const db = admin.firestore();
  const token = String(data?.challengeToken || "").trim();
  const newPassword = String(data?.newPassword || "");
  if (!token || token.length > 256 || !isValidNewPassword(newPassword)) {
    throw new functions.https.HttpsError(
      "invalid-argument",
      `La nuova password deve contenere almeno ${MIN_PASSWORD_LENGTH} caratteri.`
    );
  }
  const challengeRef = db.collection(CHALLENGE_COLLECTION).doc(opaqueId(token));
  let challenge;
  await db.runTransaction(async (transaction) => {
    const snapshot = await transaction.get(challengeRef);
    challenge = snapshot.data() || null;
    if (!snapshot.exists || challenge.state !== "issued" || timestampMs(challenge.expiresAt) <= Date.now()) {
      throw new functions.https.HttpsError("permission-denied", "Sessione di recupero scaduta. Ripeti la procedura.");
    }
    transaction.set(challengeRef, {
      state: "processing",
      processingAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  });

  try {
    const currentProfileSnapshot = await db.collection("platformUsers").doc(challenge.uid).get();
    const currentProfile = currentProfileSnapshot.data() || null;
    if (!profileIsActive(currentProfile)
      || profileIsAdministrator(currentProfile, challenge.email)
      || await emailIsConfiguredAdministrator(db, challenge.email)) {
      throw new functions.https.HttpsError("permission-denied", "Account non disponibile per il recupero con codice.");
    }
    await admin.auth().updateUser(challenge.uid, { password: newPassword });
    await admin.auth().revokeRefreshTokens(challenge.uid);
    const now = admin.firestore.FieldValue.serverTimestamp();
    await db.collection("platformUsers").doc(challenge.uid).set({
      mustChangePassword: false,
      passwordChangedAt: now,
      passwordRecoveryCodeUsedAt: now,
      updatedAt: now
    }, { merge: true });
    await db.collection(AUDIT_COLLECTION).add({
      userId: challenge.uid,
      userEmail: challenge.email,
      action: "password-recovered-with-shared-code",
      channel: "secure-callable-challenge",
      createdAt: now
    });
    await challengeRef.set({ state: "completed", completedAt: now }, { merge: true });
    return { changed: true, email: challenge.email };
  } catch (error) {
    await challengeRef.set({
      state: "issued",
      lastErrorAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    if (error instanceof functions.https.HttpsError) throw error;
    if (["auth/invalid-password", "auth/weak-password"].includes(error?.code)) {
      throw new functions.https.HttpsError("invalid-argument", "Scegli una password più sicura.");
    }
    console.error("Recupero password con codice fallito", { uid: challenge?.uid || "", code: error?.code || "" });
    throw new functions.https.HttpsError("internal", "Cambio password non riuscito. Riprova.");
  }
});
