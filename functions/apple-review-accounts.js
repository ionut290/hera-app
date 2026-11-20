"use strict";

/**
 * Admin-only account creation for App Review. This is not a general-purpose
 * account signup flow. Password is generated once, never logged or persisted
 * by this function; Firebase Authentication stores the normal password hash.
 * No production records are copied into the reviewer environment.
 */
const crypto = require("node:crypto");
const admin = require("firebase-admin");
const { onCall, HttpsError } = require("firebase-functions/v2/https");

const ADMIN_EMAIL = "ionut29019@gmail.com";
const REGION = "europe-west1";

const cleanEmail = (input) => String(input || "").trim().toLowerCase();
function randomPassword() {
  const groups = [
    "abcdefghjkmnpqrstuvwxyz",
    "ABCDEFGHJKLMNPQRSTUVWXYZ",
    "23456789",
    "!@#$%"
  ];
  const source = groups.join("");
  const chars = groups.map((group) => group[crypto.randomInt(group.length)]);
  while (chars.length < 26) chars.push(source[crypto.randomInt(source.length)]);
  for (let i = chars.length - 1; i > 0; i -= 1) {
    const j = crypto.randomInt(i + 1);
    [chars[i], chars[j]] = [chars[j], chars[i]];
  }
  return chars.join("");
}
// Creating a review credential is a privileged operation. A signed ID
// token can remain usable briefly after the administrator has been disabled
// or their email/privileges changed. Always re-check the live Auth record,
// and demand a recently authenticated administrator (never a review user).
async function assertAdmin(auth) {
  const uid = auth?.uid;
  const tokenEmail = cleanEmail(auth?.token?.email);
  if (typeof uid !== "string" || !/^[A-Za-z0-9_-]{1,128}$/.test(uid) ||
      auth?.token?.apple_review === true ||
      auth?.token?.email_verified !== true || !tokenEmail) {
    throw new HttpsError("permission-denied", "Account amministratore verificato richiesto.");
  }
  const authTime = Number(auth.token?.auth_time);
  const now = Math.floor(Date.now() / 1000);
  if (!Number.isFinite(authTime) || authTime > now + 60 ||
      now - authTime > 600) {
    throw new HttpsError("failed-precondition",
      "Accedi nuovamente come amministratore prima di creare un account di revisione.");
  }

  let live;
  try {
    live = await admin.auth().getUser(uid);
  } catch (error) {
    // An old token for a disabled/deleted identity is never a grant.
    if (error?.code !== "auth/user-not-found") {
      console.error("Apple review admin verification unavailable", {
        code: error?.code || "unknown"
      });
    }
    throw new HttpsError("permission-denied", "Amministratore non autorizzato.");
  }
  if (!live || live.uid !== uid || live.disabled === true ||
      live.emailVerified !== true ||
      cleanEmail(live.email) !== tokenEmail ||
      live.customClaims?.apple_review === true) {
    throw new HttpsError("permission-denied", "Amministratore non autorizzato.");
  }

  // The configured admin email list is checked AFTER the live Auth record.
  if (tokenEmail === ADMIN_EMAIL) return;
  let snap;
  try {
    snap = await admin.firestore().collection("appConfig").doc("adminUsers").get();
  } catch (error) {
    console.error("Apple review admin allowlist unavailable", {
      code: error?.code || "unknown"
    });
    throw new HttpsError("permission-denied", "Verifica amministratore non disponibile.");
  }
  const approved = snap.exists && Array.isArray(snap.data()?.emails)
    ? snap.data().emails.map(cleanEmail) : [];
  if (!approved.includes(tokenEmail)) {
    throw new HttpsError("permission-denied", "Solo un amministratore può creare l'account Apple.");
  }
}

exports.createAppleReviewAccount = onCall({
  region: REGION,
  invoker: "public", // callable transport; Firebase ID token + admin check still required
  timeoutSeconds: 40,
  maxInstances: 2
}, async (request) => {
  await assertAdmin(request.auth);
  const email = cleanEmail(request.data?.email);
  if (email.length > 220 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new HttpsError("invalid-argument", "Inserire una email dimostrativa valida.");
  }
  let previous;
  try {
    previous = await admin.auth().getUserByEmail(email);
  } catch (error) {
    if (error?.code !== "auth/user-not-found") throw error;
  }
  if (previous) throw new HttpsError("already-exists", "L'email appartiene già a un account. Non verrà modificato.");

  const password = randomPassword();
  let user;
  try {
    user = await admin.auth().createUser({
      email,
      password,
      emailVerified: true,
      disabled: false,
      displayName: "Apple App Review"
    });
    // This signed claim is recognized by Firestore + Storage security rules.
    // It must never be inferred from editable profile data or email strings.
    await admin.auth().setCustomUserClaims(user.uid, { apple_review: true });
    await admin.firestore().collection("platformUsers").doc(user.uid).set({
      uid: user.uid,
      email,
      displayName: "Apple App Review",
      role: "apple_review",
      ruolo: "apple_review",
      statoAccount: "attivo",
      accountStatus: "attivo",
      isAdmin: false,
      admin: false,
      banned: false,
      reviewer: true,
      createdAt: admin.firestore.FieldValue.serverTimestamp()
    });
    return {
      created: true,
      email,
      password,
      instructions: "Credenziali visibili una sola volta. Incollarle direttamente in App Store Connect. Non usarle per utenti aziendali."
    };
  } catch (error) {
    if (user?.uid) {
      try { await admin.auth().deleteUser(user.uid); } catch (_) {
        // If cleanup fails, report the failure to the administrator without
        // revealing credentials. Firebase rules still deny reviewer access.
      }
    }
    if (error instanceof HttpsError) throw error;
    throw new HttpsError("internal", "Creazione account di revisione non riuscita. Verificare i log server.");
  }
});
