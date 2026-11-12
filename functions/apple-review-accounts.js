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
async function assertAdmin(auth) {
  if (!auth || auth.token?.apple_review === true || auth.token?.email_verified !== true) {
    throw new HttpsError("permission-denied", "Account amministratore verificato richiesto.");
  }
  const email = cleanEmail(auth.token?.email);
  if (!email) throw new HttpsError("permission-denied", "Account amministratore richiesto.");

  if (email === ADMIN_EMAIL) return;
  const snap = await admin.firestore().collection("appConfig").doc("adminUsers").get();
  const approved = snap.exists && Array.isArray(snap.data()?.emails)
    ? snap.data().emails.map(cleanEmail) : [];
  if (!approved.includes(email)) {
    throw new HttpsError("permission-denied", "Solo un amministratore può creare l'account Apple.");
  }
}

exports.createAppleReviewAccount = onCall({
  region: REGION,
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
