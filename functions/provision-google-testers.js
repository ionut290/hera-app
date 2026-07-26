"use strict";

const admin = require("firebase-admin");

admin.initializeApp();

const temporaryPassword = String(process.env.TESTER_TEMP_PASSWORD || "");
if (temporaryPassword.length < 10) {
  throw new Error("TESTER_TEMP_PASSWORD mancante o troppo corta.");
}

async function listGoogleUsers() {
  const users = [];
  let pageToken;
  do {
    const page = await admin.auth().listUsers(1000, pageToken);
    users.push(...page.users.filter((user) =>
      user.email
      && user.providerData.some((provider) => provider.providerId === "google.com")
    ));
    pageToken = page.pageToken;
  } while (pageToken);
  return users;
}

async function main() {
  const users = await listGoogleUsers();
  const db = admin.firestore();
  let updated = 0;

  for (const user of users) {
    await admin.auth().updateUser(user.uid, {
      password: temporaryPassword,
      disabled: false
    });
    await db.collection("platformUsers").doc(user.uid).set({
      uid: user.uid,
      email: String(user.email).trim().toLowerCase(),
      displayName: user.displayName || String(user.email).split("@")[0],
      mustChangePassword: true,
      temporaryPasswordIssuedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    updated += 1;
  }

  console.log(`Provisioning completato per ${updated} account Google.`);
}

main().catch((error) => {
  console.error("Provisioning account tester fallito:", error.message || error);
  process.exitCode = 1;
});
