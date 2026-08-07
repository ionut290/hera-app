"use strict";

const admin = require("firebase-admin");

admin.initializeApp();

const DEFAULT_PASSWORD = String(process.env.OPERATOR_DEFAULT_PASSWORD || "12345678");
const INTERNAL_LOGIN_DOMAIN = "operatori.vargacantieri.app";

if (DEFAULT_PASSWORD.length < 6) {
  throw new Error("Password predefinita non valida: Firebase richiede almeno 6 caratteri.");
}

function text(value) {
  return String(value ?? "").trim();
}

function normalizeEmail(value) {
  return text(value).toLowerCase();
}

function validEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function firstValue(data, keys) {
  for (const key of keys) {
    const value = text(data?.[key]);
    if (value) return value;
  }
  return "";
}

function getDisplayName(data) {
  return firstValue(data, ["nomeCompleto", "fullName", "displayName", "nominativo", "operatore"])
    || [firstValue(data, ["nome", "name", "firstName"]), firstValue(data, ["cognome", "surname", "lastName"])].filter(Boolean).join(" ")
    || "Operatore";
}

function getPersonnelEmail(data) {
  return normalizeEmail(firstValue(data, ["emailAccessoApp", "EMAIL_ACCESSO_APP", "linkedUserEmail", "LINKED_USER_EMAIL", "email", "mail"]));
}

function getLinkedUid(data) {
  return firstValue(data, ["linkedUserId", "LINKED_USER_ID"]);
}

function loginSlug(value) {
  return text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ".")
    .replace(/^\.+|\.+$/g, "")
    .replace(/\.{2,}/g, ".")
    || "operatore";
}

function buildInternalLogin(displayName, personnelId, duplicateCount) {
  const base = loginSlug(displayName);
  const suffix = duplicateCount > 1 ? `.${loginSlug(personnelId).slice(0, 6)}` : "";
  const username = `${base}${suffix}`;
  return {
    username,
    email: `${username}@${INTERNAL_LOGIN_DOMAIN}`
  };
}

async function resolveAuthUser(data, email) {
  const linkedUid = getLinkedUid(data);
  if (linkedUid) {
    try {
      return { user: await admin.auth().getUser(linkedUid), matchedBy: "linkedUserId" };
    } catch (error) {
      if (error?.code !== "auth/user-not-found") throw error;
    }
  }
  if (email) {
    try {
      return { user: await admin.auth().getUserByEmail(email), matchedBy: "email" };
    } catch (error) {
      if (error?.code !== "auth/user-not-found") throw error;
    }
  }
  return { user: null, matchedBy: "" };
}

async function provisionPersonnelAccount(db, personnelDoc, duplicateCount) {
  const data = personnelDoc.data() || {};
  const displayName = getDisplayName(data);
  const realEmail = getPersonnelEmail(data);
  const internalLogin = buildInternalLogin(displayName, personnelDoc.id, duplicateCount);
  const loginEmail = validEmail(realEmail) ? realEmail : internalLogin.email;
  const generatedInternalLogin = !validEmail(realEmail);
  const loginUsername = internalLogin.username;
  const resolved = await resolveAuthUser(data, loginEmail);
  let user = resolved.user;
  let created = false;

  if (user) {
    const update = {
      password: DEFAULT_PASSWORD,
      displayName,
      disabled: false
    };
    if (!user.email) update.email = loginEmail;
    user = await admin.auth().updateUser(user.uid, update);
  } else {
    user = await admin.auth().createUser({
      email: loginEmail,
      password: DEFAULT_PASSWORD,
      displayName,
      emailVerified: true,
      disabled: false
    });
    created = true;
  }

  const effectiveEmail = normalizeEmail(user.email || loginEmail);

  await db.collection("platformUsers").doc(user.uid).set({
    uid: user.uid,
    email: effectiveEmail,
    displayName,
    loginUsername,
    generatedInternalLogin,
    mustChangePassword: false,
    defaultOperatorPasswordProvisioned: true,
    defaultOperatorPasswordProvisionedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });

  const personnelPatch = {
    linkedUserId: user.uid,
    LINKED_USER_ID: user.uid,
    linkedUserEmail: effectiveEmail,
    LINKED_USER_EMAIL: effectiveEmail,
    loginUsername,
    generatedInternalLogin,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  };
  if (!text(data.emailAccessoApp) && !text(data.EMAIL_ACCESSO_APP)) {
    personnelPatch.emailAccessoApp = effectiveEmail;
    personnelPatch.EMAIL_ACCESSO_APP = effectiveEmail;
  }
  await personnelDoc.ref.set(personnelPatch, { merge: true });

  return {
    status: created ? "created" : "updated",
    personnelId: personnelDoc.id,
    uid: user.uid,
    email: effectiveEmail,
    displayName,
    loginUsername,
    generatedInternalLogin,
    matchedBy: created ? "new" : resolved.matchedBy
  };
}

async function main() {
  const db = admin.firestore();
  const snapshot = await db.collection("personale").get();
  const nameCounts = new Map();

  for (const doc of snapshot.docs) {
    const key = loginSlug(getDisplayName(doc.data() || {}));
    nameCounts.set(key, (nameCounts.get(key) || 0) + 1);
  }

  const results = [];
  for (const personnelDoc of snapshot.docs) {
    try {
      const displayKey = loginSlug(getDisplayName(personnelDoc.data() || {}));
      results.push(await provisionPersonnelAccount(db, personnelDoc, nameCounts.get(displayKey) || 1));
    } catch (error) {
      results.push({
        status: "error",
        personnelId: personnelDoc.id,
        displayName: getDisplayName(personnelDoc.data() || {}),
        code: error?.code || "",
        message: error?.message || String(error)
      });
    }
  }

  const summary = results.reduce((acc, item) => {
    acc[item.status] = (acc[item.status] || 0) + 1;
    if (item.generatedInternalLogin) acc.generatedInternalLogin += 1;
    return acc;
  }, { total: results.length, created: 0, updated: 0, error: 0, generatedInternalLogin: 0 });

  console.log("Provisioning account operatori completato:", JSON.stringify(summary));
  for (const item of results) console.log(JSON.stringify(item));

  if (summary.error > 0) throw new Error(`Provisioning terminato con ${summary.error} errori.`);
}

main().catch((error) => {
  console.error("Provisioning massivo operatori fallito:", error);
  process.exitCode = 1;
});
