const admin = require("firebase-admin");
const functions = require("firebase-functions");
const { google } = require("googleapis");

admin.initializeApp();

const ADMIN_EMAIL = "ionut29019@gmail.com";
const CENTRAL_DRIVE_ROOT_FOLDER_NAME = "Varga Cantieri";
const MAX_UPLOAD_BYTES = 20 * 1024 * 1024;

function normalizeFolderName(value, fallback = "Generale") {
  return String(value || fallback).trim().replace(/[\\/:*?"<>|]+/g, "-").slice(0, 120) || fallback;
}

function escapeDriveQueryValue(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

async function getAdminEmails(db) {
  const snapshot = await db.collection("appConfig").doc("adminUsers").get();
  const configured = snapshot.exists && Array.isArray(snapshot.data().emails) ? snapshot.data().emails : [];
  return new Set([ADMIN_EMAIL, ...configured].map((email) => String(email || "").trim().toLowerCase()).filter(Boolean));
}

async function assertAdmin(context, db) {
  const email = String(context.auth?.token?.email || "").trim().toLowerCase();
  if (!context.auth || !email) {
    throw new functions.https.HttpsError("unauthenticated", "Login richiesto.");
  }
  const adminEmails = await getAdminEmails(db);
  if (!adminEmails.has(email)) {
    throw new functions.https.HttpsError("permission-denied", "Solo admin può configurare Drive.");
  }
}

async function buildDriveClient(db) {
  const secretSnapshot = await db.collection("appConfig").doc("driveAdminSecret").get();
  const secret = secretSnapshot.exists ? secretSnapshot.data() : null;
  if (!secret || (!secret.accessToken && !secret.refreshToken)) {
    throw new functions.https.HttpsError("failed-precondition", "Cloud amministratore non configurato");
  }

  const oauth2 = new google.auth.OAuth2(
    functions.config().google?.client_id,
    functions.config().google?.client_secret
  );
  oauth2.setCredentials({
    access_token: secret.accessToken || undefined,
    refresh_token: secret.refreshToken || undefined
  });

  return { drive: google.drive({ version: "v3", auth: oauth2 }), secret };
}

async function getOrCreateFolder(drive, name, parentId = "") {
  const clauses = [
    `name='${escapeDriveQueryValue(name)}'`,
    "mimeType='application/vnd.google-apps.folder'",
    "trashed=false"
  ];
  if (parentId) clauses.push(`'${parentId}' in parents`);
  const found = await drive.files.list({
    q: clauses.join(" and "),
    fields: "files(id,name)",
    pageSize: 1
  });
  const existing = found.data.files && found.data.files[0];
  if (existing?.id) return existing.id;

  const created = await drive.files.create({
    requestBody: {
      name,
      mimeType: "application/vnd.google-apps.folder",
      ...(parentId ? { parents: [parentId] } : {})
    },
    fields: "id"
  });
  return created.data.id;
}

exports.uploadCentralDriveFile = functions.https.onCall(async (data, context) => {
  if (!context.auth) {
    throw new functions.https.HttpsError("unauthenticated", "Login richiesto.");
  }
  const base64 = String(data?.base64 || "");
  const fileName = normalizeFolderName(data?.fileName, "file");
  const mimeType = String(data?.mimeType || "application/octet-stream");
  const commessaName = normalizeFolderName(data?.commessaName, "Generale");
  const driveType = normalizeFolderName(data?.driveType, "EXPORT").toUpperCase();
  const buffer = Buffer.from(base64, "base64");
  if (!buffer.length) {
    throw new functions.https.HttpsError("invalid-argument", "File vuoto.");
  }
  if (buffer.length > MAX_UPLOAD_BYTES) {
    throw new functions.https.HttpsError("invalid-argument", "File troppo grande per upload backend.");
  }

  const db = admin.firestore();
  const { drive, secret } = await buildDriveClient(db);
  const rootFolderId = secret.rootFolderId || await getOrCreateFolder(drive, CENTRAL_DRIVE_ROOT_FOLDER_NAME);
  const commessaFolderId = await getOrCreateFolder(drive, commessaName, rootFolderId);
  const typeFolderId = await getOrCreateFolder(drive, driveType, commessaFolderId);

  const uploaded = await drive.files.create({
    requestBody: {
      name: fileName,
      mimeType,
      parents: [typeFolderId]
    },
    media: {
      mimeType,
      body: require("stream").Readable.from(buffer)
    },
    fields: "id,name,webViewLink"
  });

  await db.collection("centralDriveUploads").add({
    fileId: uploaded.data.id || "",
    fileName,
    mimeType,
    commessaName,
    driveType,
    createdByUid: context.auth.uid,
    createdByEmail: context.auth.token.email || "",
    createdAt: admin.firestore.FieldValue.serverTimestamp()
  });

  return {
    fileId: uploaded.data.id || "",
    webViewLink: uploaded.data.webViewLink || ""
  };
});

exports.configureCentralDrive = functions.https.onCall(async (data, context) => {
  const db = admin.firestore();
  await assertAdmin(context, db);
  const ownerEmail = String(context.auth.token.email || ADMIN_EMAIL);
  const accessToken = String(data?.accessToken || "");
  const refreshToken = String(data?.refreshToken || "");
  const rootFolderId = String(data?.rootFolderId || "");
  if (!accessToken && !refreshToken) {
    throw new functions.https.HttpsError("invalid-argument", "Token Drive mancante.");
  }
  await db.collection("appConfig").doc("driveAdminSecret").set({
    ownerEmail,
    accessToken,
    refreshToken,
    rootFolderId,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  await db.collection("appConfig").doc("driveBridge").set({
    ownerEmail,
    configured: true,
    rootFolderId,
    rootFolderName: CENTRAL_DRIVE_ROOT_FOLDER_NAME,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  return { ok: true };
});
