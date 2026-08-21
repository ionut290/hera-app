"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");
const { google } = require("googleapis");
const { Readable } = require("stream");

const CENTRAL_DRIVE_ROOT_FOLDER_ID = "1s6qmv2SsiTUbCjqFX4yIk4VoPQayFrU0";
const MAX_UPLOAD_BYTES = 15 * 1024 * 1024;

function cleanFolderName(value, fallback = "Generale") {
  return String(value || fallback).trim().replace(/[\\/:*?"<>|]+/g, "-").slice(0, 120) || fallback;
}

function escapeDriveQueryValue(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

async function buildDriveClient() {
  const db = admin.firestore();
  const secretSnapshot = await db.collection("appConfig").doc("driveAdminSecret").get();
  const secret = secretSnapshot.exists ? secretSnapshot.data() : null;
  if (!secret || (!secret.accessToken && !secret.refreshToken)) {
    throw new functions.https.HttpsError("failed-precondition", "Archivio Drive centrale non configurato.");
  }
  const oauth2 = new google.auth.OAuth2(
    functions.config().google?.client_id,
    functions.config().google?.client_secret
  );
  oauth2.setCredentials({
    access_token: secret.accessToken || undefined,
    refresh_token: secret.refreshToken || undefined
  });
  return google.drive({ version: "v3", auth: oauth2 });
}

async function getOrCreateFolder(drive, name, parentId) {
  const safeName = cleanFolderName(name);
  const found = await drive.files.list({
    q: `name='${escapeDriveQueryValue(safeName)}' and mimeType='application/vnd.google-apps.folder' and trashed=false and '${parentId}' in parents`,
    fields: "files(id,name)",
    pageSize: 1
  });
  const existing = found.data.files?.[0];
  if (existing?.id) return existing.id;
  const created = await drive.files.create({
    requestBody: {
      name: safeName,
      mimeType: "application/vnd.google-apps.folder",
      parents: [parentId]
    },
    fields: "id"
  });
  return created.data.id;
}

exports.uploadWhazzupPdfToDrive = functions.region("europe-west1").https.onCall(async (data, context) => {
  if (!context.auth) {
    throw new functions.https.HttpsError("unauthenticated", "Login richiesto.");
  }
  const base64 = String(data?.base64 || "");
  const fileName = cleanFolderName(data?.fileName, "documento.pdf");
  const mimeType = String(data?.mimeType || "application/pdf");
  if (mimeType !== "application/pdf" || !fileName.toLowerCase().endsWith(".pdf")) {
    throw new functions.https.HttpsError("invalid-argument", "Sono consentiti soltanto PDF.");
  }
  const buffer = Buffer.from(base64, "base64");
  if (!buffer.length) throw new functions.https.HttpsError("invalid-argument", "PDF vuoto.");
  if (buffer.length > MAX_UPLOAD_BYTES) throw new functions.https.HttpsError("invalid-argument", "PDF oltre 15 MB.");

  const commessaName = cleanFolderName(data?.commessaName || data?.commessaId, "Generale");
  const drive = await buildDriveClient();
  const commessaFolderId = await getOrCreateFolder(drive, commessaName, CENTRAL_DRIVE_ROOT_FOLDER_ID);
  const pdfFolderId = await getOrCreateFolder(drive, "WHAZZUP PDF", commessaFolderId);

  const uploaded = await drive.files.create({
    requestBody: { name: fileName, mimeType, parents: [pdfFolderId] },
    media: { mimeType, body: Readable.from(buffer) },
    fields: "id,name,webViewLink"
  });
  const fileId = String(uploaded.data.id || "");
  if (!fileId) throw new functions.https.HttpsError("internal", "Drive non ha restituito l'identificativo del PDF.");

  await drive.permissions.create({
    fileId,
    requestBody: { type: "anyone", role: "reader", allowFileDiscovery: false },
    fields: "id"
  });

  const fileUrl = uploaded.data.webViewLink || `https://drive.google.com/file/d/${fileId}/view`;
  return { fileId, fileUrl, storageProvider: "drive" };
});

async function deleteWhazzupDriveFile(fileId) {
  const id = String(fileId || "").trim();
  if (!id) return false;
  const drive = await buildDriveClient();
  try {
    await drive.files.delete({ fileId: id });
    return true;
  } catch (error) {
    if (error?.code === 404 || error?.response?.status === 404) return true;
    throw error;
  }
}

exports.deleteWhazzupDriveFile = deleteWhazzupDriveFile;
