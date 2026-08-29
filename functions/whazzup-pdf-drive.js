"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");
const { defineJsonSecret } = require("firebase-functions/params");
const { google } = require("googleapis");
const { Readable } = require("stream");

const CENTRAL_DRIVE_ROOT_FOLDER_ID = "1s6qmv2SsiTUbCjqFX4yIk4VoPQayFrU0";
const MAX_UPLOAD_BYTES = 15 * 1024 * 1024;
const RUNTIME_CONFIG = defineJsonSecret("RUNTIME_CONFIG");

function safeName(value, fallback = "Generale") {
  return String(value || fallback)
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .slice(0, 120) || fallback;
}

function escapeDriveQueryValue(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

async function getDrive() {
  const db = admin.firestore();
  const snapshot = await db.collection("appConfig").doc("driveAdminSecret").get();
  const secret = snapshot.exists ? snapshot.data() : null;
  if (!secret || (!secret.accessToken && !secret.refreshToken)) {
    throw new functions.https.HttpsError("failed-precondition", "Drive centrale non configurato.");
  }
  const runtimeConfig = RUNTIME_CONFIG.value() || {};
  const clientId = String(runtimeConfig.google?.client_id || "").trim();
  const clientSecret = String(runtimeConfig.google?.client_secret || "").trim();
  if (!clientId || !clientSecret) {
    throw new functions.https.HttpsError("failed-precondition", "Credenziali Google Drive mancanti nel backend.");
  }
  const oauth2 = new google.auth.OAuth2(clientId, clientSecret);
  oauth2.setCredentials({
    access_token: secret.accessToken || undefined,
    refresh_token: secret.refreshToken || undefined
  });
  return google.drive({ version: "v3", auth: oauth2 });
}

async function getOrCreateFolder(drive, folderName, parentId) {
  const name = safeName(folderName);
  const query = [
    `name='${escapeDriveQueryValue(name)}'`,
    "mimeType='application/vnd.google-apps.folder'",
    "trashed=false",
    `'${parentId}' in parents`
  ].join(" and ");
  const list = await drive.files.list({ q: query, fields: "files(id,name)", pageSize: 1 });
  const existing = list.data.files?.[0];
  if (existing?.id) return existing.id;
  const created = await drive.files.create({
    requestBody: {
      name,
      mimeType: "application/vnd.google-apps.folder",
      parents: [parentId]
    },
    fields: "id"
  });
  if (!created.data.id) throw new Error(`Impossibile creare la cartella Drive ${name}.`);
  return created.data.id;
}

exports.uploadWhazzupPdfToDrive = functions
  .region("europe-west1")
  .runWith({
    timeoutSeconds: 120,
    memory: "512MB",
    secrets: [RUNTIME_CONFIG]
  })
  .https.onCall(async (data, context) => {
    if (!context.auth) {
      throw new functions.https.HttpsError("unauthenticated", "Login richiesto.");
    }

    try {
      const base64 = String(data?.base64 || "").trim();
      const fileName = safeName(data?.fileName, "documento.pdf");
      const mimeType = String(data?.mimeType || "application/pdf").toLowerCase();
      if (mimeType !== "application/pdf" || !fileName.toLowerCase().endsWith(".pdf")) {
        throw new functions.https.HttpsError("invalid-argument", "Sono consentiti soltanto file PDF.");
      }
      const buffer = Buffer.from(base64, "base64");
      if (!buffer.length) throw new functions.https.HttpsError("invalid-argument", "Il PDF è vuoto.");
      if (buffer.length > MAX_UPLOAD_BYTES) throw new functions.https.HttpsError("invalid-argument", "Il PDF supera 15 MB.");

      const drive = await getDrive();
      const commessaName = safeName(data?.commessaName || data?.commessaId, "Generale");
      const commessaFolderId = await getOrCreateFolder(drive, commessaName, CENTRAL_DRIVE_ROOT_FOLDER_ID);
      const pdfFolderId = await getOrCreateFolder(drive, "WHAZZUP PDF", commessaFolderId);

      const uploaded = await drive.files.create({
        requestBody: {
          name: fileName,
          mimeType: "application/pdf",
          parents: [pdfFolderId]
        },
        media: {
          mimeType: "application/pdf",
          body: Readable.from(buffer)
        },
        fields: "id,name,webViewLink"
      });

      const fileId = String(uploaded.data.id || "");
      if (!fileId) throw new Error("Drive non ha restituito l'identificativo del PDF.");

      try {
        await drive.permissions.create({
          fileId,
          requestBody: { type: "anyone", role: "reader", allowFileDiscovery: false },
          fields: "id"
        });
      } catch (permissionError) {
        await drive.files.delete({ fileId }).catch(() => null);
        throw permissionError;
      }

      const fileUrl = uploaded.data.webViewLink || `https://drive.google.com/file/d/${fileId}/view`;
      console.info("PDF Whazzup caricato su Drive", {
        uid: context.auth.uid,
        fileId,
        commessaName,
        bytes: buffer.length
      });
      return { fileId, fileUrl, storageProvider: "drive" };
    } catch (error) {
      if (error instanceof functions.https.HttpsError) throw error;
      console.error("Upload PDF Whazzup su Drive fallito", {
        uid: context.auth?.uid || "",
        code: error?.code || error?.response?.status || "",
        message: error?.message || String(error)
      });
      throw new functions.https.HttpsError(
        "internal",
        `Caricamento PDF su Drive non riuscito: ${error?.message || "errore backend"}`
      );
    }
  });
