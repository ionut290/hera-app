"use strict";

const { onSchedule } = require("firebase-functions/v2/scheduler");
const { defineJsonSecret } = require("firebase-functions/params");
const { getFirestore, Timestamp } = require("firebase-admin/firestore");
const { google } = require("googleapis");

const SOURCE = "whazzup-pdf-drive-v2";
const RUNTIME_CONFIG = defineJsonSecret("RUNTIME_CONFIG");

async function getDrive(db) {
  const snapshot = await db.collection("appConfig").doc("driveAdminSecret").get();
  const secret = snapshot.exists ? snapshot.data() : null;
  if (!secret || (!secret.accessToken && !secret.refreshToken)) throw new Error("Drive centrale non configurato");
  const runtimeConfig = RUNTIME_CONFIG.value() || {};
  const clientId = String(runtimeConfig.google?.client_id || "").trim();
  const clientSecret = String(runtimeConfig.google?.client_secret || "").trim();
  const oauth2 = new google.auth.OAuth2(clientId || undefined, clientSecret || undefined);
  oauth2.setCredentials({
    access_token: secret.accessToken || undefined,
    refresh_token: secret.refreshToken || undefined
  });
  return google.drive({ version: "v3", auth: oauth2 });
}

exports.cleanupExpiredWhazzupPdfs = onSchedule(
  {
    schedule: "every day 03:20",
    timeZone: "Europe/Rome",
    region: "europe-west1",
    retryCount: 1,
    secrets: [RUNTIME_CONFIG]
  },
  async () => {
    const db = getFirestore();
    const now = Timestamp.now().toMillis();
    const snapshot = await db.collection("documents").where("source", "==", SOURCE).get();
    let drive = null;
    let deleted = 0;
    let failed = 0;

    for (const doc of snapshot.docs) {
      const data = doc.data() || {};
      const expiresAtMs = data.expiresAt?.toMillis?.() || Date.parse(String(data.expiresAtIso || "")) || 0;
      if (!expiresAtMs || expiresAtMs > now) continue;

      const fileId = String(data.driveFileId || "").trim();
      try {
        if (fileId) {
          drive = drive || await getDrive(db);
          await drive.files.delete({ fileId });
        }
        await doc.ref.delete();
        deleted += 1;
      } catch (error) {
        if (error?.code === 404 || error?.response?.status === 404) {
          await doc.ref.delete();
          deleted += 1;
          continue;
        }
        failed += 1;
        console.warn("Cleanup PDF Whazzup V2 fallito", {
          documentId: doc.id,
          fileId,
          message: error?.message || String(error)
        });
      }
    }

    console.log("Cleanup PDF Whazzup V2 completato", {
      scanned: snapshot.size,
      deleted,
      failed,
      source: SOURCE
    });
  }
);
