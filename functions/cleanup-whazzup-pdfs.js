"use strict";

const { onSchedule } = require("firebase-functions/v2/scheduler");
const legacyFunctions = require("firebase-functions/v1");
const { getFirestore, Timestamp } = require("firebase-admin/firestore");
const { getStorage } = require("firebase-admin/storage");
const { google } = require("googleapis");

const SOURCE = "whazzup-impianto-pdf";

async function buildDriveClient(db) {
  const secretSnapshot = await db.collection("appConfig").doc("driveAdminSecret").get();
  const secret = secretSnapshot.exists ? secretSnapshot.data() : null;
  if (!secret || (!secret.accessToken && !secret.refreshToken)) throw new Error("Drive centrale non configurato");
  const oauth2 = new google.auth.OAuth2(
    legacyFunctions.config().google?.client_id,
    legacyFunctions.config().google?.client_secret
  );
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
    retryCount: 1
  },
  async () => {
    const db = getFirestore();
    const bucket = getStorage().bucket();
    const now = Timestamp.now().toMillis();
    const snapshot = await db.collection("documents").where("source", "==", SOURCE).get();
    let drive = null;
    let deleted = 0;

    for (const doc of snapshot.docs) {
      const data = doc.data() || {};
      const expiresAtMs = data.expiresAt?.toMillis?.() || Date.parse(String(data.expiresAtIso || "")) || 0;
      if (!expiresAtMs || expiresAtMs > now) continue;

      const provider = String(data.storageProvider || "storage").trim().toLowerCase();
      if (provider === "drive") {
        const driveFileId = String(data.driveFileId || "").trim();
        if (driveFileId) {
          try {
            drive = drive || await buildDriveClient(db);
            await drive.files.delete({ fileId: driveFileId });
          } catch (error) {
            if (!(error?.code === 404 || error?.response?.status === 404)) {
              console.warn("PDF Whazzup: file Drive non eliminato", { documentId: doc.id, driveFileId, error: error?.message || error });
              continue;
            }
          }
        }
      } else {
        const storagePath = String(data.storagePath || "").trim();
        if (storagePath) {
          await bucket.file(storagePath).delete({ ignoreNotFound: true }).catch((error) => {
            console.warn("PDF Whazzup: file Storage non eliminato", { documentId: doc.id, storagePath, error: error?.message || error });
          });
        }
      }

      await doc.ref.delete();
      deleted += 1;
    }

    console.log("Cleanup PDF Whazzup completato", { scanned: snapshot.size, deleted, source: SOURCE });
  }
);