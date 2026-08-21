"use strict";

const { onSchedule } = require("firebase-functions/v2/scheduler");
const { getFirestore, Timestamp } = require("firebase-admin/firestore");
const { getStorage } = require("firebase-admin/storage");

const SOURCE = "whazzup-impianto-pdf";

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
    const snapshot = await db.collection("documents")
      .where("source", "==", SOURCE)
      .get();

    let deleted = 0;
    for (const doc of snapshot.docs) {
      const data = doc.data() || {};
      const expiresAtMs = data.expiresAt?.toMillis?.()
        || Date.parse(String(data.expiresAtIso || ""))
        || 0;
      if (!expiresAtMs || expiresAtMs > now) continue;

      const storagePath = String(data.storagePath || "").trim();
      if (storagePath) {
        await bucket.file(storagePath).delete({ ignoreNotFound: true }).catch((error) => {
          console.warn("PDF Whazzup: file Storage non eliminato", { documentId: doc.id, storagePath, error: error?.message || error });
        });
      }
      await doc.ref.delete();
      deleted += 1;
    }

    console.log("Cleanup PDF Whazzup completato", {
      scanned: snapshot.size,
      deleted,
      source: SOURCE
    });
  }
);
